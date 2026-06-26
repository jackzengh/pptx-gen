from __future__ import annotations

import asyncio
import json
import os
import shutil
from pathlib import Path

from openai import AsyncOpenAI

from agents import (
    OpenAIResponsesModel,
    Runner,
    function_tool,
    set_default_openai_api,
    set_default_openai_client,
    set_tracing_disabled,
)
from agents.run import RunConfig
from agents.sandbox import Manifest, SandboxAgent, SandboxRunConfig
from agents.sandbox.capabilities import (
    Capabilities,
    LocalDirLazySkillSource,
    Skills,
)
from agents.sandbox.capabilities.compaction import Compaction
from agents.sandbox.capabilities.shell import Shell
from agents.sandbox.entries import LocalDir
from agents.sandbox.sandboxes.unix_local import UnixLocalSandboxClient

# Import the validator directly so the function tool can call it from the host.
from harness import check_pptx

EXAMPLE_DIR = Path(__file__).resolve().parent
WORKSPACE = EXAMPLE_DIR / "agent_workspace"
HOST_REPO_DIR = WORKSPACE / "repo"
# Pinning the manifest root to a real host dir (instead of the SDK's default
# "/workspace", which makes it mkdtemp a throwaway temp dir) means the agent's
# sandbox files persist on disk and we can recover/validate the built deck. The
# agent's `repo` mount lands at <SANDBOX_ROOT>/repo.
# RUN_TAG lets two models run in parallel without clobbering each other's
# pinned workspace (e.g. RUN_TAG=gpt vs RUN_TAG=glm).
RUN_TAG = os.environ.get("RUN_TAG", "")
HOST_SANDBOX_DIR = WORKSPACE / (f"sandbox-{RUN_TAG}" if RUN_TAG else "sandbox")
SANDBOX_REPO_DIR = HOST_SANDBOX_DIR / "repo"
HOST_OUT_DIR = WORKSPACE / (f"out-{RUN_TAG}" if RUN_TAG else "out")
DECK_NAME = "ttwo_overview.pptx"
HOST_SKILLS_DIR = WORKSPACE / "skills"

BASE_URL = os.environ.get("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1")
API_KEY = os.environ.get("OPENROUTER_API_KEY", "")
MODEL = os.environ.get("MODEL", "openai/gpt-5.5")

if not API_KEY:
    raise SystemExit("Set OPENROUTER_API_KEY in the environment.")

_client = AsyncOpenAI(base_url=BASE_URL, api_key=API_KEY)
set_default_openai_client(client=_client, use_for_tracing=False)
set_default_openai_api("responses")
set_tracing_disabled(disabled=True)


@function_tool
def validate_deck(pptx_path: str) -> str:
    p = Path(pptx_path)
    if not p.is_absolute():
        p = HOST_REPO_DIR / p
    if not p.exists():
        return json.dumps({"error": f"file not found: {p}"})
    return json.dumps(check_pptx(str(p)))


def build_agent(model: str, *, use_apply_patch: bool = True) -> SandboxAgent[None]:
    # Wrap the slug in an explicit Responses model bound to the OpenRouter
    # client. Passing the bare string (e.g. "z-ai/glm-5.2") would make the SDK's
    # MultiProvider treat the part before "/" as a provider prefix and fail with
    # "Unknown prefix: z-ai". An explicit model object bypasses that routing.
    resolved_model = OpenAIResponsesModel(model=model, openai_client=_client)

    # The Filesystem capability exposes the freeform `apply_patch` tool. Some
    # non-OpenAI models (e.g. GLM-5.2 over OpenRouter) emit apply_patch calls in
    # a shape the SDK can't match ("Model produced apply_patch call without an
    # apply_patch tool"). For those, drop Filesystem and have the agent write/
    # edit files with the Shell capability's `exec_command` (bash) instead.
    if use_apply_patch:
        base_caps = Capabilities.default()  # [Filesystem, Shell, Compaction]
        edit_hint = (
            "Paths for apply_patch and the shell are relative to the sandbox "
            "workspace root."
        )
    else:
        base_caps = [Shell(), Compaction()]  # no apply_patch
        edit_hint = (
            "You do NOT have an apply_patch tool. Create and edit files only with "
            "the shell (`exec_command`) — e.g. write the whole script with a `cat "
            "> repo/build_deck.py <<'PY' ... PY` heredoc and rewrite it wholesale "
            "to make changes. Shell paths are relative to the workspace root."
        )

    return SandboxAgent(
        name="Deck builder",
        model=resolved_model,
        instructions=(
            "You build PowerPoint decks. First read `repo/task.md`, then read the "
            "`pptx` skill it points to (SKILL.md and the files it links) before writing "
            "any code. Build the deck with python-pptx by writing a `repo/build_deck.py` "
            "script and running it with the shell. Follow the skill's rules exactly: name "
            "every shape, set auto_size=SHAPE_TO_FIT_TEXT on every text frame, use a true "
            "<p:bg> background for full-bleed color, no overlaps, consistent margins/"
            "alignment, and a caption directly above every chart and table. After each "
            "build, validate with the `validate_deck` tool (or `uv run --no-project "
            "--with python-pptx python main.py ttwo_overview.pptx` from `repo/`). If it is "
            "not ok, fix exactly the named shapes and repeat until ok is true (42/42). "
            + edit_hint
        ),
        # Mount the project repo (spec + harness + main.py) at `repo/`. Pin the
        # workspace root to a real host dir so the built deck persists (see
        # HOST_SANDBOX_DIR above).
        default_manifest=Manifest(
            root=str(HOST_SANDBOX_DIR),
            entries={
                "repo": LocalDir(src=HOST_REPO_DIR),
            },
        ),
        capabilities=base_caps
        + [
            Skills(
                lazy_from=LocalDirLazySkillSource(
                    source=LocalDir(src=HOST_SKILLS_DIR),
                )
            ),
        ],
        # A normal Agent field (SandboxAgent subclasses Agent): our own function tool.
        tools=[validate_deck],
    )


def _print_trace(result) -> None:
    """Print a compact trace of the tool calls the agent made.

    The final text output can be empty (e.g. GLM-5.2 sometimes returns nothing),
    so this makes every run auditable regardless.
    """
    print("\n--- tool-call trace ---")
    for item in result.new_items:
        kind = type(item).__name__
        if kind == "ToolCallItem":
            raw = item.raw_item
            name = getattr(raw, "name", getattr(raw, "type", "tool"))
            print(f"  → call: {name}")
        elif kind == "ToolCallOutputItem":
            out = str(item.output)
            print(f"    output: {out[:160]}{'…' if len(out) > 160 else ''}")
    print("--- end trace ---\n")


def _persist_and_validate() -> None:
    """Copy the built deck out of the pinned sandbox and validate it host-side."""
    built = SANDBOX_REPO_DIR / DECK_NAME
    if not built.exists():
        print(f"[!] No deck found at {built} — the agent did not build {DECK_NAME}.")
        return
    HOST_OUT_DIR.mkdir(parents=True, exist_ok=True)
    dest = HOST_OUT_DIR / DECK_NAME
    shutil.copy2(built, dest)
    report = check_pptx(str(dest))
    print(f"[deck] saved to {dest}")
    # `summary` is only present when there are problems in some check_pptx
    # versions; fall back gracefully so reporting never crashes the run.
    s = report.get("summary")
    if s:
        print(
            f"[check_pptx] ok={report.get('ok')} "
            f"score={s.get('score')}/{s.get('max_score')} fails={s.get('fail_count')}"
        )
    else:
        print(f"[check_pptx] ok={report.get('ok')}")
    for p in report.get("problems", [])[:10]:
        print(f"    - {p['criterion']} (slide {p['slide']}): {p['message']}")


async def run(model: str, *, use_apply_patch: bool, workflow_name: str) -> None:
    result = await Runner.run(
        build_agent(model, use_apply_patch=use_apply_patch),
        "Read `repo/task.md`, build `repo/ttwo_overview.pptx` from `repo/spec.md` using "
        "the pptx skill, validate it with check_pptx until ok is true, and summarize what "
        "you built (palette/motif, iterations, final score).",
        run_config=RunConfig(
            sandbox=SandboxRunConfig(client=UnixLocalSandboxClient()),
            workflow_name=workflow_name,
        ),
        # Building + validating + fixing a 14-slide deck takes many turns; the
        # SDK default of 10 is far too low.
        max_turns=60,
    )
    print("\n=== agent summary ===")
    print(result.final_output or "(no final text output)")
    _print_trace(result)
    _persist_and_validate()


async def main() -> None:
    await run(MODEL, use_apply_patch=True, workflow_name="Take-Two deck build")


if __name__ == "__main__":
    asyncio.run(main())
