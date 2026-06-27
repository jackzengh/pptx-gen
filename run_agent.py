from __future__ import annotations

import argparse
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
    LocalDirLazySkillSource,
    Skills,
)
from agents.sandbox.capabilities.compaction import Compaction
from agents.sandbox.capabilities.shell import Shell
from agents.sandbox.entries import LocalDir
from agents.sandbox.sandboxes.unix_local import UnixLocalSandboxClient
from harness import check_pptx

WORKSPACE = Path(__file__).resolve().parent / "agent_workspace"
HOST_REPO_DIR = WORKSPACE / "repo"
HOST_SKILLS_DIR = WORKSPACE / "skills"
DECK_NAME = "out.pptx"
LOG_NAME = "run_log.txt"
_SANDBOX_DIR: Path = WORKSPACE / "sandbox"
_OUT_DIR: Path = WORKSPACE / "out"
_client: AsyncOpenAI | None = None

DEFAULT_MODEL = "z-ai/gpt-5.5"

INSTRUCTIONS_SKILL = (
    "You build PowerPoint decks. First read `repo/task.md`, then read the "
    "`pptx` skill it points to (SKILL.md and the files it links) before writing "
    "any code. ALSO read `repo/reference_build_deck.py` — a complete, working "
    "build script for a different deck that already passes check_pptx; reuse its "
    "helper functions and structural patterns verbatim, but build the deck from "
    "`repo/spec.md` with your own palette (do not copy its Take-Two content). "
    "Build the deck with python-pptx by writing a `repo/build_deck.py` "
    "script and running it with the shell. Follow the skill's rules exactly: name "
    "every shape, use a true "
    "<p:bg> background for full-bleed color, no overlaps, consistent margins/"
    "alignment, and a caption directly above every chart and table. After each "
    "build, validate with the `validate_deck` tool (or `uv run --no-project "
    f"--with python-pptx python main.py {DECK_NAME}` from `repo/`). If it is "
    "not ok, fix exactly the named shapes and repeat until ok is true. "
    "When the deck validates ok, call the `finish_deck` tool with the deck path and a "
    "`log` of what you did (palette/motif, iterations, final result) to persist the deck "
    "and the run log. "
    "Create and edit files only with "
    "the shell (`exec_command`) — e.g. write the whole script with a `cat "
    "> repo/build_deck.py <<'PY' ... PY` heredoc and rewrite it wholesale "
    "to make changes. Shell paths are relative to the workspace root."
)

INSTRUCTIONS_NOSKILL = (
    "You build PowerPoint decks with python-pptx. First read `repo/task.md`, "
    "then build the deck from `repo/spec.md` with your own palette. "
    "ALSO read `repo/reference_build_deck.py` — a complete, working build "
    "script for a DIFFERENT deck that already passes the check_pptx validator; "
    "reuse its helper functions and structural patterns verbatim, but build the "
    "deck from `repo/spec.md` (do not copy its Take-Two content). "
    "Build the deck by writing a `repo/build_deck.py` script and running it with "
    "the shell. Make a clean, professional consulting deck: name every shape, "
    "use a true <p:bg> background "
    "for full-bleed color, no overlaps, consistent margins/alignment, and a caption "
    "directly above every chart and table. After each build, validate with the "
    "`validate_deck` tool (or `uv run --no-project --with python-pptx python main.py "
    f"{DECK_NAME}` from `repo/`). If it is not ok, fix exactly the named shapes and "
    "repeat until ok is true. When the deck validates ok, call the `finish_deck` tool "
    "with the deck path and a `log` of what you did (palette/motif, iterations, final "
    "result) to persist the deck and the run log. "
    "Create and edit files only with the shell (`exec_command`) — e.g. write the whole "
    "script with a `cat > repo/build_deck.py <<'PY' ... PY` heredoc and rewrite it "
    "wholesale to make changes. Shell paths are relative to the workspace root."
)


@function_tool
def validate_deck(pptx_path: str) -> str:
    p = Path(pptx_path)
    if not p.is_absolute():
        p = HOST_REPO_DIR / p
    if not p.exists():
        return json.dumps({"error": f"file not found: {p}"})
    return json.dumps(check_pptx(str(p)))


@function_tool
def finish_deck(deck_path: str, log: str) -> str:
    """Persist the finished deck and a run-log .txt to the host out/ dir.

    Call this once at the end. `deck_path` is the built .pptx; `log` is your
    run narrative (palette/motif, iterations, whether validation passed).
    """
    p = Path(deck_path)
    if not p.is_absolute():
        p = HOST_REPO_DIR / p
    if not p.exists():
        return json.dumps({"error": f"file not found: {p}"})

    _OUT_DIR.mkdir(parents=True, exist_ok=True)
    saved_deck = _OUT_DIR / DECK_NAME
    saved_log = _OUT_DIR / LOG_NAME
    shutil.copy2(p, saved_deck)
    saved_log.write_text(log)

    report = check_pptx(str(saved_deck))
    return json.dumps({
        "saved_deck": str(saved_deck),
        "saved_log": str(saved_log),
        "ok": report.get("ok"),
        "problem_count": len(report.get("problems", [])),
    })


def build_agent(model: str, use_skills: bool) -> SandboxAgent[None]:
    assert _client is not None, "OpenRouter client not initialised"
    resolved_model = OpenAIResponsesModel(model=model, openai_client=_client)

    # Mount the project repo (spec + harness + main.py) at `repo/`. Pin the
    # workspace root to a real host dir so the built deck persists.
    manifest = Manifest(
        root=str(_SANDBOX_DIR),
        entries={
            "repo": LocalDir(src=HOST_REPO_DIR),
        },
    )

    capabilities = [Shell(), Compaction()]  # no apply_patch, no vision
    if use_skills:
        capabilities.append(
            Skills(
                lazy_from=LocalDirLazySkillSource(
                    source=LocalDir(src=HOST_SKILLS_DIR),
                )
            )
        )

    return SandboxAgent(
        name="Deck builder" if use_skills else "Deck builder (no skills)",
        model=resolved_model,
        instructions=INSTRUCTIONS_SKILL if use_skills else INSTRUCTIONS_NOSKILL,
        default_manifest=manifest,
        capabilities=capabilities,
        # A normal Agent field (SandboxAgent subclasses Agent): our own function tool.
        tools=[validate_deck, finish_deck],
    )


def fallback_save() -> None:
    built = _SANDBOX_DIR / "repo" / DECK_NAME
    if not built.exists():
        print(f"No deck found at {built} — the agent did not build {DECK_NAME}.")
        return
    _OUT_DIR.mkdir(parents=True, exist_ok=True)
    destination_path = _OUT_DIR / DECK_NAME
    shutil.copy2(built, destination_path)
    print(f"Deck saved to {destination_path}")


async def run(
    model: str,
    *,
    workflow_name: str,
    spec_path: str | None = None,
    use_skills: bool = True,
    max_turns: int = 60,
) -> None:

    if spec_path is not None:
        src = Path(spec_path)
        if not src.exists():
            raise SystemExit(f"Spec file not found: {src}")
        shutil.copy2(src, HOST_REPO_DIR / "spec.md")

    if not use_skills:
        noskill_task = HOST_REPO_DIR / "task_noskill.md"
        if noskill_task.exists():
            shutil.copy2(noskill_task, HOST_REPO_DIR / "task.md")

    await Runner.run(
        build_agent(model, use_skills),
        f"Read `repo/task.md`, build `repo/{DECK_NAME}` from `repo/spec.md` using "
        "the pptx skill, validate it with check_pptx until ok is true, and summarize what "
        "you built (palette/motif, iterations, whether validation passed).",
        run_config=RunConfig(
            sandbox=SandboxRunConfig(client=UnixLocalSandboxClient()),
            workflow_name=workflow_name,
        ),
        max_turns=max_turns,
    )
    fallback_save()

# pass in the spec, model, tag, workflow, no-skills, max-turns
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Build a PowerPoint deck from a spec with a sandboxed agent.",
    )
    p.add_argument(
        "--spec",
        default=os.environ.get("SPEC_PATH") or None,
        help="Deck spec markdown to build (env SPEC_PATH). Omit to reuse repo/spec.md.",
    )
    p.add_argument(
        "--model",
        default=os.environ.get("MODEL", DEFAULT_MODEL),
        help=f"OpenRouter model id (env MODEL; default {DEFAULT_MODEL}).",
    )
    p.add_argument(
        "--tag",
        default=os.environ.get("RUN_TAG", ""),
        help="Suffix for sandbox-*/out-* dirs, for parallel runs (env RUN_TAG).",
    )
    p.add_argument(
        "--workflow",
        default=None,
        help="Run/workflow label (default: derived from spec + model).",
    )
    p.add_argument(
        "--no-skills",
        action="store_true",
        help="Ablation: drop the pptx skill + tracing and use task_noskill.md.",
    )
    p.add_argument(
        "--max-turns",
        type=int,
        default=60,
        help="Agent turn budget (default 60).",
    )
    return p.parse_args()


async def main() -> None:
    global _client, _SANDBOX_DIR, _OUT_DIR
    args = parse_args()

    api_key = os.environ.get("OPENROUTER_API_KEY", "")
    if not api_key:
        raise SystemExit("Set OPENROUTER_API_KEY in the environment.")

    tag = args.tag
    _SANDBOX_DIR = WORKSPACE / (f"sandbox-{tag}" if tag else "sandbox")
    _OUT_DIR = WORKSPACE / (f"out-{tag}" if tag else "out")

    use_skills = not args.no_skills

    _client = AsyncOpenAI(base_url="https://openrouter.ai/api/v1", api_key=api_key)
    set_default_openai_client(client=_client, use_for_tracing=False)
    set_default_openai_api("responses")
    set_tracing_disabled(disabled=True)

    # Derive a readable workflow label if the caller didn't supply one.
    if args.workflow:
        workflow_name = args.workflow
    else:
        spec_label = Path(args.spec).stem if args.spec else "repo-spec"
        suffix = " [no-skills]" if not use_skills else ""
        tag_suffix = f" [{tag}]" if tag else ""
        workflow_name = f"deck build: {spec_label} ({args.model}){suffix}{tag_suffix}"

    await run(
        args.model,
        workflow_name=workflow_name,
        spec_path=args.spec,
        use_skills=use_skills,
        max_turns=args.max_turns,
    )


if __name__ == "__main__":
    asyncio.run(main())
