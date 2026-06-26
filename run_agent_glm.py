"""Build the Take-Two deck with the sandbox agent, driven by GLM-5.2.

Identical to `run_agent.py` except the model is Z.ai's `z-ai/glm-5.2` instead of
`openai/gpt-5.5`, and apply_patch is disabled (GLM-5.2 over OpenRouter can't
drive that freeform hosted tool — it writes files via the shell instead).

Importing `run_agent` reuses its OpenRouter Responses-API provider setup, the
shared `build_agent` factory, the `validate_deck` tool, and the `run()` helper
that logs the tool trace, persists the built deck to agent_workspace/out/, and
validates it with check_pptx.

Run:
    OPENROUTER_API_KEY=sk-or-... uv run python run_agent_glm.py
    MODEL=z-ai/glm-5.2 OPENROUTER_API_KEY=sk-or-... uv run python run_agent_glm.py
"""

from __future__ import annotations

import asyncio
import os

# Importing run_agent runs its provider setup at import time (OpenRouter
# /responses client + set_default_openai_api("responses")). It raises SystemExit
# if OPENROUTER_API_KEY is unset.
from run_agent import run

# GLM-5.2 over OpenRouter. Override with MODEL=... in the env.
MODEL = os.environ.get("MODEL", "z-ai/glm-5.2")


async def main() -> None:
    # use_apply_patch=False: GLM-5.2 can't drive the freeform apply_patch hosted
    # tool, so it builds/edits files via the shell.
    await run(
        MODEL,
        use_apply_patch=False,
        workflow_name="Take-Two deck build (GLM-5.2)",
    )


if __name__ == "__main__":
    asyncio.run(main())
