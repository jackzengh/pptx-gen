# Ford F-Series deck — GLM-5.2, 3 harness runs

- **Model:** `z-ai/glm-5.2` (via OpenRouter, openai-agents sandbox harness)
- **Spec:** `specs/ford_fseries.md` (21 `##` sections → 21 slides)
- **Driver:** `run_ford_glm.py` (calls `run_agent.run`), one run per `RUN_TAG=ford{1,2,3}`
- **Date:** 2026-06-26

## Results

| Run | Deck                | Slides | check_pptx | Build wall-time |
|-----|---------------------|--------|------------|-----------------|
| 1   | `ford_run1.pptx`    | 21     | ok, 0 problems | ~7 min      |
| 2   | `ford_run2.pptx`    | 21     | ok, 0 problems | ~9 min      |
| 3   | `ford_run3.pptx`    | 21     | ok, 0 problems | ~9 min      |

All three validate clean (`ok=True`, 0 problems).

## Files
- `ford_run{1,2,3}.pptx` — the built decks
- `ford_run{1,2,3}_log.txt` — each run's agent narrative (palette/motif, iterations, validation)

Source outputs preserved in `agent_workspace/out-ford{1,2,3}/`; sandboxes in `agent_workspace/sandbox-ford{1,2,3}/`.


--- CLEANUP NOTE ---
This is the EARLIER GLM-5.2 Ford batch (built ~13:24, 3 clean runs, all 21-slide BCG Ford).
Distinct decks from the later ../ford_run1.pptx / ../ford_run2.pptx batch (built ~14:30+) — different MD5s.
Per the README all 3 here validate ok=True / 0 problems.
