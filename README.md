# pptx-gen

A harness for building **investment-banking / consulting-grade PowerPoint decks** with an LLM agent, and a layout validator (`check_pptx`) that enforces the visual rules those decks must follow.

The idea: hand the agent a markdown **spec** (one slide per section) and a **model**; it writes a `python-pptx` build script, runs it, validates the result against `check_pptx`, and fixes whatever the validator flags — looping until the deck passes.

---

## Quick start

Requires Python ≥3.14 and an [OpenRouter](https://openrouter.ai/) API key.

```bash
# install deps (uses uv; or pip install -e .)
uv sync

# put your key in .env  (the run script reads it from the environment)
echo 'OPENROUTER_API_KEY=sk-or-...' > .env
set -a; . ./.env; set +a

# build a deck from a spec
.venv/bin/python run_agent.py --spec specs/ford_fseries.md --model z-ai/glm-5.2 --tag run1
```

The built deck lands in `agent_workspace/out-run1/out.pptx` alongside a `run_log.txt` (the agent's own narrative of palette, iterations, and final result).

### Validate a deck by hand

```bash
python main.py path/to/deck.pptx        # prints {ok, problems, notes} as JSON
```

---

## How it works

```
spec.md ──► run_agent.py ──► sandboxed agent ──► build_deck.py ──► out.pptx
                                   │                                  │
                                   └──────── check_pptx ◄─────────────┘
                                        (validate → fix → repeat)
```

1. **`run_agent.py`** is the single entry point. It builds a sandboxed agent (OpenAI Agents SDK over an OpenRouter model), mounts the project repo and — by default — the `pptx` skill, and gives the agent two host-side tools: `validate_deck` (runs `check_pptx`) and `finish_deck` (persists the deck + log to `out-<tag>/`).
2. The agent reads `repo/task.md`, the `pptx` skill, and a working **reference build script** (`reference_build_deck.py`), then writes its own `build_deck.py` using `python-pptx` and runs it in the shell.
3. After each build it calls **`check_pptx`**, which flattens every slide into absolute-coordinate boxes and runs eight layout checks. If `ok` is false, the agent fixes exactly the named shapes and rebuilds.
4. When the deck validates clean, the agent calls `finish_deck` to save it.

### `run_agent.py` flags

| Flag | Env fallback | Default | Purpose |
|------|--------------|---------|---------|
| `--spec PATH` | `SPEC_PATH` | reuse `repo/spec.md` | deck spec to build |
| `--model NAME` | `MODEL` | `z-ai/gpt-5.5` | OpenRouter model id |
| `--tag TAG` | `RUN_TAG` | `""` | suffix for `sandbox-*`/`out-*` dirs (parallel-safe runs) |
| `--workflow NAME` | — | derived from spec + model | run/workflow label |
| `--no-skills` | — | off | ablation: drop the pptx skill + use `task_noskill.md` |
| `--max-turns N` | — | `60` | agent turn budget |

---

## The validator: `check_pptx`

`check_pptx(path) -> {ok, problems, notes}` reads only the **stored geometry** in the file (`shape.left/top/width/height`) and reports layout defects. Lives in `harness/` — `tools.py` holds the eight checks, `utils.py` the shape-flattening + geometry helpers, `text_metrics.py` the font-metric text measurement.

The eight checks:

- **outer-margin frame** — nothing bleeds off-slide or into the reserved 0.3in margin band (except `background_` panels, which are meant to bleed).
- **no overlap** — no two shapes' rectangles collide, except intentional `background`/`overlay`/`chartpart` cases.
- **left-margin consistency** — content hangs off one shared left edge, per slide and deck-wide.
- **edge & grid alignment** — peer boxes in a row share width and top.
- **consistent gutters** — gaps between 3+ peers in a row/column are even.
- **element boxing** — every chart/table has a caption directly above it.
- **text overflow** — text estimated (via real font metrics, `measure_text_emu`) to spill past its box *and* collide with the shape below.
- **y-offset variation** — title and footer sit at one fixed y across the deck.

---

## Specs

A spec is markdown with **one slide per `##` section**: a slide id, type, and title, then `**Layout:**` / `**Text:**` / bullets / tables / `_Source:_` blocks describing the content. See `specs/` for worked examples (`ford_fseries.md`, `ghp_stoptb.md`, `ttwo_overview.md`).

---

## Skills

The agent reads skills from `agent_workspace/skills/` before writing code. Three ship with the project:

### `pptx` — build a deck from a spec

The core deck-builder. `agent_workspace/skills/pptx/`:

- **`SKILL.md`** — workflow + the hard rules `check_pptx` enforces.
- **`design.md`** — palette, typography, spacing, and grader conventions.
- **`python-pptx.md`** — API reference and the reusable `textbox`/`paragraph`/`title`/`chart`/`table` helpers.
- **`naming.md`** — shape-naming convention (role prefixes: `background_`, `chartpart_`, `overlay_`, `meta_`).
- **`editing.md`** — techniques for modifying an existing `.pptx`.

### `deck-drafting` — make a deck *from scratch*

The upstream step: turn raw source material (a PDF, CSV/XLSX, report, or notes) into a content
spec in this project's dialect, which the `pptx` skill then builds into a `.pptx`. Use it when you
have sources rather than a ready `spec.md` — it reads the sources, decides the narrative and
per-slide content, prepares the chart/table data, and writes the spec (with action titles, a
layout-diversity budget, and per-slide source citations). `agent_workspace/skills/deck-drafting/`.

### `engineer` — work on the harness itself

Engineering practices for changing this codebase (the `check_pptx` validator, `run_agent.py`, the
build helpers): the `tools.py`/`utils.py` policy-vs-mechanics split, the validator contract, the
shape-role conventions, house style, and a checklist for adding or adjusting a layout check.
`agent_workspace/skills/engineer/`.

---

## Layout

```
run_agent.py              # single CLI entry point (spec + model → deck)
main.py                   # check_pptx CLI wrapper
harness/                  # the validator
  tools.py                #   the eight check_* functions + check_pptx()
  utils.py                #   shape flattening, geometry, alignment/overlap helpers
  text_metrics.py         #   measure_text_emu — real-font text width/height
specs/                    # example deck specs
agent_workspace/
  repo/                   # mounted into the sandbox as repo/
    task*.md, spec.md     #   task instructions + active spec
    reference_build_deck.py  # a complete working build script the agent reuses
    harness/              #   bundled self-contained copy of the validator
  skills/                 # skills the agent reads
    pptx/                 #   build a deck from a spec (python-pptx + check_pptx)
    deck-drafting/        #   sources → content spec (make a deck from scratch)
    engineer/             #   practices for working on the harness itself
  sandbox-<tag>/          # per-run sandbox (gitignored)
  out-<tag>/              # per-run output: out.pptx + run_log.txt (gitignored)
```
