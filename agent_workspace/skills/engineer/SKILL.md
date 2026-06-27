---
name: engineer
description: Engineering practices for working on the pptx-gen harness — the check_pptx validator, the agent runner, and the deck-build helpers. Use when changing harness/ (tools.py, utils.py, text_metrics.py), run_agent.py, the build helpers, or adding/adjusting a layout check.
---

# Engineering Practices (pptx-gen)

Best practices for maintaining the deck-building harness in this repo. This is a small,
focused Python project — a layout validator plus a sandboxed agent that builds decks. Keep
changes readable, reuse what's here, and validate against a real deck before you're done.

## The architecture

```
run_agent.py        # single CLI entry point: spec + model → sandboxed agent → built deck
main.py             # `python main.py <deck.pptx>` → prints check_pptx JSON
harness/            # the validator
  tools.py          #   the eight check_* rules + check_pptx() — read as pure policy
  utils.py          #   data model (Box/SlideInfo/DeckContext), shape flattening, geometry
  text_metrics.py   #   measure_text_emu — real-font text width/height (no renderer)
```

The split between `tools.py` and `utils.py` is the central design choice: **checks are pure
policy, helpers are mechanics.** A reader of `tools.py` should see *what* the rule is ("peers in
a row must share a width") without wading through EMU math or XML traversal — that lives in
`utils.py`, justified in place. Preserve this separation: don't inline geometry into a check, and
don't put rubric policy into a helper.

## The validator contract

`check_pptx(path) -> {ok, problems, notes}`:
- `ok` is `True` only when `problems` is empty.
- each problem is a plain dict: `{"slide", "message", "shapes", "boxes_in"}` — build it as a
  literal at the call site (there is no `problem()` helper; it was inlined).
- `notes` surfaces per-slide extraction caveats.

The eight checks (`tools.py`): `check_outer_margin_frame`, `check_no_overlap`,
`check_left_margin_consistency`, `check_edge_and_grid_alignment`, `check_consistent_gutters`,
`check_element_boxing`, `check_text_overflow`, `check_y_offset_variation`. `check_pptx` flattens
every slide into `Box`es (`collect_shapes`), wraps them in a `DeckContext`, and aggregates all
eight.

## Conventions that matter here

- **Shape roles** are read from the shape `name` by `role_of` (`utils.py`): the prefixes
  `background_`, `chartpart_`, `overlay_`, `meta_` mark shapes that are exempt from certain
  checks (a background panel may touch the edge; chartparts may overlap each other; overlays sit
  on content; meta is chrome). When adding a check, decide how it treats each role.
- **Helpers are grouped by purpose and named for it:** `alignment_*` (cluster/peer-group the grid),
  `overlap_*` (text-aware ink area), plus `role_of`/`kind_of`/`grid_boxes`. Reuse these before
  writing new geometry — most "I need to group boxes" needs are already covered by
  `alignment_peer_clusters` and `alignment_cluster_1d`.
- **Geometry is in inches**, stored from EMU via `emu_to_in`; `TOL` (0.1in) is the single shared
  alignment tolerance; `MARGIN_BAND` (0.3in) is the outer margin. Use the constants, don't
  hardcode.
- **Text sizing uses `measure_text_emu`** (real Liberation-font metrics), not character counts.
  Overflow is *estimated*, not rendered — a clean overflow check doesn't prove text fits, only
  that no collision was detected.
- **Don't rely on `auto_size` / `SHAPE_TO_FIT_TEXT`.** The autofit flag never rewrites stored
  geometry, so the validator measures against the **declared** box. Size boxes from real geometry.

## House style

- **One-line `#` comments per function**, in plain language explaining the *why* — match the
  existing density (e.g. `# which shapes need to abide by alignment? some don't — they're inside
  others, or backgrounds`). No multi-line docstrings on the check/helper functions.
- **Inline trivial helpers** rather than keeping one-line wrapper functions; name what's left for
  what it does.
- **Reuse before adding.** Check `utils.py` for an existing clustering/geometry helper before
  writing a new one. Abstractions that destroy local context aren't worth it in a file this size.

## Quick checklist when changing the harness

- [ ] **Policy vs. mechanics** — new rule logic goes in a `check_*` in `tools.py`; new geometry
      goes in `utils.py`. Don't cross the streams.
- [ ] **Reuse helpers** — did you check `utils.py` (`alignment_*`, `overlap_*`, `role_of`,
      `grid_boxes`) before adding new code?
- [ ] **Roles** — does your change handle `background_`/`chartpart_`/`overlay_`/`meta_` shapes
      correctly (exempt where intended)?
- [ ] **Constants** — used `TOL` / `MARGIN_BAND` / `CAPTION_GAP`, not magic numbers?
- [ ] **Problem shape** — any new finding is a `{slide, message, shapes, boxes_in}` dict.
- [ ] **Comments** — one-line `#` comment per function, in the house style.
- [ ] **Validate end-to-end** — run `python main.py <a real deck.pptx>` (or `check_pptx` directly)
      on a known deck and confirm the result is sane, not just that imports succeed.
