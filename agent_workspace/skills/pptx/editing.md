# Editing Existing Presentations

How to **modify an existing `.pptx`** — retitle a slide, rewrite body text, swap chart numbers,
recolor, move/resize boxes, and add/delete/reorder slides — using this project's toolchain
(**python-pptx** for the mutation, **`check_pptx`** for validation).

> This is the editing counterpart to [python-pptx.md](python-pptx.md) (build-from-scratch).
> Read [SKILL.md](SKILL.md) for `check_pptx` and the QA rules, and [design.md](design.md) for the
> palette/typography/grader conventions an edit must keep satisfying.

**This project edits *in memory* with python-pptx — there is no unpack/pack/XML-file workflow and
no thumbnail/vision QA.** You open the deck, mutate Python objects, save, and run `check_pptx`.

> **About the helpers below.** There is **no installed `build_helpers` module** — the helper
> functions (`textbox`, `paragraph`, `rectangle`, `connector`, `title`, …) live as source in
> [python-pptx.md](python-pptx.md). When an edit needs to *add* a shape, **paste those helper
> definitions into the top of your edit script** (the same way every build script does).
>
> **Importing the overflow estimator.** The only thing you `import` from the harness is
> `measure_text_emu`. The harness ships **inside the working dir** — it's at `repo/harness/`
> (`repo/` is the project the agent edits; the validator `repo/main.py` + `repo/harness/` are
> mounted there). So `from harness.text_metrics import measure_text_emu` resolves **only when you
> run from inside `repo/`** — `cd repo` first, exactly like the build/validate flow in `task.md`:
>
> ```bash
> cd repo
> uv run --no-project --with python-pptx python edit_deck.py     # your edit script
> uv run --no-project --with python-pptx python main.py deck.pptx  # validate (or the validate_deck tool)
> ```
>
> If you can't `cd repo` (running elsewhere), prepend it to `sys.path` instead:
> `import sys; sys.path.insert(0, "repo"); from harness.text_metrics import measure_text_emu`.

---

## The edit loop

```
1. Baseline   → run check_pptx on the ORIGINAL first (see "Baseline first")
2. Open       → prs = Presentation(path)
3. Locate     → find the shape(s) by name
4. Mutate     → change text / style / geometry, or add/remove shapes/slides
5. Re-measure → if text got longer, measure_text_emu to confirm it still fits its box
6. Save       → prs.save(path)
7. Validate   → python main.py path   → fix → repeat until ok:true, no overflow finding
```

Always **work on a copy**, never the user's original until the edit is verified:

```bash
cp deck.pptx deck.edited.pptx     # edit + validate this; swap in only when green
```

### Baseline first (critical)

A deck you didn't build may already have `check_pptx` problems. **Run `check_pptx` on the
unmodified file first and record its problem set.** After your edit, you only own problems that are
*new* relative to that baseline — don't chase pre-existing ones you didn't introduce (and don't let
them mask a regression you did introduce). Re-validate after every save; one fix often surfaces another.

---

## 1. Locate the shape(s) to edit

Names follow [naming.md](naming.md), so they're predictable and greppable. Two ways to see them:

```bash
python -m markitdown deck.pptx          # all text, in reading order — find WHAT to change
```

```python
from pptx import Presentation
prs = Presentation("deck.pptx")
for i, s in enumerate(prs.slides):                      # i is 0-based here;
    print(f"--- slide {i}  (check_pptx calls this slide {i+1}) ---")
    for sh in s.shapes:
        txt = sh.text[:50].replace("\n", " ") if sh.has_text_frame else ""
        print(f"  {sh.name:30} [{sh.left and round(sh.left/914400,2)},"
              f"{sh.top and round(sh.top/914400,2)}]  {txt!r}")
```

> **Indexing gotcha:** `prs.slides[i]` is **0-based**; `check_pptx` reports slides **1-based**.
> A problem on "slide 3" lives at `prs.slides[2]`.

Find a shape by name:

```python
def shape_by_name(slide, name):
    return next(sh for sh in slide.shapes if sh.name == name)

title = shape_by_name(prs.slides[2], "title")
```

---

## 2. Edit text in place

Set the **run** text, not the shape, to preserve per-span formatting:

```python
title = shape_by_name(prs.slides[2], "title")
title.text_frame.paragraphs[0].runs[0].text = "A sharper, longer action title here"
```

- `shape.text = "..."` **wipes all runs and paragraphs** down to one plain run — only use it when
  you intend to discard existing formatting. To keep fonts/colors/bold, edit `runs[i].text`.
- **Multi-item content ⇒ one paragraph per item — never concatenate** into a single string.
  Reuse the build helper so bullets/indents stay consistent:

  ```python
  # textbox/paragraph: paste from python-pptx.md into your script (no module to import)
  # rebuild a list cleanly: clear the frame, then add one paragraph per item
  tf = shape_by_name(slide, "body_drivers").text_frame
  tf.clear()                                        # leaves a single empty paragraph
  items = ["Recurrent spending grew.", "Owned IP expanded.", "Margins held."]
  for k, line in enumerate(items):
      paragraph(tf, line, 15, INK, first=(k == 0), bullet=True, level=0)
  ```

### Longer text is the #1 edit hazard — re-measure it

The autofit flag (`auto_size`) grows the box only *at render time* and does **not** rewrite stored
geometry, so `check_pptx` can't see overflow from it — and there is **no vision model** here to
eyeball it. So after making text longer, confirm it still fits its box's stored width with the same
estimator `check_pptx` uses:

```python
from harness.text_metrics import measure_text_emu
box = shape_by_name(slide, "body_drivers")
_, needed_h = measure_text_emu(new_text, box.width, font_pt=15, leading_factor=1.2,
                               font_family="Arial")
print("fits" if needed_h <= box.height else f"OVERFLOWS by {(needed_h-box.height)/914400:.2f}in")
```

**Prefer a rephrase that fits the existing box.** The cleanest edit changes only text, not layout —
so when a title/body must change, first try wording that `measure_text_emu` says already fits
`box.height`. This is verified to add zero new `check_pptx` problems.

**If you must grow the box** (the new text genuinely needs more height), do NOT just "push
everything below down by the height delta" — that naïvely shoves fixed chrome (footer, page number,
source) off the page and breaks the deck's cross-slide vertical rhythm. Two rules make it safe,
both **verified** against this project's decks:

1. **Never move fixed chrome.** Skip any shape whose name contains `footer`, `pagenum`/`pageno`,
   `source`, or `meta` — these must keep the same `top` on every slide (the `vertical-rhythm`
   check). Move only body/content shapes between the title and the footer.
2. **Clear the box's *new* bottom plus a safety pad.** `check_pptx`'s overlap check uses an
   *estimated reflow height* that runs slightly past `measure_text_emu`'s number, so matching the
   delta exactly still collides. Place each below-it shape so its `top ≥ new_bottom + ~0.12in`,
   preserving its original gap:

   ```python
   EMU = 914400; PAD = int(0.12 * EMU)
   old_bottom = box.top + box.height
   box.height = needed_h                          # grow it
   new_bottom = box.top + needed_h
   FIXED = ("footer", "pagenum", "pageno", "source", "meta")
   is_fixed = lambda n: any(x in n.lower() for x in FIXED)
   for sh in slide.shapes:
       if sh is box or sh.top is None or is_fixed(sh.name):
           continue
       if sh.top >= old_bottom - int(0.1 * EMU):   # it sat below the title
           gap = sh.top - old_bottom               # keep its original offset
           sh.top = new_bottom + PAD + max(0, gap)
   ```

`check_pptx`'s `typography-text-fits-its-box` finding also catches overflow automatically on
validate — its appearance means "rephrase to fit, or grow-and-clear as above," then re-run.

---

## 3. Edit style (color, size, weight, font)

```python
from pptx.dml.color import RGBColor
from pptx.util import Pt
run = shape_by_name(slide, "title").text_frame.paragraphs[0].runs[0]
run.font.size  = Pt(30)
run.font.bold  = True
run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)     # never a "#hex" string

card = shape_by_name(slide, "card_cover_gta")        # recolor a fill
card.fill.solid(); card.fill.fore_color.rgb = RGBColor(0x0B, 0x1A, 0x2B)
```

- **Colors are `RGBColor(0xRR,0xGG,0xBB)`**, never strings.
- Recoloring has **no automated check** (geometry-only validator, no vision). Apply
  [design.md](design.md)'s palette and keep text-on-fill contrast **≥ 4.5:1 by rule**, computed as
  you choose colors — nothing verifies it after the fact.

---

## 4. Move / resize a shape

Geometry is EMU (`1 in = 914400`). Whatever you move must still satisfy `check_pptx`:

```python
EMU = 914400
card = shape_by_name(slide, "card_cover_2k")
card.left = int(1.0 * EMU); card.top = int(2.0 * EMU)
card.width = int(2.2 * EMU); card.height = int(1.0 * EMU)
```

Keep these `check_pptx` invariants when repositioning:
- **0.3 in outer margin** — nothing within 0.3 in of any slide edge.
- **Row peers**: equal `width` and equal `top`; **equal gutters** between 3+ items in a row.
- **No overlap** — two shapes' boxes must not intersect (watch filled panels especially).
- **Vertical rhythm**: titles (and footers) share the same `top` on every slide.

When you move one box, re-validate — a moved box commonly collides with its new neighbor.

---

## 5. Add or remove shapes

**Add** through the build helpers from [python-pptx.md](python-pptx.md) so the new shape gets a
proper `name`:

```python
# textbox/paragraph/... pasted from python-pptx.md (no build_helpers module exists)
_, tf = textbox(slide, int(0.45*EMU), int(6.8*EMU), int(11*EMU), int(0.3*EMU),
                name="footer_note")
paragraph(tf, "Updated October 2025", 10, MUTED, first=True)
```

**Remove** by detaching the element from its parent:

```python
sp = shape_by_name(slide, "card_cover_usd3.4b")
sp._element.getparent().remove(sp._element)
```

> **Template-slot ≠ source-item:** if a slide has 4 cards but your new content has 3, **delete the
> 4th card's entire shape** (and any label/icon that belongs to it) — don't just blank its text, or
> you leave an orphaned empty box that `check_pptx` still counts and a reader still sees.

---

## 6. Slide operations — add / delete / reorder

python-pptx has **no public API** for these; they're done on the slide-id list
(`prs.slides._sldIdLst`, the `<p:sldIdLst>` element). These recipes are **verified** against this
project's decks. `<p:sldId>` order *is* slide order.

```python
import copy
from pptx import Presentation
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

def delete_slide(prs, idx):
    """Remove slide #idx (0-based): drop its relationship, then its sldId entry."""
    sldId = list(prs.slides._sldIdLst)[idx]
    prs.part.drop_rel(sldId.get(RID))
    prs.slides._sldIdLst.remove(sldId)

def move_slide(prs, old_idx, new_idx):
    """Reorder by repositioning the <p:sldId> element."""
    lst = prs.slides._sldIdLst
    sldId = list(lst)[old_idx]
    lst.remove(sldId)
    lst.insert(new_idx, sldId)

def duplicate_slide(prs, idx):
    """Deep-copy slide #idx onto a fresh slide from the SAME layout, appended at the end.
    Copies the shape tree; good for text/shape slides. Slides with charts or images need
    their media/chart parts relinked separately — re-add those via the build helpers."""
    src = prs.slides[idx]
    new = prs.slides.add_slide(src.slide_layout)
    for sh in list(new.shapes):                 # clear layout-injected placeholders
        sh._element.getparent().remove(sh._element)
    for sh in src.shapes:
        new.shapes._spTree.append(copy.deepcopy(sh._element))
    return new                                  # then move_slide() it into position
```

After any slide op, **re-open the saved file** (`Presentation(path)`) to confirm it isn't corrupt,
and run `check_pptx`. (`add_slide` appends to the end — call `move_slide` to place it.)

> **CRITICAL — save and re-open between a delete and a later add/duplicate.** `delete_slide` drops a
> slide's relationship but its underlying part lingers in the in-memory package, so a *subsequent*
> `add_slide`/`duplicate_slide` on the **same** `Presentation` object reuses an existing slide
> partname and writes a **corrupt** file (you'll see `UserWarning: Duplicate name:
> 'ppt/slides/slideN.xml'` on save). **Verified.** Do slide ops as: mutate → `prs.save(path)` →
> `prs = Presentation(path)` (fresh re-open) → next mutation. Re-opening rebuilds the package and
> renumbers parts cleanly. (Pure duplicates with no prior delete in the same object are fine; it's
> the delete-then-add sequence that bites — so when in doubt, save-and-reopen between every slide op.)

> **Reorder renumbers slides.** When you compare against your baseline after a delete/reorder, match
> problems by **message**, not slide number — a pre-existing issue on "slide 10" will reappear as
> "slide 11" after a reorder. That's the same problem moving, not a new regression.

> **A duplicated slide is real content.** `check_pptx` validates the copy like any slide — if the
> source had an off-grid left edge, the duplicate inherits it and may be flagged. That's a genuine
> finding about the copy, not corruption; fix the geometry as you would any other shape.

---

## Common pitfalls (this toolchain)

- **Forgot to re-validate.** Every save must be followed by `check_pptx`. Editing is iterative.
- **Longer replacement overflows.** Re-measure with `measure_text_emu`; grow the box and push the
  shape below down. (No vision model — you cannot "just look.")
- **Concatenated multi-item content** into one paragraph. Use one `paragraph()` per item.
- **Blanked text but left the shape.** Delete the whole element when an item drops; an empty box is
  still counted and still seen.
- **Chasing pre-existing problems.** Baseline the original first; own only *new* problems.
- **`shape.text = ` nuked formatting** you wanted to keep. Edit `runs[i].text` instead.
- **Stacked content on a filled panel** → overlap flag. Use the true `<p:bg>` background for
  full-bleed color, or place content beside the panel (see python-pptx.md "Backgrounds").
- **Smart quotes.** python-pptx writes the literal characters you put in the string, so just use
  real `"" ''` in Python text; there is no XML escaping step in this workflow.

---

## Verification

```bash
python main.py deck.edited.pptx        # must be ok:true, no typography-text-fits-its-box finding
python -m markitdown deck.edited.pptx  # confirm the intended text actually changed; nothing orphaned
```

**Done = `check_pptx` returns `ok:true` with no overflow finding, the intended change is present,
and (for slide ops) the saved file re-opens without error** — measured against the baseline you
recorded before editing.
