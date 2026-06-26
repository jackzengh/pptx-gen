# Design Rules (mandatory)

Concrete, non-negotiable rules that `check_pptx` cannot fully verify but a reader sees instantly. They exist because real generated decks broke each one. The target is a BCG/McKinsey-grade deck: **dense, deliberate, edge-to-edge** — not an AI deck with a chart floating in a sea of empty space.

**This is the single source for design decisions — read it before you build.** SKILL.md points here on purpose; palettes, typography, layout patterns, spacing constants, and the grader conventions all live in this file.

---

## Design ideas — start here

**Don't create boring slides.** Plain bullets on a white background won't impress anyone.

### Before starting

- **Pick a bold, content-informed palette.** If swapping your colors into a totally different deck would still "work," they aren't specific enough.
- **Dominance over equality.** One color dominates (60–70% of visual weight), with 1–2 supporting tones and one sharp accent. Never give all colors equal weight.
- **Dark/light contrast.** Dark backgrounds for title + conclusion, light for content ("sandwich"). Or commit to dark throughout for a premium feel.
- **Commit to a visual motif.** Pick ONE distinctive element and repeat it on every slide — rounded image frames, icons in colored circles, a thick single-side band.

### Color palettes

Don't default to generic blue. Use these as inspiration:

| Theme              | Primary  | Secondary | Accent   |
| ------------------ | -------- | --------- | -------- |
| Midnight Executive | `1E2761` | `CADCFC`  | `FFFFFF` |
| Forest & Moss      | `2C5F2D` | `97BC62`  | `F5F5F5` |
| Coral Energy       | `F96167` | `F9E795`  | `2F3C7E` |
| Warm Terracotta    | `B85042` | `E7E8D1`  | `A7BEAE` |
| Ocean Gradient     | `065A82` | `1C7293`  | `21295C` |
| Charcoal Minimal   | `36454F` | `F2F2F2`  | `212121` |
| Teal Trust         | `028090` | `00A896`  | `02C39A` |
| Berry & Cream      | `6D2E46` | `A26769`  | `ECE2D0` |
| Sage Calm          | `84B59F` | `69A297`  | `50808E` |
| Cherry Bold        | `990011` | `FCF6F5`  | `2F3C7E` |

> In python-pptx, write these as `RGBColor(0x1E, 0x27, 0x61)` — never `"#1E2761"`.

### Layout & data display

**Every slide needs a visual element** — image, chart, icon, or shape. Text-only slides are forgettable.

- Two-column (text left, visual right) · icon + text rows · 2×2 / 2×3 grids · half-bleed image with overlay.
- Large stat callouts (60–72pt numbers, small labels) · before/after columns · numbered process flows.

### Visual polish

- **Icons in small colored circles** next to section headers.
- **Italic accent text** for key stats or taglines.

### Typography

Choose an interesting font pairing — don't default to Arial.

| Element        | Size          |
| -------------- | ------------- |
| Slide title    | 36–44pt bold  |
| Section header | 20–24pt bold  |
| Body text      | 14–16pt       |
| Captions       | 10–12pt muted |

Pairings: Georgia/Calibri · Arial Black/Arial · Calibri/Calibri Light · Cambria/Calibri · Trebuchet MS/Calibri · Impact/Arial · Palatino/Garamond · Consolas/Calibri.

### Spacing

- 0.5" minimum margins (the outer 0.1" is a hard fail — keep well clear).
- 0.3–0.5" between content blocks. Pick one gap and use it consistently.
- Leave breathing room — don't fill every inch.

### Avoid (common mistakes)

- **Don't repeat the same layout** — vary columns, cards, and callouts.
- **Don't center body text** — left-align paragraphs and lists; center only titles.
- **Don't skimp on size contrast** — titles 36pt+ vs 14–16pt body.
- **Don't default to blue** — pick colors that reflect the topic.
- **Don't mix spacing randomly** — choose 0.3" or 0.5" gaps and stick to it.
- **Don't create text-only slides.**
- **Don't forget text-box padding** — when aligning text to a shape edge, zero the frame margins (`tf.margin_left = Pt(0)`, etc.).
- **Don't use low-contrast elements** — icons AND text need strong contrast against the background; avoid light text on light or dark text on dark (see the WCAG ≥4.5:1 rule below).
- **NEVER use accent lines under titles** — a hallmark of AI-generated slides. Use whitespace or background color instead.

---

## Grader conventions (these score the deck — follow every one)

A separate rubric grader scores typography, color, and IB conventions beyond geometry. These rules are what took a deck from 41% → 100%. Each is a hard pass/fail (and several feed binary penalties worth −10 to −15), so satisfy all:

1. **Action titles — every slide, including section dividers and a contents page.** A title must be a full sentence: **≥5 words AND contain a finite verb**. The grader matches verbs from a fixed lexicon — safe choices: *is, are, drives, leads, reached, grew, rose, holds, anchors, exceeds, shows, supports*-NO/*added*-NO. Avoid *supports, contributes, represents, measures, examines, powers, reinforces, generated, assembled* — they are NOT in the lexicon and FAIL. Rephrase: "The F-Series **drives** ~$49 billion in US GDP" (not "contributes"); "Ford **leads** US pickup assembly" (not "assembled"). Bare noun-phrase titles ("Overview", "Context") fail.
2. **Source line on EVERY data slide** (any chart/table/number). It must start with the literal marker **"Source:" or "Note:"** — and note the marker regex matches the whole word, so **"Sources:" (plural) FAILS**. Always write **"Source:"** even for multiple sources.
3. **Page numbers** on every content slide (not the cover), monotonic, bottom corner. Name them `pagenum_N`. **Keep the box fully inside the 0.3in margin band** — on a 13.333in slide, the page-number box's right edge must be ≤ 13.03in (e.g. left ≈ 12.55in, width ≈ 0.4in, right-aligned), and its bottom ≤ 7.18in. A box at x=12.78 width=0.4 (right edge 13.18) intrudes the band and triggers the −15 margin penalty.
4. **One font size per logical level deck-wide.** Pick exactly one size each for: title (all action titles same pt), body, and footnote/caption/source. The grader lumps anything named `caption*` with footnotes — so **captions and source lines must be the same size** (e.g. both 9pt). Big hero stat numbers may be their own consistent "display" tier.
5. **≤6 non-neutral hues** deck-wide (black/white/grays don't count). One brand color + a dark variant + light tint is plenty.
6. **Contrast ≥4.5:1 for every text run vs its background.** White text on a *bright* mid-tone (e.g. green #2E9E7B) is only ~3.3:1 → FAIL; use a **darker panel** (#1B6B52 gives 6.4:1) so white text passes. Never put muted grey (#666) on a colored panel — move that text onto white or make it white. Compute the WCAG ratio when unsure.
7. **Fixed title-top y on every content slide** (the cover/title slide is exempt). Eyebrow labels go in a reserved strip ABOVE the fixed title-top — they must not push the title down.
8. **Nothing in the outer 0.3in margin band.** Keep all shapes (except intentional full-bleed `background_*` panels/images) ≥0.3in from every edge. The footer/source band must end by ~7.16in on a 7.5in slide — bottom-anchor multi-line sources so they grow up, not down.
9. **One message per slide; consistent number formatting** (same decimals/currency-scale/unit per metric).

(`check_pptx` covers geometry; this rubric covers the rest. A deck that passes check_pptx can still score 40% on the rubric — both matter.)

---

## Vertical rhythm — fixed constants (use these exactly)

Pick once, apply on **every** content slide so the eye never re-finds the title (`check_pptx` enforces title `vertical-rhythm` and left-margin consistency). All in inches.

```
MARGIN          = 0.45   # left/right/top working margin (NOT 0.6 — that wastes space)
TITLE_TOP       = 0.45   # y of the title box top (same on every content slide)
TITLE_PT        = 26-30  # action-title size; sized for its REAL wrapped height
TITLE_GAP       = 0.14   # gap from the title's RENDERED bottom to the accent connector
CONN_TO_BODY    = 0.18   # gap from the connector down to where content begins
BLOCK_GAP       = 0.22   # gap between stacked content blocks
GUTTER          = 0.25   # gap between columns / cards (never < 0.15)
SRC_Y           = 7.05   # source line baseline (bottom band, same every slide)
LINE_SPACING    = 1.2    # body line spacing (1.15–1.3; single is cramped, double wastes space)
BULLET_SPACE_AFTER = 6pt # paragraph space-after between bullets — NEVER a blank-line paragraph
```

Footnote: title size and `TITLE_TOP` must hold across the deck. If one title is much longer, keep the *top* fixed and let the box grow **down** (measured) — never shift the top.

---

## Fill the slide — kill dead whitespace

The #1 failure mode is an empty-looking slide. Fix it:

- **Margins 0.4–0.5in**, not 0.6+. The content area is ~12.4 × 6.6in — use it.
- A data slide is **primary visual + supporting column + source line**, sized to nearly fill the area below the title. A chart should be ~4.2–4.8in tall and ~6.5in wide, not a small box mid-slide.
- Prefer **dense multi-column / sidebar / icon-row / half-bleed-image** layouts over one element centered in space.
- Stat cards and columns extend to the working margins; the bullet column fills its height.
- **But never crowd:** keep ≥0.15in gutters and ≥0.12in between any text and a shape edge. Dense ≠ cramped. The goal is *intentional* fullness, like the reference BCG deck where every region carries content or a deliberate color field.
- If content doesn't fit, **cut words** (tighten bullets) — never shrink body below 11pt or titles below 24pt.

---

## Titles & the accent connector

- **Size the title box for its REAL wrapped height** (measure it — `title()` helper uses `measure_text_emu`). Never a fixed 0.7in: a 2-line title that gets 0.7in overflows and collides with content below, which `check_pptx`'s `typography-text-fits-its-box` flags.
- **Accent connector sits `TITLE_GAP` (≈0.14in) below the title's rendered bottom** — a clear gap, never cramped against the descenders. ~1.6in wide, ~2.5pt, accent color. Content begins `CONN_TO_BODY` (≈0.18in) below the connector. The `title()` helper returns that content-start y — always position content from its return value, never a hardcoded number.
- This connector is the deck's spine motif and is the **one** deliberate line under a title. It is NOT the "no full-width hairline under titles" anti-pattern (a faint rule spanning the whole title that fakes structure — never do that).
- **Eyebrow label (signature BCG move, optional):** a 10–11pt ALL-CAPS muted-or-accent label directly above the action title (e.g. `F-SERIES`, `FORD`). Sits ~0.18in above the title top; shift the title down by that much on eyebrow slides but keep the eyebrow's top at the fixed `TITLE_TOP`.

## Sub-headers vs. titles (different rules)

- **Sub-headers** ("Summary of the study", "Approach", "Why it matters") are small bold labels INSIDE a column. They get a **short thin accent underline** beneath them (~0.5–0.8in wide, ~1pt). This underline-under-subheader is correct and on-brand.
- The **main action title** gets the connector (above), **not** a sub-header underline. Don't confuse the two.

---

## Charts

- **NEVER leave the default "Chart Title."** Set `chart.has_title = False`; the slide's subtitle/caption above the chart is the title. The `chart()` helper does this.
- Series colors from the palette; **highlight one bar/point in the accent color and grey the rest** when one data point proves the title. Muted axis labels, subtle gridlines, no 3-D, no chart junk.
- Bar = comparison, line = trend over time, pie = parts-of-a-whole (few slices only; color each slice + legend).
- Needs a **caption within 0.1in above it** (the subtitle) for `element-boxing`.
- Make charts **large** (≈4.2–4.8in tall) — a small chart is the main whitespace offender.

## Tables

- **Readable:** header ≥12pt bold (white on filled header row), body ≥11pt. Row heights ≥0.3in. Never tiny text.
- Alternating subtle row fill or none; header row in primary/accent color.
- Caption within 0.1in above it.

## Stat cards / KPIs

- Row of cards: **equal width, equal top, even gutters** (`check_pptx` enforces). Big number ≥28pt bold in an accent color + small label ≤12pt beneath, **both inside the card shape's own text frame** (don't stack separate overlapping textboxes on the card).
- Size each card for its content; cards extend to the working margins to fill the row.

## Section / divider / quote slides

- **Color-sidebar divider:** left ~third is a solid brand-color panel (a true named shape `background_sidebar`) listing all section names — current one white/bold, the rest greyed at lower opacity; the other two-thirds is a large image. (See Layout patterns.)
- **Pull-quote:** large quote text + attribution on a dark background. Motif anchored top-right like content slides — **never** scatter loose colored bars floating mid-slide (reads as a bug).
- Deliberate, not decorated.

## General

- Body text left-aligned; center only short titles/labels if at all.
- ≥0.4in from edges (hard fail inside 0.1in).
- One dominant color (60–70%), 1–2 supporting tones, one sharp accent. Dark cover/section/closing, light content ("sandwich").
- Every text frame via the `textbox()` helper (`auto_size = SHAPE_TO_FIT_TEXT`, zero margins). **Name every shape.**

---

## Layout patterns (the BCG/Ford visual language)

Concrete geometry for each. The reference is the Ford/BCG F-Series Economic Impact deck: full-bleed color fields, half-bleed photos, offset accent squares, annotated bridge charts, icon grids — almost no dead space.

### Eyebrow label + action title
- Eyebrow textbox at `(MARGIN, TITLE_TOP)`, 10–11pt ALL-CAPS, accent or muted, height ~0.22in.
- Title textbox just below it; connector `TITLE_GAP` under the title's measured bottom.

### Offset accent square
- A small filled square (~0.45 × 0.45in, accent color, often lower opacity) peeking from behind the top-left of the title block: place it at `(MARGIN-0.15, TITLE_TOP-0.05)` *before* the title so the title overlaps its corner. Named `motif_square`. (Watch overlap rules: give it no text so it's `kind=other`, exempt from alignment checks but still margin/overlap-checked — keep it inside the edge band.)

### Color sidebar layout
- `background_sidebar` rectangle: `x=0, y=0, w≈4.4in, h=7.5in`, brand color, true shape (NOT `<p:bg>`, since it's a partial panel). No border.
- White emphasis title or a section-nav list inside it, left-padded to ~0.45in from the panel's left.
- Right region (`x≈4.7in → 13.0in`) holds the image or content. Keep content ≥0.2in clear of the panel's right edge.

### Half-bleed image
- Image fills one half edge-to-edge: e.g. `x=6.9, y=0, w=6.43, h=7.5` (right half) via `add_picture` with explicit w/h (crop/scale the source to the box ratio first to avoid distortion).
- Content on the other half within the working margin. A small `Image: …` caption (9pt, white or muted) bottom-left of the image.

### Annotated waterfall / bridge chart
Bars step up to a total. Two build options:
- **Manual rectangles (most control, recommended):** draw each step as a named `rectangle` at computed (x, y, h) so bar *i* starts at the running cumulative total; label each bar with its value (textbox centered on the bar); add dotted connectors between step tops; final **Total** bar in a darker shade. Add grouping **brackets** above ranges (a thin connector with short down-ticks at each end) labelled "Product GDP contribution", "Multiplier effect", etc. Category labels below the axis; sub-captions under categories ("Suppliers", "Dealers (Sales)").
- **Stacked column:** an invisible base series (no fill) sized to each bar's start, plus a visible series for the step height. Less control over annotations; prefer manual for brackets/labels.

### Comparison bar chart with multiplier row
- Descending bars; **hero bar in accent green, the rest grey**. Value label on/above each bar.
- A row of relative-size multipliers ("0.8x", "1.2x", "2.2x"…) in a thin band beneath the category axis, left-labelled (e.g. "F-Series relative size"). Optional small product image atop the hero bar.

### Icon grid (2×N)
- Rows of: icon-in-colored-square (~0.6in square, accent fill, white icon) + bold header (12–14pt) + 2–4 line description (11pt) beneath.
- Left-margin **row labels** in italic muted ("Industry", "Citing patent") aligned to each row band.
- Columns equal width, even gutters; two rows of four is the canonical density.

### Stat-equation hero
- Big numbers with `×` and `=` operators across a faded full-bleed image (true `<p:bg>` image at low contrast, or an image rect with a translucent color overlay).
- e.g. `~17M  ×  26–35%  ×  2.1–2.4  =  13M`, each operand with a small caption beneath; a divider rule before the `=` result. A supporting grey sidebar on the right with a short bulleted list.

### Section divider with nav
- Color sidebar (above) listing every section; current section white/bold, others greyed (~50% lighter). Large image fills the rest. No source line needed on pure dividers.

---

## Principles from professional / consulting decks

The habits that separate a polished deck from an AI-looking one.

**Action titles (the biggest lever).** The title is a *complete sentence stating the takeaway* — "The F-Series supports ~500,000 American jobs, ~13–14 per direct employee" — not a label like "Employment." Coined by BCG; tells the reader what to conclude. The visual is the evidence. Preserve the spec's full-sentence titles verbatim.

**One idea per slide.** Two messages → two slides. The action title is that idea; everything else supports it.

**Invisible grid.** Same margins, title position, body area, and footnote line on every slide. The constants above enforce this.

**Two-column is the workhorse** (evidence + takeaways), but vary with sidebars, icon grids, full-bleed stats, and bridge charts so the deck isn't monotonous — like the reference deck, which uses a different layout almost every slide while holding the grid.

**Highlight one number, grey the rest.** Emphasize the single data point that proves the title; mute the rest.

**Typography = hierarchy through weight, not font count.** Max 2 families (e.g. a strong sans for titles + clean sans for body, or Georgia/Arial). Size + weight make hierarchy; never use underline for emphasis (reads as a link) — that's reserved for sub-headers.

**Source line on every data slide.** Small muted footnote at `SRC_Y` ("Sources: …"). Footnote superscripts (¹ ²) tie claims to notes, BCG-style.

Sources: [Deckary — Consulting Slide Standards](https://deckary.com/blog/consulting-slide-standards), [Deckary — BCG Presentation Style](https://deckary.com/blog/bcg-presentation-style), [Deckary — PowerPoint Design Guide](https://deckary.com/blog/pillar-powerpoint-design-guide), [Slidor — 15 design rules](https://www.slidor.agency/blog/15-regles-illustrees-pour-ameliorer-le-design-de-vos-presentations-powerpoint), [Flashdocs — 10 design principles](https://www.flashdocs.com/post/10-design-principles-every-slide-creator-should-know), [Microsoft — text spacing in PowerPoint](https://support.microsoft.com/en-us/office/change-text-alignment-indentation-and-spacing-in-powerpoint-36df6486-0118-4e95-b9bd-835eb047ac88).

---

## Grading conventions (consulting rubric — required to score)

A visual grader scores decks on layout, typography, color, and IB conventions. These rules are what separates a 40% deck from a 90%+ deck — apply ALL of them.

**Action titles on EVERY slide (incl. sections & contents).** Every title is a full sentence with a verb and ≥5 words stating a takeaway — never a bare noun phrase. Bad: "Context", "Contents of this report", "Employment impact". Good: "This study measures Ford's economic impact across four dimensions", "Ford and the F-Series support hundreds of thousands of US jobs". Section dividers get a headline sentence too (the nav list is secondary).

**One size per logical level, deck-wide.** Pick ONE pt size for each tier and reuse it everywhere — the grader fails the deck if titles (or body, or footnotes) use 2+ different sizes across slides. A safe tier set:
- Title = 26pt (ALL action titles: content, section, cover, closing)
- Display = 30pt (big hero stat numbers only)
- Body = 13pt (body, bullets, callouts, column headers, captions-in-content)
- Subtitle/caption/eyebrow = 11pt
- Source/footnote/page-number = 9pt
Keep title > body by ≥4pt and body > footnote by ≥2pt (size hierarchy).

**Fixed title-top y on every content slide.** The title box TOP is identical deck-wide (e.g. 0.6in). Eyebrows do NOT push the title down — place the eyebrow in a reserved strip ABOVE the fixed title-top (e.g. 0.32in). The footer/source band and page numbers also share one fixed top.

**≤6 non-neutral accent hues deck-wide.** Black, white, and grays (R≈G≈B) don't count. Use ONE accent family (e.g. green + a darker green + a light green) and render all "muted" elements in grayscale neutrals — including chart competitor bars (a gray ramp, not extra hues).

**Contrast ≥4.5:1 (WCAG).** Source/footnote text must be a dark-enough gray on light backgrounds (e.g. `#606A72` or darker, NOT `#9AA3AB`). On a colored sidebar, use white or a very light tint of the sidebar color (e.g. `#E9F4EF` on green) — never a mid gray.

**Source/Note line on EVERY data slide.** Any slide with a chart, table, or numeric data needs a text box whose first line literally starts with "Source:", "Sources:", or "Note:".

**Page numbers on every content slide** (not the cover), monotonically increasing, small + muted, bottom-right, at the fixed footer top.
