# Shape Naming Convention

`check_pptx` reads each shape's `shape.name` for **two** purposes:

1. **Legible QA** — every problem it reports echoes the names, so
   `'card_revenue' and 'card_cost' overlap by 0.42 in^2` tells you exactly what to move (the python-pptx default `'Rectangle 3'` does not).
2. **Role detection** — the *prefix before the first `_`* tells `check_pptx` how to treat the shape. A few prefixes opt a shape OUT of specific checks (see "Role prefixes" below). This is how BCG-style full-bleed panels, hand-drawn charts, and on-image labels pass while ordinary content stays strictly checked.

So the name is **load-bearing**, not cosmetic. **Name every shape the moment you create it**, `lowercase_snake_case`, as `role-or-kind_descriptor`. Add a slide suffix (`_s3`) when a descriptor repeats and you need to tell instances apart.

```python
box.name = "title"
card.name = "card_revenue"
panel.name = "background_sidebar"
```

---

## 1. Content names (checked normally)

Ordinary content — these participate fully in the layout grid, overlap, and margin checks. Use a clear `kind_descriptor`:

| Prefix | Use for | Example |
|--------|---------|---------|
| `title` | the slide title (exactly one per slide) | `title` |
| `body_*` | paragraph / bullet text blocks | `body_intro` |
| `card_*` | content cards / boxes | `card_revenue` |
| `kpi_*` | big-number stat callouts | `kpi_arr` |
| `chart_*` | native charts (`add_chart`) | `chart_revenue` |
| `table_*` | tables | `table_pricing` |
| `picture_*` | real images (`add_picture`) | `picture_logo` |

---

## 2. Role prefixes (opt OUT of specific checks)

There are **four** role prefixes. Each maps to exactly one exemption behavior. Name a special shape `<role>_<descriptor>` — e.g. `background_sidebar`, `chartpart_wf_total`, `overlay_image_credit`, `meta_pagenum_3`. If you don't use the right role prefix, the shape gets fully checked and you'll get false overlap/margin failures.

| Role | Use for | Exempt from |
|------|---------|-------------|
| **`background_*`** | full-bleed panels and image *fields* — a colored sidebar, an edge-to-edge image placeholder. Meant to touch the slide edge and have text sit on top. | outer-margin check; being an overlap *victim* (text on it isn't a collision). |
| **`chartpart_*`** | any piece of a **hand-drawn** chart you build from shapes: waterfall/bridge bars, comparison bars, scatter points + their labels, grouping brackets, the multiplier row, axis labels. The whole chart is ONE unit, so its parts legitimately differ in width/gutter and overlap each other. | grid checks (unequal bar widths/gutters are correct); the parts overlapping each other. |
| **`overlay_*`** | a label that sits ON another element: a caption on an image field, a tile number badge, an on-chart annotation/callout. | overlapping the thing it labels. |
| **`meta_*`** | non-grid chrome: eyebrow labels, footer, source line, page number, sidebar nav list, rules/dividers, legend, icon-row sub-headers. | the left-margin / row / gutter / alignment grid checks. |

Examples:
```python
sidebar.name      = "background_sidebar"        # full-bleed green panel
imgfield.name     = "background_image_cover"     # edge-to-edge photo placeholder
bar.name          = "chartpart_wf_suppliers"     # one waterfall step bar
barlabel.name     = "chartpart_wf_suppliers_val" # its value label (sits on the bar)
bracket.name      = "chartpart_bracket_multiplier"
caption.name      = "overlay_image_credit"       # "Image: Ford." on the photo
eyebrow.name      = "meta_eyebrow"
pagenum.name      = "meta_pagenum_3"
source.name       = "meta_source"
```

> The descriptor after the role still matters for QA — `chartpart_wf_total` reads far better in a failure message than `chartpart_7`. Keep it meaningful.

> **Legacy aliases:** older bare prefixes (`wf_`, `cmp_`, `plot_`, `pt_`, `caption_`, `eyebrow_`, `pagenum_`, `image_`, `sidebar_`, …) still work for backward compatibility, but new decks should use the four roles above — they're easier to remember and map 1:1 to behavior.

---

## Two more rules

1. **Full-bleed color via `<p:bg>` needs no name.** A true slide background (`slide.background.fill`) lives in `<p:bg>`, never appears in `slide.shapes`, and `check_pptx` never sees it. Use it for whole-slide color. Use a named `background_*` *shape* only for a partial panel/sidebar/band that should participate in overlap/margin checks.

2. **Name it as you create it.** An unnamed shape reports as `'shape'` and gets fully checked — so a full-bleed panel you forgot to prefix `background_` will be (wrongly) flagged for bleeding off-slide.

See [python-pptx.md](python-pptx.md) for the panel-vs-true-background code and [design.md](design.md) for when to use each layout.
