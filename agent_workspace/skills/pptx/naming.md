# Shape Naming Convention

`check_pptx` echoes each shape's `shape.name` in every problem it reports. A problem like:

```
'background_header' and 'card_revenue' overlap by 0.42 in^2  (slide 2)
```

is instantly actionable. The default — `'shape'` — is not. So **name every shape the moment you create it**:

```python
box.name = "title"
card.name = "card_revenue"
```

The name lives in the shape's `<p:cNvPr name="...">` element and has no visual effect; it is purely a label for you and for the validator.

## Conventions

Use a `kind_descriptor` form, lowercase, snake_case. Add a slide suffix (`_s3`) when the same descriptor repeats across slides and you need to tell them apart.

| Prefix | Use for | Example |
|--------|---------|---------|
| `title` | the slide title | `title` |
| `body_*` | main body / paragraph text | `body_intro` |
| `background_*` | colored panels, bands, sidebars (shapes that sit in the layout) | `background_header`, `background_sidebar_s3` |
| `card_*` | content cards / boxes | `card_revenue`, `card_cost` |
| `kpi_*` | big-number stat callouts | `kpi_arr` |
| `caption_*` | the caption above a chart/table | `caption_chart_revenue` |
| `chart_*` | charts | `chart_revenue` |
| `table_*` | tables | `table_pricing` |
| `picture_*` | images | `picture_logo` |
| `icon_*` | small icons | `icon_check` |
| `divider_*` | connector lines / rules | `divider_s1` |

## Two rules

1. **"Background" goes in the name of background shapes.** Any colored panel, band, or sidebar that is a *shape* (it appears in `slide.shapes` and is checked for overlap/margins) should have `background` in its name — e.g. `background_header`. This makes it obvious in QA output which element is the backdrop vs. the content.

2. **Full-bleed color is NOT a shape, so it needs no name.** Set it with `slide.background.fill` (it lives in `<p:bg>`). It never appears in `slide.shapes` and `check_pptx` never sees it. Reserve named shapes for elements that should actually participate in the layout grid (titles, cards, charts, captions, and intentional panels/bands).

See [python-pptx.md](python-pptx.md) for the panel-vs-true-background code.
