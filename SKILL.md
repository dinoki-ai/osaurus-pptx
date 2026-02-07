# Osaurus PPTX — Agent Skill Guide

Use this guide when creating PowerPoint presentations with the Osaurus PPTX plugin.

## Overview

This plugin **creates** PowerPoint (.pptx) presentations from scratch. It is not a full round-trip editor — treat every presentation as a fresh build.

## Workflow

Always follow this sequence:

1. **create_presentation** — get a `presentation_id`, slide dimensions, and theme info
2. **add_slide** (repeat) — add slides one at a time; each returns a `slide_number`
3. **add_text / add_image / add_shape / add_table / add_chart / set_slide_background** — populate each slide with elements
4. **save_presentation** — write the final `.pptx` file

Never skip steps. You must create slides before adding elements. You must save at the end — nothing is written to disk until `save_presentation`.

## Slide Coordinate System

All positions and sizes are in **inches** measured from the top-left corner of the slide.

### Default slide dimensions (16:9)

- Width: **13.33"**
- Height: **7.5"**

### 4:3 slides

- Width: **10.0"**
- Height: **7.5"**

The `create_presentation` response includes `width_inches` and `height_inches` — use these values for positioning calculations.

### Safe margins

Leave at least **0.5"** on all edges for a clean look. Usable area for 16:9:
- X range: 0.5" to 12.83"
- Y range: 0.5" to 7.0"

## Positioning Recipes

Use these coordinates for common slide layouts on 16:9 slides. Adjust proportionally for other sizes.

### Title Slide

```
Title:    x=1.0  y=2.0  w=11.33 h=1.5  font_size=40 bold=true  alignment=center
Subtitle: x=1.0  y=4.0  w=11.33 h=1.0  font_size=24            alignment=center
```

### Section Header

```
Heading:  x=1.0  y=2.5  w=11.33 h=1.5  font_size=36 bold=true  alignment=center
```

### Title + Body Content

```
Title:    x=0.75 y=0.5  w=11.83 h=1.0  font_size=32 bold=true
Body:     x=0.75 y=1.75 w=11.83 h=5.0  font_size=18 bullets=true
```

### Title + Two Columns

```
Title:    x=0.75 y=0.5  w=11.83 h=1.0  font_size=32 bold=true
Left:     x=0.75 y=1.75 w=5.67  h=5.0  font_size=16
Right:    x=6.67 y=1.75 w=5.67  h=5.0  font_size=16
```

### Title + Table or Chart

```
Title:    x=0.75 y=0.5  w=11.83 h=1.0  font_size=32 bold=true
Table:    x=0.75 y=1.75 w=11.83 h=5.0
Chart:    x=1.5  y=1.75 w=10.33 h=5.0
```

### Title + Image

```
Title:    x=0.75 y=0.5  w=11.83 h=1.0  font_size=32 bold=true
Image:    x=2.5  y=1.75 w=8.33  h=5.0
```

## Themes

Five built-in themes. Choose at creation time with the `theme` parameter:

| Theme | Style | Best for |
|-------|-------|----------|
| **modern** (default) | Blue/orange, Calibri | General purpose |
| **corporate** | Navy/steel blue, Georgia headings | Business, formal |
| **creative** | Pink/purple, Avenir Next | Marketing, design |
| **minimal** | Grayscale, Helvetica Neue | Clean, text-heavy |
| **dark** | Purple/teal on dark bg, SF Pro | Technical, modern |

Do not hardcode colors when possible — the theme provides defaults for text, headers, table headers, and more. Let the theme do the styling work.

When using the **dark** theme, remember that the background is dark (`121212`). Set slide backgrounds explicitly if needed, and avoid dark text colors on dark slides.

## Tool Tips

### create_presentation
- The `size` parameter controls aspect ratio, not layout. Default is `"16:9"`. Use `"4:3"` for standard or `"WxH"` for custom (e.g., `"10x7.5"`).

### add_slide
- The `layout` parameter is **metadata only**. It does not auto-generate any content. Every slide starts blank regardless of layout type. You must add all elements manually.

### add_text
- Use `\n` in the `text` parameter for line breaks / multiple paragraphs.
- Set `bullets=true` for bullet-pointed lists.
- Hex colors should omit the `#` prefix: use `"FF0000"` not `"#FF0000"`.
- For centered titles, use `alignment: "center"` and `bold: true`.

### add_image
- Paths can be relative to the workspace or absolute.
- Supported formats: PNG, JPG, GIF, SVG, BMP, TIFF.
- This tool requires user permission (`ask` policy).

### add_shape
- 21 shape types available: `rect`, `round_rect`, `ellipse`, `triangle`, `diamond`, `pentagon`, `hexagon`, `octagon`, `star4`, `star5`, `star6`, `right_arrow`, `left_arrow`, `up_arrow`, `down_arrow`, `heart`, `cloud`, `lightning`, `line`, `parallelogram`, `trapezoid`.
- Shapes can contain text via the `text` parameter — useful for labeled diagrams and flowcharts.

### add_table
- `rows` is a 2D array of strings. The first row is the header by default (`has_header: true`).
- Column widths auto-distribute evenly. Use `column_widths` for custom sizing (must match column count).

### add_chart
- Types: `bar`, `column`, `line`, `pie`, `doughnut`.
- Each series needs a `name` and `values` array. Optionally set a `color` per series.
- `categories` are the x-axis labels.

### save_presentation
- Always call this when done. Nothing is persisted until you save.
- The `.pptx` extension is added automatically if missing.
- This tool requires user permission (`ask` policy).

## Limitations

Be aware of these constraints:

1. **Creation-focused.** There are no tools to update or remove individual elements after adding them. If a slide needs correction, use `delete_slide` to remove it, re-add it with `add_slide`, and rebuild its elements.

2. **No element modification.** You cannot change text, reposition elements, or update properties on existing elements. Plan each slide fully before adding elements.

3. **No slide reordering.** Slides are ordered by insertion. Plan the slide sequence in advance.

4. **read_presentation is limited.** Reading an existing `.pptx` only preserves text elements and slide backgrounds. Images, shapes, tables, and charts are not parsed. Use `read_presentation` primarily for inspecting text content, not for full round-trip editing.

5. **One save at the end.** Build the entire presentation in memory, then call `save_presentation` once. You can save multiple times to different paths if needed.

## Strategy for Multi-Slide Presentations

1. Call `create_presentation` once.
2. Plan all slides before building — decide the content and layout for each.
3. Add slides sequentially: `add_slide` then immediately populate with elements before moving to the next slide. This keeps slide numbers predictable.
4. Call `save_presentation` once at the end.
5. If the user wants changes, use `delete_slide` on the affected slide, `add_slide` to re-add at the same position, rebuild its elements, and save again.

## Correction Strategy

Since elements cannot be modified after creation:

- **Wrong text/formatting on a slide:** `delete_slide` the slide, `add_slide` again, re-add all elements with corrections.
- **Wrong slide order:** Unfortunately slides cannot be reordered. Rebuild the presentation if order matters.
- **Want to inspect what was built:** Use `get_presentation_info` with `include_details: true` to see all elements on all slides.

## Example: Building a 3-Slide Presentation

```
1. create_presentation(title="Q4 Report", theme="corporate")
   → get presentation_id, note width=13.33, height=7.5

2. add_slide(presentation_id, layout="title")
   → slide 1
3. add_text(presentation_id, slide_number=1, text="Q4 2025 Report",
     x=1.0, y=2.0, width=11.33, height=1.5,
     font_size=40, bold=true, alignment="center")
4. add_text(presentation_id, slide_number=1, text="Annual Review",
     x=1.0, y=4.0, width=11.33, height=1.0,
     font_size=24, alignment="center")

5. add_slide(presentation_id, layout="title_content")
   → slide 2
6. add_text(presentation_id, slide_number=2, text="Key Metrics",
     x=0.75, y=0.5, width=11.83, height=1.0,
     font_size=32, bold=true)
7. add_chart(presentation_id, slide_number=2, chart_type="column",
     categories=["Oct", "Nov", "Dec"],
     series=[{name: "Revenue", values: [120, 135, 150]}],
     title="Monthly Revenue",
     x=1.5, y=1.75, width=10.33, height=5.0)

8. add_slide(presentation_id, layout="blank")
   → slide 3
9. add_text(presentation_id, slide_number=3, text="Thank You",
     x=1.0, y=2.5, width=11.33, height=1.5,
     font_size=36, bold=true, alignment="center")

10. save_presentation(presentation_id, path="Q4_Report.pptx")
```
