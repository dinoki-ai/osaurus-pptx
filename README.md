# osaurus-pptx

An [Osaurus](https://osaurus.ai) plugin for creating, reading, and modifying PowerPoint (.pptx) presentations. Supports text, images, shapes, tables, charts, themes, backgrounds, and more — with no external dependencies.

## Tools

| Tool                    | Description                                                                                         |
| ----------------------- | --------------------------------------------------------------------------------------------------- |
| `create_presentation`   | Create a new presentation with a title, layout (16:9, 4:3, custom), and theme                       |
| `add_slide`             | Add a slide with a layout type (blank, title, title_content, section_header, etc.)                  |
| `add_text`              | Add a text box with rich formatting — font, size, color, bold, italic, alignment, bullets, rotation |
| `add_image`             | Add an image from a file (PNG, JPG, GIF, SVG, BMP, TIFF)                                            |
| `add_shape`             | Add a geometric shape (21 types including rect, ellipse, arrows, stars, heart, cloud)               |
| `add_table`             | Add a data table with header styling, alternating row colors, and cell merging                      |
| `add_chart`             | Add a chart (bar, column, line, pie, doughnut) with series data                                     |
| `set_slide_background`  | Set a solid color or gradient background on a slide                                                 |
| `delete_slide`          | Remove a slide by number                                                                            |
| `read_presentation`     | Read an existing .pptx file into memory                                                             |
| `get_presentation_info` | Get metadata and content summary, with optional detailed element info                               |
| `save_presentation`     | Save a presentation as a .pptx file                                                                 |

## Themes

Five built-in themes: **Modern** (default), **Corporate**, **Creative**, **Minimal**, and **Dark**. Each theme controls colors, fonts, and styling across all elements.

## Development

### Build

```bash
swift build -c release
```

### Test

```bash
swift test
```

### Install locally

```bash
osaurus manifest extract .build/release/libosaurus-pptx.dylib
osaurus tools package osaurus.pptx 0.1.0
osaurus tools install ./osaurus.pptx-0.1.0.zip
```

## Publishing

This project includes a GitHub Actions workflow (`.github/workflows/release.yml`) that automatically builds and releases the plugin when you push a version tag.

```bash
git tag v0.1.0
git push origin v0.1.0
```

## License

MIT
