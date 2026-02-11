# Copy PPTX Slides

Copy slides from multiple source PPTX files into a template-based presentation, automatically applying the template's theme (fonts, colors, backgrounds, layouts).

## Features

- **Template-based output** — creates a new presentation inheriting all themes/layouts from a template
- **Multi-source support** — combine slides from multiple PPTX files in one output
- **Font remapping** — replaces hardcoded fonts with template theme fonts (major/minor)
- **Color remapping** — converts hardcoded colors back to theme color references so they adapt to the destination theme
- **Background control** — apply template background or keep the original source background
- **Placeholder mapping** — optionally map source content into template placeholders

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

**Requirements:** `python-pptx >= 1.0.2`

## Usage

### Command line

```bash
venv/bin/python src/copy_slide.py
```

### As a library

```python
from src.copy_slide import copy_slides_to_template, first_n, last_n, slide_range

copy_slides_to_template(
    template_path='template.pptx',
    slide_selections=[
        ('file_upload.pptx', first_n('file_upload.pptx', 5)),   # First 5 slides
        ('file_upload.pptx', [5]),                                # Slide 6 (0-based index)
        ('file_upload.pptx', last_n('file_upload.pptx', 2)),     # Last 2 slides
        ('other_file.pptx', slide_range(2, 7)),                   # Slides 3-8
    ],
    output_path='output.pptx',
    layout_index=0,             # Template layout to apply (0-based)
    apply_template_bg=True,     # True = template background, False = keep source
    remap_fonts=True,           # Remap fonts to template theme fonts
    remap_colors=True,          # Remap hardcoded colors to theme color refs
    use_placeholders=False,     # Map content into template placeholders
)
```

### Slide selection helpers

| Function | Description | Example |
|---|---|---|
| `first_n(path, n)` | First N slides | `first_n('f.pptx', 5)` → `[0,1,2,3,4]` |
| `last_n(path, n)` | Last N slides | `last_n('f.pptx', 2)` → `[9,10]` |
| `slide_range(start, end)` | Range (inclusive, 0-based) | `slide_range(2, 5)` → `[2,3,4,5]` |
| `[i, j, ...]` | Specific slides | `[0, 3, 7]` |

### API reference

#### `copy_slides_to_template()`

| Parameter | Type | Default | Description |
|---|---|---|---|
| `template_path` | `str` | — | Path to the template PPTX |
| `slide_selections` | `list[tuple]` | — | List of `(source_path, slide_indices)` |
| `output_path` | `str` | — | Output file path |
| `layout_index` | `int` | `0` | Template layout index (0-based) |
| `apply_template_bg` | `bool` | `True` | Apply template background to all slides |
| `remap_fonts` | `bool` | `True` | Replace hardcoded fonts with theme fonts |
| `remap_colors` | `bool` | `True` | Replace hardcoded colors with theme color refs |
| `use_placeholders` | `bool` | `False` | Map content into template placeholders |
