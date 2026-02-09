# Copy PPTX Slides

Copy slides from multiple source PPTX files into a template-based presentation.

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Usage

### Command line

```bash
venv/bin/python copy_slide.py
```

### As a library

```python
from copy_slide import copy_slides_to_template, first_n, last_n, slide_range

copy_slides_to_template(
    template_path='template.pptx',
    slide_selections=[
        ('file_upload.pptx', first_n('file_upload.pptx', 5)),  # First 5 slides
        ('file_upload.pptx', [5]),                               # Slide 6 (0-based index)
        ('file_upload.pptx', last_n('file_upload.pptx', 2)),    # Last 2 slides
        ('other_file.pptx', slide_range(2, 7)),                  # Slides 3-8
    ],
    output_path='output.pptx',
    layout_index=0,          # Template layout to apply
    apply_template_bg=True,  # True = template background, False = keep source background
)
```

### Helper functions

| Function | Description | Example |
|---|---|---|
| `first_n(path, n)` | First N slides | `first_n('file.pptx', 5)` -> `[0,1,2,3,4]` |
| `last_n(path, n)` | Last N slides | `last_n('file.pptx', 2)` -> `[9,10]` |
| `slide_range(start, end)` | Range (inclusive, 0-based) | `slide_range(2, 5)` -> `[2,3,4,5]` |
| `[i, j, ...]` | Specific slides | `[0, 3, 7]` |
