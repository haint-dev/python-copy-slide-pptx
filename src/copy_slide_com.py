#!/usr/bin/env python3
"""
Copy slides from source PPTX files into a template-based presentation
using Microsoft PowerPoint COM Automation (Windows only).

Requirements:
    - Windows OS with Microsoft PowerPoint installed
    - pip install pywin32

Usage:
    python copy_slide_com.py
"""

import win32com.client
import os
import sys


def copy_slides_to_template(template_path, slide_selections, output_path):
    """
    Copy selected slides from source files and apply template design.

    PowerPoint handles all formatting: fonts, colors, backgrounds are
    automatically remapped to match the template's theme.

    Args:
        template_path: Path to the template PPTX file
        slide_selections: List of tuples (source_path, slide_indices)
            - source_path: Path to a source PPTX file
            - slide_indices: List of 1-based slide indices to copy
        output_path: Path for the output PPTX file

    Returns:
        Path to the saved output file
    """
    template_path = os.path.abspath(template_path)
    output_path = os.path.abspath(output_path)

    app = win32com.client.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = False

    try:
        # Open template as destination - theme/layouts are already in place
        import shutil
        shutil.copy2(template_path, output_path)
        dst_prs = app.Presentations.Open(output_path)

        # Remove all existing slides from template
        while dst_prs.Slides.Count > 0:
            dst_prs.Slides(1).Delete()

        print(f"Template: {os.path.basename(template_path)} (cleared)")

        slide_count = 0

        for src_path, indices in slide_selections:
            src_path = os.path.abspath(src_path)
            print(f"Source: {os.path.basename(src_path)}")

            src_prs = app.Presentations.Open(src_path, WithWindow=False)
            total = src_prs.Slides.Count

            for idx in indices:
                if idx < 1 or idx > total:
                    print(f"  [SKIP] Slide {idx} out of range (1-{total})")
                    continue

                src_prs.Slides(idx).Copy()
                # Paste and keep destination theme (ppPasteDefault=2)
                dst_prs.Slides.Paste()
                pasted_slide = dst_prs.Slides(dst_prs.Slides.Count)

                # Apply template's first layout to the pasted slide
                pasted_slide.Layout = dst_prs.Slides(dst_prs.Slides.Count).CustomLayout
                # Override: use template's slide layout
                if dst_prs.SlideMaster.CustomLayouts.Count > 0:
                    pasted_slide.CustomLayout = dst_prs.SlideMaster.CustomLayouts(1)

                slide_count += 1
                print(f"  Copied slide {idx} -> destination slide {slide_count}")

            src_prs.Close()

        # Save
        dst_prs.Save()
        print(f"\nSaved {slide_count} slides to: {output_path}")

        dst_prs.Close()

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        raise
    finally:
        app.Quit()

    return output_path


# ---------------------------------------------------------------------------
# Helper functions for building slide index lists (1-based for PowerPoint)
# ---------------------------------------------------------------------------

def first_n(pptx_path, n):
    """Return 1-based indices of the first N slides."""
    app = win32com.client.Dispatch("PowerPoint.Application")
    prs = app.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
    total = prs.Slides.Count
    prs.Close()
    app.Quit()
    return list(range(1, min(n, total) + 1))


def last_n(pptx_path, n):
    """Return 1-based indices of the last N slides."""
    app = win32com.client.Dispatch("PowerPoint.Application")
    prs = app.Presentations.Open(os.path.abspath(pptx_path), WithWindow=False)
    total = prs.Slides.Count
    prs.Close()
    app.Quit()
    return list(range(max(1, total - n + 1), total + 1))


def slide_range(start, end):
    """Return 1-based indices from start to end (inclusive)."""
    return list(range(start, end + 1))


# ---------------------------------------------------------------------------
# Example / Demo
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(BASE_DIR, 'template.pptx')
    upload_path = os.path.join(BASE_DIR, 'uploaded.pptx')
    output_path = os.path.join(BASE_DIR, 'output.pptx')

    slide_selections = [
        (upload_path, slide_range(1, 6)),       # Slides 1-6
        (upload_path, [24, 25, 26, 28]),        # Specific slides
    ]

    copy_slides_to_template(
        template_path=template_path,
        slide_selections=slide_selections,
        output_path=output_path,
    )
