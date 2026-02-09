#!/usr/bin/env python3
"""
Copy slides from multiple source PPTX files into a template-based presentation.

Usage:
    Flexible API to:
    1. Create a new presentation from a template (keeping theme/layouts, removing slides)
    2. Copy selected slides from any number of source files
    3. Apply the template's layout/background to copied slides
"""

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from copy import deepcopy
import os
import re

# Counter for generating unique media filenames
_media_counter = 0


def create_from_template(template_path):
    """
    Create a new empty presentation that inherits all themes/layouts from the template.
    All existing slides are removed.

    Args:
        template_path: Path to the template PPTX file

    Returns:
        A Presentation object with no slides but all template themes/layouts intact
    """
    prs = Presentation(template_path)

    # Remove all existing slides (keep themes/layouts/masters)
    sldIdLst = prs.element.find(qn('p:sldIdLst'))
    for sldId in list(sldIdLst):
        rId = sldId.get(qn('r:id'))
        prs.part.drop_rel(rId)
        sldIdLst.remove(sldId)

    return prs


def _next_media_partname(content_type):
    """Generate a unique partname for a media file."""
    global _media_counter
    _media_counter += 1
    ext_map = {
        'image/png': '.png',
        'image/jpeg': '.jpg',
        'image/gif': '.gif',
        'image/bmp': '.bmp',
        'image/tiff': '.tiff',
        'image/svg+xml': '.svg',
        'image/x-emf': '.emf',
        'image/x-wmf': '.wmf',
    }
    ext = ext_map.get(content_type, '.bin')
    return PackURI(f'/ppt/media/copied_image{_media_counter}{ext}')


def _copy_relationships(src_slide, dst_slide):
    """
    Copy all non-layout, non-notes relationships from source slide to destination.
    Creates new Part copies for media to avoid duplicate partname conflicts.
    Returns a mapping of old rId -> new rId.
    """
    rid_map = {}

    for rel in src_slide.part.rels.values():
        # Skip layout and notes relationships (destination has its own)
        if 'slideLayout' in rel.reltype or 'notesSlide' in rel.reltype:
            continue

        if rel.is_external:
            # External relationships (hyperlinks, etc.)
            new_rid = dst_slide.part.rels.get_or_add_ext_rel(
                rel.reltype, rel.target_ref
            )
            rid_map[rel.rId] = new_rid
        else:
            # Internal relationships (images, charts, embedded objects, etc.)
            # Create a fresh Part copy with unique name to avoid ZIP duplicate warnings
            src_part = rel.target_part
            if 'image' in rel.reltype:
                new_partname = _next_media_partname(src_part.content_type)
                new_part = Part(
                    new_partname, src_part.content_type,
                    dst_slide.part.package, src_part.blob
                )
                new_rid = dst_slide.part.relate_to(new_part, rel.reltype)
            else:
                new_rid = dst_slide.part.relate_to(src_part, rel.reltype)
            rid_map[rel.rId] = new_rid

    return rid_map


def _update_rids_in_tree(tree, rid_map):
    """Replace old rId references with new ones throughout an XML tree.
    Uses single-pass replacement to avoid cascading (e.g. rId4->rId5 then rId5->rId4).
    """
    for elem in tree.iter():
        for attr_name in list(elem.attrib.keys()):
            val = elem.get(attr_name)
            if val in rid_map and rid_map[val] != val:
                elem.set(attr_name, rid_map[val])


def copy_slide(src_prs, src_slide_index, dst_prs, layout_index=0, apply_template_bg=True):
    """
    Copy a single slide from source presentation to destination.

    Args:
        src_prs: Source Presentation object
        src_slide_index: 0-based index of the slide to copy
        dst_prs: Destination Presentation object
        layout_index: Index of the slide layout in destination to apply (0-based)
        apply_template_bg: If True, use template background; if False, copy source background

    Returns:
        The newly created slide in the destination presentation
    """
    src_slide = list(src_prs.slides)[src_slide_index]

    # Create new slide with template layout
    dst_layout = dst_prs.slide_layouts[layout_index]
    dst_slide = dst_prs.slides.add_slide(dst_layout)

    # Step 1: Copy relationships (images, charts, etc.) and build rId mapping
    rid_map = _copy_relationships(src_slide, dst_slide)

    # Step 2: Deep copy the shape tree from source
    new_spTree = deepcopy(src_slide.shapes._spTree)

    # Step 3: Update rId references in the copied shapes
    _update_rids_in_tree(new_spTree, rid_map)

    # Step 4: Replace destination's shape tree with the copied one
    dst_spTree = dst_slide.shapes._spTree
    dst_spTree.getparent().replace(dst_spTree, new_spTree)

    # Step 5: Handle background
    if not apply_template_bg:
        # Copy source slide's background (if it has one)
        src_cSld = src_slide.element.find(qn('p:cSld'))
        src_bg = src_cSld.find(qn('p:bg')) if src_cSld is not None else None
        if src_bg is not None:
            dst_cSld = dst_slide.element.find(qn('p:cSld'))
            new_bg = deepcopy(src_bg)
            # Insert background before spTree
            dst_spTree_elem = dst_cSld.find(qn('p:spTree'))
            dst_cSld.insert(list(dst_cSld).index(dst_spTree_elem), new_bg)

    return dst_slide


def copy_slides_to_template(template_path, slide_selections, output_path,
                            layout_index=0, apply_template_bg=True):
    """
    Main function: create presentation from template and copy selected slides.

    Args:
        template_path: Path to the template PPTX file
        slide_selections: List of tuples (source_path, slide_indices)
            - source_path: Path to a source PPTX file
            - slide_indices: List of 0-based slide indices to copy
        output_path: Path for the output PPTX file
        layout_index: Template layout index to apply (0-based)
        apply_template_bg: If True, template background is applied to all slides

    Returns:
        Path to the saved output file
    """
    print(f"Creating presentation from template: {os.path.basename(template_path)}")
    dst_prs = create_from_template(template_path)
    print(f"  Available layouts: {[l.name for l in dst_prs.slide_layouts]}")
    print(f"  Using layout [{layout_index}]: \"{dst_prs.slide_layouts[layout_index].name}\"")
    print()

    slide_count = 0
    for src_path, indices in slide_selections:
        src_prs = Presentation(src_path)
        total = len(list(src_prs.slides))
        print(f"Source: {os.path.basename(src_path)} ({total} slides)")

        for idx in indices:
            if idx < 0 or idx >= total:
                print(f"  [SKIP] Slide {idx} out of range (0-{total-1})")
                continue

            copy_slide(src_prs, idx, dst_prs, layout_index, apply_template_bg)
            slide_count += 1
            print(f"  Copied slide {idx} -> destination slide {slide_count}")

    dst_prs.save(output_path)
    print(f"\nSaved {slide_count} slides to: {output_path}")
    return output_path


# ---------------------------------------------------------------------------
# Helper functions for building slide index lists
# ---------------------------------------------------------------------------

def first_n(pptx_path, n):
    """Return indices of the first N slides."""
    total = len(list(Presentation(pptx_path).slides))
    return list(range(min(n, total)))


def last_n(pptx_path, n):
    """Return indices of the last N slides."""
    total = len(list(Presentation(pptx_path).slides))
    return list(range(max(0, total - n), total))


def slide_range(start, end):
    """Return indices from start to end (inclusive, 0-based)."""
    return list(range(start, end + 1))


# ---------------------------------------------------------------------------
# Example / Demo
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    template_path = os.path.join(BASE_DIR, 'template.pptx')
    upload_path = os.path.join(BASE_DIR, 'file_upload.pptx')
    output_path = os.path.join(BASE_DIR, 'output.pptx')

    # Task: copy first 5 slides + slide 6 + last 2 slides from file_upload
    slide_selections = [
        (upload_path, first_n(upload_path, 5)),   # Slides 0,1,2,3,4
        (upload_path, [5]),                        # Slide 6 (index 5, 0-based)
        #(upload_path, last_n(upload_path, 2)),     # Slides 9,10
    ]

    copy_slides_to_template(
        template_path=template_path,
        slide_selections=slide_selections,
        output_path=output_path,
        layout_index=0,           # "Cover slide layout"
        apply_template_bg=True,   # Apply template background
    )
