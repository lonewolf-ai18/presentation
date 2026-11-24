import argparse
from pathlib import Path
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT


def clear_explicit_background(slide_or_layout) -> None:
    """Remove any explicit <p:bg> element so the object inherits from its parent."""
    bg_el = slide_or_layout._element.cSld.find(qn("p:bg"))
    if bg_el is not None:
        slide_or_layout._element.cSld.remove(bg_el)


def clear_all_elements(slide_layout) -> None:
    """Remove all shapes/placeholders/elements from a slide layout."""
    for shape in list(slide_layout.shapes):
        sp = shape._element
        parent = sp.getparent()
        if parent is not None:
            parent.remove(sp)


def find_or_prepare_blank_layout(slide_master):
    """Find or pick a slide layout, name it 'blank', and strip all elements from it."""
    blank_layout = None

    # Look for an existing layout already named 'blank' (case-insensitive)
    for layout in slide_master.slide_layouts:
        if layout.name.strip().lower() == "blank":
            blank_layout = layout
            break

    # Otherwise, pick the last available layout
    if blank_layout is None:
        if len(slide_master.slide_layouts) == 0:
            raise RuntimeError("Slide master has no slide layouts available.")
        blank_layout = slide_master.slide_layouts[-1]

    blank_layout.name = "blank"
    clear_all_elements(blank_layout)
    clear_explicit_background(blank_layout)
    return blank_layout


def set_slide_layout(slide, target_layout):
    """Reassign an existing slide's RT.SLIDE_LAYOUT relationship to target_layout."""
    part = slide.part
    # Remove existing relationship(s) of type SLIDE_LAYOUT
    for rel in list(part.rels.values()):
        if rel.reltype == RT.SLIDE_LAYOUT:
            part.drop_rel(rel.rId)
    # Create new relationship pointing to target_layout's part
    part.relate_to(target_layout.part, RT.SLIDE_LAYOUT)


def update_presentation(pptx_path: Path) -> None:
    if not pptx_path.is_file():
        raise FileNotFoundError(f"Presentation not found: {pptx_path}")

    prs = Presentation(str(pptx_path))

    for master in prs.slide_masters:
        # 1. Change master slide background to black
        bg = master.background
        bg.fill.solid()
        bg.fill.fore_color.rgb = RGBColor(0, 0, 0)

        # 2. Prepare a slide layout named 'blank' with zero elements
        blank_layout = find_or_prepare_blank_layout(master)

        # 3. Change all slides using this master's layouts to 'blank'
        master_layout_parts = {layout.part for layout in master.slide_layouts}
        for slide in prs.slides:
            if slide.slide_layout.part in master_layout_parts:
                set_slide_layout(slide, blank_layout)
                # Remove per-slide background override so the black master background shows
                clear_explicit_background(slide)

        # 4. Delete all slide layouts except 'blank'
        for layout in list(master.slide_layouts):
            if layout != blank_layout:
                master.slide_layouts.remove(layout)

    # Overwrite the original file
    prs.save(str(pptx_path))


def main():
    parser = argparse.ArgumentParser(
        description="Overwrite a .pptx file with a black master background, an empty 'blank' layout, and all slides assigned to it."
    )
    parser.add_argument("pptx_file", type=Path, help="Path to .pptx file to overwrite")
    args = parser.parse_args()

    update_presentation(args.pptx_file)
    print(f"Overwritten presentation: {args.pptx_file}")


if __name__ == "__main__":
    main()
