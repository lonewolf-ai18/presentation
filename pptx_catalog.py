import argparse
import os
import shutil
import sys
import tempfile
from pathlib import Path

# Required pywin32: pip install pywin32
import win32com.client as win32


# Layout constants in points (1 inch = 72 points)
POS_X = [0.3 * 72, 3.5 * 72, 6.7 * 72, 9.9 * 72]
POS_Y = [0.2 * 72, 2.0 * 72, 3.8 * 72, 5.6 * 72]
IMG_WIDTH = 3.1 * 72
IMG_HEIGHT = 1.74 * 72

PP_LAYOUT_BLANK = 12  # ppLayoutBlank
MSO_FALSE = 0
MSO_TRUE = -1


def create_one_page_catalog(source_file_path: str) -> Path:
    source_path = Path(source_file_path).resolve()
    if not source_path.is_file():
        raise FileNotFoundError(f"PowerPoint file not found: {source_path}")

    output_png_path = source_path.with_suffix(".png")
    temp_dir = Path(tempfile.mkdtemp(prefix="pptx_catalog_"))

    ppt_app = None
    ppt_pres = None
    catalog_pres = None

    try:
        ppt_app = win32.Dispatch("PowerPoint.Application")
        # Note: PowerPoint sometimes restricts hiding the window in COM automation;
        # WithWindow:=True on Presentations.Add mirrors your VBA code.
        ppt_pres = ppt_app.Presentations.Open(str(source_path), WithWindow=False)
        slides_count = ppt_pres.Slides.Count

        if slides_count == 0:
            raise ValueError("The presentation contains no slides.")

        # 1. Export each slide to a temporary PNG
        exported_images: list[Path] = []
        for idx in range(1, slides_count + 1):
            img_path = temp_dir / f"Slide_{idx}.png"
            ppt_pres.Slides(idx).Export(str(img_path), "PNG")
            exported_images.append(img_path)

        ppt_pres.Close()
        ppt_pres = None

        # 2. Create a new presentation for the 4x4 grid catalog
        catalog_pres = ppt_app.Presentations.Add(WithWindow=True)
        current_slide = catalog_pres.Slides.Add(1, PP_LAYOUT_BLANK)

        for img_idx, img_path in enumerate(exported_images):
            col_idx = img_idx % 4
            row_idx = (img_idx // 4) % 4

            if img_idx > 0 and img_idx % 16 == 0:
                next_slide_num = catalog_pres.Slides.Count + 1
                current_slide = catalog_pres.Slides.Add(next_slide_num, PP_LAYOUT_BLANK)

            current_slide.Shapes.AddPicture(
                FileName=str(img_path),
                LinkToFile=MSO_FALSE,
                SaveWithDocument=MSO_TRUE,
                Left=float(POS_X[col_idx]),
                Top=float(POS_Y[row_idx]),
                Width=float(IMG_WIDTH),
                Height=float(IMG_HEIGHT),
            )

        # 3. Export the catalog slide as <same_filename>.png
        catalog_pres.Slides(1).Export(str(output_png_path), "PNG")
        return output_png_path

    finally:
        if catalog_pres is not None:
            try:
                catalog_pres.Close()
            except Exception:
                pass
        if ppt_pres is not None:
            try:
                ppt_pres.Close()
            except Exception:
                pass
        if ppt_app is not None:
            try:
                ppt_app.Quit()
            except Exception:
                pass
        shutil.rmtree(temp_dir, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a 4x4 slide catalog PNG from a PowerPoint (.pptx) file."
    )
    parser.add_argument("pptx_file", help="Path to the input .pptx file")
    args = parser.parse_args()

    try:
        output_file = create_one_page_catalog(args.pptx_file)
        print(f"Saved catalog image to: {output_file}")
        return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
