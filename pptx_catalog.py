import argparse
import io
import sys
import textwrap
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# Layout constants in points (1 inch = 72 points)
POS_X_PT = [0.3 * 72, 3.5 * 72, 6.7 * 72, 9.9 * 72]
POS_Y_PT = [0.2 * 72, 2.0 * 72, 3.8 * 72, 5.6 * 72]
CELL_WIDTH_PT = 3.1 * 72
CELL_HEIGHT_PT = 1.74 * 72

# Standard widescreen canvas in points (13.333 in x 7.5 in)
CANVAS_WIDTH_PT = 13.333 * 72
CANVAS_HEIGHT_PT = 7.5 * 72


def pt_to_px(points: float, dpi: int) -> int:
    return round(points * dpi / 72.0)


def load_font(size: int) -> ImageFont.ImageFont:
    candidates = [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for path in candidates:
        if Path(path).is_file():
            try:
                return ImageFont.truetype(path, size=size)
            except OSError:
                pass
    return ImageFont.load_default()


def extract_first_picture(slide) -> Image.Image | None:
    """Return the first embedded raster picture on a slide, if present."""
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                return Image.open(io.BytesIO(shape.image.blob)).convert("RGB")
            except Exception:
                continue
    return None


def extract_slide_text(slide, max_chars: int = 180) -> str:
    """Collect readable text from slide shapes as a fallback preview."""
    chunks: list[str] = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            txt = " ".join(shape.text_frame.text.split())
            if txt:
                chunks.append(txt)
    combined = " — ".join(chunks).strip()
    if len(combined) > max_chars:
        return combined[: max_chars - 3].rstrip() + "..."
    return combined or "(No text extracted)"


def render_slide_card(
    slide,
    slide_number: int,
    cell_w: int,
    cell_h: int,
    title_font: ImageFont.ImageFont,
    body_font: ImageFont.ImageFont,
) -> Image.Image:
    """Build a cell-sized preview using an embedded image if available, else a text card."""
    card = Image.new("RGB", (cell_w, cell_h), color=(248, 250, 252))
    draw = ImageDraw.Draw(card)

    # Subtle border
    draw.rectangle([0, 0, cell_w - 1, cell_h - 1], outline=(203, 213, 225), width=2)

    picture = extract_first_picture(slide)
    if picture is not None:
        inner_pad = 6
        avail_w = max(1, cell_w - 2 * inner_pad)
        avail_h = max(1, cell_h - 2 * inner_pad)
        pic_copy = picture.copy()
        pic_copy.thumbnail((avail_w, avail_h), Image.Resampling.LANCZOS)
        x = inner_pad + (avail_w - pic_copy.width) // 2
        y = inner_pad + (avail_h - pic_copy.height) // 2
        card.paste(pic_copy, (x, y))

        # Small slide badge in top-left corner
        badge_text = f"#{slide_number}"
        draw.rectangle([6, 6, 44, 26], fill=(15, 23, 42))
        draw.text((12, 9), badge_text, fill=(255, 255, 255), font=body_font)
        return card

    # Text fallback card
    pad_x = max(8, cell_w // 22)
    pad_y = max(6, cell_h // 22)
    draw.text((pad_x, pad_y), f"Slide {slide_number}", fill=(15, 23, 42), font=title_font)

    body_text = extract_slide_text(slide)
    wrapped = textwrap.fill(body_text, width=36)
    draw.text(
        (pad_x, pad_y + max(18, cell_h // 6)),
        wrapped,
        fill=(51, 65, 85),
        font=body_font,
    )
    return card


def build_catalog_pages(prs: Presentation, dpi: int) -> list[Image.Image]:
    slides = list(prs.slides)
    if not slides:
        raise ValueError("The presentation contains no slides.")

    canvas_w = pt_to_px(CANVAS_WIDTH_PT, dpi)
    canvas_h = pt_to_px(CANVAS_HEIGHT_PT, dpi)
    pos_x_px = [pt_to_px(x, dpi) for x in POS_X_PT]
    pos_y_px = [pt_to_px(y, dpi) for y in POS_Y_PT]
    cell_w_px = pt_to_px(CELL_WIDTH_PT, dpi)
    cell_h_px = pt_to_px(CELL_HEIGHT_PT, dpi)

    title_font = load_font(max(12, cell_h_px // 11))
    body_font = load_font(max(10, cell_h_px // 14))

    pages: list[Image.Image] = []
    current_page: Image.Image | None = None

    for idx, slide in enumerate(slides):
        slot_in_page = idx % 16
        if slot_in_page == 0:
            current_page = Image.new("RGB", (canvas_w, canvas_h), color=(255, 255, 255))
            pages.append(current_page)

        col_idx = slot_in_page % 4
        row_idx = slot_in_page // 4

        card = render_slide_card(
            slide=slide,
            slide_number=idx + 1,
            cell_w=cell_w_px,
            cell_h=cell_h_px,
            title_font=title_font,
            body_font=body_font,
        )
        assert current_page is not None
        current_page.paste(card, (pos_x_px[col_idx], pos_y_px[row_idx]))

    return pages


def create_catalog(source_file_path: str, dpi: int = 150) -> list[Path]:
    source_path = Path(source_file_path).resolve()
    if not source_path.is_file():
        raise FileNotFoundError(f"PowerPoint file not found: {source_path}")

    prs = Presentation(str(source_path))
    pages = build_catalog_pages(prs, dpi=dpi)

    saved_paths: list[Path] = []
    for page_num, page_img in enumerate(pages, start=1):
        if page_num == 1:
            out_path = source_path.with_suffix(".png")
        else:
            out_path = source_path.with_name(f"{source_path.stem}_page{page_num}.png")
        page_img.save(out_path, format="PNG")
        saved_paths.append(out_path)

    return saved_paths


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Create a 4x4 slide catalog PNG from a .pptx file using python-pptx and Pillow."
    )
    parser.add_argument("pptx_file", help="Path to the input .pptx file")
    parser.add_argument(
        "--dpi",
        type=int,
        default=150,
        help="Output resolution in DPI (default: 150)",
    )
    args = parser.parse_args()

    try:
        output_files = create_catalog(args.pptx_file, dpi=args.dpi)
        for path in output_files:
            print(f"Saved catalog image to: {path}")
        return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
if __name__ == "__main__":
    sys.exit(main())
