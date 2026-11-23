#!/usr/bin/env python3
"""Compress raster images inside a .pptx file without breaking OOXML."""

from __future__ import annotations

import argparse
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime
from pathlib import Path
from PIL import Image, ImageColor, ImageOps

RASTER_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".webp"}
SKIP_EXTENSIONS = {".svg", ".emf", ".wmf", ".mp4", ".mov", ".avi", ".mp3", ".wav"}
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def resize_if_needed(img: Image.Image, max_dim: int) -> Image.Image:
    """Downscale image proportionally if width or height exceeds max_dim."""
    width, height = img.size
    largest_side = max(width, height)
    if largest_side <= max_dim:
        return img

    scale = max_dim / largest_side
    new_size = (max(1, int(round(width * scale))), max(1, int(round(height * scale))))
    return img.resize(new_size, Image.Resampling.LANCZOS)


def to_rgb_with_background(img: Image.Image, bg_color: tuple[int, int, int]) -> Image.Image:
    """Convert any image mode (including RGBA/palette with transparency) to RGB."""
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        rgba = img.convert("RGBA")
        background = Image.new("RGBA", rgba.size, (*bg_color, 255))
        composited = Image.alpha_composite(background, rgba)
        return composited.convert("RGB")
    return img.convert("RGB")


def next_available_jpg_path(media_dir: Path, source_path: Path) -> Path:
    """Return a non-colliding .jpg path inside media_dir."""
    candidate = source_path.with_suffix(".jpg")
    if candidate == source_path or not candidate.exists():
        return candidate

    idx = 1
    while True:
        alt = media_dir / f"{source_path.stem}_c{idx}.jpg"
        if not alt.exists():
            return alt
        idx += 1


def update_rels_references(extract_dir: Path, old_name: str, new_name: str) -> None:
    """Update media filename references inside all .rels XML files."""
    if old_name == new_name:
        return

    for rels_path in extract_dir.rglob("*.rels"):
        content = rels_path.read_text(encoding="utf-8")
        updated = content.replace(f"/media/{old_name}", f"/media/{new_name}")
        if updated != content:
            rels_path.write_text(updated, encoding="utf-8")


def update_content_types(
    extract_dir: Path,
    old_rel_part: str,
    new_rel_part: str,
    used_extensions: set[str],
) -> None:
    """
    Update [Content_Types].xml so PowerPoint recognizes converted .jpg parts
    and does not complain about missing/mismatched content types.
    """
    ct_path = extract_dir / "[Content_Types].xml"
    if not ct_path.exists():
        return

    ET.register_namespace("", CONTENT_TYPES_NS)
    tree = ET.parse(ct_path)
    root = tree.getroot()

    default_tag = f"{{{CONTENT_TYPES_NS}}}Default"
    override_tag = f"{{{CONTENT_TYPES_NS}}}Override"

    existing_defaults: dict[str, ET.Element] = {}
    for elem in root.findall(default_tag):
        ext_attr = elem.attrib.get("Extension", "").lower()
        if ext_attr:
            existing_defaults[ext_attr] = elem

    # Ensure jpg default content type exists if any .jpg is used
    if "jpg" in used_extensions and "jpg" not in existing_defaults:
        new_default = ET.SubElement(root, default_tag)
        new_default.set("Extension", "jpg")
        new_default.set("ContentType", "image/jpeg")
        existing_defaults["jpg"] = new_default

    # Ensure png default content type exists if any .png is used
    if "png" in used_extensions and "png" not in existing_defaults:
        new_default = ET.SubElement(root, default_tag)
        new_default.set("Extension", "png")
        new_default.set("ContentType", "image/png")
        existing_defaults["png"] = new_default

    # Update any Override pointing to the old part name (e.g. /ppt/media/image1.png)
    old_part_norm = old_rel_part if old_rel_part.startswith("/") else f"/{old_rel_part}"
    new_part_norm = new_rel_part if new_rel_part.startswith("/") else f"/{new_rel_part}"

    for elem in list(root.findall(override_tag)):
        part_name = elem.attrib.get("PartName", "")
        if part_name == old_part_norm:
            if new_part_norm.lower().endswith((".jpg", ".jpeg")):
                elem.set("PartName", new_part_norm)
                elem.set("ContentType", "image/jpeg")
            elif new_part_norm.lower().endswith(".png"):
                elem.set("PartName", new_part_norm)
                elem.set("ContentType", "image/png")
            else:
                root.remove(elem)

    # If a .png was converted to .jpg and there are no .png files left, clean up png Override
    if old_part_norm.lower().endswith(".png") and "png" not in used_extensions:
        for elem in list(root.findall(override_tag)):
            if elem.attrib.get("PartName", "").lower().endswith(".png"):
                root.remove(elem)

    tree.write(ct_path, encoding="utf-8", xml_declaration=True)


def compress_image_file(
    image_path: Path,
    media_dir: Path,
    extract_dir: Path,
    max_dim: int = 2000,
    jpeg_quality: int = 80,
    bg_color: tuple[int, int, int] = (255, 255, 255),
    keep_png: bool = False,
) -> tuple[Path, int, int, str | None, str | None]:
    """Compress a raster image in ppt/media; convert PNG to .jpg unless keep_png=True."""
    original_bytes = image_path.stat().st_size
    ext = image_path.suffix.lower()

    if ext in SKIP_EXTENSIONS or ext not in RASTER_EXTENSIONS:
        return image_path, original_bytes, original_bytes, None, None

    if ext == ".png" and keep_png:
        with Image.open(image_path) as raw_img:
            img = ImageOps.exif_transpose(raw_img)
            resized_img = resize_if_needed(img, max_dim=max_dim)
            was_resized = resized_img.size != img.size

            temp_png = media_dir / f".tmp_{image_path.name}"
            resized_img.save(temp_png, format="PNG", optimize=True)

        new_bytes = temp_png.stat().st_size
        if was_resized or new_bytes < original_bytes:
            temp_png.replace(image_path)
            return image_path, original_bytes, new_bytes, None, None

        temp_png.unlink(missing_ok=True)
        return image_path, original_bytes, original_bytes, None, None

    with Image.open(image_path) as raw_img:
        img = ImageOps.exif_transpose(raw_img)
        img = resize_if_needed(img, max_dim=max_dim)
        rgb_img = to_rgb_with_background(img, bg_color=bg_color)

        jpg_path = next_available_jpg_path(media_dir, image_path)
        rgb_img.save(
            jpg_path,
            format="JPEG",
            quality=jpeg_quality,
            optimize=True,
            progressive=True,
        )

    new_bytes = jpg_path.stat().st_size

    if jpg_path != image_path:
        old_rel = f"/{image_path.relative_to(extract_dir).as_posix()}"
        new_rel = f"/{jpg_path.relative_to(extract_dir).as_posix()}"
        update_rels_references(extract_dir, image_path.name, jpg_path.name)
        if image_path.exists():
            image_path.unlink()
        return jpg_path, original_bytes, new_bytes, old_rel, new_rel

    if new_bytes < original_bytes:
        return jpg_path, original_bytes, new_bytes, None, None

    return image_path, original_bytes, original_bytes, None, None


def repack_zip(source_dir: Path, output_zip: Path) -> None:
    """Create a .pptx (ZIP) archive from all files inside source_dir."""
    with zipfile.ZipFile(output_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for file_path in sorted(source_dir.rglob("*")):
            if file_path.is_file():
                arcname = file_path.relative_to(source_dir).as_posix()
                zf.write(file_path, arcname)


def default_output_path(input_path: Path, now: datetime | None = None) -> Path:
    """Build default output path using _yymm suffix (e.g., _2611 for Nov 2026)."""
    stamp = (now or datetime.now()).strftime("%y%m")
    return input_path.with_name(f"{input_path.stem}_{stamp}{input_path.suffix}")


def compress_pptx(
    input_path: Path,
    output_path: Path | None = None,
    max_dim: int = 2000,
    jpeg_quality: int = 80,
    bg_color: tuple[int, int, int] = (255, 255, 255),
    keep_png: bool = False,
) -> Path:
    """Extract a .pptx archive, compress raster images, fix metadata, and repack."""
    if not input_path.exists() or not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    if input_path.suffix.lower() != ".pptx":
        raise ValueError(f"Expected a .pptx file, got: {input_path.suffix}")

    if output_path is None:
        output_path = default_output_path(input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp_dir:
        extract_dir = Path(tmp_dir) / "extracted"
        with zipfile.ZipFile(input_path, "r") as zf:
            zf.extractall(extract_dir)

        media_dir = extract_dir / "ppt" / "media"
        total_orig = 0
        total_new = 0
        processed_count = 0
        renames: list[tuple[str, str]] = []

        if media_dir.exists() and media_dir.is_dir():
            original_files = [p for p in sorted(media_dir.iterdir()) if p.is_file()]
            for file_path in original_files:
                if not file_path.exists():
                    continue
                _, orig_sz, new_sz, old_rel, new_rel = compress_image_file(
                    image_path=file_path,
                    media_dir=media_dir,
                    extract_dir=extract_dir,
                    max_dim=max_dim,
                    jpeg_quality=jpeg_quality,
                    bg_color=bg_color,
                    keep_png=keep_png,
                )
                total_orig += orig_sz
                total_new += new_sz
                processed_count += 1
                if old_rel and new_rel:
                    renames.append((old_rel, new_rel))

            used_extensions = {
                p.suffix.lower().lstrip(".")
                for p in media_dir.iterdir()
                if p.is_file() and p.suffix
            }
            for old_rel, new_rel in renames:
                update_content_types(extract_dir, old_rel, new_rel, used_extensions)

        temp_out = Path(tmp_dir) / "output.pptx"
        repack_zip(extract_dir, temp_out)
        output_path.write_bytes(temp_out.read_bytes())

    orig_mb = input_path.stat().st_size / (1024 * 1024)
    out_mb = output_path.stat().st_size / (1024 * 1024)
    media_saved_kb = (total_orig - total_new) / 1024

    print(f"Processed media items : {processed_count}")
    print(f"Renamed to .jpg       : {len(renames)}")
    print(f"Media size delta      : {-media_saved_kb:+.1f} KB")
    print(f"Original .pptx size   : {orig_mb:.2f} MB")
    print(f"Compressed .pptx size : {out_mb:.2f} MB")
    print(f"Saved to              : {output_path}")

    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Compress raster images inside a .pptx file (converting PNG to JPG by default)."
    )
    parser.add_argument("pptx_path", type=Path, help="Path to the input .pptx file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Path to the output .pptx file (default: <name>_<yymm>.pptx)",
    )
    parser.add_argument(
        "--keep-png",
        action="store_true",
        help="Do not convert PNG images to JPG (default: convert PNG to JPG)",
    )
    parser.add_argument(
        "--max-dim",
        type=int,
        default=2000,
        help="Maximum width or height in pixels (default: 2000)",
    )
    parser.add_argument(
        "--quality",
        type=int,
        default=80,
        help="JPEG compression quality from 1 to 95 (default: 80)",
    )
    parser.add_argument(
        "--bg-color",
        type=str,
        default="#FFFFFF",
        help="Background color used when flattening transparent PNGs to JPG (default: #FFFFFF)",
    )

    args = parser.parse_args()
    rgb_bg = ImageColor.getrgb(args.bg_color)
    if len(rgb_bg) != 3:
        raise ValueError(f"Invalid background color: {args.bg_color}")

    compress_pptx(
        input_path=args.pptx_path,
        output_path=args.output,
        max_dim=args.max_dim,
        jpeg_quality=args.quality,
        bg_color=rgb_bg,
        keep_png=args.keep_png,
    )


if __name__ == "__main__":
    main()
