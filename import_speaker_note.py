import argparse
import re
import sys
from pathlib import Path

from docx import Document
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt


SLIDE_HEADER_RE = re.compile(r"^Slide\s+(\d+)\s*:", re.IGNORECASE)
BLANK_LINES_BEFORE_ORIGINAL = 2


def _clean_text(text: str) -> str:
    return text.replace("\r\n", "\n").replace("\r", "").replace("\n", "")


def _copy_run_format(src_run, dst_run) -> None:
    """Copy basic character formatting from a docx run to a pptx run."""
    font = dst_run.font

    if src_run.bold is not None:
        font.bold = src_run.bold
    if src_run.italic is not None:
        font.italic = src_run.italic
    if src_run.underline is not None:
        font.underline = bool(src_run.underline)

    if src_run.font.name:
        font.name = src_run.font.name

    if src_run.font.size is not None:
        # python-docx Length is in EMUs; Pt(points) converts cleanly
        font.size = Pt(src_run.font.size.pt)

    color = src_run.font.color
    if color is not None and color.rgb is not None:
        rgb_bytes = bytes(color.rgb)
        font.color.rgb = RGBColor(rgb_bytes[0], rgb_bytes[1], rgb_bytes[2])


def _copy_pptx_run_format(src_run, dst_run) -> None:
    """Copy formatting from an existing pptx run to a new pptx run."""
    src_font = src_run.font
    dst_font = dst_run.font

    if src_font.bold is not None:
        dst_font.bold = src_font.bold
    if src_font.italic is not None:
        dst_font.italic = src_font.italic
    if src_font.underline is not None:
        dst_font.underline = src_font.underline
    if src_font.name:
        dst_font.name = src_font.name
    if src_font.size is not None:
        dst_font.size = src_font.size
    try:
        if src_font.color.rgb is not None:
            dst_font.color.rgb = src_font.color.rgb
    except AttributeError:
        pass


def _para_is_blank(para) -> bool:
    return _clean_text(para.text).strip() == ""


def parse_notes_from_docx(docx_path: Path) -> dict[int, list]:
    """
    Scan paragraphs for 'Slide #:' headers and collect subsequent python-docx
    Paragraph objects until the next 'Slide #:' marker.
    """
    doc = Document(str(docx_path))
    notes_by_slide: dict[int, list] = {}
    current_slide_num: int | None = None
    current_paragraphs: list = []

    def flush_current() -> None:
        if current_slide_num is not None:
            cleaned = list(current_paragraphs)
            while cleaned and _para_is_blank(cleaned[0]):
                cleaned.pop(0)
            while cleaned and _para_is_blank(cleaned[-1]):
                cleaned.pop()
            notes_by_slide[current_slide_num] = cleaned

    for para in doc.paragraphs:
        stripped = _clean_text(para.text).strip()
        match = SLIDE_HEADER_RE.match(stripped)
        if match:
            flush_current()
            current_slide_num = int(match.group(1))
            current_paragraphs = []
        elif current_slide_num is not None:
            current_paragraphs.append(para)

    flush_current()
    return notes_by_slide


def capture_existing_notes(slide) -> list[list[tuple[str, object | None]]]:
    """
    Capture existing speaker note paragraphs and runs from a slide so they can
    be re-appended with formatting preserved.
    Returns a list of paragraphs; each paragraph is a list of (text, source_pptx_run).
    """
    if not slide.has_notes_slide:
        return []

    notes_frame = slide.notes_slide.notes_text_frame
    captured: list[list[tuple[str, object | None]]] = []

    for p in notes_frame.paragraphs:
        if p.runs:
            runs_data = [(_clean_text(r.text), r) for r in p.runs]
        else:
            txt = _clean_text(p.text)
            runs_data = [(txt, None)] if txt else []
        captured.append(runs_data)

    def is_empty_p(p_runs: list[tuple[str, object | None]]) -> bool:
        return "".join(t for t, _ in p_runs).strip() == ""

    while captured and is_empty_p(captured[0]):
        captured.pop(0)
    while captured and is_empty_p(captured[-1]):
        captured.pop()

    return captured


def _populate_pptx_paragraph_from_docx(dst_p, docx_para) -> None:
    if docx_para.runs:
        for src_run in docx_para.runs:
            txt = _clean_text(src_run.text)
            if txt:
                dst_run = dst_p.add_run()
                dst_run.text = txt
                _copy_run_format(src_run, dst_run)
    else:
        txt = _clean_text(docx_para.text)
        if txt:
            dst_p.add_run().text = txt


def _populate_pptx_paragraph_from_existing(dst_p, runs_data: list[tuple[str, object | None]]) -> None:
    for txt, src_run in runs_data:
        if txt:
            dst_run = dst_p.add_run()
            dst_run.text = txt
            if src_run is not None:
                _copy_pptx_run_format(src_run, dst_run)


def write_notes_to_slide(
    slide,
    docx_paragraphs: list,
    append_original: bool = False,
    blank_lines: int = BLANK_LINES_BEFORE_ORIGINAL,
) -> None:
    """
    Write formatted docx paragraphs into a slide's speaker notes.
    If append_original is True (Slide 1), existing notes are appended after
    `blank_lines` empty paragraphs.
    """
    existing_paras = capture_existing_notes(slide) if append_original else []

    # Build a unified sequence of entries:
    # ("docx", docx_para) | ("blank", None) | ("existing", runs_data)
    entries: list[tuple[str, object]] = [("docx", p) for p in docx_paragraphs]

    if existing_paras:
        if entries:
            for _ in range(blank_lines):
                entries.append(("blank", None))
        for p_runs in existing_paras:
            entries.append(("existing", p_runs))

    notes_slide = slide.notes_slide
    text_frame = notes_slide.notes_text_frame
    text_frame.clear()

    if not entries:
        return

    for idx, (kind, payload) in enumerate(entries):
        dst_p = text_frame.paragraphs[0] if idx == 0 else text_frame.add_paragraph()
        if kind == "docx":
            _populate_pptx_paragraph_from_docx(dst_p, payload)
        elif kind == "existing":
            _populate_pptx_paragraph_from_existing(dst_p, payload)
        # ("blank", None) leaves dst_p empty as a blank line


def import_speaker_notes(
    ppt_path_str: str,
    docx_path_str: str | None = None,
    output_path_str: str | None = None,
    blank_lines: int = BLANK_LINES_BEFORE_ORIGINAL,
) -> Path:
    ppt_path = Path(ppt_path_str).resolve()
    if not ppt_path.is_file():
        raise FileNotFoundError(f"PowerPoint file not found: {ppt_path}")

    docx_path = (
        Path(docx_path_str).resolve()
        if docx_path_str
        else ppt_path.with_suffix(".docx")
    )
    if not docx_path.is_file():
        raise FileNotFoundError(
            f"Matching Word (.docx) file not found: {docx_path}"
        )

    output_path = (
        Path(output_path_str).resolve() if output_path_str else ppt_path
    )

    notes_map = parse_notes_from_docx(docx_path)
    prs = Presentation(str(ppt_path))
    slides = list(prs.slides)

    for slide_num, docx_paras in notes_map.items():
        if 1 <= slide_num <= len(slides):
            write_notes_to_slide(
                slides[slide_num - 1],
                docx_paras,
                append_original=(slide_num == 1),
                blank_lines=blank_lines,
            )

    prs.save(str(output_path))
    return output_path


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Import formatted speaker notes from a matching .docx file into a .pptx presentation."
    )
    parser.add_argument("pptx_file", help="Path to the input .pptx file")
    parser.add_argument(
        "--docx",
        default=None,
        help="Optional explicit path to .docx file (defaults to same directory and name as .pptx)",
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Optional output .pptx path (defaults to overwriting the input .pptx)",
    )
    parser.add_argument(
        "--blank-lines",
        type=int,
        default=BLANK_LINES_BEFORE_ORIGINAL,
        help="Number of blank lines before appending Slide 1's original speaker note (default: 2)",
    )
    args = parser.parse_args()

    try:
        saved_path = import_speaker_notes(
            ppt_path_str=args.pptx_file,
            docx_path_str=args.docx,
            output_path_str=args.output,
            blank_lines=max(0, args.blank_lines),
        )
        print(f"Updated presentation saved to: {saved_path}")
        return 0
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
