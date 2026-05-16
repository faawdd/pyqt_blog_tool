from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Tuple

import html2text
import mammoth


def html_to_markdown(html_text: str) -> str:
    markdown_converter = html2text.HTML2Text()
    markdown_converter.body_width = 0
    markdown_converter.ignore_links = False
    markdown_converter.ignore_images = False
    markdown_converter.ignore_emphasis = False

    markdown_text = markdown_converter.handle(html_text).strip()
    if markdown_text:
        markdown_text += "\n"
    return markdown_text


def _extension_from_content_type(content_type: str | None) -> str:
    mapping = {
        "image/png": "png",
        "image/jpeg": "jpg",
        "image/jpg": "jpg",
        "image/gif": "gif",
        "image/webp": "webp",
        "image/bmp": "bmp",
        "image/tiff": "tiff",
        "image/svg+xml": "svg",
    }
    if not content_type:
        return "png"
    return mapping.get(content_type.lower(), "png")


def _to_workspace_relative_posix(path: Path) -> str:
    try:
        relative_path = path.resolve().relative_to(Path.cwd().resolve())
        return relative_path.as_posix()
    except ValueError:
        return path.resolve().as_posix()


def convert_docx_to_html_and_markdown(docx_path: str, output_img_dir: str) -> Tuple[str, str]:
    """Convert a .docx file to HTML and Markdown, extracting images to output_img_dir."""
    docx_file = Path(docx_path)
    image_dir = Path(output_img_dir)
    image_dir.mkdir(parents=True, exist_ok=True)

    timestamp_base = datetime.now().strftime("%Y%m%d_%H%M")
    image_counter = 0

    def handle_image(image: mammoth.images.Image) -> dict:
        nonlocal image_counter
        image_counter += 1
        extension = _extension_from_content_type(getattr(image, "content_type", None))
        filename = f"img_{timestamp_base}_{image_counter:02d}.{extension}"
        output_path = image_dir / filename

        with image.open() as image_bytes:
            output_path.write_bytes(image_bytes.read())

        return {"src": _to_workspace_relative_posix(output_path)}

    with docx_file.open("rb") as file_obj:
        result = mammoth.convert_to_html(
            file_obj,
            convert_image=mammoth.images.inline(handle_image),
        )

    html_text = result.value

    markdown_text = html_to_markdown(html_text)

    return html_text, markdown_text


def convert_docx_to_markdown(docx_path: str, output_img_dir: str) -> str:
    """Convert a .docx file to Markdown text.

    Args:
        docx_path: Path to the .docx file.
        output_img_dir: Directory used for extracted images.

    Returns:
        markdown_text: Converted Markdown text.
    """
    _, markdown_text = convert_docx_to_html_and_markdown(docx_path, output_img_dir)
    return markdown_text
