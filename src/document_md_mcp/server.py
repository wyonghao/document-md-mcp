from __future__ import annotations

import hashlib
import json
import mimetypes
import re
import shutil
import stat
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Literal

import fitz
from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from mcp.server.fastmcp import FastMCP


mcp = FastMCP("document-md")

PdfImageStrategy = Literal["embedded", "pages", "both"]


@dataclass
class ConversionPaths:
    source: Path
    output_dir: Path
    markdown: Path
    image_dir: Path
    manifest: Path


def _safe_stem(path: Path) -> str:
    stem = path.stem.strip()
    if not stem:
        return "document"
    return re.sub(r"[^\w.\- ]+", "_", stem, flags=re.UNICODE)


def _paths_for(input_path: str, output_dir: str | None) -> ConversionPaths:
    source = Path(input_path).expanduser().resolve()
    if not source.exists():
        raise FileNotFoundError(f"Input file does not exist: {source}")
    if not source.is_file():
        raise ValueError(f"Input path is not a file: {source}")

    target_dir = Path(output_dir).expanduser().resolve() if output_dir else source.parent
    stem = _safe_stem(source)
    return ConversionPaths(
        source=source,
        output_dir=target_dir,
        markdown=target_dir / f"{stem}.md",
        image_dir=target_dir / f"{stem}md",
        manifest=target_dir / f"{stem}.manifest.json",
    )


def _prepare_output(paths: ConversionPaths, overwrite: bool) -> None:
    paths.output_dir.mkdir(parents=True, exist_ok=True)

    existing = [paths.markdown, paths.manifest, paths.image_dir]
    conflicts = [str(path) for path in existing if path.exists()]
    if conflicts and not overwrite:
        raise FileExistsError(
            "Output already exists. Pass overwrite=true to replace: "
            + ", ".join(conflicts)
        )

    if paths.image_dir.exists():
        _clear_directory_contents(paths.image_dir)
    else:
        paths.image_dir.mkdir(parents=True, exist_ok=True)


def _clear_directory_contents(path: Path) -> None:
    for child in path.iterdir():
        if child.is_dir():
            _remove_tree(child)
        else:
            child.unlink(missing_ok=True)


def _remove_tree(path: Path) -> None:
    def make_writable(function, target, excinfo) -> None:
        try:
            Path(target).chmod(stat.S_IWRITE)
            function(target)
        except Exception:
            raise excinfo

    last_error: Exception | None = None
    for _ in range(5):
        try:
            shutil.rmtree(path, onexc=make_writable)
            return
        except PermissionError as error:
            last_error = error
            time.sleep(0.2)
    if last_error is not None:
        raise last_error


def _write_text(path: Path, text: str) -> None:
    path.write_text(text.rstrip() + "\n", encoding="utf-8")


def _write_manifest(paths: ConversionPaths, manifest: dict[str, Any]) -> None:
    paths.manifest.write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _relative_markdown_link(target: Path, base_file: Path) -> str:
    rel = target.relative_to(base_file.parent)
    return rel.as_posix()


def _markdown_escape_cell(value: str) -> str:
    return value.replace("|", "\\|").replace("\n", "<br>")


def _iter_docx_blocks(parent: DocxDocument):
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, parent)
        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)


def _paragraph_prefix(paragraph: Paragraph) -> str:
    style_name = paragraph.style.name if paragraph.style is not None else ""
    heading_match = re.match(r"Heading ([1-6])$", style_name or "")
    if heading_match:
        return "#" * int(heading_match.group(1)) + " "
    if "List Bullet" in style_name:
        return "- "
    if "List Number" in style_name:
        return "1. "
    return ""


def _format_run_text(text: str, bold: bool | None, italic: bool | None) -> str:
    if not text:
        return ""
    escaped = text.replace("\\", "\\\\").replace("[", "\\[").replace("]", "\\]")
    if bold and italic:
        return f"***{escaped}***"
    if bold:
        return f"**{escaped}**"
    if italic:
        return f"*{escaped}*"
    return escaped


def _docx_image_extension(blob: bytes, content_type: str | None) -> str:
    if content_type:
        ext = mimetypes.guess_extension(content_type)
        if ext:
            return ".jpg" if ext == ".jpe" else ext
    if blob.startswith(b"\x89PNG"):
        return ".png"
    if blob.startswith(b"\xff\xd8"):
        return ".jpg"
    if blob.startswith(b"GIF8"):
        return ".gif"
    return ".bin"


def _convert_docx(paths: ConversionPaths) -> dict[str, Any]:
    doc = Document(str(paths.source))
    image_counter = 0
    images: list[dict[str, Any]] = []
    lines: list[str] = [f"# {paths.source.stem}", ""]

    def extract_run_images(run) -> list[str]:
        nonlocal image_counter
        links: list[str] = []
        blips = run._element.xpath(".//a:blip")
        for blip in blips:
            rel_id = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
            if not rel_id or rel_id not in doc.part.related_parts:
                continue
            part = doc.part.related_parts[rel_id]
            blob = part.blob
            image_counter += 1
            ext = _docx_image_extension(blob, getattr(part, "content_type", None))
            image_path = paths.image_dir / f"image{image_counter:03d}{ext}"
            image_path.write_bytes(blob)
            rel_link = _relative_markdown_link(image_path, paths.markdown)
            links.append(f"![image {image_counter}]({rel_link})")
            images.append(
                {
                    "kind": "embedded",
                    "path": str(image_path),
                    "markdown_path": rel_link,
                    "content_type": getattr(part, "content_type", None),
                    "source_relation_id": rel_id,
                }
            )
        return links

    for block in _iter_docx_blocks(doc):
        if isinstance(block, Paragraph):
            content = []
            for run in block.runs:
                content.append(_format_run_text(run.text, run.bold, run.italic))
                content.extend(extract_run_images(run))
            paragraph_text = "".join(content).strip()
            if paragraph_text:
                lines.append(_paragraph_prefix(block) + paragraph_text)
                lines.append("")
        elif isinstance(block, Table):
            rows = []
            for row in block.rows:
                rows.append([_markdown_escape_cell(cell.text.strip()) for cell in row.cells])
            if rows:
                width = max(len(row) for row in rows)
                normalized = [row + [""] * (width - len(row)) for row in rows]
                lines.append("| " + " | ".join(normalized[0]) + " |")
                lines.append("| " + " | ".join(["---"] * width) + " |")
                for row in normalized[1:]:
                    lines.append("| " + " | ".join(row) + " |")
                lines.append("")

    _write_text(paths.markdown, "\n".join(lines))
    manifest = _base_manifest(paths, "docx")
    manifest["images"] = images
    manifest["markdown"] = str(paths.markdown)
    _write_manifest(paths, manifest)
    return manifest


def _pdf_page_markdown(page: fitz.Page) -> str:
    try:
        text = page.get_text("markdown")
    except Exception:
        text = page.get_text("text")
    return text.strip()


def _convert_pdf(
    paths: ConversionPaths,
    image_strategy: PdfImageStrategy,
    page_dpi: int,
) -> dict[str, Any]:
    if image_strategy not in {"embedded", "pages", "both"}:
        raise ValueError("pdf_image_strategy must be one of: embedded, pages, both")
    if page_dpi < 72 or page_dpi > 600:
        raise ValueError("pdf_page_dpi must be between 72 and 600")

    document = fitz.open(str(paths.source))
    images: list[dict[str, Any]] = []
    pages: list[dict[str, Any]] = []
    lines: list[str] = [f"# {paths.source.stem}", ""]

    for page_index, page in enumerate(document, start=1):
        lines.append(f"## Page {page_index}")
        lines.append("")

        page_entry: dict[str, Any] = {"page": page_index, "images": []}

        if image_strategy in {"pages", "both"}:
            page_image_path = paths.image_dir / f"page{page_index:03d}.png"
            pixmap = page.get_pixmap(dpi=page_dpi, alpha=False)
            pixmap.save(str(page_image_path))
            rel_link = _relative_markdown_link(page_image_path, paths.markdown)
            image_info = {
                "kind": "rendered_page",
                "page": page_index,
                "path": str(page_image_path),
                "markdown_path": rel_link,
                "dpi": page_dpi,
                "width": pixmap.width,
                "height": pixmap.height,
            }
            images.append(image_info)
            page_entry["images"].append(image_info)
            lines.append(f"![page {page_index}]({rel_link})")
            lines.append("")

        text = _pdf_page_markdown(page)
        if text:
            lines.append(text)
            lines.append("")

        if image_strategy in {"embedded", "both"}:
            for image_index, image in enumerate(page.get_images(full=True), start=1):
                xref = image[0]
                extracted = document.extract_image(xref)
                blob = extracted["image"]
                ext = extracted.get("ext", "bin")
                image_path = paths.image_dir / f"page{page_index:03d}-image{image_index:03d}.{ext}"
                image_path.write_bytes(blob)
                rel_link = _relative_markdown_link(image_path, paths.markdown)
                image_info = {
                    "kind": "embedded",
                    "page": page_index,
                    "path": str(image_path),
                    "markdown_path": rel_link,
                    "xref": xref,
                    "width": extracted.get("width"),
                    "height": extracted.get("height"),
                    "colorspace": extracted.get("colorspace"),
                }
                images.append(image_info)
                page_entry["images"].append(image_info)
                lines.append(f"![page {page_index} image {image_index}]({rel_link})")
                lines.append("")

        pages.append(page_entry)

    document.close()
    _write_text(paths.markdown, "\n".join(lines))
    manifest = _base_manifest(paths, "pdf")
    manifest["page_count"] = len(pages)
    manifest["pages"] = pages
    manifest["images"] = images
    manifest["markdown"] = str(paths.markdown)
    _write_manifest(paths, manifest)
    return manifest


def _base_manifest(paths: ConversionPaths, source_type: str) -> dict[str, Any]:
    return {
        "source": str(paths.source),
        "source_type": source_type,
        "source_sha256": _sha256(paths.source),
        "output_dir": str(paths.output_dir),
        "image_dir": str(paths.image_dir),
        "manifest": str(paths.manifest),
    }


@mcp.tool()
def supported_document_types() -> dict[str, list[str]]:
    """Return supported document extensions."""
    return {"extensions": [".docx", ".pdf"]}


@mcp.tool()
def convert_document(
    input_path: str,
    output_dir: str | None = None,
    overwrite: bool = False,
    pdf_image_strategy: PdfImageStrategy = "both",
    pdf_page_dpi: int = 180,
) -> dict[str, Any]:
    """Convert a DOCX or PDF into Markdown, images, and a manifest JSON file."""
    paths = _paths_for(input_path, output_dir)
    suffix = paths.source.suffix.lower()
    if suffix not in {".docx", ".pdf"}:
        raise ValueError(f"Unsupported file type: {suffix}")

    _prepare_output(paths, overwrite)
    if suffix == ".docx":
        return _convert_docx(paths)
    return _convert_pdf(paths, pdf_image_strategy, pdf_page_dpi)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
