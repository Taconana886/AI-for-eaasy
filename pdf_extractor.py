import logging
import io
import re
from pathlib import Path
from typing import Optional

logger = logging.getLogger(__name__)


def extract_text_and_images(pdf_path: str | Path) -> dict:
    """Extract text, images, and metadata from PDF using PyMuPDF.

    Returns:
        {
            "text": str,              # full extracted text with layout
            "pages": [str],           # per-page text
            "images": [               # extracted images
                {"data": bytes, "x": float, "y": float, "w": float, "h": float,
                 "page": int, "alt_text": str, "type": "figure"|"table"}
            ],
            "metadata": dict,         # PDF metadata
        }
    """
    import fitz

    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    doc = fitz.open(str(pdf_path))
    result = {
        "text": "",
        "pages": [],
        "images": [],
        "metadata": doc.metadata or {},
    }

    for page_num in range(len(doc)):
        page = doc[page_num]
        # Text with layout preservation
        page_text = page.get_text("text")
        result["pages"].append(page_text)
        result["text"] += f"\n--- Page {page_num + 1} ---\n{page_text}"

        # Extract images
        image_list = page.get_images(full=True)
        for img_idx, img_info in enumerate(image_list):
            xref = img_info[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            if len(image_bytes) < 1024:
                continue

            # Find image position on page
            bbox = _find_image_bbox(page, xref)
            result["images"].append({
                "data": image_bytes,
                "x": bbox.x0 if bbox else 0,
                "y": bbox.y0 if bbox else 0,
                "w": bbox.width if bbox else 0,
                "h": bbox.height if bbox else 0,
                "ext": base_image["ext"],
                "page": page_num,
                "alt_text": _guess_image_caption(page_text, bbox),
                "type": "figure",
            })

    doc.close()
    return result


def _find_image_bbox(page, xref):
    import fitz
    for block in page.get_text("dict")["blocks"]:
        if block["type"] == 1:
            for img in block.get("images", []):
                if img.get("xref") == xref:
                    return fitz.Rect(block["bbox"])
    return None


def _guess_image_caption(page_text: str, bbox) -> str:
    if not bbox:
        return ""
    lines = [l.strip() for l in page_text.split("\n") if l.strip()]
    for line in lines:
        if re.search(r"(?:Fig(?:ure)?|Table|图|表)\s*[\.\:：\s]*\d", line, re.I):
            return line[:200]
    for line in lines[-5:]:
        if re.search(r"(?:Fig(?:ure)?|Table|图|表)", line, re.I):
            return line[:200]
    return ""


def extract_images_from_pdf(pdf_path: str | Path) -> list[dict]:
    """Quick image extraction, returns list of {data, ext, page, caption}."""
    result = extract_text_and_images(pdf_path)
    return result["images"]


def extract_text_from_pdf(pdf_path: str | Path) -> str:
    result = extract_text_and_images(pdf_path)
    return result["text"]
