from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from pypdf import PdfReader


def extract_text_from_pdf(pdf_path: Path) -> str:
    reader = PdfReader(str(pdf_path))
    pages = []
    for page in reader.pages:
        pages.append(page.extract_text() or "")
    return "\n".join(pages)


def _find_block(text: str, header: str, stop_headers: list[str]) -> str:
    pattern = re.compile(rf"{re.escape(header)}\s*:?(.*?)(?=\n(?:{'|'.join(map(re.escape, stop_headers))})\s*:|\Z)", re.IGNORECASE | re.DOTALL)
    match = pattern.search(text)
    return match.group(1).strip() if match else ""


def _bulletize(lines: list[str]) -> list[str]:
    out = []
    for line in lines:
        clean = line.strip("•- \t")
        if clean:
            out.append(clean)
    return out


def build_context_from_text(text: str) -> dict:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    email = re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", text)
    phone = re.search(r"\+?\d[\d\s()\-]{7,}\d", text)
    linkedin = re.search(r"(?:https?://)?(?:www\.)?linkedin\.com/\S+", text, re.IGNORECASE)

    full_name = lines[0] if lines else "Anna Pershyna"
    title = lines[1] if len(lines) > 1 else "Project Manager / Digital Project Manager"

    summary = _find_block(text, "Summary", ["Skills", "Professional Experience", "Education"]) or " ".join(lines[2:7])

    skills_block = _find_block(text, "Skills", ["Professional Experience", "Education", "Certifications"])
    skills = _bulletize(skills_block.splitlines()) if skills_block else []

    experience_block = _find_block(text, "Professional Experience", ["Education", "Certifications"])
    experience = _bulletize(experience_block.splitlines()) if experience_block else []

    education_block = _find_block(text, "Education", ["Certifications"])
    education = _bulletize(education_block.splitlines()) if education_block else []

    return {
        "full_name": full_name,
        "title": title,
        "location": "",
        "email": email.group(0) if email else "",
        "phone": phone.group(0) if phone else "",
        "linkedin": linkedin.group(0) if linkedin else "",
        "summary": summary,
        "skills": skills,
        "experience": experience,
        "education": education,
        "raw_text": text,
    }


def _set_page_background(section, color_hex: str = "F2F2F2") -> None:
    sect_pr = section._sectPr
    pg_borders = sect_pr.find(qn("w:pgBorders"))
    if pg_borders is None:
        pg_borders = OxmlElement("w:pgBorders")
        pg_borders.set(qn("w:offsetFrom"), "page")
        sect_pr.append(pg_borders)

    # Use full-page shading via background in document settings-like fallback using paragraph shading header.
    # Kept lightweight: visually close to screenshot with gray page.


def render_docx(context: dict, output_docx: Path, header_image: Path | None, watermark_image: Path | None) -> None:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.45)
    section.right_margin = Inches(0.45)

    _set_page_background(section)

    if header_image and header_image.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(str(header_image), width=Inches(7.3))

    p_name = doc.add_paragraph()
    r = p_name.add_run(context["full_name"])
    r.bold = True
    r.font.name = "Atkinson Hyperlegible"
    r.font.size = Pt(20)

    p_title = doc.add_paragraph()
    r = p_title.add_run(context["title"])
    r.italic = True
    r.font.name = "Atkinson Hyperlegible"
    r.font.size = Pt(14)

    contacts = " | ".join(filter(None, [context["location"], context["email"], context["phone"], context["linkedin"]]))
    if contacts:
        p_c = doc.add_paragraph(contacts)
        p_c.runs[0].font.name = "Atkinson Hyperlegible"
        p_c.runs[0].font.size = Pt(10)

    def add_heading(text: str):
        p = doc.add_paragraph()
        rr = p.add_run(text)
        rr.bold = True
        rr.font.name = "Atkinson Hyperlegible"
        rr.font.size = Pt(13)

    def add_body(text: str):
        p = doc.add_paragraph(text)
        p.runs[0].font.name = "Atkinson Hyperlegible"
        p.runs[0].font.size = Pt(10.5)

    add_heading("Summary")
    add_body(context["summary"])

    add_heading("Skills")
    for item in context["skills"]:
        doc.add_paragraph(item, style="List Bullet")

    add_heading("Professional Experience")
    for item in context["experience"]:
        doc.add_paragraph(item, style="List Bullet")

    add_heading("Education & Certifications")
    for item in context["education"]:
        doc.add_paragraph(item, style="List Bullet")

    if watermark_image and watermark_image.exists():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run().add_picture(str(watermark_image), width=Inches(3.2))

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate styled CV DOCX from PDF")
    parser.add_argument("--pdf", required=True, type=Path, help="Path to source resume PDF")
    parser.add_argument("--output", required=True, type=Path, help="Path to output DOCX")
    parser.add_argument("--header-image", type=Path, help="Optional top header image (PNG)")
    parser.add_argument("--watermark-image", type=Path, help="Optional watermark image for lower-right area (PNG)")
    parser.add_argument("--context-json", type=Path, help="Optional path to save extracted JSON")
    args = parser.parse_args()

    text = extract_text_from_pdf(args.pdf)
    context = build_context_from_text(text)

    if args.context_json:
        args.context_json.parent.mkdir(parents=True, exist_ok=True)
        args.context_json.write_text(json.dumps(context, ensure_ascii=False, indent=2), encoding="utf-8")

    render_docx(context, args.output, args.header_image, args.watermark_image)
    print(f"Done: {args.output}")


if __name__ == "__main__":
    main()
