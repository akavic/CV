#!/usr/bin/env python3
"""Generate ATS-friendly DOCX and PDF resume exports (full and 1-page)."""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

WORKSPACE = Path(__file__).resolve().parent.parent

FONT_NAME = "Calibri"
ACCENT = RGBColor(0x1A, 0x1A, 0x1A)
MUTED = RGBColor(0x44, 0x44, 0x44)


@dataclass
class Layout:
    condensed: bool
    margin: Inches
    body_size: Pt
    name_size: Pt
    section_size: Pt
    subhead_size: Pt
    line_spacing: float
    section_space_before: Pt
    section_space_after: Pt
    job_space_before: Pt
    bullet_space_after: Pt


def layout_for(condensed: bool) -> Layout:
    if condensed:
        return Layout(
            condensed=True,
            margin=Inches(0.5),
            body_size=Pt(9.5),
            name_size=Pt(20),
            section_size=Pt(10),
            subhead_size=Pt(9.5),
            line_spacing=1.0,
            section_space_before=Pt(5),
            section_space_after=Pt(2),
            job_space_before=Pt(4),
            bullet_space_after=Pt(1),
        )
    return Layout(
        condensed=False,
        margin=Inches(0.65),
        body_size=Pt(10.5),
        name_size=Pt(22),
        section_size=Pt(11),
        subhead_size=Pt(10.5),
        line_spacing=1.08,
        section_space_before=Pt(10),
        section_space_after=Pt(4),
        job_space_before=Pt(6),
        bullet_space_after=Pt(2),
    )


SKILLS_FULL = {
    "Languages": "JavaScript, TypeScript, Go (Golang), PHP, SQL, HTML, CSS, Shell scripting",
    "Frontend": "React 17, Redux, state management, custom hooks, form validation, REST API consumption",
    "Backend": "Go REST APIs, GORM, PHP, Node.js, RESTful API design, CRUD, transactional processing, API integration, authentication, microservices",
    "Databases": "PostgreSQL, relational modeling, multi-table transactional writes, schema alignment, Oracle SQL",
    "Cloud & DevOps": "GitLab CI/CD, Git, cross-compilation (Go/Linux), Apache, staging and production deployments, cloud deployment workflows, Azure fundamentals (AZ-900)",
    "Tools & Platforms": "Linux, Unix, Agile/Scrum, Jira, Eclipse, XML, SailPoint, Oracle Identity Analytics (OIA), IDAM, ITIL, BMC Remedy, distributed systems, production support",
}

SKILLS_CONDENSED = {
    "Languages": "JavaScript, TypeScript, Go, PHP, SQL, HTML, CSS, Shell",
    "Frontend": "React 17, Redux, hooks, state management, REST APIs",
    "Backend": "Go, GORM, PHP, Node.js, REST, CRUD, microservices, API integration, authentication",
    "Databases": "PostgreSQL, transactional writes, Oracle SQL",
    "Cloud & DevOps": "GitLab CI/CD, Git, Apache, staging/production deploys, AZ-900",
    "Tools & Platforms": "Linux, Agile, Jira, SailPoint, OIA, IDAM, ITIL, BMC Remedy, distributed systems",
}

SUMMARY = [
    "Software Engineer with production **full-stack** and **backend** experience delivering features across **React**, **Go**, **PHP**, **PostgreSQL**, **REST APIs**, and **microservices**—including **API integration**, **transactional processing**, **authentication**, **GitLab CI/CD** deployments, and **cross-stack debugging**.",
    "Enterprise **identity management** background (SailPoint, Oracle Identity Analytics) with **Agile** delivery; targeting Software Engineer, Full-Stack, Backend, Platform, API, React, and Golang roles.",
]

WILINE_FULL = [
    "Delivered an enterprise **Vendor Contract Management System** end to end across **React**, **Go**, and **PHP**: dynamic multi-type contract workflows (Type II, Real Estate, and others) with tailored validation, reusable hooks, and **state management** via custom hooks and **Redux**.",
    "Implemented production **RESTful APIs** in **Go** for full **CRUD** on vendor contracts, including **transactional processing** across related **PostgreSQL** entities (vendor contacts, NOC contacts, addresses) with **GORM** and strict struct/schema alignment.",
    "Built secure **server-to-server integration** between Go services and legacy **PHP** contract file storage using shared **API keys** and **Base64 Basic Authentication**; added standalone PHP endpoints for external upload/download without PHP session dependency.",
    "Resolved complex **cross-stack production bugs** spanning JavaScript, Go/database field mismatches, JSON unmarshal errors, and API edge cases—restoring stability for live users.",
    "Operated **GitLab CI/CD** pipelines for test and production: React builds, **Go binary** cross-compilation for Linux, and PHP releases—supporting repeatable **cloud deployment** practices.",
    "Developed and maintained **full-stack** web applications across **PHP**, **React**, **JavaScript**, and **PostgreSQL**; shipped enhancements and performance fixes through structured **debugging** and verification.",
    "Designed, built, and deployed standalone **Node.js**, **React**, and **PostgreSQL** operational tools to production and staging with ownership of business logic and data access.",
    "Partnered with cross-functional stakeholders in **Agile development** to plan work, unblock integrations, and support **production support** for released features.",
]

WILINE_CONDENSED = [
    "Delivered enterprise **Vendor Contract Management** across **React**, **Go**, and **PHP** microservices—multi-type workflows, **Redux** state management, **REST CRUD APIs**, and **PostgreSQL/GORM** transactional writes.",
    "Built secure **Go–PHP** server-to-server integration (API keys, Basic Auth) and standalone file endpoints; resolved **cross-stack production bugs** across JavaScript, Go, JSON, and API layers.",
    "Operated **GitLab CI/CD** for React, **Go** (Linux cross-compile), and **PHP** releases to test and production; maintained **full-stack** PHP/React/PostgreSQL apps and **Node.js** tooling in **Agile** teams.",
    "Shipped **Customer Satisfaction Voting System** (**React/PostgreSQL**) and internal automation utilities to staging and production.",
]

WIPRO_SAILPOINT_FULL = [
    "Diagnosed and fixed application defects with development teams via **Jira**, delivering **testable** code changes validated through **unit**, **integration**, and **UAT** testing.",
    "Supported client-facing **UAT** and integration test cycles with engineers and customers to meet release quality bars.",
    "Applied **Git**, **Eclipse**, and **XML** in development environments to implement and verify fixes on IAM application codebases.",
    "Authored defect resolution documentation for engineering and support knowledge reuse.",
    "Raised **ITIL**-aligned change requests through **BMC Remedy** for controlled production change.",
    "Led **batch run** coordination using **Unix shell scripting**, PuTTY, and WinSCP across dependent teams.",
]

WIPRO_SAILPOINT_CONDENSED = [
    "Resolved IAM defects via **Jira** with **unit/integration/UAT** testing; implemented fixes using **Git**, **Eclipse**, and **XML**; supported **UAT** cycles and **ITIL/BMC Remedy** change control.",
    "Coordinated **batch operations** with **Unix shell scripting**, PuTTY, and WinSCP across dependent teams.",
]

OIA_FULL = [
    "Integrated new resources into OIA using **Oracle SQL** and cross-team delivery for identity management initiatives.",
    "Executed **server-side testing** after patches—confirming application, web, and NFS/logging health on Linux servers before release approval.",
]

OIA_CONDENSED = [
    "Integrated OIA resources with **Oracle SQL**; performed **server-side testing** and Linux/NFS health validation after patches.",
]

PROJECTS_FULL = [
    ("Customer Satisfaction Voting System", "December 2022 – February 2023", [
        "Built a **React** web application with **PostgreSQL** persistence to collect customer satisfaction votes and feedback for service improvement.",
        "Owned front-end implementation and database-backed storage for reliable feedback capture in production.",
    ]),
    ("Operational Automation Utilities", "2022 – Present", [
        "Created standalone **Node.js**, **React**, and **PostgreSQL** scripts for internal workflows including data handling, application logic, and deployment to staging/production environments.",
    ]),
]

HIGHLIGHTS_FULL = [
    "**Full-stack ownership:** UI (React/Redux), API layers (Go/PHP), and PostgreSQL data models in the same delivery cycle.",
    "**Legacy + modern integration:** Bridging new **Go** services with established **PHP** platforms without disrupting existing workflows.",
    "**Production engineering:** Deployment discipline, integration testing, and live defect triage across distributed application tiers.",
    "**Enterprise IAM exposure:** SailPoint and OIA experience with security-sensitive systems, change control, and auditability.",
]


def strip_md(text: str) -> str:
    return re.sub(r"\*\*(.+?)\*\*", r"\1", text)


def add_run_bold(paragraph, text: str, size: Pt) -> None:
    for part in re.split(r"(\*\*.+?\*\*)", text):
        m = re.match(r"\*\*(.+?)\*\*", part)
        run = paragraph.add_run(m.group(1) if m else part)
        run.bold = bool(m)
        run.font.name = FONT_NAME
        run.font.size = size


def set_document_defaults(doc: Document, lay: Layout) -> None:
    for section in doc.sections:
        section.top_margin = lay.margin
        section.bottom_margin = lay.margin
        section.left_margin = lay.margin
        section.right_margin = lay.margin
    style = doc.styles["Normal"]
    style.font.name = FONT_NAME
    style.font.size = lay.body_size
    style.paragraph_format.space_after = Pt(1 if lay.condensed else 2)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    style.paragraph_format.line_spacing = lay.line_spacing


def add_section_heading(doc: Document, title: str, lay: Layout) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = lay.section_space_before
    p.paragraph_format.space_after = lay.section_space_after
    run = p.add_run(title.upper())
    run.bold = True
    run.font.name = FONT_NAME
    run.font.size = lay.section_size
    run.font.color.rgb = ACCENT
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "1A1A1A")
    pBdr.append(bottom)


def add_contact(doc: Document, lay: Layout) -> None:
    name = doc.add_paragraph()
    name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    nr = name.add_run("Victor Akinla")
    nr.bold = True
    nr.font.name = FONT_NAME
    nr.font.size = lay.name_size
    nr.font.color.rgb = ACCENT

    contact = doc.add_paragraph()
    contact.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cr = contact.add_run(
        "Dublin, Ireland  |  Irish  |  +353 86 073 1698  |  vakinla.va@gmail.com"
    )
    cr.font.name = FONT_NAME
    cr.font.size = Pt(9.5 if lay.condensed else 10)
    cr.font.color.rgb = MUTED

    links = doc.add_paragraph()
    links.alignment = WD_ALIGN_PARAGRAPH.CENTER
    lr = links.add_run(
        "LinkedIn: linkedin.com/in/victor-akinla-21061a88  |  GitHub: github.com/akavic"
    )
    lr.font.name = FONT_NAME
    lr.font.size = Pt(9 if lay.condensed else 9.5)
    lr.font.color.rgb = MUTED
    links.paragraph_format.space_after = Pt(2 if lay.condensed else 4)


def add_body_paragraph(doc: Document, text: str, lay: Layout, space_after: int | None = None) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after if space_after is not None else (2 if lay.condensed else 4))
    add_run_bold(p, text, lay.body_size)


def add_skill_line(doc: Document, label: str, value: str, lay: Layout) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1 if lay.condensed else 2)
    lr = p.add_run(f"{label}: ")
    lr.bold = True
    lr.font.name = FONT_NAME
    lr.font.size = lay.body_size
    vr = p.add_run(value)
    vr.font.name = FONT_NAME
    vr.font.size = lay.body_size


def add_job_header(doc: Document, company: str, role: str, dates: str, lay: Layout, location: str = "") -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = lay.job_space_before
    p.paragraph_format.space_after = Pt(0)
    c = p.add_run(f"{company} — {role}")
    c.bold = True
    c.font.name = FONT_NAME
    c.font.size = lay.subhead_size
    meta = doc.add_paragraph()
    meta.paragraph_format.space_after = Pt(2 if lay.condensed else 3)
    mr = meta.add_run(f"{dates}" + (f"  |  {location}" if location else ""))
    mr.italic = True
    mr.font.name = FONT_NAME
    mr.font.size = Pt(9 if lay.condensed else 10)
    mr.font.color.rgb = MUTED


def add_subrole(doc: Document, title: str, lay: Layout) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(3 if lay.condensed else 4)
    p.paragraph_format.space_after = Pt(1 if lay.condensed else 2)
    r = p.add_run(title)
    r.bold = True
    r.font.name = FONT_NAME
    r.font.size = lay.body_size


def add_bullet(doc: Document, text: str, lay: Layout) -> None:
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = lay.bullet_space_after
    p.paragraph_format.left_indent = Inches(0.15 if lay.condensed else 0.2)
    add_run_bold(p, text, lay.body_size)


def build_docx(condensed: bool = False) -> Document:
    lay = layout_for(condensed)
    doc = Document()
    set_document_defaults(doc, lay)
    add_contact(doc, lay)

    add_section_heading(doc, "Professional Summary", lay)
    for i, s in enumerate(SUMMARY):
        add_body_paragraph(doc, s, lay, space_after=1 if condensed and i else None)

    add_section_heading(doc, "Technical Skills", lay)
    skills = SKILLS_CONDENSED if condensed else SKILLS_FULL
    for label, value in skills.items():
        add_skill_line(doc, label, value, lay)

    add_section_heading(doc, "Professional Experience", lay)

    add_job_header(doc, "WiLine", "Software Developer", "August 2022 – Present", lay, "Dublin, Ireland")
    if not condensed:
        add_body_paragraph(
            doc,
            "Engineer on customer and internal platforms combining **React** frontends, **Go REST APIs**, **PHP** legacy services, and **PostgreSQL** under a **microservices** architecture. Owns feature delivery, **backend development**, **API integration**, production defect resolution, and **GitLab CI/CD** deployments.",
            lay,
            space_after=3,
        )
    for b in WILINE_CONDENSED if condensed else WILINE_FULL:
        add_bullet(doc, b, lay)

    add_job_header(doc, "Wipro", "Application Engineer", "April 2019 – July 2022", lay, "Dublin, Ireland")
    if not condensed:
        add_body_paragraph(
            doc,
            "Application engineer for global **identity access management** programs, focused on defect resolution, integration testing, and reliable enterprise releases.",
            lay,
            space_after=3,
        )
    add_subrole(doc, "SailPoint | 2020 – 2022", lay)
    for b in WIPRO_SAILPOINT_CONDENSED if condensed else WIPRO_SAILPOINT_FULL:
        add_bullet(doc, b, lay)
    add_subrole(doc, "Oracle Identity Analytics (OIA) | 2019 – 2020", lay)
    for b in OIA_CONDENSED if condensed else OIA_FULL:
        add_bullet(doc, b, lay)

    add_job_header(doc, "Host Ireland", "Network Engineer", "September 2018 – March 2019", lay)
    if condensed:
        add_body_paragraph(
            doc,
            "Provisioned customer IP networks and configured **Juniper** routing; resolved help-desk connectivity incidents.",
            lay,
            space_after=1,
        )
    else:
        for b in [
            "Provisioned customer IP subnets and resolved network incidents through help-desk support.",
            "Configured customer routing paths on **Juniper** network devices.",
        ]:
            add_bullet(doc, b, lay)

    if not condensed:
        add_section_heading(doc, "Projects", lay)
        for title, dates, bullets in PROJECTS_FULL:
            add_job_header(doc, "WiLine", title, dates, lay)
            for b in bullets:
                add_bullet(doc, b, lay)

    add_section_heading(doc, "Education & Certifications" if condensed else "Education", lay)
    p = doc.add_paragraph()
    r1 = p.add_run("BSc Computer Science — Dublin City University")
    r1.bold = True
    r1.font.name = FONT_NAME
    r1.font.size = lay.body_size
    meta = doc.add_paragraph()
    m = meta.add_run("August 2013 – May 2017  |  Dublin, Ireland")
    m.font.name = FONT_NAME
    m.font.size = Pt(9 if condensed else 10)
    m.font.color.rgb = MUTED
    if condensed:
        cert = doc.add_paragraph()
        cr = cert.add_run("Certifications: Microsoft AZ-900 (Azure Fundamentals), TOGAF 9 Certified")
        cr.font.name = FONT_NAME
        cr.font.size = lay.body_size
    else:
        add_body_paragraph(doc, "Relevant study: data structures and algorithms, databases, cryptography.", lay, space_after=2)
        add_section_heading(doc, "Certifications", lay)
        for cert in [
            "Microsoft Certified: Azure Fundamentals (AZ-900)",
            "TOGAF 9 Certified — The Open Group",
        ]:
            add_bullet(doc, cert, lay)

    if not condensed:
        add_section_heading(doc, "Additional Technical Highlights", lay)
        for b in HIGHLIGHTS_FULL:
            add_bullet(doc, b, lay)

    return doc


def build_html(condensed: bool = False) -> str:
    skills = SKILLS_CONDENSED if condensed else SKILLS_FULL
    skills_html = "".join(
        f'<p><span class="label">{k}:</span> {v}</p>' for k, v in skills.items()
    )
    page_css = (
        "@page { size: A4; margin: 0.45in 0.5in; }"
        if condensed
        else "@page { size: A4; margin: 0.55in 0.6in; }"
    )
    body_size = "9.5pt" if condensed else "10.5pt"
    h1_size = "20pt" if condensed else "22pt"
    h2_margin = "8px 0 4px 0" if condensed else "12px 0 6px 0"

    def bullets(items: list[str]) -> str:
        return "<ul>" + "".join(f"<li>{strip_md(i)}</li>" for i in items) + "</ul>"

    wiline = bullets([strip_md(b) for b in (WILINE_CONDENSED if condensed else WILINE_FULL)])
    sail = bullets([strip_md(b) for b in (WIPRO_SAILPOINT_CONDENSED if condensed else WIPRO_SAILPOINT_FULL)])
    oia = bullets([strip_md(b) for b in (OIA_CONDENSED if condensed else OIA_FULL)])

    host = (
        "<p>Provisioned customer IP networks and configured Juniper routing; resolved help-desk connectivity incidents.</p>"
        if condensed
        else bullets([
            "Provisioned customer IP subnets and resolved network incidents through help-desk support.",
            "Configured customer routing paths on Juniper network devices.",
        ])
    )

    projects = ""
    if not condensed:
        proj_parts = []
        for title, dates, blts in PROJECTS_FULL:
            proj_parts.append(f'<p class="job-title">{title} — WiLine</p><p class="job-meta">{dates}</p>{bullets([strip_md(b) for b in blts])}')
        projects = f"<h2>Projects</h2>{''.join(proj_parts)}"

    edu = (
        "<p><strong>BSc Computer Science — Dublin City University</strong><br/>"
        '<span class="job-meta">August 2013 – May 2017 | Dublin, Ireland</span></p>'
        "<p>Certifications: Microsoft AZ-900 (Azure Fundamentals), TOGAF 9 Certified</p>"
        if condensed
        else (
            "<p><strong>BSc Computer Science — Dublin City University</strong><br/>"
            '<span class="job-meta">August 2013 – May 2017 | Dublin, Ireland</span></p>'
            "<p>Relevant study: data structures and algorithms, databases, cryptography.</p>"
            "<h2>Certifications</h2>"
            + bullets(["Microsoft Certified: Azure Fundamentals (AZ-900)", "TOGAF 9 Certified — The Open Group"])
        )
    )

    highlights = ""
    if not condensed:
        highlights = "<h2>Additional Technical Highlights</h2>" + bullets([strip_md(b) for b in HIGHLIGHTS_FULL])

    summary = "".join(f"<p>{strip_md(s)}</p>" for s in SUMMARY)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<title>Victor Akinla — Software Engineer Resume</title>
<style>
  {page_css}
  body {{ font-family: Calibri, Arial, Helvetica, sans-serif; font-size: {body_size}; line-height: 1.15; color: #1a1a1a; margin: 0 auto; }}
  h1 {{ font-size: {h1_size}; text-align: center; margin: 0 0 3px 0; }}
  .contact, .links {{ text-align: center; color: #444; margin: 0; font-size: 9.5pt; }}
  .links {{ margin-bottom: 6px; }}
  h2 {{ font-size: 10pt; text-transform: uppercase; border-bottom: 1px solid #1a1a1a; padding-bottom: 1px; margin: {h2_margin}; }}
  .job-title {{ font-weight: bold; margin: 5px 0 0 0; }}
  .job-meta {{ font-style: italic; color: #444; font-size: 9pt; margin: 0 0 3px 0; }}
  .subrole {{ font-weight: bold; margin: 4px 0 1px 0; font-size: 9.5pt; }}
  p {{ margin: 0 0 3px 0; }}
  ul {{ margin: 1px 0 4px 0; padding-left: 16px; }}
  li {{ margin-bottom: 1px; }}
  .label {{ font-weight: bold; }}
</style>
</head>
<body>
<h1>Victor Akinla</h1>
<p class="contact">Dublin, Ireland | Irish | +353 86 073 1698 | vakinla.va@gmail.com</p>
<p class="links">LinkedIn: linkedin.com/in/victor-akinla-21061a88 | GitHub: github.com/akavic</p>
<h2>Professional Summary</h2>
{summary}
<h2>Technical Skills</h2>
{skills_html}
<h2>Professional Experience</h2>
<p class="job-title">WiLine — Software Developer</p>
<p class="job-meta">August 2022 – Present | Dublin, Ireland</p>
{wiline}
<p class="job-title">Wipro — Application Engineer</p>
<p class="job-meta">April 2019 – July 2022 | Dublin, Ireland</p>
<p class="subrole">SailPoint | 2020 – 2022</p>
{sail}
<p class="subrole">Oracle Identity Analytics (OIA) | 2019 – 2020</p>
{oia}
<p class="job-title">Host Ireland — Network Engineer</p>
<p class="job-meta">September 2018 – March 2019</p>
{host}
{projects}
<h2>Education{" & Certifications" if condensed else ""}</h2>
{edu}
{highlights}
</body>
</html>"""


def export_pdf(html: str, path: Path) -> bool:
    try:
        from weasyprint import HTML

        HTML(string=html).write_pdf(str(path))
        return True
    except Exception as exc:
        print(f"PDF export skipped ({path.name}): {exc}")
        return False


def page_count(path: Path) -> int | None:
    try:
        from pypdf import PdfReader

        return len(PdfReader(str(path)).pages)
    except Exception:
        return None


def generate_variant(condensed: bool) -> None:
    suffix = "-1page" if condensed else ""
    docx_path = WORKSPACE / f"victor-akinla-software-engineer-resume{suffix}.docx"
    pdf_path = WORKSPACE / f"victor-akinla-software-engineer-resume{suffix}.pdf"
    html_path = WORKSPACE / f"victor-akinla-software-engineer-resume{suffix}.html"

    doc = build_docx(condensed=condensed)
    doc.save(docx_path)
    print(f"DOCX: {docx_path}")

    html = build_html(condensed=condensed)
    html_path.write_text(html, encoding="utf-8")
    print(f"HTML: {html_path}")

    if export_pdf(html, pdf_path):
        pages = page_count(pdf_path)
        print(f"PDF:  {pdf_path} ({pages} page{'s' if pages != 1 else ''})")


def main() -> None:
    print("=== Full resume (2 pages) ===")
    generate_variant(condensed=False)
    print("\n=== Condensed 1-page resume ===")
    generate_variant(condensed=True)


if __name__ == "__main__":
    main()
