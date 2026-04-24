#!/usr/bin/env python3
"""Generate basic Office-like documents from JSON specs.

This MVP has no third-party dependencies. It writes:
- .docx-compatible HTML documents
- .xls-compatible HTML workbooks
- .pptx-compatible HTML slide decks

Use real OOXML renderers later when python-docx/python-pptx/openpyxl are available.
"""

import argparse
import html
import json
import os
from datetime import datetime


def ensure_dir(path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)


def esc(value):
    return html.escape(str(value), quote=True)


def read_spec(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def html_page(title, body, extra_style=''):
    style = f"""
body {{ font-family: Aptos, Calibri, Arial, sans-serif; line-height: 1.55; color: #1f2933; margin: 48px; }}
h1 {{ color: #12355b; border-bottom: 2px solid #d7e3f4; padding-bottom: 8px; }}
h2 {{ color: #1f4f7a; margin-top: 28px; }}
table {{ border-collapse: collapse; width: 100%; margin: 16px 0; }}
th, td {{ border: 1px solid #b8c6d8; padding: 8px 10px; vertical-align: top; }}
th {{ background: #eaf2fb; }}
.slide {{ page-break-after: always; min-height: 560px; padding: 28px; border: 1px solid #d5dee8; margin-bottom: 28px; }}
.subtitle {{ color: #64748b; font-size: 18px; }}
.meta {{ color: #64748b; font-size: 12px; }}
{extra_style}
"""
    return f"<!doctype html><html><head><meta charset='utf-8'><title>{esc(title)}</title><style>{style}</style></head><body>{body}</body></html>"


def render_docx(spec, output):
    title = spec.get('title', 'Document')
    parts = [f"<h1>{esc(title)}</h1>", f"<p class='meta'>Generated: {datetime.now():%Y-%m-%d %H:%M}</p>"]
    for section in spec.get('sections', []):
        parts.append(f"<h2>{esc(section.get('heading', 'Section'))}</h2>")
        for para in section.get('paragraphs', []):
            parts.append(f"<p>{esc(para)}</p>")
        table = section.get('table')
        if table:
            parts.append('<table>')
            headers = table.get('headers', [])
            if headers:
                parts.append('<tr>' + ''.join(f"<th>{esc(h)}</th>" for h in headers) + '</tr>')
            for row in table.get('rows', []):
                parts.append('<tr>' + ''.join(f"<td>{esc(c)}</td>" for c in row) + '</tr>')
            parts.append('</table>')
        for item in section.get('bullets', []):
            parts.append(f"<ul><li>{esc(item)}</li></ul>")
    ensure_dir(output)
    with open(output, 'w', encoding='utf-8') as f:
        f.write(html_page(title, '\n'.join(parts)))


def render_xlsx(spec, output):
    title = spec.get('title', 'Workbook')
    parts = [f"<h1>{esc(title)}</h1>"]
    for sheet in spec.get('sheets', []):
        parts.append(f"<h2>{esc(sheet.get('name', 'Sheet'))}</h2><table>")
        headers = sheet.get('headers', [])
        if headers:
            parts.append('<tr>' + ''.join(f"<th>{esc(h)}</th>" for h in headers) + '</tr>')
        for row in sheet.get('rows', []):
            parts.append('<tr>' + ''.join(f"<td>{esc(c)}</td>" for c in row) + '</tr>')
        parts.append('</table>')
    ensure_dir(output)
    with open(output, 'w', encoding='utf-8') as f:
        f.write(html_page(title, '\n'.join(parts)))


def render_pptx(spec, output):
    title = spec.get('title', 'Slide Deck')
    parts = [f"<h1>{esc(title)}</h1>"]
    for index, slide in enumerate(spec.get('slides', []), start=1):
        parts.append('<section class="slide">')
        parts.append(f"<p class='meta'>Slide {index}</p>")
        parts.append(f"<h2>{esc(slide.get('title', 'Untitled'))}</h2>")
        if slide.get('subtitle'):
            parts.append(f"<p class='subtitle'>{esc(slide['subtitle'])}</p>")
        if slide.get('body'):
            parts.append(f"<p>{esc(slide['body'])}</p>")
        bullets = slide.get('bullets', [])
        if bullets:
            parts.append('<ul>' + ''.join(f"<li>{esc(b)}</li>" for b in bullets) + '</ul>')
        parts.append('</section>')
    ensure_dir(output)
    with open(output, 'w', encoding='utf-8') as f:
        f.write(html_page(title, '\n'.join(parts), extra_style='@page { size: landscape; }'))


def main():
    parser = argparse.ArgumentParser(description='Generate Office-like documents from JSON specs.')
    parser.add_argument('kind', choices=['docx', 'xlsx', 'pptx'])
    parser.add_argument('spec')
    parser.add_argument('output')
    args = parser.parse_args()
    spec = read_spec(args.spec)
    {'docx': render_docx, 'xlsx': render_xlsx, 'pptx': render_pptx}[args.kind](spec, args.output)
    print(args.output)


if __name__ == '__main__':
    main()
