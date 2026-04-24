#!/usr/bin/env python3
"""Safely modify Office documents by writing a new output file."""

import argparse
import json
import os


def ensure_dir(path):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)


def read_spec(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def assert_safe_output(input_path, output_path):
    if os.path.abspath(input_path) == os.path.abspath(output_path):
        raise SystemExit('Refusing to overwrite the input file; choose a different output path.')
    if os.path.exists(output_path):
        raise SystemExit(f'Refusing to overwrite existing output: {output_path}')


def replace_in_paragraph(paragraph, old, new):
    if old not in paragraph.text:
        return False
    # Rebuild runs to keep behavior predictable for generated documents.
    paragraph.text = paragraph.text.replace(old, new)
    return True


def modify_docx(input_path, spec, output_path):
    from docx import Document

    document = Document(input_path)
    for op in spec.get('operations', []):
        name = op.get('op')
        if name == 'replace_text':
            old = str(op.get('old', ''))
            new = str(op.get('new', ''))
            if not old:
                continue
            for paragraph in document.paragraphs:
                replace_in_paragraph(paragraph, old, new)
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            replace_in_paragraph(paragraph, old, new)
        elif name == 'append_section':
            document.add_heading(op.get('heading', 'Section'), level=int(op.get('level', 1)))
            for para in op.get('paragraphs', []):
                document.add_paragraph(str(para))
            for item in op.get('bullets', []):
                document.add_paragraph(str(item), style='List Bullet')
            table = op.get('table')
            if table:
                headers = table.get('headers', [])
                rows = table.get('rows', [])
                cols = max(len(headers), *(len(row) for row in rows), 1)
                doc_table = document.add_table(rows=1 if headers else 0, cols=cols)
                doc_table.style = 'Table Grid'
                if headers:
                    for i, value in enumerate(headers):
                        doc_table.rows[0].cells[i].text = str(value)
                for row in rows:
                    cells = doc_table.add_row().cells
                    for i, value in enumerate(row):
                        cells[i].text = str(value)
    ensure_dir(output_path)
    document.save(output_path)


def get_or_create_sheet(workbook, name):
    if name in workbook.sheetnames:
        return workbook[name]
    return workbook.create_sheet(title=str(name)[:31])


def modify_xlsx(input_path, spec, output_path):
    from openpyxl import load_workbook

    workbook = load_workbook(input_path)
    for op in spec.get('operations', []):
        name = op.get('op')
        sheet_name = op.get('sheet') or workbook.sheetnames[0]
        sheet = get_or_create_sheet(workbook, sheet_name)
        if name == 'set_cell':
            sheet[str(op.get('cell', 'A1'))] = op.get('value')
        elif name == 'append_rows':
            for row in op.get('rows', []):
                sheet.append(row)
        elif name == 'replace_text':
            old = str(op.get('old', ''))
            new = str(op.get('new', ''))
            if not old:
                continue
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and old in cell.value:
                        cell.value = cell.value.replace(old, new)
    ensure_dir(output_path)
    workbook.save(output_path)


def modify_pptx(input_path, spec, output_path):
    from pptx import Presentation
    from pptx.util import Pt

    prs = Presentation(input_path)
    for op in spec.get('operations', []):
        name = op.get('op')
        if name == 'replace_text':
            old = str(op.get('old', ''))
            new = str(op.get('new', ''))
            if not old:
                continue
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, 'text') and old in shape.text:
                        shape.text = shape.text.replace(old, new)
        elif name == 'append_slide':
            layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = str(op.get('title', 'Untitled'))
            body = slide.placeholders[1].text_frame
            body.clear()
            body_text = op.get('body')
            if body_text:
                body.paragraphs[0].text = str(body_text)
                body.paragraphs[0].font.size = Pt(20)
            for item in op.get('bullets', []):
                p = body.add_paragraph()
                p.text = str(item)
                p.level = 0
                p.font.size = Pt(20)
            notes = op.get('notes')
            if notes:
                slide.notes_slide.notes_text_frame.text = str(notes)
    ensure_dir(output_path)
    prs.save(output_path)


def main():
    parser = argparse.ArgumentParser(description='Modify Office documents without overwriting originals.')
    parser.add_argument('input')
    parser.add_argument('spec')
    parser.add_argument('output')
    args = parser.parse_args()
    assert_safe_output(args.input, args.output)
    spec = read_spec(args.spec)
    ext = os.path.splitext(args.input)[1].lower()
    if ext == '.docx':
        modify_docx(args.input, spec, args.output)
    elif ext == '.xlsx':
        modify_xlsx(args.input, spec, args.output)
    elif ext == '.pptx':
        modify_pptx(args.input, spec, args.output)
    else:
        raise SystemExit(f'Unsupported input extension: {ext}')
    print(args.output)


if __name__ == '__main__':
    main()
