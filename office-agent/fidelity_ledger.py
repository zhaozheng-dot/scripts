#!/usr/bin/env python3
"""Build fidelity ledger files for PPTX to DOCX conversion."""

import argparse
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_common import read_json, write_json, write_text


def slide_title(slide):
    return slide.get('title') or f"Slide {slide.get('slide_no')}"


def ledger_rows(extracted, plan):
    mode = plan.get('selected_mode', 'generic_raw')
    rows = []
    for slide in extracted.get('slides', []):
        title = slide_title(slide)
        if mode in {'generic_raw', 'editable_material', 'raw_transcript'}:
            location = f"Slide {slide['slide_no']} section"
            method = 'preserved in slide order'
        elif mode == 'generic_reading':
            location = f"Reading section {slide['slide_no']}"
            method = 'reflowed into readable Word paragraphs with source order preserved'
        elif mode == 'generic_visual_report':
            location = guess_generic_location(title, slide)
            method = 'converted into generic report components and page appendix'
        else:
            location = guess_professional_location(title, slide)
            method = 'restructured by user-confirmed professional template'
        complex_count = sum(1 for item in slide.get('items', []) if item.get('semantic_guess') == 'complex_visual')
        image_count = sum(1 for item in slide.get('items', []) if item.get('type') == 'image')
        notes = []
        if complex_count:
            notes.append(f'{complex_count} complex visual item(s) require manual review')
        if image_count and not plan.get('include_images'):
            notes.append(f'{image_count} image item(s) not embedded by selected plan')
        if not notes:
            notes.append('processed')
        rows.append({
            'slide_no': slide.get('slide_no'),
            'original_title': title,
            'word_location': location,
            'handling': method,
            'item_count': len(slide.get('items', [])),
            'notes': '; '.join(notes),
        })
    return rows


def guess_generic_location(title, slide):
    text = (title + ' ' + ' '.join(item.get('text', '') for item in slide.get('items', []))).lower()
    if any(key in text for key in ['summary', '摘要', '结论', '建议']):
        return 'Generic visual report - summary/recommendations'
    if any(key in text for key in ['risk', '风险', 'pending', '警示']):
        return 'Generic visual report - risks/attention items'
    if any(key in text for key in ['source', '来源', 'disclaimer', '免责声明']):
        return 'Generic visual report - sources/disclaimer'
    return 'Generic visual report - page appendix'


def guess_professional_location(title, slide):
    text = (title + ' ' + ' '.join(item.get('text', '') for item in slide.get('items', []))).lower()
    if 'executive' in text or '摘要' in text:
        return 'Executive summary'
    if 'company' in text or '公司概况' in text:
        return 'Company and product overview'
    if 'team' in text or '团队' in text:
        return 'Management team and governance'
    if 'regulatory' in text or 'commercial' in text or '监管' in text or '商业化' in text:
        return 'Regulatory and commercialization'
    if 'competitive' in text or 'business model' in text or '竞争' in text or '商业模式' in text:
        return 'Market, competition, and business model'
    if 'financial' in text or 'swot' in text or 'risk' in text or '财务' in text or '风险' in text:
        return 'Financials, SWOT, and risk matrix'
    if 'source' in text or 'appendix' in text or '来源' in text or '附录' in text:
        return 'Sources and appendix'
    return 'Professional template body'


def markdown(rows):
    lines = ['# Fidelity Ledger', '', '| PPT Page | Original Title | Word Location | Handling | Notes |', '|---|---|---|---|---|']
    for row in rows:
        lines.append(f"| {row['slide_no']} | {row['original_title']} | {row['word_location']} | {row['handling']} | {row['notes']} |")
    lines.append('')
    return '\n'.join(lines)


def main():
    parser = argparse.ArgumentParser(description='Build fidelity ledger from extracted PPTX data and plan.')
    parser.add_argument('extract_json')
    parser.add_argument('plan_json')
    parser.add_argument('ledger_json')
    parser.add_argument('ledger_md')
    args = parser.parse_args()
    extracted = read_json(args.extract_json)
    plan = read_json(args.plan_json)
    rows = ledger_rows(extracted, plan)
    write_json(args.ledger_json, {'source': extracted.get('source'), 'rows': rows})
    write_text(args.ledger_md, markdown(rows))
    print(args.ledger_json)
    print(args.ledger_md)


if __name__ == '__main__':
    main()
