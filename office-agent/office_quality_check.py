#!/usr/bin/env python3
"""Basic quality checks for generated Office Agent DOCX output."""

import argparse
import os
import sys

from docx import Document

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_common import read_json, write_json


def extracted_has_attention_or_source(extracted):
    for slide in extracted.get('slides', []):
        for item in slide.get('items', []):
            if item.get('semantic_guess') in {'risk', 'source_or_disclaimer', 'summary_or_recommendation'}:
                return True
            text = item.get('text', '')
            if any(key in text.lower() for key in ['risk', 'source', 'disclaimer']):
                return True
            if any(key in text for key in ['风险', '来源', '免责声明', '建议']):
                return True
    return False


def check(docx_path, extract_json=None, ledger_json=None):
    doc = Document(docx_path)
    text = '\n'.join(p.text for p in doc.paragraphs)
    extracted = read_json(extract_json) if extract_json else None
    has_attention_or_source = any(
        key in text for key in ['风险', 'risk', 'Risk', '来源', 'source', 'Source', '免责声明', '建议']
    )
    if not has_attention_or_source and extracted:
        has_attention_or_source = extracted_has_attention_or_source(extracted)
    result = {
        'file': docx_path,
        'exists': os.path.exists(docx_path),
        'size_bytes': os.path.getsize(docx_path),
        'paragraphs': len(doc.paragraphs),
        'tables': len(doc.tables),
        'technical_ok': True,
        'experience_checks': {
            'has_cover_like_title': bool(doc.paragraphs and doc.paragraphs[0].text.strip()),
            'has_summary_or_navigation': any(key in text for key in ['执行摘要', 'Executive Summary', '内容导航', '一页摘要']),
            'has_attention_or_source_content': has_attention_or_source,
            'avoids_slide_by_slide': 'Slide 1:' not in text[:2000],
        },
        'content_checks': {},
        'warnings': [],
    }
    if extracted:
        slides = len(extracted.get('slides', []))
        result['content_checks']['source_slides'] = slides
    if ledger_json:
        ledger = read_json(ledger_json)
        rows = ledger.get('rows', [])
        result['content_checks']['ledger_rows'] = len(rows)
        if extract_json and len(rows) != result['content_checks'].get('source_slides'):
            result['warnings'].append('Ledger row count does not match source slide count.')
    failed = [k for k, v in result['experience_checks'].items() if not v]
    if failed:
        result['warnings'].append('Experience checks failed: ' + ', '.join(failed))
    return result


def main():
    parser = argparse.ArgumentParser(description='Run basic quality checks for DOCX output.')
    parser.add_argument('docx')
    parser.add_argument('output_json')
    parser.add_argument('--extract-json', default='')
    parser.add_argument('--ledger-json', default='')
    args = parser.parse_args()
    result = check(args.docx, args.extract_json or None, args.ledger_json or None)
    write_json(args.output_json, result)
    print(args.output_json)


if __name__ == '__main__':
    main()
