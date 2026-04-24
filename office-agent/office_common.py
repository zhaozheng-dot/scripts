#!/usr/bin/env python3
"""Shared helpers for Office Agent scripts."""

import json
import os
from datetime import datetime


def ensure_dir(path):
    directory = os.path.dirname(os.path.abspath(path))
    if directory:
        os.makedirs(directory, exist_ok=True)


def read_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def write_json(path, data):
    ensure_dir(path)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def write_text(path, text):
    ensure_dir(path)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text)


def safe_stem(path):
    return os.path.splitext(os.path.basename(path))[0]


def unique_output(path):
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    stamp = datetime.now().strftime('%Y%m%d-%H%M%S')
    return f'{stem}-{stamp}{ext}'


def refuse_overwrite_input(input_path, output_path):
    if os.path.abspath(input_path) == os.path.abspath(output_path):
        raise SystemExit('Refusing to overwrite the input file; choose a different output path.')


def classify_text_density(slides, text_count):
    if slides == 0:
        return 'none'
    per_slide = text_count / slides
    if per_slide >= 25:
        return 'high'
    if per_slide >= 10:
        return 'medium'
    return 'low'


def detect_document_type(titles, texts):
    blob = ' '.join(titles + texts).lower()
    if any(word in blob for word in ['investment', '投资', 'swot', 'risk assessment', '风险评估', 'funding', '融资']):
        return 'investment_review'
    if any(word in blob for word in ['product', '产品', 'manual', '手册']):
        return 'product_manual'
    if any(word in blob for word in ['project', '项目', 'progress', '进展']):
        return 'project_update'
    if any(word in blob for word in ['business', '商业', 'market', '市场']):
        return 'business_report'
    return 'general_presentation'


def risk_level_for(document_type):
    if document_type in {'investment_review'}:
        return 'high'
    if document_type in {'business_report', 'project_update'}:
        return 'medium'
    return 'low'
