#!/usr/bin/env python3
"""Run Office Agent conversion regression cases."""

import argparse
import json
import os
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_common import write_json, write_text
from office_convert import create_plan, run_confirmed_plan


def discover_cases(input_dir):
    return sorted(
        os.path.join(input_dir, name)
        for name in os.listdir(input_dir)
        if name.lower().endswith('.pptx')
    )


def run_case(input_path, output_root):
    case_name = os.path.splitext(os.path.basename(input_path))[0]
    workspace = os.path.join(output_root, case_name)
    os.makedirs(workspace, exist_ok=True)
    paths, plan = create_plan(input_path, workspace=workspace)
    result = run_confirmed_plan(paths['plan'], confirm=True)
    quality = {}
    try:
        with open(result['quality'], 'r', encoding='utf-8') as f:
            quality = json.load(f)
    except Exception as exc:
        quality = {'status': 'fail', 'warnings': [str(exc)]}
    return {
        'case': case_name,
        'input': input_path,
        'workspace': workspace,
        'plan': paths['plan'],
        'output': result.get('docx'),
        'quality': result.get('quality'),
        'quality_md': result.get('quality_md'),
        'status': quality.get('status', 'unknown'),
        'warnings': quality.get('warnings', []),
    }


def summary_markdown(summary):
    lines = [
        '# Office Agent Regression Summary',
        '',
        f"- Run at: `{summary['run_at']}`",
        f"- Status: `{summary['status']}`",
        f"- Total cases: `{len(summary['cases'])}`",
        '',
        '| Case | Status | Output |',
        '|---|---|---|',
    ]
    for case in summary['cases']:
        lines.append(f"| {case['case']} | `{case['status']}` | `{case.get('output')}` |")
    lines.extend(['', '## Warnings', ''])
    any_warning = False
    for case in summary['cases']:
        for warning in case.get('warnings', []):
            any_warning = True
            lines.append(f"- {case['case']}: {warning}")
    if not any_warning:
        lines.append('- None')
    lines.append('')
    return '\n'.join(lines)


def main():
    parser = argparse.ArgumentParser(description='Run Office Agent regression suite.')
    parser.add_argument('--input-dir', default=os.path.join(SCRIPT_DIR, 'examples', 'regression_inputs'))
    parser.add_argument('--output-root', default='')
    parser.add_argument('--make-fixtures', action='store_true')
    args = parser.parse_args()

    if args.make_fixtures:
        from make_regression_fixtures import main as make_fixtures
        make_fixtures()

    if not os.path.isdir(args.input_dir):
        raise SystemExit(f'Missing input directory: {args.input_dir}')
    output_root = args.output_root or os.path.join('/mnt/f/office-output/regression', datetime.now().strftime('%Y%m%d-%H%M%S'))
    os.makedirs(output_root, exist_ok=True)
    cases = [run_case(path, output_root) for path in discover_cases(args.input_dir)]
    statuses = {case['status'] for case in cases}
    overall = 'fail' if 'fail' in statuses else 'warn' if 'warn' in statuses else 'pass'
    summary = {'run_at': datetime.now().isoformat(timespec='seconds'), 'status': overall, 'output_root': output_root, 'cases': cases}
    summary_json = os.path.join(output_root, 'summary.json')
    summary_md = os.path.join(output_root, 'summary.md')
    write_json(summary_json, summary)
    write_text(summary_md, summary_markdown(summary))
    print(summary_json)
    print(summary_md)
    if overall == 'fail':
        raise SystemExit(1)


if __name__ == '__main__':
    main()
