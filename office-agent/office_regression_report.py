#!/usr/bin/env python3
"""Create Obsidian-friendly reports from Office Agent regression summaries."""

import argparse
import json
import os
from datetime import datetime

DEFAULT_OBSIDIAN_REPORT = '/mnt/f/obsidian_repository/scienc-project-repo/AI/AGI/opencode/Office Agent 回归报告.md'


def read_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def write_text(path, text):
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(text)


def status_counts(cases):
    counts = {}
    for case in cases:
        status = case.get('status', 'unknown')
        counts[status] = counts.get(status, 0) + 1
    return counts


def classify_failed_check(case, check):
    name = check.get('name', '')
    task_type = (case.get('result') or {}).get('task_type') or ''
    if 'event' in name:
        return 'service_event_lifecycle'
    if 'quality' in name:
        return 'quality_gate'
    if 'output' in name:
        return 'artifact_delivery'
    if 'ledger' in name:
        return 'conversion_fidelity'
    if 'change_log' in name:
        return 'modify_auditability'
    if 'error' in name or 'failure' in name:
        return 'error_contract'
    return task_type or 'unknown'


def failure_roadmap(cases):
    buckets = {}
    for case in cases:
        failed_checks = [check for check in case.get('checks', []) if not check.get('ok')]
        if not failed_checks and case.get('status') not in {'fail'}:
            continue
        for check in failed_checks or [{'name': 'case_status', 'value': case.get('status')}]:
            bucket = classify_failed_check(case, check)
            buckets.setdefault(bucket, []).append({'case_id': case.get('case_id'), 'check': check})
    return buckets


def roadmap_markdown(buckets):
    if not buckets:
        return ['- 当前无失败项；路线图保持为增强项，而非阻断项。']
    lines = []
    priority = {
        'service_event_lifecycle': 'P0 服务生命周期/事件日志',
        'error_contract': 'P0 错误契约',
        'artifact_delivery': 'P1 产物交付',
        'quality_gate': 'P1 质量门控',
        'conversion_fidelity': 'P3 转换保真',
        'modify_auditability': 'P3 修改审计',
    }
    for key, items in sorted(buckets.items()):
        title = priority.get(key, key)
        lines.append(f'- {title}: {len(items)} issue(s)')
        for item in items[:5]:
            check = item['check']
            lines.append(f"  - `{item['case_id']}` / `{check.get('name')}` -> `{check.get('value')}`")
    return lines


def summary_table(cases):
    lines = ['| Case | Status | Quality | Events | Output |', '|---|---|---|---:|---|']
    for case in cases:
        lines.append(
            f"| {case.get('case_id')} | `{case.get('status')}` | `{case.get('quality_status')}` | {case.get('events_count', 0)} | `{case.get('output')}` |"
        )
    return lines


def build_report(service_summary, conversion_summary=None):
    cases = service_summary.get('cases', [])
    counts = status_counts(cases)
    buckets = failure_roadmap(cases)
    lines = [
        '# Office Agent 回归报告',
        '',
        f"> 更新时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        '',
        '## 服务级回归',
        '',
        f"- Summary: `{service_summary.get('output_root')}`",
        f"- Run at: `{service_summary.get('run_at')}`",
        f"- Overall: `{service_summary.get('status')}`",
        f"- Total cases: `{len(cases)}`",
        f"- Status counts: `{json.dumps(counts, ensure_ascii=False)}`",
        '',
        '## Case Matrix',
        '',
    ]
    lines.extend(summary_table(cases))
    lines.extend(['', '## 失败分类路线图', ''])
    lines.extend(roadmap_markdown(buckets))
    if conversion_summary:
        conversion_cases = conversion_summary.get('cases', [])
        lines.extend([
            '',
            '## 转换回归',
            '',
            f"- Summary: `{conversion_summary.get('output_root')}`",
            f"- Run at: `{conversion_summary.get('run_at')}`",
            f"- Overall: `{conversion_summary.get('status')}`",
            f"- Total cases: `{len(conversion_cases)}`",
            '',
            '| Case | Status | Output |',
            '|---|---|---|',
        ])
        for case in conversion_cases:
            lines.append(f"| {case.get('case')} | `{case.get('status')}` | `{case.get('output')}` |")
    lines.extend([
        '',
        '## 下一步建议',
        '',
        '- 若回归失败，优先修复 P0/P1 类问题，再进入 Office 能力增强。',
        '- 当前无失败时，继续推进 P3 Modify 能力：Word 目录、PPT 主题色、Excel 公式。',
        '- 后续拿到真实脱敏材料后，追加到 `examples/service_regression_cases/` 或私有输入目录。',
        '',
    ])
    return '\n'.join(lines)


def main():
    parser = argparse.ArgumentParser(description='Write Office Agent regression report for Obsidian.')
    parser.add_argument('service_summary_json')
    parser.add_argument('--conversion-summary-json', default='')
    parser.add_argument('--output', default=DEFAULT_OBSIDIAN_REPORT)
    args = parser.parse_args()

    service_summary = read_json(args.service_summary_json)
    conversion_summary = read_json(args.conversion_summary_json) if args.conversion_summary_json else None
    report = build_report(service_summary, conversion_summary)
    write_text(args.output, report)
    print(args.output)


if __name__ == '__main__':
    main()
