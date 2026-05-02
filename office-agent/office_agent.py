#!/usr/bin/env python3
"""Unified Office Agent entry point for generate, convert, and modify tasks."""

import argparse
import os
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_common import read_json, safe_stem, unique_output, write_json, write_text
from office_convert import create_plan as create_convert_plan
from office_convert import run_confirmed_plan
from office_generate import (
    render_docx_ooxml,
    render_pptx_ooxml,
    render_xlsx_ooxml,
)
from office_modify import modify_docx, modify_pptx, modify_xlsx
from office_quality_check import check as quality_check
from office_quality_check import markdown as quality_markdown


def now_stamp():
    return datetime.now().strftime('%Y%m%d-%H%M%S')


def workspace_for(task_type, name, workspace=None):
    if workspace:
        root = workspace
    else:
        root = os.path.join('/mnt/f/office-output/agent-runs', f'{task_type}-{safe_stem(name)}-{now_stamp()}')
    os.makedirs(root, exist_ok=True)
    return root


def default_generate_output(kind, spec, workspace):
    title = safe_stem(spec.get('title', f'generated-{kind}'))
    return os.path.join(workspace, f'{title}.{kind}')


def default_modify_output(input_path, workspace):
    stem, ext = os.path.splitext(os.path.basename(input_path))
    return os.path.join(workspace, f'{stem}-modified{ext}')


def plan_markdown(plan):
    lines = [
        '# Office Agent Plan',
        '',
        f"- Task type: `{plan.get('task_type')}`",
        f"- Kind: `{plan.get('kind', '')}`",
        f"- Risk level: `{plan.get('risk_level')}`",
        f"- Requires confirmation: `{plan.get('requires_user_confirmation')}`",
        f"- Output: `{plan.get('output')}`",
        '',
        '## Steps',
        '',
    ]
    for step in plan.get('steps', []):
        lines.append(f'- {step}')
    lines.extend(['', '## Quality Artifacts', ''])
    for key in ['quality_report', 'quality_markdown', 'change_log']:
        if plan.get(key):
            lines.append(f'- {key}: `{plan[key]}`')
    lines.extend(['', '## Warnings', ''])
    warnings = plan.get('warnings', [])
    if warnings:
        for warning in warnings:
            lines.append(f'- {warning}')
    else:
        lines.append('- None')
    return '\n'.join(lines) + '\n'


def build_generate_plan(kind, spec_path, workspace=None, output=None):
    spec = read_json(spec_path)
    root = workspace_for('generate', spec.get('title', kind), workspace)
    output = output or default_generate_output(kind, spec, root)
    plan = {
        'schema_version': '1.0',
        'task_type': 'generate',
        'kind': kind,
        'source_spec': spec_path,
        'title': spec.get('title', ''),
        'risk_level': spec.get('risk_level', 'low'),
        'requires_user_confirmation': True,
        'confirmed': False,
        'output': output,
        'quality_report': os.path.join(root, 'quality.json'),
        'quality_markdown': os.path.join(root, 'quality.md'),
        'steps': [
            'Read structured request JSON.',
            f'Render real OOXML {kind.upper()} file.',
            'Run structured quality check.',
            'Write quality report without overwriting source files.',
        ],
        'warnings': [],
    }
    plan_path = os.path.join(root, 'plan.json')
    plan_md = os.path.join(root, 'plan.md')
    write_json(plan_path, plan)
    write_text(plan_md, plan_markdown(plan))
    return plan_path, plan_md, plan


def run_generate_plan(plan_path, confirm=False):
    plan = read_json(plan_path)
    if confirm:
        plan['confirmed'] = True
    if not plan.get('confirmed'):
        raise SystemExit('Plan is not confirmed. Re-run with --confirm after user approval.')
    spec = read_json(plan['source_spec'])
    renderers = {'docx': render_docx_ooxml, 'pptx': render_pptx_ooxml, 'xlsx': render_xlsx_ooxml}
    renderers[plan['kind']](spec, plan['output'])
    write_json(plan_path, plan)
    quality = quality_check(plan['output'], plan_json=plan_path, risk_level=plan.get('risk_level'))
    write_json(plan['quality_report'], quality)
    write_text(plan['quality_markdown'], quality_markdown(quality))
    return {'output': plan['output'], 'quality': plan['quality_report'], 'quality_md': plan['quality_markdown'], 'status': quality.get('status')}


def build_modify_plan(input_path, spec_path, workspace=None, output=None):
    spec = read_json(spec_path)
    ext = os.path.splitext(input_path)[1].lower().lstrip('.')
    root = workspace_for('modify', input_path, workspace)
    output = output or default_modify_output(input_path, root)
    operations = spec.get('operations', [])
    plan = {
        'schema_version': '1.0',
        'task_type': 'modify',
        'kind': ext,
        'source': input_path,
        'instruction_spec': spec_path,
        'risk_level': spec.get('risk_level', 'low'),
        'requires_user_confirmation': True,
        'confirmed': False,
        'output': output,
        'change_log': os.path.join(root, 'change-log.json'),
        'quality_report': os.path.join(root, 'quality.json'),
        'quality_markdown': os.path.join(root, 'quality.md'),
        'steps': [
            'Read source Office file.',
            f'Apply {len(operations)} requested operations to a new output file.',
            'Write change log.',
            'Run structured quality check.',
        ],
        'operations': operations,
        'warnings': ['The source file will not be overwritten.'],
    }
    plan_path = os.path.join(root, 'plan.json')
    plan_md = os.path.join(root, 'plan.md')
    write_json(plan_path, plan)
    write_text(plan_md, plan_markdown(plan))
    return plan_path, plan_md, plan


def run_modify_plan(plan_path, confirm=False):
    plan = read_json(plan_path)
    if confirm:
        plan['confirmed'] = True
    if not plan.get('confirmed'):
        raise SystemExit('Plan is not confirmed. Re-run with --confirm after user approval.')
    if os.path.abspath(plan['source']) == os.path.abspath(plan['output']):
        raise SystemExit('Refusing to overwrite the input file; choose a different output path.')
    if os.path.exists(plan['output']):
        plan['output'] = unique_output(plan['output'])

    spec = read_json(plan['instruction_spec'])
    ext = os.path.splitext(plan['source'])[1].lower()
    if ext == '.docx':
        modify_docx(plan['source'], spec, plan['output'])
    elif ext == '.pptx':
        modify_pptx(plan['source'], spec, plan['output'])
    elif ext == '.xlsx':
        modify_xlsx(plan['source'], spec, plan['output'])
    else:
        raise SystemExit(f'Unsupported input extension: {ext}')

    changes = []
    for index, op in enumerate(spec.get('operations', []), start=1):
        changes.append({'index': index, 'operation': op.get('op'), 'details': op})
    change_log = {
        'source': plan['source'],
        'output': plan['output'],
        'changed_at': datetime.now().isoformat(timespec='seconds'),
        'changes': changes,
    }
    write_json(plan['change_log'], change_log)
    write_json(plan_path, plan)
    quality = quality_check(plan['output'], plan_json=plan_path, change_log_json=plan['change_log'], risk_level=plan.get('risk_level'))
    write_json(plan['quality_report'], quality)
    write_text(plan['quality_markdown'], quality_markdown(quality))
    return {
        'output': plan['output'],
        'change_log': plan['change_log'],
        'quality': plan['quality_report'],
        'quality_md': plan['quality_markdown'],
        'status': quality.get('status'),
    }


def print_plan(plan_path, plan_md, plan):
    print('Plan generated. Review before running.')
    print(f'Plan JSON: {plan_path}')
    print(f'Plan Markdown: {plan_md}')
    print(f"Task: {plan.get('task_type')} {plan.get('kind')} -> {plan.get('output')}")
    print(f"Run after confirmation: python3 office_agent.py run {plan_path} --confirm")


def print_result(result):
    print('Task completed.')
    for key, value in result.items():
        print(f'{key}: {value}')


def main():
    parser = argparse.ArgumentParser(description='Office Agent unified entry point.')
    sub = parser.add_subparsers(dest='command', required=True)

    gen = sub.add_parser('generate', help='Plan a generate task.')
    gen.add_argument('kind', choices=['docx', 'pptx', 'xlsx'])
    gen.add_argument('spec')
    gen.add_argument('--workspace', default='')
    gen.add_argument('--output', default='')
    gen.add_argument('--confirm', action='store_true', help='Plan and run immediately after explicit user approval.')

    conv = sub.add_parser('convert', help='Plan a PPTX to DOCX conversion task.')
    conv.add_argument('conversion', choices=['pptx-to-docx'])
    conv.add_argument('input')
    conv.add_argument('--workspace', default='')
    conv.add_argument('--output', default='')
    conv.add_argument('--mode', default='')
    conv.add_argument('--fidelity', choices=['F1', 'F2', 'F3'])
    conv.add_argument('--include-images', action='store_true')
    conv.add_argument('--confirm', action='store_true', help='Plan and run immediately after explicit user approval.')

    mod = sub.add_parser('modify', help='Plan a modify task.')
    mod.add_argument('input')
    mod.add_argument('spec')
    mod.add_argument('--workspace', default='')
    mod.add_argument('--output', default='')
    mod.add_argument('--confirm', action='store_true', help='Plan and run immediately after explicit user approval.')

    run = sub.add_parser('run', help='Run a confirmed Office Agent plan.')
    run.add_argument('plan_json')
    run.add_argument('--confirm', action='store_true')

    args = parser.parse_args()

    if args.command == 'generate':
        plan_path, plan_md, plan = build_generate_plan(args.kind, args.spec, args.workspace or None, args.output or None)
        print_plan(plan_path, plan_md, plan)
        if args.confirm:
            print_result(run_generate_plan(plan_path, confirm=True))
    elif args.command == 'convert':
        paths, plan = create_convert_plan(
            args.input,
            mode=args.mode or None,
            fidelity=args.fidelity,
            include_images=args.include_images,
            output=args.output or None,
            workspace=args.workspace or None,
        )
        print('Plan generated. Review before running conversion.')
        print(f"Plan JSON: {paths['plan']}")
        print(f"Plan Markdown: {paths['plan_md']}")
        print(f"Run after confirmation: python3 office_agent.py run {paths['plan']} --confirm")
        if args.confirm:
            print_result(run_confirmed_plan(paths['plan'], confirm=True))
    elif args.command == 'modify':
        plan_path, plan_md, plan = build_modify_plan(args.input, args.spec, args.workspace or None, args.output or None)
        print_plan(plan_path, plan_md, plan)
        if args.confirm:
            print_result(run_modify_plan(plan_path, confirm=True))
    elif args.command == 'run':
        plan = read_json(args.plan_json)
        task_type = plan.get('task_type')
        if task_type == 'generate':
            print_result(run_generate_plan(args.plan_json, confirm=args.confirm))
        elif task_type == 'modify':
            print_result(run_modify_plan(args.plan_json, confirm=args.confirm))
        else:
            print_result(run_confirmed_plan(args.plan_json, confirm=args.confirm))


if __name__ == '__main__':
    main()
