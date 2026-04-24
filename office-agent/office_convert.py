#!/usr/bin/env python3
"""Unified gated PPTX to DOCX conversion entry point."""

import argparse
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from confirm_plan import confirm_plan
from convert_plan import make_plan, plan_markdown
from fidelity_ledger import ledger_rows, markdown as ledger_markdown
from office_common import read_json, safe_stem, write_json, write_text
from office_quality_check import check as quality_check
from pptx_extract import extract
from pptx_preflight import preflight
from pptx_to_report_docx import render_from_plan


def default_workspace(input_path, workspace=None):
    if workspace:
        return workspace
    return os.path.join('/mnt/f/office-output/extracted', safe_stem(input_path))


def paths_for(input_path, workspace=None):
    base = safe_stem(input_path)
    root = default_workspace(input_path, workspace)
    return {
        'workspace': root,
        'preflight': os.path.join(root, f'{base}-preflight.json'),
        'plan': os.path.join(root, f'{base}-plan.json'),
        'plan_md': os.path.join(root, f'{base}-plan.md'),
        'confirmed_plan': os.path.join(root, f'{base}-plan-confirmed.json'),
        'extract': os.path.join(root, f'{base}-source-map.json'),
        'assets_dir': os.path.join(root, 'assets'),
        'ledger_json': os.path.join(root, f'{base}-fidelity-ledger.json'),
        'ledger_md': os.path.join(root, f'{base}-fidelity-ledger.md'),
        'quality': os.path.join(root, f'{base}-quality-report.json'),
    }


def update_plan_paths(plan, paths, output=None):
    plan['extract_output'] = paths['extract']
    plan['ledger_json'] = paths['ledger_json']
    plan['ledger_md'] = paths['ledger_md']
    plan['quality_report'] = paths['quality']
    if output:
        plan['output'] = output
    return plan


def create_plan(input_path, mode=None, fidelity=None, include_images=False, output=None, workspace=None):
    paths = paths_for(input_path, workspace)
    preflight_data = preflight(input_path)
    plan = make_plan(preflight_data, mode=mode, fidelity=fidelity, include_images=include_images)
    plan = update_plan_paths(plan, paths, output=output)
    write_json(paths['preflight'], preflight_data)
    write_json(paths['plan'], plan)
    write_text(paths['plan_md'], plan_markdown(plan))
    return paths, plan


def run_confirmed_plan(plan_path, confirm=False, template=None, output=None, include_images=None, fidelity=None, assets=False):
    plan = read_json(plan_path)
    if confirm:
        plan = confirm_plan(
            plan,
            template=template or None,
            output=output or None,
            include_images=include_images,
            fidelity=fidelity,
        )
    elif not plan.get('confirmed'):
        raise SystemExit('Plan is not confirmed. Re-run with --confirm after user approval, or pass a confirmed plan JSON.')

    input_path = plan['source']
    paths = paths_for(input_path, os.path.dirname(plan_path))
    plan = update_plan_paths(plan, paths, output=output or plan.get('output'))
    confirmed_plan_path = paths['confirmed_plan'] if confirm else plan_path
    write_json(confirmed_plan_path, plan)

    extracted = extract(input_path, paths['assets_dir'] if assets or plan.get('include_images') else None)
    write_json(paths['extract'], extracted)

    rows = ledger_rows(extracted, plan)
    write_json(paths['ledger_json'], {'source': extracted.get('source'), 'rows': rows})
    write_text(paths['ledger_md'], ledger_markdown(rows))

    docx_path = render_from_plan(extracted, plan, require_confirmed=True)
    quality = quality_check(docx_path, paths['extract'], paths['ledger_json'])
    write_json(paths['quality'], quality)
    return {
        'plan': confirmed_plan_path,
        'source_map': paths['extract'],
        'ledger_json': paths['ledger_json'],
        'ledger_md': paths['ledger_md'],
        'docx': docx_path,
        'quality': paths['quality'],
        'quality_warnings': quality.get('warnings', []),
    }


def print_plan_result(paths, plan):
    print('Plan generated. Review before running conversion.')
    print(f"Preflight: {paths['preflight']}")
    print(f"Plan JSON: {paths['plan']}")
    print(f"Plan Markdown: {paths['plan_md']}")
    print(f"Recommended: {plan.get('mode_code')} {plan.get('selected_mode')} / template={plan.get('template')} / fidelity={plan.get('fidelity_level')}")
    reasons = plan.get('recommendation_reasons', [])
    if reasons:
        print('Reasons:')
        for reason in reasons[:4]:
            print(f'- {reason}')
    print('Run after confirmation:')
    print(f"  python3 office_convert.py run {paths['plan']} --confirm")


def print_run_result(result):
    print('Conversion completed.')
    print(f"Confirmed plan: {result['plan']}")
    print(f"Source map: {result['source_map']}")
    print(f"Fidelity ledger: {result['ledger_md']}")
    print(f"DOCX: {result['docx']}")
    print(f"Quality report: {result['quality']}")
    if result['quality_warnings']:
        print('Quality warnings:')
        for warning in result['quality_warnings']:
            print(f'- {warning}')
    else:
        print('Quality warnings: none')


def main():
    parser = argparse.ArgumentParser(description='Unified gated PPTX to DOCX conversion workflow.')
    sub = parser.add_subparsers(dest='command', required=True)

    plan_parser = sub.add_parser('plan', help='Generate preflight and user-confirmable plan only.')
    plan_parser.add_argument('input')
    plan_parser.add_argument('--workspace', default='')
    plan_parser.add_argument('--mode', choices=['generic_raw', 'generic_reading', 'generic_visual_report', 'professional_report', 'editable_material', 'raw_transcript', 'reading_layout'])
    plan_parser.add_argument('--fidelity', choices=['F1', 'F2', 'F3'])
    plan_parser.add_argument('--include-images', action='store_true')
    plan_parser.add_argument('--output', default='')

    run_parser = sub.add_parser('run', help='Run conversion from a confirmed or explicitly confirmed plan.')
    run_parser.add_argument('plan_json')
    run_parser.add_argument('--confirm', action='store_true', help='Mark the plan confirmed in this command after user approval.')
    run_parser.add_argument('--template', default='')
    run_parser.add_argument('--output', default='')
    run_parser.add_argument('--include-images', choices=['true', 'false'], default='')
    run_parser.add_argument('--fidelity', choices=['F1', 'F2', 'F3'])
    run_parser.add_argument('--assets', action='store_true', help='Extract image assets even when they are not embedded.')

    args = parser.parse_args()
    if args.command == 'plan':
        paths, plan = create_plan(
            args.input,
            mode=args.mode,
            fidelity=args.fidelity,
            include_images=args.include_images,
            output=args.output or None,
            workspace=args.workspace or None,
        )
        print_plan_result(paths, plan)
    elif args.command == 'run':
        include_images = None
        if args.include_images:
            include_images = args.include_images == 'true'
        result = run_confirmed_plan(
            args.plan_json,
            confirm=args.confirm,
            template=args.template or None,
            output=args.output or None,
            include_images=include_images,
            fidelity=args.fidelity,
            assets=args.assets,
        )
        print_run_result(result)


if __name__ == '__main__':
    main()
