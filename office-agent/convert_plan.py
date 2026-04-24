#!/usr/bin/env python3
"""Create a user-confirmable conversion plan from preflight data."""

import argparse
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_common import read_json, safe_stem, write_json, write_text
from template_registry import (
    MODE_REGISTRY,
    default_fidelity_for_mode,
    default_template_for_mode,
    mode_code,
    mode_label,
    normalize_mode,
    recommended_modes,
)


def default_output(base, mode):
    suffix = mode_label(mode)
    return f'/mnt/f/office-output/word/{base}-{suffix}.docx'


def choose_default_mode(preflight):
    source_modes = [normalize_mode(mode) for mode in preflight.get('recommended_modes', [])]
    for mode in source_modes:
        if mode and mode in MODE_REGISTRY and mode != 'professional_report':
            return mode
    modes = recommended_modes(preflight)
    if preflight.get('text_density') in {'medium', 'high'}:
        return 'generic_reading'
    return modes[0]


def make_plan(preflight, mode=None, fidelity=None, include_images=False):
    base = safe_stem(preflight['file'])
    mode = normalize_mode(mode) or choose_default_mode(preflight)
    risk_level = preflight.get('risk_level')
    detected_type = preflight.get('detected_type')
    fidelity = fidelity or default_fidelity_for_mode(mode, risk_level)
    template = default_template_for_mode(mode, detected_type)
    requires_confirmation = True if risk_level in {'medium', 'high'} else preflight.get('requires_confirmation', True)
    if mode == 'professional_report':
        requires_confirmation = True
    plan = {
        'source': preflight['file'],
        'detected_type': detected_type,
        'risk_level': risk_level,
        'selected_mode': mode,
        'mode_code': mode_code(mode),
        'mode_label': mode_label(mode),
        'template': template,
        'candidate_modes': recommended_modes(preflight),
        'fidelity_level': fidelity,
        'include_images': include_images,
        'generate_source_map': True,
        'generate_fidelity_ledger': True,
        'requires_user_confirmation': requires_confirmation,
        'confirmed': False,
        'output': default_output(base, mode),
        'extract_output': f'/mnt/f/office-output/extracted/{base}-source-map.json',
        'ledger_json': f'/mnt/f/office-output/extracted/{base}-fidelity-ledger.json',
        'ledger_md': f'/mnt/f/office-output/extracted/{base}-fidelity-ledger.md',
        'quality_report': f'/mnt/f/office-output/extracted/{base}-quality-report.json',
        'allowed_operations': allowed_operations(mode, fidelity),
        'warnings': build_warnings(preflight, mode, fidelity, template),
    }
    return plan


def allowed_operations(mode, fidelity):
    ops = ['do_not_overwrite_source', 'preserve_sources', 'preserve_disclaimer', 'generate_source_map']
    if mode == 'generic_raw':
        ops += ['preserve_slide_order', 'minimal_rewrite']
    elif mode == 'generic_reading':
        ops += ['preserve_slide_order', 'improve_heading_hierarchy', 'convert_obvious_tables']
    elif mode == 'generic_visual_report':
        ops += ['convert_cards_to_tables', 'create_summary_cards', 'keep_page_appendix']
    elif mode == 'professional_report':
        ops += ['use_user_confirmed_business_template', 'restructure_sections', 'convert_cards_to_tables']
    elif mode == 'editable_material':
        ops += ['preserve_assets', 'preserve_slide_order', 'minimal_rewrite']

    if fidelity == 'F1':
        ops += ['no_rewrite', 'no_fact_merge']
    elif fidelity == 'F2':
        ops += ['merge_duplicates', 'light_rewording', 'no_new_facts']
    elif fidelity == 'F3':
        ops += ['professional_rewrite', 'no_unsupported_facts', 'requires_explicit_user_authorization']
    return ops


def build_warnings(preflight, mode, fidelity, template):
    warnings = list(preflight.get('warnings', []))
    if mode == 'professional_report':
        warnings.append('Professional template is a plugin path; use only after user confirms document type and template.')
    if mode == 'professional_report' and not template:
        warnings.append('No plugin template is registered for the detected document type; choose a generic mode or add a plugin.')
    if fidelity == 'F3':
        warnings.append('F3 professional rewriting must be explicitly authorized; do not add unsupported facts.')
    if preflight.get('risk_level') in {'medium', 'high'}:
        warnings.append('Medium/high-risk material requires user confirmation before DOCX generation.')
    return list(dict.fromkeys(warnings))


def candidate_markdown(plan):
    lines = []
    for mode in plan.get('candidate_modes', []):
        data = MODE_REGISTRY.get(mode, {})
        lines.append(f"- {data.get('code', '')} `{mode}`：{data.get('name', mode)}；默认模板 `{data.get('template') or 'plugin'}`；{data.get('description', '')}")
    return '\n'.join(lines) or '- None'


def plan_markdown(plan):
    warnings = '\n'.join(f'- {w}' for w in plan.get('warnings', [])) or '- None'
    return f"""# Office Conversion Plan

Source: `{plan['source']}`
Detected type: `{plan['detected_type']}`
Risk level: `{plan['risk_level']}`

## Candidate Modes

{candidate_markdown(plan)}

## Recommended Execution

- Mode: `{plan['mode_code']} / {plan['selected_mode']} / {plan['mode_label']}`
- Template: `{plan['template']}`
- Fidelity: `{plan['fidelity_level']}`
- Include images: `{plan['include_images']}`
- Source map: `{plan['generate_source_map']}`
- Fidelity ledger: `{plan['generate_fidelity_ledger']}`
- Requires confirmation: `{plan['requires_user_confirmation']}`

## Output

- DOCX: `{plan['output']}`
- Source map: `{plan['extract_output']}`
- Ledger JSON: `{plan['ledger_json']}`
- Ledger Markdown: `{plan['ledger_md']}`
- Quality report: `{plan['quality_report']}`

## Allowed Operations

{chr(10).join(f'- {op}' for op in plan['allowed_operations'])}

## Warnings

{warnings}

## Confirmation Required

Confirm the selected mode, fidelity level, image handling, template, and output path before generation.
"""


def main():
    parser = argparse.ArgumentParser(description='Create a confirmable conversion plan.')
    parser.add_argument('preflight_json')
    parser.add_argument('plan_json')
    parser.add_argument('--plan-md', default='')
    parser.add_argument('--mode', choices=list(MODE_REGISTRY.keys()) + ['raw_transcript', 'reading_layout'])
    parser.add_argument('--fidelity', choices=['F1', 'F2', 'F3'])
    parser.add_argument('--include-images', action='store_true')
    args = parser.parse_args()
    preflight = read_json(args.preflight_json)
    plan = make_plan(preflight, args.mode, args.fidelity, args.include_images)
    write_json(args.plan_json, plan)
    if args.plan_md:
        write_text(args.plan_md, plan_markdown(plan))
    print(args.plan_json)


if __name__ == '__main__':
    main()
