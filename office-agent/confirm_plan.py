#!/usr/bin/env python3
"""Mark an Office Agent conversion plan as confirmed after user approval."""

import argparse
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_common import read_json, write_json
from template_registry import default_template_for_mode, validate_template


def confirm_plan(plan, template=None, output=None, include_images=None, fidelity=None):
    if fidelity:
        plan['fidelity_level'] = fidelity
    if include_images is not None:
        plan['include_images'] = include_images
    if output:
        plan['output'] = output
    if template:
        plan['template'] = template
    if not plan.get('template'):
        plan['template'] = default_template_for_mode(plan.get('selected_mode'), plan.get('detected_type'))
    validate_template(plan['template'], plan.get('selected_mode'), plan.get('detected_type'), confirmed=True)
    if plan.get('fidelity_level') == 'F3':
        plan['f3_explicitly_authorized'] = True
    plan['confirmed'] = True
    plan['confirmation_note'] = 'Confirmed by user before generation.'
    return plan


def parse_bool(value):
    if value is None:
        return None
    normalized = value.strip().lower()
    if normalized in {'1', 'true', 'yes', 'y', '是'}:
        return True
    if normalized in {'0', 'false', 'no', 'n', '否'}:
        return False
    raise argparse.ArgumentTypeError('Expected true/false or yes/no.')


def main():
    parser = argparse.ArgumentParser(description='Confirm a conversion plan after user approval.')
    parser.add_argument('plan_json')
    parser.add_argument('--output-json', default='')
    parser.add_argument('--template', default='')
    parser.add_argument('--output', default='')
    parser.add_argument('--include-images', type=parse_bool, default=None)
    parser.add_argument('--fidelity', choices=['F1', 'F2', 'F3'])
    args = parser.parse_args()
    plan = read_json(args.plan_json)
    plan = confirm_plan(plan, args.template or None, args.output or None, args.include_images, args.fidelity)
    write_json(args.output_json or args.plan_json, plan)
    print(args.output_json or args.plan_json)


if __name__ == '__main__':
    main()
