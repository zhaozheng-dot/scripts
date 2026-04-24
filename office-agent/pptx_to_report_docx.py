#!/usr/bin/env python3
"""Render a DOCX from extracted PPTX data and a confirmed conversion plan."""

import argparse
import importlib
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_common import read_json, refuse_overwrite_input, unique_output
from template_registry import default_template_for_mode, validate_template

DEFAULT_TEMPLATE = 'generic_reading'


def load_template(name):
    module_name = f'templates.{name}'
    return importlib.import_module(module_name)


def choose_template(args_template, plan):
    if args_template:
        return args_template
    return plan.get('template') or default_template_for_mode(plan.get('selected_mode'), plan.get('detected_type')) or DEFAULT_TEMPLATE


def render_from_plan(extracted, plan, template_name=None, output=None, require_confirmed=False):
    if require_confirmed and not plan.get('confirmed'):
        raise SystemExit('Plan is not confirmed. Set confirmed=true in plan JSON after user approval.')
    template_name = template_name or choose_template('', plan)
    try:
        validate_template(template_name, plan.get('selected_mode'), plan.get('detected_type'), plan.get('confirmed', False))
    except ValueError as exc:
        raise SystemExit(str(exc))
    output = output or plan['output']
    refuse_overwrite_input(extracted.get('source', ''), output)
    output = unique_output(output)
    template = load_template(template_name)
    return template.render(extracted, plan, output)


def main():
    parser = argparse.ArgumentParser(description='Render generic or plugin DOCX from extracted PPTX data.')
    parser.add_argument('extract_json')
    parser.add_argument('plan_json')
    parser.add_argument('--template', default='')
    parser.add_argument('--output', default='')
    parser.add_argument('--require-confirmed', action='store_true')
    args = parser.parse_args()
    extracted = read_json(args.extract_json)
    plan = read_json(args.plan_json)
    print(render_from_plan(extracted, plan, args.template or None, args.output or None, args.require_confirmed))


if __name__ == '__main__':
    main()
