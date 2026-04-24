#!/usr/bin/env python3
"""Generate a faithful generic raw DOCX transcript from extracted PPTX content."""

import argparse
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_common import read_json, refuse_overwrite_input, unique_output
from templates import generic_raw


def render(extracted, plan, output):
    source = extracted.get('source', '')
    refuse_overwrite_input(source, output)
    output = unique_output(output)
    return generic_raw.render(extracted, plan, output)


def main():
    parser = argparse.ArgumentParser(description='Render generic raw transcript DOCX from extracted PPTX data.')
    parser.add_argument('extract_json')
    parser.add_argument('plan_json')
    parser.add_argument('--output', default='')
    args = parser.parse_args()
    extracted = read_json(args.extract_json)
    plan = read_json(args.plan_json)
    output = args.output or plan['output']
    print(render(extracted, plan, output))


if __name__ == '__main__':
    main()
