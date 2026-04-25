#!/usr/bin/env python3
"""Hermes-style client for Office Agent HTTP service."""

import argparse
import json
import os
import time
from urllib.error import HTTPError
from urllib.request import Request, urlopen


def read_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def request_json(base_url, method, path, body=None):
    data = None
    headers = {}
    if body is not None:
        data = json.dumps(body, ensure_ascii=False).encode('utf-8')
        headers['Content-Type'] = 'application/json'
    req = Request(base_url.rstrip('/') + path, data=data, headers=headers, method=method)
    try:
        with urlopen(req, timeout=30) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except HTTPError as exc:
        payload = exc.read().decode('utf-8')
        try:
            return json.loads(payload)
        except Exception:
            raise RuntimeError(payload) from exc


def print_plan(plan_md):
    if plan_md and os.path.exists(plan_md):
        print('--- PLAN START ---')
        with open(plan_md, 'r', encoding='utf-8') as f:
            print(f.read().strip())
        print('--- PLAN END ---')
    else:
        print(f'Plan markdown not readable: {plan_md}')


def wait_task(base_url, task_id, interval=0.5, timeout=120):
    deadline = time.time() + timeout
    while time.time() < deadline:
        payload = request_json(base_url, 'GET', f'/office/tasks/{task_id}')
        task = payload.get('task', {})
        status = task.get('status')
        print(f'poll task_id={task_id} status={status}')
        if status in {'succeeded', 'failed', 'cancelled'}:
            return task
        time.sleep(interval)
    raise TimeoutError(f'Task timed out: {task_id}')


def run_flow(base_url, plan_body, auto_confirm=False, show_plan=True):
    planned = request_json(base_url, 'POST', '/office/plan', plan_body)
    if not planned.get('ok'):
        raise RuntimeError(planned)
    task = planned['task']
    task_id = task['task_id']
    print(f'planned task_id={task_id} type={task.get("task_type")}')
    if show_plan:
        print_plan(task.get('plan_md'))
    if not auto_confirm:
        print('auto_confirm=false; stop after planning.')
        return task

    started = request_json(base_url, 'POST', '/office/run', {'task_id': task_id, 'confirm': True})
    if not started.get('ok'):
        raise RuntimeError(started)
    final = wait_task(base_url, task_id)
    events = request_json(base_url, 'GET', f'/office/tasks/{task_id}/events')
    print('--- ARTIFACTS ---')
    for key in ['output', 'quality_report', 'quality_markdown', 'plan_json', 'plan_md', 'events']:
        print(f'{key}: {final.get(key)}')
    if final.get('result'):
        print('result:', json.dumps(final['result'], ensure_ascii=False, indent=2))
    print('events_count:', len(events.get('events', [])))
    return final


def sample_body(kind):
    if kind == 'generate':
        return {
            'task_type': 'generate',
            'kind': 'docx',
            'spec': {
                'title': 'Hermes Client Validation',
                'sections': [
                    {'heading': 'Overview', 'paragraphs': ['Generated through office_service_client.py.']},
                    {'heading': 'Result', 'paragraphs': ['This validates plan, confirm, run, poll, and artifact delivery.']},
                ],
            },
        }
    if kind == 'convert':
        return {
            'task_type': 'convert',
            'conversion': 'pptx-to-docx',
            'input': '/mnt/f/scripts/office-agent/examples/regression_inputs/high_density.pptx',
            'fidelity': 'F2',
            'include_images': False,
        }
    if kind == 'modify':
        return {
            'task_type': 'modify',
            'input': '/mnt/f/office-output/agent-validation/generate-docx/四位一体系统简报.docx',
            'instruction_path': '/mnt/f/scripts/office-agent/schemas/modify.example.json',
        }
    raise ValueError(f'Unknown sample: {kind}')


def main():
    parser = argparse.ArgumentParser(description='Hermes-style Office Agent HTTP client.')
    parser.add_argument('--base-url', default='http://127.0.0.1:8765')
    parser.add_argument('--request-json', default='')
    parser.add_argument('--sample', choices=['generate', 'convert', 'modify'], default='')
    parser.add_argument('--auto-confirm', action='store_true')
    parser.add_argument('--no-show-plan', action='store_true')
    args = parser.parse_args()

    if args.request_json:
        body = read_json(args.request_json)
    elif args.sample:
        body = sample_body(args.sample)
    else:
        raise SystemExit('Use --request-json or --sample.')
    run_flow(args.base_url, body, auto_confirm=args.auto_confirm, show_plan=not args.no_show_plan)


if __name__ == '__main__':
    main()
