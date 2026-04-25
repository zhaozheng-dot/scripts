#!/usr/bin/env python3
"""Run Office Agent service-level regression cases through the HTTP API."""

import argparse
import json
import os
import subprocess
import sys
import time
from datetime import datetime
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_common import write_json, write_text
from office_generate import render_docx_ooxml, render_pptx_ooxml, render_xlsx_ooxml
import make_service_regression_fixtures

TERMINAL_STATUSES = {'succeeded', 'failed', 'cancelled'}
REQUIRED_SUCCESS_EVENTS = ['planned', 'queued', 'running', 'succeeded']


def read_json(path):
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)


def request_json(base_url, method, path, body=None, timeout=30):
    data = None
    headers = {}
    if body is not None:
        data = json.dumps(body, ensure_ascii=False).encode('utf-8')
        headers['Content-Type'] = 'application/json'
    req = Request(base_url.rstrip('/') + path, data=data, headers=headers, method=method)
    try:
        with urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode('utf-8'))
    except HTTPError as exc:
        payload = exc.read().decode('utf-8')
        try:
            return json.loads(payload)
        except Exception:
            raise RuntimeError(payload) from exc


def wait_for_service(base_url, timeout=15):
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        try:
            payload = request_json(base_url, 'GET', '/health', timeout=3)
            if payload.get('ok'):
                return True
        except (URLError, RuntimeError, TimeoutError, OSError) as exc:
            last_error = exc
        time.sleep(0.5)
    raise RuntimeError(f'Office service did not become ready: {last_error}')


def wait_task(base_url, task_id, interval=0.5, timeout=180):
    deadline = time.time() + timeout
    while time.time() < deadline:
        payload = request_json(base_url, 'GET', f'/office/tasks/{task_id}')
        task = payload.get('task', {})
        if task.get('status') in TERMINAL_STATUSES:
            return task
        time.sleep(interval)
    raise TimeoutError(f'Task timed out: {task_id}')


def discover_cases(case_dir):
    paths = []
    for name in os.listdir(case_dir):
        if name.lower().endswith('.json'):
            paths.append(os.path.join(case_dir, name))
    return sorted(paths)


def ensure_setup(case):
    setup = case.get('setup')
    if not setup:
        return None
    output = setup.get('output')
    spec = setup.get('spec') or {}
    kind = setup.get('kind')
    if not output or not kind:
        raise ValueError(f"Invalid setup for case {case.get('case_id')}")
    os.makedirs(os.path.dirname(output), exist_ok=True)
    renderers = {'docx': render_docx_ooxml, 'pptx': render_pptx_ooxml, 'xlsx': render_xlsx_ooxml}
    renderers[kind](spec, output)
    return output


def load_quality(path):
    if not path or not os.path.exists(path):
        return {'status': 'missing', 'warnings': ['quality report not found']}
    try:
        return read_json(path)
    except Exception as exc:
        return {'status': 'fail', 'warnings': [str(exc)]}


def error_payload(payload):
    return (payload or {}).get('error') or {}


def error_matches(error, expected):
    if not expected:
        return True
    code = expected.get('code')
    message_contains = expected.get('message_contains')
    if code and error.get('code') != code:
        return False
    if message_contains and message_contains not in error.get('message', ''):
        return False
    return True


def validate_case(task, events):
    checks = []
    event_names = [event.get('event') for event in events]
    output = task.get('output')
    quality = load_quality(task.get('quality_report'))
    quality_status = quality.get('status', 'unknown')

    checks.append({'name': 'task_succeeded', 'ok': task.get('status') == 'succeeded', 'value': task.get('status')})
    checks.append({'name': 'quality_acceptable', 'ok': quality_status in {'pass', 'warn'}, 'value': quality_status})
    checks.append({'name': 'output_exists', 'ok': bool(output and os.path.exists(output)), 'value': output})
    checks.append({'name': 'events_complete', 'ok': all(name in event_names for name in REQUIRED_SUCCESS_EVENTS), 'value': event_names})

    result = task.get('result') or {}
    if task.get('task_type') == 'modify':
        checks.append({'name': 'change_log_exists', 'ok': bool(result.get('change_log') and os.path.exists(result.get('change_log'))), 'value': result.get('change_log')})
    if task.get('task_type') == 'convert':
        checks.append({'name': 'ledger_exists', 'ok': bool(result.get('ledger_json') and os.path.exists(result.get('ledger_json'))), 'value': result.get('ledger_json')})

    status = 'pass' if all(check['ok'] for check in checks) else 'fail'
    if status == 'pass' and quality_status == 'warn':
        status = 'warn'
    return status, quality, checks


def base_case_result(case, case_id, case_path, setup_output=None):
    return {
        'case_id': case_id,
        'description': case.get('description', ''),
        'case_path': case_path,
        'setup_output': setup_output,
        'task_id': None,
        'task_status': None,
        'status': 'fail',
        'quality_status': 'not_applicable',
        'warnings': [],
        'events_count': 0,
        'event_names': [],
        'workspace': None,
        'output': None,
        'quality_report': None,
        'quality_markdown': None,
        'events': None,
        'result': None,
        'error': None,
        'checks': [],
        'expectations': case.get('expectations', []),
    }


def run_case(base_url, case_path):
    case = read_json(case_path)
    case_id = case.get('case_id') or os.path.splitext(os.path.basename(case_path))[0]
    setup_output = ensure_setup(case)
    result = base_case_result(case, case_id, case_path, setup_output)
    expected_failure = case.get('expected_failure') or {}

    planned = request_json(base_url, 'POST', '/office/plan', case['request'])
    if not planned.get('ok'):
        error = error_payload(planned)
        result['error'] = error
        result['checks'] = [
            {'name': 'expected_plan_failure', 'ok': expected_failure.get('stage') == 'plan', 'value': expected_failure.get('stage')},
            {'name': 'error_matches', 'ok': error_matches(error, expected_failure), 'value': error},
        ]
        result['status'] = 'pass' if all(check['ok'] for check in result['checks']) else 'fail'
        return result
    if expected_failure.get('stage') == 'plan':
        result['task_id'] = planned.get('task', {}).get('task_id')
        result['checks'] = [{'name': 'expected_plan_failure', 'ok': False, 'value': 'plan succeeded'}]
        return result

    task_id = planned['task']['task_id']
    result['task_id'] = task_id
    if case.get('cancel_after_plan'):
        cancelled = request_json(base_url, 'POST', '/office/cancel', {'task_id': task_id})
        final = cancelled.get('task', {})
        events_payload = request_json(base_url, 'GET', f'/office/tasks/{task_id}/events')
        events = events_payload.get('events', [])
        event_names = [event.get('event') for event in events]
        checks = [
            {'name': 'cancelled', 'ok': final.get('status') == 'cancelled', 'value': final.get('status')},
            {'name': 'cancel_event', 'ok': 'cancel_requested' in event_names, 'value': event_names},
            {'name': 'not_queued', 'ok': 'queued' not in event_names, 'value': event_names},
        ]
        result.update({
            'task_status': final.get('status'),
            'status': 'pass' if all(check['ok'] for check in checks) else 'fail',
            'events_count': len(events),
            'event_names': event_names,
            'workspace': final.get('workspace'),
            'output': final.get('output'),
            'events': final.get('events'),
            'result': final.get('result'),
            'error': final.get('error'),
            'checks': checks,
        })
        return result

    confirm = bool(case.get('run_confirm', True))
    started = request_json(base_url, 'POST', '/office/run', {'task_id': task_id, 'confirm': confirm})
    if not started.get('ok'):
        error = error_payload(started)
        result['error'] = error
        result['checks'] = [
            {'name': 'expected_run_request_failure', 'ok': expected_failure.get('stage') == 'run_request', 'value': expected_failure.get('stage')},
            {'name': 'error_matches', 'ok': error_matches(error, expected_failure), 'value': error},
        ]
        result['status'] = 'pass' if all(check['ok'] for check in result['checks']) else 'fail'
        return result

    final = wait_task(base_url, task_id)
    events_payload = request_json(base_url, 'GET', f'/office/tasks/{task_id}/events')
    events = events_payload.get('events', [])
    event_names = [event.get('event') for event in events]
    if expected_failure.get('stage') == 'run':
        checks = [
            {'name': 'terminal_status', 'ok': final.get('status') == expected_failure.get('terminal_status'), 'value': final.get('status')},
            {'name': 'expected_event', 'ok': expected_failure.get('event') in event_names, 'value': event_names},
            {'name': 'message_matches', 'ok': expected_failure.get('message_contains', '') in ((final.get('error') or {}).get('message', '')), 'value': final.get('error')},
        ]
        result.update({
            'task_status': final.get('status'),
            'status': 'pass' if all(check['ok'] for check in checks) else 'fail',
            'events_count': len(events),
            'event_names': event_names,
            'workspace': final.get('workspace'),
            'output': final.get('output'),
            'events': final.get('events'),
            'result': final.get('result'),
            'error': final.get('error'),
            'checks': checks,
        })
        return result

    status, quality, checks = validate_case(final, events)
    result.update({
        'task_status': final.get('status'),
        'status': status,
        'quality_status': quality.get('status', 'unknown'),
        'warnings': quality.get('warnings', []),
        'events_count': len(events),
        'event_names': event_names,
        'workspace': final.get('workspace'),
        'output': final.get('output'),
        'quality_report': final.get('quality_report'),
        'quality_markdown': final.get('quality_markdown'),
        'events': final.get('events'),
        'result': final.get('result'),
        'error': final.get('error'),
        'checks': checks,
    })
    return result


def summary_markdown(summary):
    lines = [
        '# Office Agent Service Regression Summary',
        '',
        f"- Run at: `{summary['run_at']}`",
        f"- Status: `{summary['status']}`",
        f"- Total cases: `{len(summary['cases'])}`",
        f"- Output root: `{summary['output_root']}`",
        '',
        '| Case | Status | Quality | Events | Output |',
        '|---|---|---|---:|---|',
    ]
    for case in summary['cases']:
        lines.append(
            f"| {case['case_id']} | `{case['status']}` | `{case['quality_status']}` | {case['events_count']} | `{case.get('output')}` |"
        )
    lines.extend(['', '## Failed Checks', ''])
    any_failed = False
    for case in summary['cases']:
        failed = [check for check in case.get('checks', []) if not check.get('ok')]
        if not failed:
            continue
        any_failed = True
        lines.append(f"### {case['case_id']}")
        for check in failed:
            lines.append(f"- `{check['name']}`: `{check.get('value')}`")
        lines.append('')
    if not any_failed:
        lines.append('- None')
        lines.append('')
    lines.extend(['## Warnings', ''])
    any_warning = False
    for case in summary['cases']:
        for warning in case.get('warnings', []):
            any_warning = True
            lines.append(f"- {case['case_id']}: {warning}")
    if not any_warning:
        lines.append('- None')
    lines.append('')
    return '\n'.join(lines)


def start_service(args):
    cmd = [sys.executable, os.path.join(SCRIPT_DIR, 'office_service.py'), '--host', args.host, '--port', str(args.port)]
    return subprocess.Popen(cmd, cwd=SCRIPT_DIR, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)


def main():
    parser = argparse.ArgumentParser(description='Run Office Agent service regression suite.')
    parser.add_argument('--case-dir', default=os.path.join(SCRIPT_DIR, 'examples', 'service_regression_cases'))
    parser.add_argument('--output-root', default='')
    parser.add_argument('--base-url', default='http://127.0.0.1:8765')
    parser.add_argument('--host', default='127.0.0.1')
    parser.add_argument('--port', type=int, default=8765)
    parser.add_argument('--use-existing-service', action='store_true')
    parser.add_argument('--make-fixtures', action='store_true')
    args = parser.parse_args()

    if args.make_fixtures:
        from make_regression_fixtures import main as make_fixtures
        make_fixtures()
        make_service_regression_fixtures.main([])

    output_root = args.output_root or os.path.join('/mnt/f/office-output/service-regression', datetime.now().strftime('%Y%m%d-%H%M%S'))
    os.makedirs(output_root, exist_ok=True)

    service = None
    if not args.use_existing_service:
        service = start_service(args)
    try:
        wait_for_service(args.base_url)
        cases = []
        for case_path in discover_cases(args.case_dir):
            case_result = run_case(args.base_url, case_path)
            cases.append(case_result)
            print(f"{case_result['case_id']}: {case_result['status']} ({case_result['task_status']})")
        statuses = {case['status'] for case in cases}
        overall = 'fail' if 'fail' in statuses else 'warn' if 'warn' in statuses else 'pass'
        summary = {
            'run_at': datetime.now().isoformat(timespec='seconds'),
            'status': overall,
            'output_root': output_root,
            'case_dir': args.case_dir,
            'cases': cases,
        }
        summary_json = os.path.join(output_root, 'summary.json')
        summary_md = os.path.join(output_root, 'summary.md')
        write_json(summary_json, summary)
        write_text(summary_md, summary_markdown(summary))
        print(summary_json)
        print(summary_md)
        if overall == 'fail':
            raise SystemExit(1)
    finally:
        if service:
            service.terminate()
            try:
                service.wait(timeout=5)
            except subprocess.TimeoutExpired:
                service.kill()


if __name__ == '__main__':
    main()
