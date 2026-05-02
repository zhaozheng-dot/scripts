#!/usr/bin/env python3
"""HTTP service wrapper for Office Agent."""

import argparse
import json
import os
import sys
import threading
import traceback
import uuid
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_agent import (
    build_generate_plan,
    build_modify_plan,
    run_generate_plan,
    run_modify_plan,
)
from office_common import read_json, write_json
from office_convert import create_plan as create_convert_plan
from office_convert import run_confirmed_plan

DEFAULT_TASK_ROOT = os.environ.get('OFFICE_AGENT_TASK_ROOT', '/mnt/f/office-output/office-service/tasks')
TASK_LOCK = threading.Lock()
TASK_THREADS = {}
TERMINAL_STATUSES = {'succeeded', 'failed', 'cancelled'}


def now_iso():
    return datetime.now().isoformat(timespec='seconds')


def normalize_path(path):
    if not path or os.name != 'posix':
        return path
    if len(path) >= 3 and path[1:3] == ':\\':
        drive = path[0].lower()
        rest = path[3:].replace('\\', '/')
        return f'/mnt/{drive}/{rest}'
    return path


def task_root():
    os.makedirs(DEFAULT_TASK_ROOT, exist_ok=True)
    return DEFAULT_TASK_ROOT


def new_task_id():
    return datetime.now().strftime('%Y%m%d-%H%M%S') + '-' + uuid.uuid4().hex[:8]


def task_dir(task_id):
    return os.path.join(task_root(), task_id)


def task_path(task_id):
    return os.path.join(task_dir(task_id), 'task.json')


def events_path(task_id):
    return os.path.join(task_dir(task_id), 'events.jsonl')


def append_event(task_id, event, details=None):
    record = {'ts': now_iso(), 'event': event, 'details': details or {}}
    os.makedirs(task_dir(task_id), exist_ok=True)
    with TASK_LOCK:
        with open(events_path(task_id), 'a', encoding='utf-8') as f:
            f.write(json.dumps(record, ensure_ascii=False) + '\n')


def read_events(task_id, limit=None):
    path = events_path(task_id)
    if not os.path.exists(path):
        return []
    with open(path, 'r', encoding='utf-8') as f:
        rows = [json.loads(line) for line in f if line.strip()]
    return rows[-limit:] if limit else rows


def save_task(task):
    with TASK_LOCK:
        write_json(task_path(task['task_id']), task)


def load_task(task_id):
    path = task_path(task_id)
    if not os.path.exists(path):
        return None
    return read_json(path)


def update_task(task_id, **changes):
    task = load_task(task_id)
    if not task:
        return None
    task.update(changes)
    task['updated_at'] = now_iso()
    save_task(task)
    return task


def public_task(task):
    return {k: v for k, v in task.items() if k != 'traceback'}


def error_response(code, message, details=None):
    return {'ok': False, 'error': {'code': code, 'message': message, 'details': details or {}}}


def make_plan_record(body):
    task_id = new_task_id()
    root = task_dir(task_id)
    os.makedirs(root, exist_ok=True)
    task_type = body.get('task_type')
    if task_type not in {'generate', 'convert', 'modify'}:
        raise ValueError('task_type must be generate, convert, or modify')

    if task_type == 'generate':
        kind = body.get('kind')
        if kind not in {'docx', 'pptx', 'xlsx'}:
            raise ValueError('generate.kind must be docx, pptx, or xlsx')
        spec_path = normalize_path(body.get('spec_path', ''))
        if not spec_path:
            spec = body.get('spec')
            if not isinstance(spec, dict):
                raise ValueError('generate requires spec_path or inline spec object')
            spec_path = os.path.join(root, 'request.json')
            write_json(spec_path, spec)
        plan_path, plan_md, plan = build_generate_plan(kind, spec_path, workspace=root, output=normalize_path(body.get('output', '')) or None)
    elif task_type == 'convert':
        conversion = body.get('conversion', 'pptx-to-docx')
        if conversion != 'pptx-to-docx':
            raise ValueError('Only conversion=pptx-to-docx is currently supported')
        input_path = normalize_path(body.get('input', ''))
        if not input_path:
            raise ValueError('convert.input is required')
        paths, plan = create_convert_plan(
            input_path,
            mode=body.get('mode') or None,
            fidelity=body.get('fidelity') or None,
            include_images=bool(body.get('include_images', False)),
            output=normalize_path(body.get('output', '')) or None,
            workspace=root,
        )
        plan_path = paths['plan']
        plan_md = paths['plan_md']
    else:
        input_path = normalize_path(body.get('input', ''))
        if not input_path:
            raise ValueError('modify.input is required')
        instruction_path = normalize_path(body.get('instruction_path', ''))
        if not instruction_path:
            instruction = body.get('instruction')
            if not isinstance(instruction, dict):
                raise ValueError('modify requires instruction_path or inline instruction object')
            instruction_path = os.path.join(root, 'instruction.json')
            write_json(instruction_path, instruction)
        plan_path, plan_md, plan = build_modify_plan(input_path, instruction_path, workspace=root, output=normalize_path(body.get('output', '')) or None)

    task = {
        'task_id': task_id,
        'task_type': task_type,
        'status': 'planned',
        'created_at': now_iso(),
        'updated_at': now_iso(),
        'workspace': root,
        'events': events_path(task_id),
        'plan_json': plan_path,
        'plan_md': plan_md,
        'output': plan.get('output'),
        'quality_report': plan.get('quality_report'),
        'quality_markdown': plan.get('quality_markdown') or plan.get('quality_report', '').replace('.json', '.md'),
        'requires_user_confirmation': plan.get('requires_user_confirmation', True),
        'confirmed': False,
        'cancel_requested': False,
        'result': None,
        'error': None,
    }
    save_task(task)
    append_event(task_id, 'planned', {'task_type': task_type, 'plan_json': plan_path, 'plan_md': plan_md})
    return task


def run_task(task_id, confirm=False):
    task = load_task(task_id)
    if not task:
        raise ValueError(f'Unknown task_id: {task_id}')
    if not confirm:
        task = update_task(task_id, status='failed', confirmed=False, error={'message': 'confirm=true is required to run an Office Agent task', 'type': 'ValueError'})
        append_event(task_id, 'run_rejected', {'reason': 'confirm=true is required'})
        return task
    if task.get('cancel_requested'):
        task = update_task(task_id, status='cancelled', confirmed=False)
        append_event(task_id, 'cancelled', {'stage': 'before_run'})
        return task
    if task.get('status') in TERMINAL_STATUSES:
        append_event(task_id, 'run_skipped', {'status': task.get('status')})
        return task

    task = update_task(task_id, status='running', confirmed=True, error=None)
    append_event(task_id, 'running', {'confirmed': True})
    try:
        if task['task_type'] == 'generate':
            result = run_generate_plan(task['plan_json'], confirm=True)
        elif task['task_type'] == 'modify':
            result = run_modify_plan(task['plan_json'], confirm=True)
        else:
            result = run_confirmed_plan(task['plan_json'], confirm=True)

        task = load_task(task_id) or task
        if task.get('cancel_requested'):
            task['status'] = 'cancelled'
            task['error'] = {'message': 'Task completed after cancellation request; artifacts may exist.', 'type': 'CancelledAfterRun'}
            append_event(task_id, 'cancelled_after_run', {'result': result})
        else:
            task['status'] = 'succeeded'
            task['result'] = result
            if result.get('output'):
                task['output'] = result.get('output')
            if result.get('docx'):
                task['output'] = result.get('docx')
            append_event(task_id, 'succeeded', {'result': result})
    except Exception as exc:
        task = load_task(task_id) or task
        task['status'] = 'failed'
        task['error'] = {'message': str(exc), 'type': type(exc).__name__}
        task['traceback'] = traceback.format_exc()
        append_event(task_id, 'failed', task['error'])
    task['updated_at'] = now_iso()
    save_task(task)
    return task


def start_task(task_id, confirm=False):
    task = load_task(task_id)
    if not task:
        return None
    if task.get('status') in TERMINAL_STATUSES:
        append_event(task_id, 'run_skipped', {'status': task.get('status')})
        return task
    if task.get('cancel_requested'):
        task = update_task(task_id, status='cancelled', confirmed=False)
        append_event(task_id, 'cancelled', {'stage': 'before_queue'})
        return task
    task = update_task(task_id, status='queued')
    append_event(task_id, 'queued', {'confirm': confirm})
    thread = threading.Thread(target=run_task, args=(task_id, confirm), daemon=True)
    TASK_THREADS[task_id] = thread
    thread.start()
    return task


def cancel_task(task_id):
    task = load_task(task_id)
    if not task:
        return None
    if task.get('status') in TERMINAL_STATUSES:
        append_event(task_id, 'cancel_ignored', {'status': task.get('status')})
        return task
    status = 'cancelled' if task.get('status') in {'planned', 'queued'} else task.get('status')
    task['cancel_requested'] = True
    task['status'] = status
    task['updated_at'] = now_iso()
    save_task(task)
    append_event(task_id, 'cancel_requested', {'status': status})
    return task


class OfficeHandler(BaseHTTPRequestHandler):
    server_version = 'OfficeAgentHTTP/1.1'

    def log_message(self, fmt, *args):
        print(f'{self.address_string()} - {fmt % args}', file=sys.stderr)

    def send_json(self, status, payload):
        data = json.dumps(payload, ensure_ascii=False, indent=2).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def read_body(self):
        length = int(self.headers.get('Content-Length', '0'))
        if length <= 0:
            return {}
        raw = self.rfile.read(length).decode('utf-8')
        return json.loads(raw) if raw else {}

    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/health':
            self.send_json(200, {'ok': True, 'service': 'office-agent', 'status': 'ready', 'task_root': task_root()})
            return
        if path.startswith('/office/tasks/'):
            parts = path.strip('/').split('/')
            task_id = parts[2] if len(parts) >= 3 else ''
            task = load_task(task_id)
            if not task:
                self.send_json(404, error_response('not_found', 'Task not found'))
                return
            if len(parts) == 4 and parts[3] == 'events':
                self.send_json(200, {'ok': True, 'task_id': task_id, 'events': read_events(task_id)})
                return
            self.send_json(200, {'ok': True, 'task': public_task(task)})
            return
        self.send_json(404, error_response('not_found', 'Unknown endpoint'))

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            body = self.read_body()
            if path == '/office/plan':
                task = make_plan_record(body)
                self.send_json(200, {'ok': True, 'task': public_task(task)})
                return
            if path == '/office/run':
                task_id = body.get('task_id')
                if not task_id:
                    raise ValueError('task_id is required')
                task = load_task(task_id)
                if not task:
                    self.send_json(404, error_response('not_found', 'Task not found'))
                    return
                if task.get('status') in {'queued', 'running'}:
                    self.send_json(409, error_response('already_running', 'Task is already queued or running'))
                    return
                task = start_task(task_id, confirm=bool(body.get('confirm', False)))
                status = 202 if task.get('status') == 'queued' else 200
                self.send_json(status, {'ok': True, 'task': public_task(task)})
                return
            if path == '/office/cancel':
                task_id = body.get('task_id')
                if not task_id:
                    raise ValueError('task_id is required')
                task = cancel_task(task_id)
                if not task:
                    self.send_json(404, error_response('not_found', 'Task not found'))
                    return
                self.send_json(200, {'ok': True, 'task': public_task(task)})
                return
            parts = path.strip('/').split('/')
            if len(parts) == 4 and parts[:2] == ['office', 'tasks'] and parts[3] == 'cancel':
                task = cancel_task(parts[2])
                if not task:
                    self.send_json(404, error_response('not_found', 'Task not found'))
                    return
                self.send_json(200, {'ok': True, 'task': public_task(task)})
                return
            self.send_json(404, error_response('not_found', 'Unknown endpoint'))
        except Exception as exc:
            self.send_json(400, error_response('bad_request', str(exc), {'type': type(exc).__name__}))


def main():
    parser = argparse.ArgumentParser(description='Run Office Agent HTTP service.')
    parser.add_argument('--host', default='127.0.0.1')
    parser.add_argument('--port', type=int, default=8765)
    args = parser.parse_args()
    server = ThreadingHTTPServer((args.host, args.port), OfficeHandler)
    print(f'Office Agent HTTP service listening on http://{args.host}:{args.port}')
    server.serve_forever()


if __name__ == '__main__':
    main()
