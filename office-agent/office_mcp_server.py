#!/usr/bin/env python3
"""MCP-style JSON-RPC stdio bridge for Office Agent.

This lightweight bridge exposes Office Agent operations as tool calls without
requiring external MCP packages. It follows the JSON-RPC request/response shape
used by MCP clients closely enough for local Agent orchestration wrappers.
"""

import json
import os
import sys
import traceback

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

from office_service import make_plan_record, run_task, load_task, public_task

TOOLS = [
    {
        'name': 'office_plan',
        'description': 'Create a confirmable Office Agent plan for generate, convert, or modify.',
        'inputSchema': {
            'type': 'object',
            'properties': {
                'task_type': {'type': 'string', 'enum': ['generate', 'convert', 'modify']},
                'kind': {'type': 'string', 'enum': ['docx', 'pptx', 'xlsx']},
                'conversion': {'type': 'string', 'enum': ['pptx-to-docx']},
                'input': {'type': 'string'},
                'spec': {'type': 'object'},
                'spec_path': {'type': 'string'},
                'instruction': {'type': 'object'},
                'instruction_path': {'type': 'string'},
                'output': {'type': 'string'},
                'mode': {'type': 'string'},
                'fidelity': {'type': 'string', 'enum': ['F1', 'F2', 'F3']},
                'include_images': {'type': 'boolean'},
            },
            'required': ['task_type'],
        },
    },
    {
        'name': 'office_run',
        'description': 'Run a planned Office Agent task. Requires explicit confirm=true.',
        'inputSchema': {
            'type': 'object',
            'properties': {
                'task_id': {'type': 'string'},
                'confirm': {'type': 'boolean'},
            },
            'required': ['task_id', 'confirm'],
        },
    },
    {
        'name': 'office_get_task',
        'description': 'Get Office Agent task status and artifacts.',
        'inputSchema': {
            'type': 'object',
            'properties': {'task_id': {'type': 'string'}},
            'required': ['task_id'],
        },
    },
]


def ok(request_id, result):
    return {'jsonrpc': '2.0', 'id': request_id, 'result': result}


def err(request_id, code, message, data=None):
    payload = {'jsonrpc': '2.0', 'id': request_id, 'error': {'code': code, 'message': message}}
    if data is not None:
        payload['error']['data'] = data
    return payload


def content_payload(data):
    return {'content': [{'type': 'text', 'text': json.dumps(data, ensure_ascii=False, indent=2)}]}


def handle_tool(name, args):
    if name == 'office_plan':
        task = make_plan_record(args or {})
        return content_payload({'ok': True, 'task': public_task(task)})
    if name == 'office_run':
        task_id = (args or {}).get('task_id')
        confirm = bool((args or {}).get('confirm', False))
        if not task_id:
            raise ValueError('task_id is required')
        if not confirm:
            raise ValueError('confirm=true is required')
        run_task(task_id, confirm=True)
        task = load_task(task_id)
        return content_payload({'ok': True, 'task': public_task(task)})
    if name == 'office_get_task':
        task_id = (args or {}).get('task_id')
        task = load_task(task_id)
        if not task:
            raise ValueError(f'Unknown task_id: {task_id}')
        return content_payload({'ok': True, 'task': public_task(task)})
    raise ValueError(f'Unknown tool: {name}')


def handle(request):
    method = request.get('method')
    request_id = request.get('id')
    try:
        if method == 'initialize':
            return ok(request_id, {'protocolVersion': '2024-11-05', 'serverInfo': {'name': 'office-agent', 'version': '1.0.0'}, 'capabilities': {'tools': {}}})
        if method == 'tools/list':
            return ok(request_id, {'tools': TOOLS})
        if method == 'tools/call':
            params = request.get('params') or {}
            return ok(request_id, handle_tool(params.get('name'), params.get('arguments') or {}))
        if method == 'ping':
            return ok(request_id, {})
        return err(request_id, -32601, f'Method not found: {method}')
    except Exception as exc:
        return err(request_id, -32000, str(exc), {'type': type(exc).__name__, 'traceback': traceback.format_exc()})


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            request = json.loads(line.lstrip('\ufeff'))
            response = handle(request)
        except Exception as exc:
            response = err(None, -32700, str(exc))
        sys.stdout.write(json.dumps(response, ensure_ascii=False) + '\n')
        sys.stdout.flush()


if __name__ == '__main__':
    main()
