#!/usr/bin/env python3
"""Template and conversion-mode registry for Office Agent."""

MODE_REGISTRY = {
    'generic_raw': {
        'code': 'M1',
        'name': 'generic raw transcript',
        'template': 'generic_raw',
        'default_fidelity_low': 'F1',
        'default_fidelity_high': 'F1',
        'description': 'Maximize content preservation and source traceability.',
    },
    'generic_reading': {
        'code': 'M2',
        'name': 'generic reading layout',
        'template': 'generic_reading',
        'default_fidelity_low': 'F1',
        'default_fidelity_high': 'F2',
        'description': 'Preserve source order while improving Word readability.',
    },
    'generic_visual_report': {
        'code': 'M3',
        'name': 'generic visual report',
        'template': 'generic_visual_report',
        'default_fidelity_low': 'F2',
        'default_fidelity_high': 'F2',
        'description': 'Translate cards, matrices, and visual emphasis into native Word structures.',
    },
    'professional_report': {
        'code': 'M4',
        'name': 'professional plugin report',
        'template': None,
        'default_fidelity_low': 'F2',
        'default_fidelity_high': 'F2',
        'description': 'Use a user-confirmed business plugin template.',
    },
    'editable_material': {
        'code': 'M5',
        'name': 'editable source material',
        'template': 'generic_raw',
        'default_fidelity_low': 'F1',
        'default_fidelity_high': 'F1',
        'description': 'Keep source order and assets for manual editing.',
    },
}

TEMPLATE_REGISTRY = {
    'generic_raw': {
        'type': 'generic',
        'supported_modes': ['generic_raw', 'editable_material'],
        'document_types': ['*'],
        'requires_confirmation': False,
        'default_fidelity': 'F1',
    },
    'generic_reading': {
        'type': 'generic',
        'supported_modes': ['generic_reading'],
        'document_types': ['*'],
        'requires_confirmation': False,
        'default_fidelity': 'F1/F2',
    },
    'generic_visual_report': {
        'type': 'generic',
        'supported_modes': ['generic_visual_report'],
        'document_types': ['*'],
        'requires_confirmation': False,
        'default_fidelity': 'F2',
    },
    'investment_review': {
        'type': 'plugin',
        'supported_modes': ['professional_report'],
        'document_types': ['investment_review'],
        'requires_confirmation': True,
        'default_fidelity': 'F2',
        'requires_source_map': True,
        'requires_fidelity_ledger': True,
    },
}

LEGACY_MODE_ALIASES = {
    'raw_transcript': 'generic_raw',
    'reading_layout': 'generic_reading',
}

DOCUMENT_TYPE_PLUGIN = {
    'investment_review': 'investment_review',
}


def normalize_mode(mode):
    if not mode:
        return mode
    return LEGACY_MODE_ALIASES.get(mode, mode)


def mode_label(mode):
    data = MODE_REGISTRY.get(normalize_mode(mode), {})
    return data.get('name', mode)


def mode_code(mode):
    data = MODE_REGISTRY.get(normalize_mode(mode), {})
    return data.get('code', '')


def default_template_for_mode(mode, detected_type=None):
    mode = normalize_mode(mode)
    if mode == 'professional_report':
        return DOCUMENT_TYPE_PLUGIN.get(detected_type)
    data = MODE_REGISTRY.get(mode, {})
    return data.get('template')


def default_fidelity_for_mode(mode, risk_level=None):
    data = MODE_REGISTRY.get(normalize_mode(mode), {})
    if risk_level in {'medium', 'high'}:
        return data.get('default_fidelity_high', 'F2')
    return data.get('default_fidelity_low', 'F1')


def recommended_modes(preflight):
    density = preflight.get('text_density')
    risk = preflight.get('risk_level')
    detected_type = preflight.get('detected_type')
    modes = ['generic_raw', 'generic_reading']
    if density in {'medium', 'high'} or preflight.get('group_shapes') or preflight.get('smartart_like_shapes'):
        modes.append('generic_visual_report')
    if detected_type in DOCUMENT_TYPE_PLUGIN and risk in {'medium', 'high'}:
        modes.append('professional_report')
    modes.append('editable_material')
    return list(dict.fromkeys(modes))


def validate_template(template_name, mode, detected_type=None, confirmed=False):
    template = TEMPLATE_REGISTRY.get(template_name)
    if not template:
        raise ValueError(f'Unknown template: {template_name}')
    mode = normalize_mode(mode)
    if mode not in template['supported_modes']:
        raise ValueError(f'Template {template_name} does not support mode {mode}')
    doc_types = template.get('document_types', ['*'])
    if '*' not in doc_types and detected_type not in doc_types:
        raise ValueError(f'Template {template_name} does not support detected_type={detected_type}')
    if template.get('type') == 'plugin' and not confirmed:
        raise ValueError(f'Plugin template {template_name} requires confirmed=true in the plan')
    return True
