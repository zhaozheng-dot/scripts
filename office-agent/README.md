# Office Agent

Office Agent is a gated Office-document automation toolkit. Its current primary workflow is generic PPTX to DOCX conversion with user-confirmed plans, source mapping, fidelity ledgers, template rendering, and quality checks.

The design principle is:

```text
Generic core first -> generic templates first -> professional templates as plugins -> examples only validate the architecture
```

Lupin Dental is a validation fixture, not the center of the architecture.

## Current Scope

Implemented and validated:

- Unified `office_agent.py` entry point for generate / convert / modify.
- Real OOXML generation for DOCX / PPTX / XLSX.
- Safe modification of DOCX / PPTX / XLSX with change logs.
- Structured quality reports with `pass` / `warn` / `fail`.
- Regression runner and synthetic PPTX fixture set.
- PPTX preflight analysis.
- Confirmable conversion plan generation.
- Explicit user-confirmation gate.
- PPTX content extraction and source map output.
- Fidelity ledger output.
- Generic DOCX rendering templates.
- Professional plugin template registration.
- Basic quality checks.
- Unified two-stage CLI entry point.

The unified entry points are:

```text
office_agent.py   # generate / convert / modify
office_convert.py # focused PPTX -> DOCX compatibility entry point
```

## Two-Stage Gated Workflow

Office Agent must not generate final DOCX output before the user confirms the conversion plan.

Correct workflow:

```text
preflight -> plan -> user confirmation -> extract -> source map -> fidelity ledger -> render -> quality check -> delivery
```

Wrong workflow:

```text
preflight -> generate multiple finished versions -> ask user to choose afterward
```

### Stage 1: Generate A Plan Only

Run this first. It creates preflight and plan files, but does not generate a DOCX.

```bash
python3 office_convert.py plan input.pptx --mode generic_reading --fidelity F2
```

Optional parameters:

```bash
python3 office_convert.py plan input.pptx \
  --workspace /mnt/f/office-output/extracted/my-case \
  --mode generic_visual_report \
  --fidelity F2 \
  --include-images \
  --output /mnt/f/office-output/word/my-output.docx
```

Generated files usually include:

```text
*-preflight.json
*-plan.json
*-plan.md
```

The plan includes `recommendation_reasons`, which explain why a mode, template, fidelity level, or manual-review path was suggested. The user should review `*-plan.md` or `*-plan.json` before continuing.

### Stage 2: Run After Confirmation

After the user confirms the plan, run:

```bash
python3 office_convert.py run plan.json --confirm
```

Optional parameters:

```bash
python3 office_convert.py run plan.json \
  --confirm \
  --template generic_reading \
  --output /mnt/f/office-output/word/final.docx \
  --include-images false
```

Generated files include:

```text
*-plan-confirmed.json
*-source-map.json
*-fidelity-ledger.json
*-fidelity-ledger.md
*.docx
*-quality-report.json
```

## Unified Agent CLI

Generate real Office files after plan confirmation:

```bash
python3 office_agent.py generate docx schemas/docx.example.json --confirm
python3 office_agent.py generate pptx schemas/pptx.example.json --confirm
python3 office_agent.py generate xlsx schemas/xlsx.example.json --confirm
```

Convert PPTX to DOCX through the same gated interface:

```bash
python3 office_agent.py convert pptx-to-docx input.pptx --confirm
```

Modify existing files without overwriting the source:

```bash
python3 office_agent.py modify input.docx schemas/modify.example.json --confirm
python3 office_agent.py modify input.pptx schemas/modify.example.json --confirm
python3 office_agent.py modify input.xlsx schemas/modify.example.json --confirm
```

Richer modify operations include `add_toc` / `set_heading_style` for DOCX, `set_theme` for PPTX, and `insert_formula` / `style_header` for XLSX. See `schemas/modify.rich.example.json`.

Run regression fixtures:

```bash
python3 make_regression_fixtures.py
python3 run_regression.py
```

Run service-level regression cases through HTTP:

```bash
python3 run_service_regression.py --make-fixtures
```

The service regression suite reads sanitized case files from `examples/service_regression_cases/`, creates synthetic business fixtures in `examples/service_regression_inputs/`, runs success, cancellation, and expected-failure cases through the HTTP API, then writes `summary.json` and `summary.md` under `/mnt/f/office-output/service-regression/`.

Run HTTP service for Hermes / Operit:

```bash
python3 office_service.py --host 127.0.0.1 --port 8765
```

Run MCP-style stdio bridge:

```bash
python3 office_mcp_server.py
```

Run Hermes-style service client validation:

```bash
python3 office_service_client.py --sample generate --auto-confirm
python3 office_service_client.py --sample convert --auto-confirm
python3 office_service_client.py --sample modify --auto-confirm
```

The service writes per-task audit events to `events.jsonl` and exposes them at `GET /office/tasks/{task_id}/events`. It also supports cancellation through `POST /office/cancel` or `POST /office/tasks/{task_id}/cancel`.

See `HERMES_OPERIT_PROTOCOL.md` for request schemas, task lifecycle, and upper-layer responsibilities.

## Conversion Modes

| Mode | Name | Default template | Default fidelity | Use case |
|---|---|---|---|---|
| M1 | `generic_raw` | `generic_raw` | F1 | Raw transcript, archive, audit, content preservation |
| M2 | `generic_reading` | `generic_reading` | F1/F2 | Preserve source order while improving Word readability |
| M3 | `generic_visual_report` | `generic_visual_report` | F2 | Translate cards, matrices, flows, and visual emphasis into Word report structures |
| M4 | `professional_report` | plugin template | F2/F3 | Use a domain-specific template after explicit confirmation |
| M5 | `editable_material` | `generic_raw` | F1 | Preserve source order and assets for manual editing |

If the user does not explicitly request a professional template, choose a generic mode first.

## Fidelity Levels

| Level | Name | Allowed | Forbidden |
|---|---|---|---|
| F1 | Faithful organization | Heading, paragraph, spacing, basic table cleanup | Fact rewriting, fact merging, body deletion |
| F2 | Light restructuring | Merge repeated headings, reorganize sections, convert cards to tables | Add unsupported facts |
| F3 | Professional rewriting | Rewrite wording, compress content, restructure report logic | Fabricate facts, change risk levels, remove critical sources |

Default rules:

- Low-risk material defaults to F1 or F2.
- Medium/high-risk material defaults to F2 and requires confirmation.
- F3 requires explicit user authorization.

## Templates

Templates are registered in:

```text
template_registry.py
```

### Generic Templates

```text
templates/generic_raw.py
templates/generic_reading.py
templates/generic_visual_report.py
```

Generic templates must not assume a specific domain such as investment review, legal memo, product manual, or project update.

### Professional Plugin Templates

```text
templates/investment_review.py
```

Professional templates are plugins. They may only be used when:

- The plan selects `professional_report`.
- The detected document type matches the plugin contract.
- The user has confirmed the plan.
- The fidelity level is explicit.
- Source map and fidelity ledger are generated.

Professional templates must not invent unsupported facts.

## Important Scripts

| File | Role |
|---|---|
| `office_agent.py` | Unified generate / convert / modify entry point |
| `office_convert.py` | Focused PPTX -> DOCX gated workflow entry point |
| `office_generate.py` | Low-level DOCX / PPTX / XLSX generation renderers |
| `office_modify.py` | Low-level DOCX / PPTX / XLSX modification operations |
| `run_regression.py` | Batch conversion regression runner |
| `run_service_regression.py` | HTTP service-level generate / convert / modify regression runner |
| `make_regression_fixtures.py` | Synthetic PPTX fixture generator |
| `make_service_regression_fixtures.py` | Synthetic business Office fixture generator |
| `office_service.py` | Dependency-free HTTP service wrapper |
| `office_service_client.py` | Hermes-style HTTP orchestration client |
| `office_mcp_server.py` | Dependency-free MCP-style JSON-RPC stdio bridge |
| `HERMES_OPERIT_PROTOCOL.md` | Hermes / Operit calling contract |
| `office_common.py` | Shared helpers |
| `template_registry.py` | Conversion mode and template registry |
| `pptx_preflight.py` | Lightweight PPTX analysis |
| `convert_plan.py` | User-confirmable plan generation |
| `confirm_plan.py` | Confirmation gate |
| `pptx_extract.py` | Full extraction and source map generation |
| `fidelity_ledger.py` | Fidelity ledger generation |
| `pptx_to_report_docx.py` | Template dispatch and DOCX rendering |
| `pptx_to_docx_raw.py` | Raw transcript compatibility entry point |
| `office_quality_check.py` | Basic quality checks |

## Output Files

| Output | Meaning |
|---|---|
| `preflight.json` | Lightweight file diagnosis and recommended modes |
| `plan.json` | User-confirmable conversion plan |
| `plan.md` | Human-readable plan |
| `plan-confirmed.json` | Confirmed execution plan |
| `source-map.json` | Extracted slide content with source locations |
| `fidelity-ledger.json` | Structured content-handling ledger |
| `fidelity-ledger.md` | Human-readable ledger |
| `*.docx` | Generated Word document |
| `quality-report.json` | Structured technical/content/experience checks with pass/warn/fail |
| `quality-report.md` | Human-readable quality report |
| `change-log.json` | Modify-task change log |
| `events.jsonl` | Per-task service audit event log |
| `summary.json` | Regression summary |
| `summary.md` | Human-readable regression summary |

## Example: Lupin Dental Validation

From WSL:

```bash
cd /mnt/f/scripts/office-agent

python3 office_convert.py plan \
  "/mnt/f/office-input/Lupin Dental (DIGICUTO SAS) 投资评审报告.pptx" \
  --workspace /mnt/f/office-output/extracted/unified-lupin \
  --mode generic_reading \
  --fidelity F2 \
  --output /mnt/f/office-output/word/unified-lupin-reading-validation.docx

python3 office_convert.py run \
  "/mnt/f/office-output/extracted/unified-lupin/Lupin Dental (DIGICUTO SAS) 投资评审报告-plan.json" \
  --confirm
```

Expected result:

```text
Quality warnings: none
```

Fixture outputs are stored in:

```text
examples/lupin_dental/
```

## Calling From Operit / Hermes / OpenCode

Recommended orchestration:

```text
1. Operit receives the user request and source path.
2. Hermes calls `POST /office/plan`.
3. The returned `plan_md` is shown to the user.
4. The user confirms mode, fidelity, image handling, template, and output path.
5. Hermes calls `POST /office/run` with `confirm=true`.
6. Hermes polls `GET /office/tasks/{task_id}` until a terminal status.
7. Hermes can read `GET /office/tasks/{task_id}/events` for audit/debugging.
8. The output, ledger/change log, and quality report are returned to the user.
9. Important decisions and results can be written back to Obsidian.
```

High-risk documents must always stop at the plan step until the user confirms.

## Known Limitations

- SmartArt, grouped shapes, and complex charts may not be fully reconstructed as editable Word structures.
- OCR for image-only tables is not implemented.
- Visual quality checks are heuristic and do not replace human review.
- F3 professional rewriting is risky for investment, legal, medical, and financial documents.
- The current real-world validation set is still small.

## Recommended Next Work

1. Define a stricter Operit/Hermes calling contract.
2. Add MCP or HTTP service wrapper after CLI behavior remains stable.
3. Add more real-world regression cases beyond synthetic fixtures.
4. Expand professional plugins only after generic regression stays green.
5. Add richer modify operations such as style normalization, chart insertion, and section-level rewrite plans.

## Development Notes

Before committing changes, run:

```bash
python3 -m py_compile office_common.py template_registry.py pptx_preflight.py pptx_extract.py convert_plan.py confirm_plan.py fidelity_ledger.py pptx_to_docx_raw.py pptx_to_report_docx.py office_quality_check.py office_convert.py office_agent.py office_generate.py office_modify.py office_service.py office_service_client.py office_mcp_server.py run_regression.py run_service_regression.py make_regression_fixtures.py make_service_regression_fixtures.py templates/*.py
```

Then validate at least one full flow:

```bash
python3 office_agent.py generate docx schemas/docx.example.json --confirm
python3 office_agent.py convert pptx-to-docx input.pptx --confirm
python3 office_agent.py modify output.docx schemas/modify.example.json --confirm
python3 run_regression.py
python3 run_service_regression.py --make-fixtures
```
