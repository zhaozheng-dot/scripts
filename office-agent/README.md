# Office Agent

Office Agent is a gated Office-document automation toolkit. Its current primary workflow is generic PPTX to DOCX conversion with user-confirmed plans, source mapping, fidelity ledgers, template rendering, and quality checks.

The design principle is:

```text
Generic core first -> generic templates first -> professional templates as plugins -> examples only validate the architecture
```

Lupin Dental is a validation fixture, not the center of the architecture.

## Current Scope

Implemented and validated:

- PPTX preflight analysis.
- Confirmable conversion plan generation.
- Explicit user-confirmation gate.
- PPTX content extraction and source map output.
- Fidelity ledger output.
- Generic DOCX rendering templates.
- Professional plugin template registration.
- Basic quality checks.
- Unified two-stage CLI entry point.

The unified entry point is:

```text
office_convert.py
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

The user should review `*-plan.md` or `*-plan.json` before continuing.

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
| `office_convert.py` | Unified gated workflow entry point |
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
| `quality-report.json` | Technical/content/experience checks |

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
1. Operit receives the user request and PPTX path.
2. Hermes schedules the long-running conversion task.
3. OpenCode runs `office_convert.py plan`.
4. The plan markdown is shown to the user.
5. The user confirms mode, fidelity, image handling, template, and output path.
6. OpenCode runs `office_convert.py run plan.json --confirm`.
7. The DOCX path, fidelity ledger, and quality report are returned to the user.
8. Important decisions and results can be written back to Obsidian.
```

High-risk documents must always stop at the plan step until the user confirms.

## Known Limitations

- SmartArt, grouped shapes, and complex charts may not be fully reconstructed as editable Word structures.
- OCR for image-only tables is not implemented.
- Visual quality checks are heuristic and do not replace human review.
- F3 professional rewriting is risky for investment, legal, medical, and financial documents.
- The current real-world validation set is still small.

## Recommended Next Work

1. Add `recommendation_reasons` to conversion plans.
2. Upgrade `office_quality_check.py` from warnings to `pass/warn/fail`.
3. Add `run_regression.py` for batch validation.
4. Add minimum test fixtures: text-only, visual-mix, high-density.
5. Define a stricter Operit/Hermes calling contract.
6. Add more professional plugins only after generic regression is stable.

## Development Notes

Before committing changes, run:

```bash
python -m py_compile office_common.py template_registry.py pptx_preflight.py pptx_extract.py convert_plan.py confirm_plan.py fidelity_ledger.py pptx_to_docx_raw.py pptx_to_report_docx.py office_quality_check.py office_convert.py templates/*.py
```

Then validate at least one full flow:

```bash
python3 office_convert.py plan input.pptx --mode generic_reading --fidelity F2
python3 office_convert.py run plan.json --confirm
```
