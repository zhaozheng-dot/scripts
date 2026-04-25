# Service Regression Cases

These JSON files are sanitized service-level regression cases for Office Agent.

Each case contains:

- `case_id`: stable identifier used in summaries.
- `description`: human-readable intent.
- `setup`: optional generated source Office file for modify cases.
- `request`: HTTP `/office/plan` request body.
- `expectations`: manual acceptance hints for review.

Run them with:

```bash
python3 run_service_regression.py --make-fixtures
```

`--make-fixtures` creates both the conversion PPTX fixtures and synthetic business Office files in `examples/service_regression_inputs/`. Cases may be success cases, cancellation cases, or expected-failure cases.

Outputs are written to `/mnt/f/office-output/service-regression/<timestamp>/`.

To add real business examples, copy a sanitized Office file into a safe folder and add a new case JSON that points to it. Do not commit confidential source files.
