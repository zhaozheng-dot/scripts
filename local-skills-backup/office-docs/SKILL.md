---
name: office-docs
description: Generate and inspect real Word, Excel, and PowerPoint documents from structured specs with safe confirmation boundaries.
version: 0.2.0
author: Operit-Hermes Cluster
metadata:
  hermes:
    tags: [office, docx, xlsx, pptx, documents, productivity]
    related_skills: [obsidian-write, git-workflow]
---

# Office Docs

## Trigger
Use this skill when the user asks to generate, modify, inspect, or summarize Word, Excel, or PowerPoint documents.

## Current Capability
The current implementation generates real OOXML files when dependencies are available:
- Word: `python-docx` / `python3-docx`
- Excel: `openpyxl` / `python3-openpyxl`
- PowerPoint: `python-pptx`

If a dependency is missing, `office_generate.py --format auto` falls back to HTML-compatible output. Use `--format ooxml` when real Office packages are required.

## Paths
- Scripts: `/mnt/f/scripts/office-agent/`
- Output: `/mnt/f/office-output/`
- Word output: `/mnt/f/office-output/word/`
- Excel output: `/mnt/f/office-output/excel/`
- PPT output: `/mnt/f/office-output/ppt/`
- Extracted text: `/mnt/f/office-output/extracted/`
- Backup: `F:\scripts\local-skills-backup\office-docs\SKILL.md`

## Tools
```bash
python3 /mnt/f/scripts/office-agent/office_generate.py docx spec.json output.docx --format ooxml
python3 /mnt/f/scripts/office-agent/office_generate.py xlsx spec.json output.xlsx --format ooxml
python3 /mnt/f/scripts/office-agent/office_generate.py pptx spec.json output.pptx --format ooxml
python3 /mnt/f/scripts/office-agent/office_extract.py input.docx output.json
python3 /mnt/f/scripts/office-agent/office_modify.py input.docx modify.json output.docx
```

## Workflow
1. Convert the user request into a structured JSON spec.
2. Show the outline or sheet/slide plan before generating important documents.
3. Generate into `/mnt/f/office-output/`, never into an Obsidian repo by default.
4. Extract text or inspect workbook/slide content to verify basic output.
5. Report output paths.
6. Do not commit generated Office outputs unless the user explicitly asks.

## Safety Rules
- Never overwrite an existing Office file; write a timestamped copy instead.
- Do not execute macros.
- Do not open unknown Office files in an automated GUI session.
- Do not store API keys, private documents, or sensitive personal data in Git by default.
- Commit scripts, schemas, and templates; avoid committing generated Office binaries.
- `git push` requires explicit user confirmation.

## Limitations
- Charts, complex themes, animations, tracked changes, and advanced formulas are not fully implemented yet.
- Excel formulas can be written, but are not calculated by Python.
- Existing document modification should extract or copy to a new file first; do not mutate the original.
