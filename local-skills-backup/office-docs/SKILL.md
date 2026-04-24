---
name: office-docs
description: Generate and inspect Word, Excel, and PowerPoint-style documents from structured specs with safe confirmation boundaries.
version: 0.1.0
author: Operit-Hermes Cluster
metadata:
  hermes:
    tags: [office, docx, xlsx, pptx, documents, productivity]
    related_skills: [obsidian-write, git-workflow]
---

# Office Docs

## Trigger
Use this skill when the user asks to generate, modify, inspect, or summarize Word, Excel, or PowerPoint documents.

## Current MVP
The current implementation is dependency-free and writes HTML documents with Office file extensions. These files are suitable for basic opening/editing in Office-compatible tools, but they are not full OOXML packages yet. Upgrade to python-docx, openpyxl, and python-pptx when Python package installation is available.

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
python3 /mnt/f/scripts/office-agent/office_generate.py docx spec.json output.docx
python3 /mnt/f/scripts/office-agent/office_generate.py xlsx spec.json output.xlsx
python3 /mnt/f/scripts/office-agent/office_generate.py pptx spec.json output.pptx
python3 /mnt/f/scripts/office-agent/office_extract.py input.docx output.json
```

## Workflow
1. Convert the user request into a structured JSON spec.
2. Show the outline or sheet/slide plan before generating important documents.
3. Generate into `/mnt/f/office-output/`, never into an Obsidian repo by default.
4. Read back or extract text to verify basic content.
5. Report output paths.
6. Do not commit binary/Office outputs unless the user explicitly asks.

## Safety Rules
- Never overwrite an existing Office file; write a timestamped copy instead.
- Do not execute macros.
- Do not open unknown Office files in an automated GUI session.
- Do not store API keys, private documents, or sensitive personal data in Git by default.
- Commit scripts, schemas, and templates; avoid committing generated Office binaries.
- `git push` requires explicit user confirmation.

## Limitations
- MVP output has basic formatting only.
- Complex PowerPoint layouts, animations, charts, formulas, and tracked changes require later OOXML or Office COM support.
- Excel formulas are not calculated by this MVP.
