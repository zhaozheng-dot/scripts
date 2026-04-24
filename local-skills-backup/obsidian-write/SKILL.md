---
name: obsidian-write
description: Safely create or append Obsidian Markdown notes in the knowledge vault without deleting, overwriting, or auto-pushing changes.
version: 1.0.0
author: Operit-Hermes Cluster
metadata:
  hermes:
    tags: [obsidian, write, markdown, knowledge, safety]
    related_skills: [obsidian-operations, git-workflow]
---

# Obsidian Write

## Trigger
Use this skill when the user asks to record, archive, save, append, or write information into Obsidian or the knowledge vault.

## Scope
This skill only covers safe Markdown writes to the local Obsidian knowledge vault. It does not commit, push, delete, rename, or reorganize repository files unless the user explicitly requests a separate Git or refactor operation.

## Paths
- Knowledge vault: `/mnt/f/obsidian_repository/knowledge_repo/`
- Daily logs: `/mnt/f/obsidian_repository/knowledge_repo/AI/Daily-Log/`
- Learned notes: `/mnt/f/obsidian_repository/knowledge_repo/AI/Learned/`
- Project notes: `/mnt/f/obsidian_repository/knowledge_repo/AI/Projects/`
- Git workflow reference: `/home/alex/.hermes/skills/dev-workflow/git-workflow/SKILL.md`

## Required Write Policy
1. Never write outside `/mnt/f/obsidian_repository/knowledge_repo/`.
2. Never modify `.obsidian/`, binary files, archives, or generated cache files.
3. Never delete existing notes.
4. Never overwrite an existing note unless the user explicitly says overwrite and the old content has been read first.
5. Prefer append mode for existing notes and create mode for new notes.
6. Do not run `git commit`, `git push`, or auto-commit scripts from this skill.
7. If a target path is ambiguous, ask the user before writing.

## Write Workflow
1. Classify the note type: `daily-log`, `learned`, `project`, or `analysis`.
2. Resolve the target directory inside the knowledge vault.
3. Check whether the target file already exists.
4. For a new file, create Markdown with YAML frontmatter.
5. For an existing file, append a timestamped section.
6. Read back the written file or appended tail to verify the write.
7. Report the file path and whether the operation was create or append.

## Frontmatter Template
```yaml
---
date: YYYY-MM-DD
type: learned
tags: [obsidian, knowledge]
generated_by: hermes
source: conversation
---
```

## Append Template
```markdown

## YYYY-MM-DD HH:mm - Update

Content goes here.
```

## Safety Checks
Before writing, confirm all of these are true:
- The resolved path starts with `/mnt/f/obsidian_repository/knowledge_repo/`.
- The resolved path ends with `.md`.
- The resolved path does not contain `/.obsidian/`.
- The operation is either create-new or append-existing.

## Failure Behavior
If any safety check fails, stop and report the reason. Do not attempt automatic path correction that could write to the wrong location.
