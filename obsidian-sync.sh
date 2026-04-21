#!/bin/bash
# Obsidian 知识库 Git 同步脚本
# 用法: bash obsidian-sync.sh [commit message]

REPO_DIR="/mnt/f/obsidian_repository/scienc-project-repo"
DEFAULT_MSG="sync: update knowledge base"
MSG="${1:-$DEFAULT_MSG}"

cd "$REPO_DIR" || exit 1

STATUS=$(git status --short)
if [ -z "$STATUS" ]; then
  echo "No changes to commit."
  exit 0
fi

echo "=== Changes ==="
echo "$STATUS"
echo ""

git add -A
git commit -m "$MSG"

echo ""
echo "=== Pushing to GitHub ==="
git push origin main 2>&1

echo ""
echo "=== Done ==="
git log --oneline -1
