#!/bin/bash
# Scripts 仓库自动同步
# 用法: bash scripts-sync.sh [commit message]

SCRIPTS_DIR="/mnt/f/scripts"
DEFAULT_MSG="sync: update scripts"
MSG="${1:-$DEFAULT_MSG}"

cd "$SCRIPTS_DIR" || exit 1

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
