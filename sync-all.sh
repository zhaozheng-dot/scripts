#!/bin/bash
# 全仓库一键同步
# 用法: bash sync-all.sh [commit message]
#
# 支持的仓库:
#   --obsidian    同步 scienc-project-repo 知识库
#   --scripts     同步 scripts 脚本仓库
#   --all         同步所有仓库
#   无参数        同步所有仓库

MSG="${1:-sync: update}"

# 如果参数是 --flag 形式，MSG 用默认值
if [[ "$1" == --* ]]; then
  MSG="sync: update"
fi

sync_repo() {
  local name="$1"
  local dir="$2"
  local msg="$3"

  echo ""
  echo "=============================="
  echo "  Syncing: $name"
  echo "=============================="

  cd "$dir" || { echo "  [SKIP] Directory not found: $dir"; return 1; }

  STATUS=$(git status --short)
  if [ -z "$STATUS" ]; then
    echo "  No changes."
    return 0
  fi

  echo "  Changes:"
  echo "$STATUS" | sed 's/^/    /'

  git add -A
  git commit -m "$msg"
  git push origin main 2>&1

  echo "  Done: $(git log --oneline -1)"
}

case "$1" in
  --obsidian)
    sync_repo "Obsidian (scienc-project-repo)" "/mnt/f/obsidian_repository/scienc-project-repo" "$MSG"
    ;;
  --scripts)
    sync_repo "Scripts" "/mnt/f/scripts" "$MSG"
    ;;
  --all|""|*)
    sync_repo "Obsidian (scienc-project-repo)" "/mnt/f/obsidian_repository/scienc-project-repo" "$MSG"
    sync_repo "Scripts" "/mnt/f/scripts" "$MSG"
    ;;
esac

echo ""
echo "=== All done ==="
