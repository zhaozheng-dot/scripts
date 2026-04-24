#!/bin/bash
# ~/.hermes/scripts/auto-commit-obsidian.sh
# Obsidian repository sync audit script. Default mode is dry-run; use --push to commit and push.
# Cron disabled on 2026-04-24 after audit found auto-push side effects.

BRIDGE="/mnt/c/Users/zhao/hermes-bridge"
LOG_DIR="$BRIDGE/logs"
FAIL_MARKER="$BRIDGE/data/cron-failed-auto-commit-obsidian.flag"
LOG_FILE="$LOG_DIR/auto-commit-$(date +%Y%m%d).log"
MODE="dry-run"
FAILED=0

if [ "${1:-}" = "--push" ]; then
  MODE="push"
fi

REPOS=(
  "/mnt/f/obsidian_repository/scienc-project-repo"
  "/mnt/f/obsidian_repository/science-tech-repo"
  "/mnt/f/obsidian_repository/science-business-reop"
  "/mnt/f/obsidian_repository/social-life-repo"
)

EXCLUDES=(
  '.obsidian/'
  '*.zip'
)

mkdir -p "$LOG_DIR" "$BRIDGE/data"

log() {
  echo "[$(date '+%Y-%m-%d %H:%M:%S')] $*" >> "$LOG_FILE"
}

status_filtered() {
  git status --porcelain -- . ':!.obsidian' ':!*.zip' 2>/dev/null
}

log "START auto-commit mode=$MODE"

for OBSIDIAN in "${REPOS[@]}"; do
  REPO_NAME=$(basename "$OBSIDIAN")
  if [ ! -d "$OBSIDIAN" ]; then
    log "SKIP missing repo: $OBSIDIAN"
    continue
  fi

  if ! cd "$OBSIDIAN"; then
    log "ERROR cannot enter repo: $OBSIDIAN"
    FAILED=1
    continue
  fi

  STATUS=$(status_filtered)
  if [ -z "$STATUS" ]; then
    log "OK no included changes: $REPO_NAME"
    continue
  fi

  log "CHANGES $REPO_NAME:"
  printf '%s
' "$STATUS" | sed 's/^/  /' >> "$LOG_FILE"

  if [ "$MODE" = "dry-run" ]; then
    echo "[DRY-RUN] $REPO_NAME has included changes:"
    printf '%s
' "$STATUS" | sed 's/^/  /'
    continue
  fi

  log "START push: $REPO_NAME"

  if ! git add -A -- . ':!.obsidian' ':!*.zip' >> "$LOG_FILE" 2>&1; then
    log "ERROR git add failed: $REPO_NAME"
    FAILED=1
    continue
  fi

  if ! git commit -m "auto: sync reviewed changes - $(date '+%Y-%m-%d %H:%M')" >> "$LOG_FILE" 2>&1; then
    log "ERROR git commit failed: $REPO_NAME"
    FAILED=1
    continue
  fi

  if ! git push >> "$LOG_FILE" 2>&1; then
    log "ERROR git push failed: $REPO_NAME"
    FAILED=1
    continue
  fi

  log "OK pushed: $REPO_NAME"
done

if [ "$FAILED" -ne 0 ]; then
  printf '{"last_error":"%s","script":"auto-commit-obsidian.sh","exit_code":1,"mode":"%s"}\n' \
    "$(date '+%Y-%m-%d %H:%M:%S')" "$MODE" > "$FAIL_MARKER"
  log "ERROR completed with failures mode=$MODE"
  exit 1
fi

rm -f "$FAIL_MARKER"
log "OK completed mode=$MODE"
