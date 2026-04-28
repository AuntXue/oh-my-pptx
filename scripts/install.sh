#!/usr/bin/env bash
set -euo pipefail

usage() {
  cat <<'EOF'
Usage:
  ./scripts/install.sh --all
  ./scripts/install.sh --claude
  ./scripts/install.sh --codex

Copies this skill folder into ~/.claude/skills/oh-my-pptx and/or
~/.codex/skills/oh-my-pptx.
EOF
}

install_one() {
  local target_root="$1"
  local source_dir="$2"
  local target_dir="${target_root}/oh-my-pptx"

  mkdir -p "$target_root"
  rm -rf "$target_dir"
  mkdir -p "$target_dir"
  tar \
    --exclude='.git' \
    --exclude='examples/*.pptx' \
    --exclude='examples/*.task.json' \
    --exclude='examples/image-*.png' \
    -cf - -C "$source_dir" . | tar -xf - -C "$target_dir"
  printf 'Installed oh-my-pptx to %s\n' "$target_dir"
}

main() {
  if [[ $# -eq 0 ]]; then
    usage
    exit 1
  fi

  local script_dir
  local skill_dir
  script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
  skill_dir="$(cd "${script_dir}/.." && pwd)"

  local do_claude=0
  local do_codex=0

  while [[ $# -gt 0 ]]; do
    case "$1" in
      --all)
        do_claude=1
        do_codex=1
        ;;
      --claude)
        do_claude=1
        ;;
      --codex)
        do_codex=1
        ;;
      --help|-h)
        usage
        exit 0
        ;;
      *)
        printf 'Unknown option: %s\n' "$1" >&2
        usage
        exit 1
        ;;
    esac
    shift
  done

  if [[ "$do_claude" -eq 1 ]]; then
    install_one "${HOME}/.claude/skills" "$skill_dir"
  fi
  if [[ "$do_codex" -eq 1 ]]; then
    install_one "${HOME}/.codex/skills" "$skill_dir"
  fi
}

main "$@"
