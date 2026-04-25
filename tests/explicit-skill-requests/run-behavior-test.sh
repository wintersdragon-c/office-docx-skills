#!/usr/bin/env bash
set -euo pipefail

if [ "$#" -lt 2 ]; then
  echo "Usage: $0 <prompt-file> <expected-skills-csv-or-dash> [forbidden-skills-csv-or-dash] [max-turns] [forbidden-mode]" >&2
  echo "forbidden-mode: not-triggered | not-applied" >&2
  exit 1
fi

prompt_file="$1"
expected_skills="$2"
forbidden_skills="${3:--}"
max_turns="${4:-3}"
forbidden_mode="${5:-not-triggered}"
repo_root="$(cd "$(dirname "$0")/../.." && pwd)"

if ! command -v claude >/dev/null 2>&1; then
  echo "SKIP: claude CLI not found"
  exit 0
fi

prompt="$(cat "$prompt_file")"
output_dir="$(mktemp -d "/tmp/office-docx-skill-behavior.XXXXXX")"
project_dir="$output_dir/project"
mkdir -p "$project_dir"
log_file="$output_dir/claude-output.json"

(
  cd "$project_dir"
  claude_cmd=(
    claude -p "$prompt"
    --plugin-dir "$repo_root"
    --dangerously-skip-permissions
    --max-turns "$max_turns"
    --output-format stream-json
    --verbose
  )
  if command -v timeout >/dev/null 2>&1; then
    timeout 300 "${claude_cmd[@]}" >"$log_file" 2>&1 || true
  elif command -v python3 >/dev/null 2>&1; then
    python3 - "$log_file" "${claude_cmd[@]}" <<'PY' || true
import subprocess
import sys

log_file = sys.argv[1]
cmd = sys.argv[2:]
with open(log_file, "w", encoding="utf-8") as handle:
    try:
        completed = subprocess.run(
            cmd,
            stdout=handle,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=300,
            check=False,
        )
        raise SystemExit(completed.returncode)
    except subprocess.TimeoutExpired:
        handle.write("\nTIMEOUT: claude command exceeded 300 seconds\n")
        raise SystemExit(124)
PY
  else
    "${claude_cmd[@]}" >"$log_file" 2>&1 || true
  fi
)

if grep -q 'API Error: Unable to connect to API' "$log_file"; then
  echo "SKIP: Claude API unavailable"
  echo "Log: $log_file"
  exit 0
fi

assistant_text_matches() {
  local pattern="$1"
  python3 - "$log_file" "$pattern" <<'PY'
import json
import re
import sys

log_file = sys.argv[1]
pattern = re.compile(sys.argv[2], re.IGNORECASE)

with open(log_file, encoding="utf-8", errors="ignore") as handle:
    for line in handle:
        try:
            event = json.loads(line)
        except json.JSONDecodeError:
            continue
        if event.get("type") != "assistant":
            continue
        message = event.get("message") or {}
        for item in message.get("content") or []:
            if item.get("type") != "text":
                continue
            if pattern.search(item.get("text") or ""):
                raise SystemExit(0)
raise SystemExit(1)
PY
}

if [ "$expected_skills" != "-" ]; then
  IFS=',' read -ra expected <<< "$expected_skills"
  for skill_name in "${expected[@]}"; do
    skill_pattern="\"skill\":\"([^\"]*:)?${skill_name}\""
    if grep -q '"name":"Skill"' "$log_file" && grep -qE "$skill_pattern" "$log_file"; then
      echo "PASS: Skill '$skill_name' was triggered"
    else
      echo "FAIL: Skill '$skill_name' was not triggered"
      echo "Log: $log_file"
      exit 1
    fi
  done
fi

if [ "$forbidden_skills" != "-" ]; then
  IFS=',' read -ra forbidden <<< "$forbidden_skills"
  for skill_name in "${forbidden[@]}"; do
    skill_pattern="\"skill\":\"([^\"]*:)?${skill_name}\""
    if grep -qE "$skill_pattern" "$log_file"; then
      if [ "$forbidden_mode" = "not-applied" ] && [ "$skill_name" = "word-default-formatting" ]; then
        reject_pattern="default formatting skill explicitly says not to apply|set that aside|will not apply.{0,120}default|won't apply.{0,120}default|do not apply.{0,120}default|default profile.{0,80}(does not|doesn't|should not).{0,80}apply"
        if assistant_text_matches "$reject_pattern"; then
          echo "PASS: Forbidden skill '$skill_name' was checked and default formatting was rejected"
          continue
        fi
      fi
      echo "FAIL: Forbidden skill '$skill_name' was triggered"
      echo "Log: $log_file"
      exit 1
    fi
    echo "PASS: Forbidden skill '$skill_name' was not triggered"
  done
fi

if [ "$expected_skills" != "-" ]; then
  first_skill_line="$(grep -n '"name":"Skill"' "$log_file" | head -1 | cut -d: -f1 || true)"
  if [ -n "$first_skill_line" ]; then
    premature_tools="$(head -n "$first_skill_line" "$log_file" | grep '"type":"tool_use"' | grep -v '"name":"Skill"' || true)"
    if [ -n "$premature_tools" ]; then
      echo "FAIL: tool invocation occurred before requested skill"
      echo "$premature_tools" | head -5
      exit 1
    fi
  fi
fi
