#!/usr/bin/env python3
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    print("Missing dependency: PyYAML. Install with: python3 -m pip install pyyaml")
    sys.exit(1)


ROOT = Path(__file__).resolve().parents[2]
EXPECTED_SKILLS = {
    "docx-bilingual-translation",
    "docx-tracked-changes",
    "word-default-formatting",
    "word-formula-writing",
}
EXPECTED_HELPERS = {
    "docx-bilingual-translation": [
        "translation_docx_helpers.py",
    ],
    "docx-tracked-changes": [
        "tracked_change_editor.py",
        "verify_tracked_changes.py",
    ],
    "word-default-formatting": [
        "formatting_helpers.py",
    ],
    "word-formula-writing": [
        "formula_writer.py",
    ],
}
EXPECTED_PRESSURE_PROMPTS = {
    "docx-tracked-changes.txt",
    "multi-all.txt",
    "multi-formatting-formulas.txt",
    "word-default-formatting.txt",
    "word-formula-writing.txt",
    "formula-only-no-formatting.txt",
    "template-no-default-formatting.txt",
    "tracked-formula-combo.txt",
}
LOCAL_PATH_PATTERNS = [
    "/" + "Users/",
    "/" + "home/",
    "wordeq" + "-venv",
    "/" + "Documents/KEVIN/",
]


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    sys.exit(1)


def read_frontmatter(path: Path) -> dict:
    text = path.read_text(encoding="utf-8")
    match = re.match(r"^---\n(.*?)\n---", text, re.DOTALL)
    if not match:
        fail(f"{path} missing YAML frontmatter")
    data = yaml.safe_load(match.group(1))
    if not isinstance(data, dict):
        fail(f"{path} frontmatter is not a mapping")
    return data


def validate_skills() -> None:
    skills_dir = ROOT / "skills"
    if not skills_dir.is_dir():
        fail("skills/ directory missing")
    actual = {path.name for path in skills_dir.iterdir() if path.is_dir() and path.name != "归档"}
    missing = EXPECTED_SKILLS - actual
    if missing:
        fail(f"missing skills: {sorted(missing)}")
    if (skills_dir / "归档").exists():
        fail("skills/归档 should not remain in the package")
    for skill in EXPECTED_SKILLS:
        skill_dir = skills_dir / skill
        skill_md = skill_dir / "SKILL.md"
        if not skill_md.exists():
            fail(f"{skill}/SKILL.md missing")
        frontmatter = read_frontmatter(skill_md)
        if frontmatter.get("name") != skill:
            fail(f"{skill} frontmatter name does not match folder")
        description = frontmatter.get("description")
        if not isinstance(description, str) or not description.startswith("Use when"):
            fail(f"{skill} description must start with 'Use when'")
        openai_yaml = skill_dir / "agents" / "openai.yaml"
        if not openai_yaml.exists():
            fail(f"{skill}/agents/openai.yaml missing")
        ui = yaml.safe_load(openai_yaml.read_text(encoding="utf-8"))
        prompt = ui.get("interface", {}).get("default_prompt", "")
        if f"${skill}" not in prompt:
            fail(f"{openai_yaml} default_prompt must mention ${skill}")
        for helper in EXPECTED_HELPERS[skill]:
            if not (skill_dir / helper).exists():
                fail(f"{skill}/{helper} missing")


def validate_json(path: Path) -> dict:
    if not path.exists():
        fail(f"{path} missing")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        fail(f"{path} invalid JSON: {exc}")


def validate_metadata() -> None:
    codex = validate_json(ROOT / ".codex-plugin" / "plugin.json")
    if codex.get("name") != "office-docx-skills":
        fail("Codex plugin name mismatch")
    if codex.get("skills") != "./skills/":
        fail("Codex plugin skills path must be ./skills/")
    interface = codex.get("interface", {})
    for key in [
        "displayName",
        "shortDescription",
        "longDescription",
        "developerName",
        "category",
        "capabilities",
        "defaultPrompt",
    ]:
        if key not in interface:
            fail(f"Codex interface missing {key}")
    for asset_key in ["composerIcon", "logo"]:
        if asset_key in interface and not (ROOT / interface[asset_key]).exists():
            fail(f"Codex interface {asset_key} points to missing file")

    claude = validate_json(ROOT / ".claude-plugin" / "plugin.json")
    if claude.get("name") != "office-docx-skills":
        fail("Claude plugin name mismatch")
    if not isinstance(claude.get("author"), dict):
        fail("Claude plugin author must be an object")

    marketplace = validate_json(ROOT / ".claude-plugin" / "marketplace.json")
    metadata = marketplace.get("metadata", {})
    if metadata.get("description") != "Development marketplace for Office DOCX Skills":
        fail("Claude marketplace metadata.description mismatch")
    plugins = marketplace.get("plugins", [])
    if not plugins or plugins[0].get("source") != "./":
        fail("Claude marketplace plugins[0].source must be ./")


def validate_docs() -> None:
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    required = [
        "Office DOCX Skills",
        "Codex",
        "Claude Code",
        "Explicit Skill Triggering",
        "Combining Skills",
        "Marketplace status",
        "claude plugin marketplace add https://github.com/wintersdragon-c/office-docx-skills.git",
        "claude plugin install office-docx-skills@office-docx-skills-dev",
        "office-docx-skills:word-default-formatting",
        "claude plugin update office-docx-skills@office-docx-skills-dev",
    ]
    for token in required:
        if token not in readme:
            fail(f"README missing {token!r}")
    if "claude plugin update office-docx-skills\n" in readme:
        fail("README must update Claude plugin with marketplace-qualified name")
    obsolete_claude_symlink = "ln -s ~/.claude/office-docx-skills/skills ~/.claude/skills/office-docx-skills"
    if obsolete_claude_symlink in readme:
        fail("README must use Claude plugin marketplace install, not ~/.claude/skills symlink")
    install = (ROOT / ".codex" / "INSTALL.md").read_text(encoding="utf-8")
    for token in ["Installation", "Verify", "Updating", "Uninstalling"]:
        if token not in install:
            fail(f".codex/INSTALL.md missing {token}")


def validate_pressure_fixtures() -> None:
    prompt_dir = ROOT / "tests" / "explicit-skill-requests" / "prompts"
    if not prompt_dir.is_dir():
        fail("explicit skill request prompts directory missing")
    actual = {path.name for path in prompt_dir.glob("*.txt")}
    missing = EXPECTED_PRESSURE_PROMPTS - actual
    if missing:
        fail(f"missing explicit skill pressure prompts: {sorted(missing)}")
    runner = ROOT / "tests" / "explicit-skill-requests" / "run-behavior-test.sh"
    if not runner.exists():
        fail("explicit skill behavior test runner missing")
    run_all = ROOT / "tests" / "explicit-skill-requests" / "run-all.sh"
    if not run_all.exists():
        fail("explicit skill behavior run-all script missing")


def validate_no_local_paths() -> None:
    excluded_parts = {".git", "ref", "__pycache__"}
    for path in ROOT.rglob("*"):
        if not path.is_file():
            continue
        if excluded_parts & set(path.parts):
            continue
        if path.suffix.lower() in {".png", ".jpg", ".jpeg", ".gif", ".docx"}:
            continue
        text = path.read_text(encoding="utf-8", errors="ignore")
        for pattern in LOCAL_PATH_PATTERNS:
            if pattern in text:
                fail(f"hard-coded local path pattern {pattern!r} in {path}")


def main() -> None:
    validate_skills()
    validate_metadata()
    validate_docs()
    validate_pressure_fixtures()
    validate_no_local_paths()
    print("PASS: office-docx-skills package structure is valid")


if __name__ == "__main__":
    main()
