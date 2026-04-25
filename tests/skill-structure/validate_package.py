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
    "docx-tracked-changes",
    "word-default-formatting",
    "word-formula-writing",
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
    ]
    for token in required:
        if token not in readme:
            fail(f"README missing {token!r}")
    unpublished_install = "/plugin install " + "office-docx-skills@"
    if unpublished_install in readme:
        fail("README must not include unpublished Claude marketplace install command")
    install = (ROOT / ".codex" / "INSTALL.md").read_text(encoding="utf-8")
    for token in ["Installation", "Verify", "Updating", "Uninstalling"]:
        if token not in install:
            fail(f".codex/INSTALL.md missing {token}")


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
    validate_no_local_paths()
    print("PASS: office-docx-skills package structure is valid")


if __name__ == "__main__":
    main()
