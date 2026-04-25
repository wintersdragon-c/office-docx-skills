# Installing Office DOCX Skills for Codex

Enable Office DOCX Skills in Codex via native skill discovery. Clone this repository and symlink its `skills/` directory.

## Prerequisites

- Git
- OpenAI Codex CLI or Codex app with skill discovery

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/wintersdragon-c/office-docx-skills.git ~/.codex/office-docx-skills
   ```

2. Create the skills symlink:

   ```bash
   mkdir -p ~/.agents/skills
   ln -s ~/.codex/office-docx-skills/skills ~/.agents/skills/office-docx-skills
   ```

   Windows PowerShell:

   ```powershell
   New-Item -ItemType Directory -Force -Path "$env:USERPROFILE\.agents\skills"
   cmd /c mklink /J "$env:USERPROFILE\.agents\skills\office-docx-skills" "$env:USERPROFILE\.codex\office-docx-skills\skills"
   ```

3. Restart Codex so it discovers the skills.

## Verify

```bash
ls -la ~/.agents/skills/office-docx-skills
ls ~/.agents/skills/office-docx-skills
```

Expected skills:

```text
docx-tracked-changes
word-default-formatting
word-formula-writing
```

## Updating

```bash
cd ~/.codex/office-docx-skills
git pull
```

Skills update through the symlink after Codex restarts.

## Uninstalling

```bash
rm ~/.agents/skills/office-docx-skills
```

Optionally remove the clone:

```bash
rm -rf ~/.codex/office-docx-skills
```
