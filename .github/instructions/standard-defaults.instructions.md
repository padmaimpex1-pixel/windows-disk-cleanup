---
applyTo: "**"
description: "Use when writing code, generating files, organizing folders, or pushing to GitHub for this user. Applies coding style, folder conventions, and GitHub defaults."
---

# Standard Defaults — padmaimpex1-pixel

## 1. GitHub

- Default GitHub owner/account: `padmaimpex1-pixel`
- Default repo host: `github.com/padmaimpex1-pixel`
- When pushing, if repository is not specified, ask for repo name once and then push there.
- **Never push without explicit user intent.**
- Commit messages must be concise and in imperative form (e.g. `Add script`, `Fix scraper`).
- Always include the Co-authored-by trailer in commits:
  `Co-authored-by: Copilot <223556219+Copilot@users.noreply.github.com>`

## 2. Folder Conventions — D:\GitRepos

Repos are organized into category folders under `D:\GitRepos`:

| Folder | Content |
|---|---|
| `AI-Agents` | AI/ML models, agents, bots, image/photo tools |
| `Communication` | WhatsApp, Exotel, voice, presentation, messaging tools |
| `Forensics-Security` | Security, password managers, forensics |
| `Healthcare` | Medical, physio, health-related projects |
| `Infrastructure` | Scripts, automation, DevOps, utilities |
| `Legal` | Legal tools, attorney apps |
| `Media-Video` | Video tools, slideshows, music, dance, ffmpeg |
| `Misc` | Uncategorized or archive projects |
| `Perl-Web` | Perl web projects |
| `Website Templates` | WordPress, Shopify, website clones and templates |
| `starter-repo` | Active working/scratch workspace |

- Python utility scripts go into `_python-scripts/` inside the relevant category folder.
- Input files (images, data) go into `D:\GitRepos\input`.
- Generated output files go to:
  `D:\Generated-Outputs`
  unless the user specifies otherwise.

## 3. Coding Style

- **Language preference**: Python for scripts/automation, JavaScript/TypeScript for web.
- Keep functions small and single-purpose.
- Only add comments where logic genuinely needs clarification — no obvious comments.
- Use descriptive variable names; avoid abbreviations.
- Prefer ecosystem tools (`pip`, `npm`, `venv`) over manual setup.
- Never add linters, test frameworks, or build tools unless already present in the project.

## 4. File Generation

- Default output path for exports, CSVs, JSONs, and session artifacts:
  `D:\Generated-Outputs`
- Always ask before overwriting existing files.
- Clean up temporary files at end of task.

## 5. Python Script Workflow

- **Always implement tasks as Python scripts**, not one-off shell commands.
- Maintain a **master Python script** per project/workspace: `master.py` at the repo root (or category root).
- For every new task, **add a new function** to `master.py` rather than creating scattered standalone scripts.
  - Function names should be descriptive and verb-led (e.g. `scan_pictures()`, `move_to_media_video()`, `organise_gitrepos()`).
  - Each function should be independently callable.
  - Add a `if __name__ == "__main__":` block that lists/calls available functions.
- **Always push scripts to GitHub after every task** so version history is maintained.
  - Default push target: `padmaimpex1-pixel` (ask for repo name if not specified).
  - Commit message should describe the task the function performs.
- Standalone scripts are allowed only when the task is genuinely isolated and unrelated to any existing master script.
- Keep `D:\Generated-Outputs` as the default output path for any files the scripts produce.

## 6. General Agent Behaviour

- Make precise, surgical changes — do not touch unrelated code.
- Verify changes work before declaring a task done.
- Ask one focused clarifying question at a time when requirements are ambiguous.
- Never generate or commit secrets or credentials.
