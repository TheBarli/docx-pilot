---
trigger: model_decision
description: Use when creating PRs, inspecting GitHub issues, checking PR status, or running GitHub CLI (gh) commands
---

# GitHub CLI Automation Instructions

You are executing GitHub CLI (`gh`) operations inside a headless, non-interactive shell environment.

### 1. Terminal Execution Safeguards & TTY Guards
- **Zero Interactivity:** NEVER run commands without complete inline arguments. Commands missing parameters will hang the headless shell indefinitely waiting for terminal input.
  - FORBIDDEN: `gh auth login`
  - FORBIDDEN: bare `gh pr create` (without `--title` and `--body`)
  - FORBIDDEN: bare `gh issue create` (without `--title` and `--body`)
  - FORBIDDEN: browser flags (`-w`, `--web`)
- **Structured Reads:** Always parse read-only data with `--json` specifying explicit fields to prevent unparseable text blobs (e.g., `gh pr list --json number,title,state,headRefName`).
- **Fail-Fast on Auth/API Errors:** If `gh` returns authentication failure, rate limiting, or `403 Forbidden`, halt execution immediately and print stderr.

### 2. Pull Request Workflows (Native Default Branch)
- **Native Default Branch (Drop `--base`):** Do NOT specify `--base` unless explicitly requested by the user. `gh pr create` automatically resolves and targets the remote repository's default branch natively.
- **Pre-check:** Verify the local branch has an upstream tracking branch (`git push -u origin <branch>`) before creating a PR.
- **PR Creation:** Always specify inline title and body:
  ```bash
  gh pr create --title "<type>(<scope>): <summary>" --body "<body>"
  ```
- **Draft PRs:** Add `--draft` for work-in-progress or pending reviews.

### 3. Human-in-the-Loop (HITL) Safety Gates
- **Autonomous Allowed:**
  - Querying issues, PRs, and repository metadata via `--json`.
  - Checking PR CI/review status via `gh pr status` / `gh pr checks`.
  - Opening draft PRs (`gh pr create --draft`).
- **Mandatory Confirmation Gate:** STOP and wait for explicit human approval before running:
  - Any PR merge command (`gh pr merge`). Present the PR number, URL, and merge strategy (`--squash`, `--rebase`, or `--merge`) for human choice.
  - Any remote branch deletion (`gh pr merge --delete-branch`, `gh repo delete`).