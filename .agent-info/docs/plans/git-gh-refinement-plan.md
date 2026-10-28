# Implementation Plan: Refinement of `git-usage` & `gh-usage` (Simplified & Pragmatic)

## 1. Objective & Philosophy
Strip away all fragile, over-engineered shell pipelines. By dropping `--base` entirely, `gh pr create` natively targets the repository's default branch on GitHub with zero discovery scripts, zero POSIX vs. PowerShell friction, and zero risk of pipeline exit-code masking.

---

## 2. File Specifications

### A. Refined `git-usage.md` ([.agents/rules/git-usage.md](file:///c:/Users/levine/Documents/Github/MsWord-ai-integration/.agents/rules/git-usage.md))

```markdown
---
trigger: model_decision
description: When inspecting repository status, staging changes, writing commit messages, or creating git commits
---

# Git Automation & Non-Interactive Execution Rules

You are executing Git commands inside a headless, non-interactive shell environment. Adhere strictly to the following constraints:

### 1. Terminal Execution Safeguards & TTY Guards
- **Zero Interactivity:** NEVER execute commands that prompt for user interaction or open an editor.
  - FORBIDDEN: `git add -i`, `git add -p`
  - FORBIDDEN: bare `git commit` (without `-m`)
  - FORBIDDEN: bare `git merge` or `git rebase -i`
  - All commits must provide explicit inline messages: `git commit -m "<header>" -m "<body>"`.
- **Explicit Path Staging:** Stage target files explicitly by path (`git add <file_path>`). NEVER construct synthetic `.patch` files or run `git apply --cached diff.patch`. For granular separation of concerns, perform code edits and commits sequentially.
- **Fail-Fast:** If any git command exits with a non-zero exit code or hook failure, halt immediately, print the exact stderr, and wait for guidance.

### 2. Context & Token Protection
- **Pre-check with Stat:** Always run `git diff --stat` before inspecting diffs to assess change volume.
- **Hard 300-Line Ceiling:** If `git diff --stat` reports >300 changed lines on a single file, NEVER dump the full diff into context. Instead, inspect changes via AST/symbol tools, targeted grep, or head/tail snippets.
- **Lockfile & Asset Protection:** NEVER inspect full diffs for generated assets, vendor bundles, or lockfiles (`package-lock.json`, `pnpm-lock.yaml`, `yarn.lock`, `Cargo.lock`, `poetry.lock`). Stage lockfiles by path without printing their full diff.
- **Targeted Diffs:** Inspect tracked files individually using `git diff -- <file_path>` only when change volume is under 300 lines.

### 3. Commit Message Standards (Conventional Commits)
- **Header:** `<type>(<scope>): <summary>`
  - Types: `feat`, `fix`, `refactor`, `perf`, `test`, `chore`, `docs`, `style`
  - Max 72 characters, imperative mood ("add" not "added"), no trailing period.
- **Body:** Explain WHAT changed and WHY for non-trivial logic. Wrap at 72 characters.
- **Footer:** Reference issues (`Closes #123`) or breaking changes (`BREAKING CHANGE: ...`).

### 4. Human-in-the-Loop (HITL) Safety Gates
- **Autonomous Allowed:**
  - Creating local feature branches, inspecting status, staging files by path.
  - Unstaging index modifications (`git restore --staged <path>`, `git restore --staged .`) as part of atomic commit grouping.
  - Committing on isolated feature branches.
  - Pushing feature branches upstream (`git push -u origin <branch>`, `git push origin <branch>`).
- **Mandatory Confirmation Gate:** STOP and obtain explicit human approval before running:
  - Permanent working tree discards: `git checkout -- .`, `git restore .`
  - Untracked file deletions: `git clean -fd`, `git clean -f`
  - Branch deletions: `git branch -D`, `git branch -d`
  - Force-pushes: `git push --force`, `git push -f`
  - Direct commits or pushes to shared protected branches (`main`, `master`, `develop`)
```

---

### B. Refined `gh-usage.md` ([.agents/rules/gh-usage.md](file:///c:/Users/levine/Documents/Github/MsWord-ai-integration/.agents/rules/gh-usage.md))

```markdown
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
```

---

## 3. Execution Steps

1. **Step 1: Write `.agents/rules/git-usage.md`**:
   - Apply clean TTY bans, 300-line diff ceiling, autonomous upstream pushes, and destructive-only confirmation gates.
2. **Step 2: Write `.agents/rules/gh-usage.md`**:
   - Drop `--base` completely, rely on native GitHub CLI default branch resolution, ban interactive flags, and enforce `--json` queries.
3. **Step 3: Verification**:
   - Verify rules format, triggers, and absence of synthetic patches or complex shell discovery pipelines.
