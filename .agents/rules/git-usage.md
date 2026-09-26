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