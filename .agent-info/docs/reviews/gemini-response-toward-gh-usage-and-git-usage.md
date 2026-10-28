### How to Improve the Rules

* **Eliminate Synthetic Patch Staging:** Stop generating artificial patch files and running `git apply --cached diff.patch`. LLM patch generation frequently breaks on whitespace, hunks, and line drift. Stage files directly by path using `git add <file_path>`, and handle granular separation by making code edits sequentially.


* **Remove Hardcoded Target Branches:** Drop forced targets like `--base main`. Use the CLI to dynamically resolve the repository's default branch (`gh repo view --json defaultBranchRef`) or upstream tracking ref.


* **Add Context and Token Guardrails:** Stop running full `git diff` blindly across modified files. Run `git diff --stat` first to assess change volume, and explicitly skip minified bundles, build artifacts, and lockfiles (`package-lock.json`, `pnpm-lock.yaml`, `Cargo.lock`).


* **Calibrate Human-in-the-Loop (HITL) Gating:** Remove blocking gates for standard commits and non-destructive PR drafts on isolated feature branches. Reserve mandatory human confirmation strictly for destructive or shared-state changes (`gh pr merge`, force-pushes, and branch deletions).


* **Remove Hardcoded Merge Strategies:** Stop enforcing `--squash` globally. Repositories often dictate their own merge history policies (`--rebase`, `--merge`, or `--squash`).



---

### Global vs. Local Configuration Split

#### 1. What Belongs in Global Rules (`~/.config/...`)

Global rules must only define **headless runtime mechanics, token preservation, and destructive-action safeguards**. These principles apply universally across every project without breaking repository-specific standards:

* Headless CLI constraints (never use interactive commands like `git add -p` or bare `gh pr create` that hang the terminal without a TTY).


* Browser flag prohibitions (no `-w` or `--web`).


* Structured machine output requirements (`--json` with explicit fields for queries).


* Safe inspection practices (`git diff --stat` checks, excluding lockfiles from deep diffs).
* Universal safety gates (human confirmation required strictly for PR merging, branch deletion, or hard resets).


* Error halt triggers (immediate stop on non-zero exit codes or auth failures).



#### 2. What Belongs in Local Rules (`.cursorrules`, `.github/`, workspace config)

Local rules govern **team policy, Git history styling, and repository architecture**. These vary from project to project and must never be applied globally:

* Commit message specifications (Conventional Commits, custom team formats, ticket prefixes, allowed scopes).


* Default base branch names (`main`, `master`, `develop`).


* Merge strategy enforcement (`--squash`, `--rebase`, `--merge`).


* Branch naming conventions (`feat/*`, `fix/*`, `chore/*`).
* Project-specific pre-commit hooks, linter passes, or test suites required before committing.

---

### Ready-to-Use Configurations

#### The Global Configuration (`global-git-safety.md`)

```markdown
---
trigger: model_decision
description: Core headless execution safety and context rules for Git and GitHub CLI.
---

# Non-Interactive Git & GitHub CLI Execution Rules

You are running inside a headless, non-interactive shell environment. Adhere strictly to the following execution constraints:

### 1. Terminal Execution Safeguards
- **Zero Interactivity:** NEVER execute commands that require user interaction, terminal prompts, or editor input (e.g., bare `git commit`, `git add -i`, `git add -p`, or bare `gh pr create`). All commands must be provided with complete inline parameters.
- **No Browser Triggers:** NEVER pass browser flags (`-w`, `--web`).
- **Structured Reads:** Always parse read-only data with `--json` specifying explicit fields (e.g., `gh pr list --json number,title,state`).
- **Fail-Fast:** If a command exits with a non-zero status, hook failure, or authentication error, stop execution immediately and print stderr.

### 2. Context Protection & Token Limits
- Always run `git diff --stat` before inspecting file changes.
- Never output full diffs for generated files, vendor bundles, minified assets, or lockfiles (`package-lock.json`, `pnpm-lock.yaml`, `Cargo.lock`, `yarn.lock`).
- Stage target files explicitly by path (`git add <path>`). Do not construct synthetic `.patch` files to stage partial hunks.

### 3. Execution Boundaries & Human Confirmation
- **Autonomous Allowed:** Creating local branches, staging changed files by path, committing on feature branches, and opening draft pull requests.
- **Mandatory Confirmation Gate:** STOP and wait for explicit human approval before running:
  - Any PR merge command (`gh pr merge`)
  - Any destructive update (`git reset --hard`, `git push --force`)
  - Any branch deletion (`git branch -D`, `gh pr merge --delete-branch`)

```

#### The Local Project Template (`.cursorrules` or project-level rule)

```markdown
# Repository Git & PR Conventions

- **Default Base Branch:** Discover dynamically, or use `main`.
- **Commit Formatting:** Follow Conventional Commits: `<type>(<scope>): <summary>`
  - Types: `feat`, `fix`, `refactor`, `perf`, `test`, `chore`, `docs`, `style`
  - Max 72 characters, imperative mood, no trailing period.
  - Include detailed body explaining WHAT and WHY for complex logic changes.
- **PR Merge Strategy:** Use `--squash` (or adjust to project standard: `--rebase` / `--merge`).
- **Branch Naming:** `<type>/<short-description>`

```