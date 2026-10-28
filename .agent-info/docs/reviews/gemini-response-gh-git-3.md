Fair point. That is completely on me. I sent you down a rabbit hole trying to engineer bulletproof cross-platform shell pipelines to resolve the base branch, when the sane answer was just dropping `--base` entirely and letting GitHub CLI handle it natively. Classic case of over-engineering a fix for a problem I introduced in the first place.

Let's cut the bloat and strip out all the shell-pipeline mess. The clean, finalized setup needs just this:

* **Git Rule (`git-usage.md`):**
* Ban interactive flags (`-i`, `-p`, bare commits, interactive rebase).
* Enforce the 300-line ceiling from `git diff --stat` (no full diffs on massive files or lockfiles).
* Allow routine upstream pushes (`git push -u origin <branch>`) autonomously.
* Gate destructive commands only (`git reset --hard`, `git clean -fd`, `git branch -D`, force pushes).


* **GH Rule (`gh-usage.md`):**
* Ban interactive commands (require `--title` and `--body` inline).
* **Drop `--base` entirely**—let `gh pr create` target the repo's default branch natively without running any discovery scripts.
* Require `--json` for queries.
* Gate `gh pr merge` and remote branch deletion.



Zero POSIX vs. PowerShell nonsense, no synthetic patch files, and no shell crashes.