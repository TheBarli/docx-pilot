---
trigger: model_decision
description: When editing, formatting, or inspecting a .docx with the word-docx MCP tools while the file may be open in Microsoft Word
---

# Live Word Mode Rule

1. **Backend is automatic.** The first line of every `word-docx` tool response shows `[backend: live]` or `[backend: headless]`. Never switch backends or set `WORD_BACKEND` yourself.
2. **When live:** edits appear in the open Word window as tracked changes, and each tool call is one undo step. Never close, save, or reopen the document. Never touch the same file through any other method (python-docx, scripts, shell).
3. **Indices:** after any `mutate_element`, use the returned local outline for the next `[p:N]`. Do not reuse earlier indices. Re-read only if your target is not in the returned outline.
4. **If a tool reports** the file is read-only, protected, cloud-hosted, or that Word looks blocked by a dialog: follow `problem-resolve` (stop, report, wait).
