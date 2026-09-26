---
trigger: model_decision
description: When saving informational notes, scratchpads, summaries, or research into text or markdown files
---

# Auxiliary Information Directory

- Store all agent-generated notes, summaries, logs, research outputs, and scratch files inside `./.agent-info/`.
- Ensure the directory exists (`./.agent-info/`) before writing to it.
- **Scope**: This applies to informational and temporary agent files only; actual project source code and official documentation belong in their standard project locations.