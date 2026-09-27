# Agent Auxiliary Information Directory (`.agent-info/`)

This directory contains persistent notes, architectural plans, reviews, and reference pipelines for the **Word DOCX FastMCP Server** project.

## Directory Structure

```text
.agent-info/
├── README.md             # This guide
├── docs/                 # Persistent Markdown documentation
│   ├── plans/            # Architecture specifications & task roadmaps
│   ├── analysis/         # Problem root-cause analyses & debugging notes
│   ├── reviews/          # Agent/human review responses & discussions
│   └── workspace-summary.md
└── pipelines/            # Specialized standalone batch-formatting scripts
```

## Guidelines

1. **MCP First**: Use the registered FastMCP server tools (`read_document`, `compose_document`, `mutate_element`, `find_and_replace`) directly via JSON-RPC instead of creating loose `.py` inspection or scratchpad scripts.
2. **Subfolder Enforcement**: Store all future agent notes in appropriate subfolders inside `docs/` (`plans/`, `analysis/`, `reviews/`). Do not leave loose files in `.agent-info/`.
3. **Formal Tests**: Automated tests belong in `tests/test_mutations.py`, run via `pytest`.
