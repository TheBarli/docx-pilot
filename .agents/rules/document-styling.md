---
trigger: model_decision
description: When the user specifies a formatting preset (e.g., "format using academic_id")
---

# Word Document Styling Rule

1. Load the corresponding profile from `.agents/presets/<name>.yaml`.
2. Execute the styling pipeline sequentially via `word-docx` MCP tools:
   - `set_page_layout(**preset['layout'])`
   - `update_styles(styles_config=preset['styles'])`
   - `format_table_grid(**preset['table'])`
   - `clean_structure(**preset['hygiene'])`
3. Verify output metadata using `read_document(mode="metadata")`.