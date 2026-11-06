"""Read and inspection tools for Live Word COM backend."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Literal, Optional, Union

from src.core.constants import OUTLINE_LINE_LIMIT
from src.live.indexer import _build_visible_map
from src.live.worker import com_worker


class ReadTools:
    """Read and inspection mixin for LiveWordEngine."""

    @staticmethod
    def read_document(
        doc: Any,
        file_path: Union[str, Path],
        mode: Literal["outline", "full", "table", "metadata"] = "full",
        table_index: Optional[int] = None,
    ) -> Union[str, Dict[str, Any]]:
        """Read and inspect open Word document structure and content."""
        def _read() -> Union[str, Dict[str, Any]]:
            path = Path(file_path)
            vmap = _build_visible_map(doc)
            body_indices = vmap["indices"]
            tables = doc.Tables
            num_tables = tables.Count
            num_paragraphs = len(body_indices)

            if mode == "metadata":
                props = {}
                try:
                    built_in = doc.BuiltInDocumentProperties
                    for key in ["Title", "Author", "Subject", "Keywords", "Last author"]:
                        try:
                            props[key] = str(built_in.Item(key).Value)
                        except Exception:
                            pass
                except Exception:
                    pass
                return {
                    "file_path": str(path),
                    "total_paragraphs": num_paragraphs,
                    "total_tables": num_tables,
                    "properties": props,
                }

            if mode == "table":
                if table_index is None:
                    raise ValueError("table_index (0-based) is required when mode='table'")
                if table_index < 0 or table_index >= num_tables:
                    raise IndexError(
                        f"Table index {table_index} is out of range. Document has {num_tables} tables (indices 0 to {num_tables - 1})."
                    )
                tbl = tables.Item(table_index + 1)
                matrix: List[List[str]] = []
                for r_idx in range(1, tbl.Rows.Count + 1):
                    row = tbl.Rows.Item(r_idx)
                    row_vals = [
                        row.Cells.Item(c).Range.Text.rstrip("\r\n\x07").replace("\n", " ").strip()
                        for c in range(1, row.Cells.Count + 1)
                    ]
                    matrix.append(row_vals)

                md_lines = [f"### Table `[table:{table_index}]` ({tbl.Rows.Count} rows x {tbl.Columns.Count} cols)", ""]
                if matrix:
                    header = matrix[0]
                    md_lines.append("| " + " | ".join(header) + " |")
                    md_lines.append("| " + " | ".join(["---"] * len(header)) + " |")
                    for row in matrix[1:]:
                        md_lines.append("| " + " | ".join(row) + " |")
                return {
                    "table_index": table_index,
                    "rows_count": tbl.Rows.Count,
                    "columns_count": tbl.Columns.Count,
                    "matrix": matrix,
                    "markdown": "\n".join(md_lines),
                }

            if mode == "outline":
                lines = [
                    f"# Outline: {path.name}",
                    f"- **Total Paragraphs**: {num_paragraphs} (indices: [p:0] to [p:{max(0, num_paragraphs - 1)}])",
                    f"- **Total Tables**: {num_tables}",
                    "",
                    "### Structural Elements (0-based indices):",
                ]

                # Map table positions (1-based COM paragraph index where table starts)
                tbl_positions = []
                for t_idx in range(1, num_tables + 1):
                    try:
                        tr = tables.Item(t_idx).Range
                        p_start = doc.Range(0, tr.Start).Paragraphs.Count + 1
                        tbl_positions.append((p_start, t_idx))
                    except Exception:
                        tbl_positions.append((float("inf"), t_idx))

                # Combine body paragraphs and tables into a single document-order list
                elements = []
                for zero_idx, com_i in enumerate(body_indices):
                    elements.append((com_i, "p", zero_idx, com_i))
                for p_start, t_idx in tbl_positions:
                    elements.append((p_start, "tbl", t_idx - 1, t_idx))

                # Sort elements in document order
                elements.sort(key=lambda x: (x[0], 0 if x[1] == "tbl" else 1))

                # Fast heading detection: query doc.GetCrossReferenceItems(1)
                heading_com_indices = set()
                try:
                    items = doc.GetCrossReferenceItems(1)  # wdRefTypeHeading = 1
                except Exception:
                    items = ()

                if items:
                    rng = doc.Content
                    f = rng.Find
                    for heading_text in items:
                        clean = heading_text.strip()
                        if not clean:
                            continue
                        f.ClearFormatting()
                        if f.Execute(FindText=clean, Forward=True):
                            p_obj = rng.Paragraphs.Item(1)
                            c_i = doc.Range(0, p_obj.Range.End).Paragraphs.Count
                            heading_com_indices.add(c_i)
                            rng.Collapse(0)

                for elem in elements:
                    elem_type = elem[1]
                    if elem_type == "tbl":
                        tbl_idx = elem[2]
                        t_idx = elem[3]
                        try:
                            tbl = tables.Item(t_idx)
                            row_cnt = tbl.Rows.Count
                            col_cnt = tbl.Columns.Count
                            header_preview = []
                            if row_cnt > 0:
                                first_row = tbl.Rows.Item(1)
                                for c_i in range(1, min(col_cnt, 4) + 1):
                                    try:
                                        c_txt = first_row.Cells.Item(c_i).Range.Text.rstrip("\r\n\x07").strip()
                                        header_preview.append(c_txt)
                                    except Exception:
                                        pass
                            header_str = " | ".join(header_preview)
                            lines.append(f"- `[table:{tbl_idx}]` **Table ({row_cnt} rows x {col_cnt} cols)**: [{header_str}]")
                        except Exception:
                            lines.append(f"- `[table:{tbl_idx}]` **Table**")

                    elif elem_type == "p":
                        zero_idx = elem[2]
                        com_i = elem[3]

                        is_known_heading = com_i in heading_com_indices
                        if is_known_heading or len(lines) < OUTLINE_LINE_LIMIT or not items:
                            try:
                                p = doc.Paragraphs.Item(com_i)
                                style_name = str(p.Style.NameLocal) if p.Style else "Normal"
                                is_heading = is_known_heading or style_name.lower().startswith("heading") or style_name.lower() == "title"
                                text_snippet = p.Range.Text.rstrip("\r\n\x07").strip()
                                if is_heading or (len(lines) < OUTLINE_LINE_LIMIT and text_snippet):
                                    preview = (text_snippet[:80] + "...") if len(text_snippet) > 80 else text_snippet
                                    preview_text = f": \"{preview}\"" if preview else " (empty)"
                                    lines.append(f"- `[p:{zero_idx}]` **({style_name})**{preview_text}")
                            except Exception:
                                continue

                return "\n".join(lines)

            # Default: "full"
            lines = [f"# Document Content: {path.name}", ""]
            for zero_idx, com_i in enumerate(body_indices):
                p = doc.Paragraphs.Item(com_i)
                style_name = str(p.Style.NameLocal)
                txt = p.Range.Text.rstrip("\r\n\x07")
                lines.append(f"`[p:{zero_idx}]` ({style_name}): {txt}")

            for t_idx in range(1, num_tables + 1):
                tbl = tables.Item(t_idx)
                lines.append(f"\n`[table:{t_idx - 1}]` Table ({tbl.Rows.Count}x{tbl.Columns.Count}):")
                for r_idx in range(1, tbl.Rows.Count + 1):
                    row = tbl.Rows.Item(r_idx)
                    row_vals = [row.Cells.Item(c).Range.Text.rstrip("\r\n\x07") for c in range(1, row.Cells.Count + 1)]
                    lines.append(f"  row {r_idx - 1}: | " + " | ".join(row_vals) + " |")
                lines.append("")

            return "\n".join(lines)

        return com_worker.run(_read)

