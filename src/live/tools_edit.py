"""Edit and mutation tools for Live Word COM backend."""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Literal, Optional, Union

from src.core.exceptions import ERROR_MESSAGES, LiveWordError
from src.live.indexer import WD_REVISION_DELETE, _build_visible_map, _get_local_outline
from src.live.undo import live_call
from src.live.worker import com_worker


class EditTools:
    """Edit and mutation mixin for LiveWordEngine."""

    @staticmethod
    def mutate_element(
        doc: Any,
        file_path: Union[str, Path],
        target_index: int,
        action: Literal["insert_before", "insert_after", "delete", "replace"],
        content: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Perform surgical mutation at target 0-based paragraph index with Track Changes."""
        def _mutate() -> Dict[str, Any]:
            vmap = _build_visible_map(doc)
            body_indices = vmap["indices"]
            if target_index < 0 or target_index >= len(body_indices):
                valid = f"0..{len(body_indices) - 1}" if body_indices else "empty"
                raise LiveWordError(
                    ERROR_MESSAGES["index_out_of_range"].format(target_index=target_index, valid_range=valid)
                )

            com_i = body_indices[target_index]
            p = doc.Paragraphs.Item(com_i)
            path = Path(file_path)

            with live_call(doc, "mutate_element", f"[p:{target_index}] {action}", track=True):
                if action == "delete":
                    p.Range.Delete()
                    fresh_outline = _get_local_outline(doc, target_index)
                    return {
                        "status": "success",
                        "action": "delete",
                        "target_index": target_index,
                        "file_path": str(path),
                        "local_outline": fresh_outline,
                    }

                if action == "replace":
                    c = content or {}
                    new_text = str(c.get("text", ""))
                    rng = p.Range
                    # Shrink range end by 1 to retain paragraph mark '\r'
                    if rng.End > rng.Start:
                        rng.End = rng.End - 1
                    rng.Text = new_text

                    if "style" in c:
                        try:
                            p.Style = c["style"]
                        except Exception:
                            pass
                    if c.get("bold") is not None:
                        rng.Font.Bold = bool(c["bold"])
                    if c.get("italic") is not None:
                        rng.Font.Italic = bool(c["italic"])

                    fresh_outline = _get_local_outline(doc, target_index)
                    return {
                        "status": "success",
                        "action": "replace",
                        "target_index": target_index,
                        "file_path": str(path),
                        "updated_text": new_text,
                        "local_outline": fresh_outline,
                    }

                if action in ("insert_before", "insert_after"):
                    c = content or {}
                    new_text = str(c.get("text", ""))
                    style = c.get("style", "Normal")

                    if action == "insert_before":
                        new_p = doc.Paragraphs.Add(p.Range)
                    else:
                        rng = p.Range
                        collapse_rng = doc.Range(rng.End, rng.End)
                        new_p = doc.Paragraphs.Add(collapse_rng)

                    try:
                        new_p.Style = style
                    except Exception:
                        pass

                    rng_new = new_p.Range
                    if rng_new.End > rng_new.Start:
                        rng_new.End = rng_new.End - 1
                    rng_new.Text = new_text

                    if c.get("bold") is not None:
                        rng_new.Font.Bold = bool(c["bold"])
                    if c.get("italic") is not None:
                        rng_new.Font.Italic = bool(c["italic"])

                    fresh_outline = _get_local_outline(doc, target_index)
                    return {
                        "status": "success",
                        "action": action,
                        "target_index": target_index,
                        "file_path": str(path),
                        "local_outline": fresh_outline,
                    }

            raise ValueError(f"Unknown action: {action}")

        return com_worker.run(_mutate)

    @staticmethod
    def find_and_replace(
        doc: Any,
        file_path: Union[str, Path],
        find_text: str,
        replace_text: str,
        match_case: bool = True,
        scope: Literal["all", "body", "tables", "headers_footers"] = "body",
    ) -> Dict[str, Any]:
        """Locate-then-set find and replace avoiding 255-char replacement limits and skipping deletions."""
        if not find_text:
            raise ValueError("find_text cannot be empty.")
        if len(find_text) > 255:
            raise ValueError("find_text exceeds Word 255 character limit.")

        def _fnr() -> Dict[str, Any]:
            path = Path(file_path)
            count = 0
            with live_call(doc, "find_and_replace", f"'{find_text[:20]}'", track=True):
                ranges = []
                if scope in ("all", "body"):
                    ranges.append(doc.Content)
                elif scope == "tables":
                    for t_i in range(1, doc.Tables.Count + 1):
                        ranges.append(doc.Tables.Item(t_i).Range)

                if scope in ("all", "headers_footers"):
                    for s_i in range(1, doc.Sections.Count + 1):
                        sec = doc.Sections.Item(s_i)
                        for hf_group in (sec.Headers, sec.Footers):
                            for hf_i in range(1, hf_group.Count + 1):
                                try:
                                    hf = hf_group.Item(hf_i)
                                    if getattr(hf, "Exists", False):
                                        ranges.append(hf.Range)
                                except Exception:
                                    pass

                for curr_rng in ranges:
                    find_obj = curr_rng.Find
                    find_obj.ClearFormatting()
                    find_obj.Text = find_text
                    find_obj.MatchCase = match_case
                    find_obj.Forward = True
                    find_obj.Wrap = 0  # wdFindStop

                    while find_obj.Execute():
                        in_del = False
                        try:
                            for rev in curr_rng.Revisions:
                                if rev.Type == WD_REVISION_DELETE:
                                    in_del = True
                                    break
                        except Exception:
                            pass

                        if not in_del:
                            curr_rng.Text = replace_text
                            count += 1

                        curr_rng.Collapse(0)  # wdCollapseEnd

            return {
                "status": "success",
                "file_path": str(path),
                "matches_replaced": count,
                "find_text": find_text,
                "replace_text": replace_text,
            }

        return com_worker.run(_fnr)

    @staticmethod
    def compose_document(
        doc: Any,
        file_path: Union[str, Path],
        actions: List[Dict[str, Any]],
        create_if_missing: bool = True,
    ) -> Dict[str, Any]:
        """Append a batch of formatted elements to the active Word document."""
        def _compose() -> Dict[str, Any]:
            path = Path(file_path)
            executed: List[str] = []

            with live_call(doc, "compose_document", f"{len(actions)} actions"):
                for idx, act in enumerate(actions):
                    act_type = str(act.get("type", "")).lower()

                    if act_type == "heading":
                        text = str(act.get("text", ""))
                        level = int(act.get("level", 1))
                        end_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                        p = doc.Paragraphs.Add(Range=end_rng)
                        p.Range.Text = text + "\n"
                        try:
                            p.Style = f"Heading {level}"
                        except Exception:
                            pass
                        executed.append(f"#{idx} Added Heading (level {level}): \"{text[:40]}\"")

                    elif act_type in ("paragraph", "list_item"):
                        text = str(act.get("text", ""))
                        default_style = "List Bullet" if act_type == "list_item" else "Normal"
                        style = str(act.get("style", default_style))
                        end_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                        p = doc.Paragraphs.Add(Range=end_rng)
                        p.Range.Text = text + "\n"
                        try:
                            p.Style = style
                        except Exception:
                            pass
                        if act.get("bold") is not None:
                            p.Range.Font.Bold = bool(act["bold"])
                        if act.get("italic") is not None:
                            p.Range.Font.Italic = bool(act["italic"])
                        executed.append(f"#{idx} Added Paragraph: \"{text[:40]}\"")

                    elif act_type == "table":
                        headers = act.get("headers", [])
                        rows = act.get("rows", [])
                        num_rows = len(rows) + (1 if headers else 0)
                        num_cols = len(headers) if headers else (len(rows[0]) if rows else 1)
                        end_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                        tbl = doc.Tables.Add(Range=end_rng, NumRows=num_rows, NumColumns=num_cols)
                        try:
                            tbl.Style = act.get("style", "Table Grid")
                        except Exception:
                            pass
                        curr_r = 1
                        if headers:
                            for c_i, h in enumerate(headers):
                                tbl.Cell(1, c_i + 1).Range.Text = str(h)
                            curr_r = 2
                        for r_vals in rows:
                            for c_i, val in enumerate(r_vals):
                                if c_i < num_cols:
                                    tbl.Cell(curr_r, c_i + 1).Range.Text = str(val)
                            curr_r += 1
                        executed.append(f"#{idx} Added Table ({num_rows}x{num_cols})")

                    elif act_type == "page_break":
                        end_rng = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                        end_rng.InsertBreak(7)  # wdPageBreak = 7
                        executed.append(f"#{idx} Added Page Break")

            return {
                "status": "success",
                "file_path": str(path),
                "total_actions": len(actions),
                "actions_executed": executed,
                "paragraphs_count": doc.Paragraphs.Count,
                "tables_count": doc.Tables.Count,
            }

        return com_worker.run(_compose)
