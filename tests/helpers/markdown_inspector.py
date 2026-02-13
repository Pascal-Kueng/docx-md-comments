from __future__ import annotations

import json
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path

COMMENT_START_RE = re.compile(r"\{\.comment-start(?P<attrs>[^}]*)\}", re.DOTALL)
KV_ATTR_RE = re.compile(r'([A-Za-z_:][-A-Za-z0-9_:.]*)="([^"]*)"')
CARD_META_INLINE_RE = re.compile(
    r"<!--\s*CARD_META\s*\{\s*#(?P<id>[A-Za-z0-9][A-Za-z0-9_-]*)\s*(?P<attrs>.*?)\}\s*-->",
    re.DOTALL,
)
CARD_HEADER_RE = re.compile(
    r"^\[!\s*(?P<kind>COMMENT|REPLY)\s+(?P<id>[A-Za-z0-9][A-Za-z0-9_-]*)\s*:\s*(?P<author>.+?)\s*\((?P<state>active|resolved)\)\s*\]$",
    re.IGNORECASE,
)
INLINE_IMAGE_RE = re.compile(
    r'!\[(?P<alt>[^\]]*)\]\((?P<src>[^)\s]+)(?:\s+"(?P<title>[^"]*)")?\)\{(?P<attrs>[^}]*)\}',
    re.DOTALL,
)
MILESTONE_TOKEN_RE = re.compile(
    r"(?:(?P<markeq>==)\s*)?"
    r"(?:/{3}\s*C(?P<id3c>[0-9][A-Za-z0-9_-]*)\s*\.\s*(?P<edge3c>[sSeE]|[Ss][Tt][Aa][Rr][Tt]|[Ee][Nn][Dd])\s*/{3}"
    r"|/{3}\s*(?P<id3>[A-Za-z0-9][A-Za-z0-9_-]*)\s*\.\s*(?P<edge3>[sSeE]|[Ss][Tt][Aa][Rr][Tt]|[Ee][Nn][Dd])\s*/{3})"
    r"(?(markeq)\s*==)"
)


def normalize_milestone_edge(edge_token: str) -> str:
    token = str(edge_token or "").strip().lower()
    if token in {"s", "start"}:
        return "s"
    if token in {"e", "end"}:
        return "e"
    return ""


def milestone_match_id_edge(match: re.Match) -> tuple[str, str]:
    group_dict = match.groupdict()
    comment_id = str(
        group_dict.get("id3c")
        or group_dict.get("id3")
        or ""
    ).strip()
    edge_token = str(
        group_dict.get("edge3c")
        or group_dict.get("edge3")
        or ""
    )
    return comment_id, normalize_milestone_edge(edge_token)


def parse_card_meta_marker(raw_html: str) -> tuple[str, dict[str, str]]:
    match = CARD_META_INLINE_RE.search(raw_html or "")
    if not match:
        return "", {}
    comment_id = str(match.group("id") or "").strip()
    attrs_raw = str(match.group("attrs") or "").strip()
    meta: dict[str, str] = {}
    if attrs_raw:
        try:
            payload = json.loads("{" + attrs_raw + "}")
        except json.JSONDecodeError:
            payload = {}
        if isinstance(payload, dict):
            for key, value in payload.items():
                key_str = str(key or "").strip()
                if key_str:
                    meta[key_str] = str(value or "").strip()
    return comment_id, meta


@dataclass(frozen=True)
class MarkdownCommentStart:
    id: str
    order: int
    text: str
    author: str
    date: str
    parent: str
    state: str

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "order": self.order,
            "text": self.text,
            "author": self.author,
            "date": self.date,
            "parent": self.parent,
            "state": self.state,
        }


@dataclass(frozen=True)
class MarkdownCommentSnapshot:
    path: str
    starts: list[MarkdownCommentStart]
    start_ids_order: list[str]
    end_ids_order: list[str]
    parent_by_id: dict[str, str]
    child_ids: list[str]
    root_ids_order: list[str]
    own_text_by_id: dict[str, str]
    flattened_by_root: dict[str, str]
    state_by_id: dict[str, str]
    placeholder_shape_match_count: int
    none_line_count: int

    def to_dict(self) -> dict:
        return {
            "path": self.path,
            "starts": [s.to_dict() for s in self.starts],
            "start_ids_order": self.start_ids_order,
            "end_ids_order": self.end_ids_order,
            "parent_by_id": self.parent_by_id,
            "child_ids": self.child_ids,
            "root_ids_order": self.root_ids_order,
            "own_text_by_id": self.own_text_by_id,
            "flattened_by_root": self.flattened_by_root,
            "state_by_id": self.state_by_id,
            "placeholder_shape_match_count": self.placeholder_shape_match_count,
            "none_line_count": self.none_line_count,
        }


def normalize_comment_text(text: str) -> str:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\\+[ \t]*\n", "\n", text)
    text = re.sub(r"\\\\[ \t]+", "\n", text)
    text = (
        text.replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
    )
    text = re.sub(r"(?m)^[\u2014\u2015]\s*$", "---", text)
    return text.strip()


def inlines_to_text(inlines) -> str:
    parts: list[str] = []

    def emit(value: str) -> None:
        if value:
            parts.append(value)

    def walk(node) -> None:
        if not isinstance(node, dict):
            return
        t = node.get("t")
        c = node.get("c")
        if t == "Str":
            emit(c or "")
        elif t == "Space":
            emit(" ")
        elif t == "SoftBreak":
            if parts and parts[-1].endswith("\\"):
                parts[-1] = parts[-1].rstrip("\\")
                emit("\n")
            else:
                emit(" ")
        elif t == "LineBreak":
            if parts and parts[-1].endswith("\\"):
                parts[-1] = parts[-1].rstrip("\\")
            emit("\n")
        elif t in {"Code", "Math"}:
            if isinstance(c, list) and c:
                emit(c[-1] or "")
            elif isinstance(c, str):
                emit(c)
        elif t == "Span":
            if isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk(item)
        elif t in {"Emph", "Strong", "Strikeout", "Superscript", "Subscript", "SmallCaps", "Underline"}:
            if isinstance(c, list):
                for item in c:
                    walk(item)
        elif t == "Quoted":
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                quote_type = c[0]
                quote_name = (
                    str(quote_type.get("t") or "").strip()
                    if isinstance(quote_type, dict)
                    else str(quote_type or "").strip()
                ).lower()
                if "single" in quote_name:
                    open_quote, close_quote = "'", "'"
                else:
                    open_quote, close_quote = '"', '"'
                emit(open_quote)
                for item in c[1]:
                    walk(item)
                emit(close_quote)
        elif t == "Cite":
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk(item)
        elif t in {"Link", "Image"}:
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk(item)
        elif t == "RawInline":
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], str):
                emit(c[1])
        elif isinstance(c, list):
            for item in c:
                if isinstance(item, dict):
                    walk(item)

    for item in inlines or []:
        walk(item)
    text = "".join(parts)
    text = re.sub(r"[ \t]+\n", "\n", text)
    return normalize_comment_text(text)


def inlines_to_card_text(inlines) -> str:
    parts: list[str] = []

    def emit(value: str) -> None:
        if value:
            parts.append(value)

    def walk(node) -> None:
        if not isinstance(node, dict):
            return
        t = node.get("t")
        c = node.get("c")
        if t == "Str":
            emit(c or "")
        elif t == "Space":
            emit(" ")
        elif t in {"SoftBreak", "LineBreak"}:
            emit("\n")
        elif t in {"Code", "Math"}:
            if isinstance(c, list) and c:
                emit(c[-1] or "")
            elif isinstance(c, str):
                emit(c)
        elif t == "RawInline":
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], str):
                emit(c[1])
        elif t == "Span":
            if isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk(item)
        elif t in {"Emph", "Strong", "Strikeout", "Superscript", "Subscript", "SmallCaps", "Underline"}:
            if isinstance(c, list):
                for item in c:
                    walk(item)
        elif t == "Quoted":
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                quote_type = c[0]
                quote_name = (
                    str(quote_type.get("t") or "").strip()
                    if isinstance(quote_type, dict)
                    else str(quote_type or "").strip()
                ).lower()
                if "single" in quote_name:
                    open_quote, close_quote = "'", "'"
                else:
                    open_quote, close_quote = '"', '"'
                emit(open_quote)
                for item in c[1]:
                    walk(item)
                emit(close_quote)
        elif t == "Cite":
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk(item)
        elif t in {"Link", "Image"}:
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk(item)
        elif isinstance(c, list):
            for item in c:
                if isinstance(item, dict):
                    walk(item)

    for item in inlines or []:
        walk(item)
    text = "".join(parts).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    return text


def reply_header(author: str, date: str) -> str:
    safe_author = (author or "Unknown").strip() or "Unknown"
    safe_date = (date or "").strip()
    if safe_date:
        return f"---\nReply from: {safe_author} ({safe_date})\n---"
    return f"---\nReply from: {safe_author}\n---"


def normalize_state_token(value: str) -> str:
    return "resolved" if (value or "").strip().lower() == "resolved" else "active"


def run_pandoc_json(markdown_path: Path) -> dict:
    cmd = ["pandoc", str(markdown_path), "-f", "markdown", "-t", "json"]
    out = subprocess.check_output(cmd, text=True, encoding="utf-8")
    return json.loads(out)


def inspect_markdown_comments(markdown_path: Path) -> MarkdownCommentSnapshot:
    markdown_path = Path(markdown_path)
    doc = run_pandoc_json(markdown_path)

    starts: list[MarkdownCommentStart] = []
    end_ids_order: list[str] = []
    own_text_by_id: dict[str, str] = {}
    metadata_by_id: dict[str, dict[str, str]] = {}
    parent_candidate_by_id: dict[str, str] = {}
    state_by_id: dict[str, str] = {}
    card_by_id: dict[str, dict[str, str]] = {}
    start_order = 0

    def extract_comment_card_text(blocks) -> str:
        parts = []

        def normalize_card_line(text: str) -> str:
            line = str(text or "").strip()
            if not line:
                return ""
            if CARD_HEADER_RE.match(line):
                return ""
            if CARD_META_INLINE_RE.search(line):
                return ""
            return line

        for block in blocks or []:
            if not isinstance(block, dict):
                continue
            t = block.get("t")
            c = block.get("c")
            if t in {"Para", "Plain"} and isinstance(c, list):
                text = normalize_card_line(inlines_to_card_text(c))
                if text:
                    parts.append(text)
            elif t == "Header" and isinstance(c, list) and len(c) >= 3 and isinstance(c[2], list):
                text = inlines_to_card_text(c[2])
                if text:
                    parts.append(text)
            elif t == "RawBlock" and isinstance(c, list) and len(c) == 2:
                raw = str(c[1] or "")
                if CARD_META_INLINE_RE.search(raw):
                    continue
                text = normalize_card_line(raw)
                if text:
                    parts.append(text)
            elif t == "Div" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                nested = extract_comment_card_text(c[1])
                if nested:
                    parts.append(nested)
        return "\n\n".join(parts).strip()

    def parse_comment_callout_header(line: str) -> dict[str, str]:
        match = CARD_HEADER_RE.match(str(line or "").strip())
        if not match:
            return {}
        return {
            "kind": str(match.group("kind") or "").strip().upper(),
            "id": str(match.group("id") or "").strip(),
            "author": str(match.group("author") or "").strip(),
            "state": normalize_state_token(match.group("state")),
        }

    def parse_comment_card_payload_text(payload_text: str, parent_hint="") -> tuple[str, dict[str, str], str]:
        lines = str(payload_text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
        header_idx = None
        header: dict[str, str] = {}
        for idx, line in enumerate(lines):
            if not str(line or "").strip():
                continue
            parsed = parse_comment_callout_header(line)
            if parsed:
                header_idx = idx
                header = parsed
                break
        if header_idx is None:
            return "", {}, ""

        meta_line_idx = None
        meta_comment_id = ""
        parsed_meta: dict[str, str] = {}
        for idx, line in enumerate(lines):
            cid, meta = parse_card_meta_marker(line)
            if cid:
                meta_line_idx = idx
                meta_comment_id = cid
                parsed_meta = meta
                break

        comment_id = str(meta_comment_id or header.get("id") or "").strip()
        if not comment_id:
            return "", {}, ""

        out_meta: dict[str, str] = {}
        for key in [
            "author",
            "date",
            "parent",
            "state",
            "paraId",
            "durableId",
            "presenceProvider",
            "presenceUserId",
        ]:
            value = str(parsed_meta.get(key) or "").strip()
            if value:
                out_meta[key] = value

        out_meta.setdefault("author", str(header.get("author") or "").strip())
        out_meta["state"] = normalize_state_token(out_meta.get("state") or header.get("state"))
        if header.get("kind") == "REPLY" and parent_hint and not out_meta.get("parent"):
            out_meta["parent"] = str(parent_hint).strip()

        body_lines = []
        for idx, line in enumerate(lines):
            if idx == header_idx or idx == meta_line_idx:
                continue
            body_lines.append(line)
        out_meta["text"] = normalize_comment_text("\n".join(body_lines))
        return comment_id, out_meta, str(header.get("kind") or "")

    def parse_comment_card_blockquote(block, parent_hint="") -> list[tuple[str, dict[str, str]]]:
        if not (isinstance(block, dict) and block.get("t") == "BlockQuote"):
            return []
        quote_blocks = block.get("c")
        if not isinstance(quote_blocks, list) or not quote_blocks:
            return []

        primary_idx = None
        primary_block = None
        for idx, candidate in enumerate(quote_blocks):
            if not isinstance(candidate, dict):
                continue
            if candidate.get("t") in {"Para", "Plain", "Header", "RawBlock"}:
                primary_idx = idx
                primary_block = candidate
                break
        if primary_block is None:
            return []

        if primary_block.get("t") == "Header":
            payload = inlines_to_card_text(primary_block.get("c", [None, None, []])[2])
        elif primary_block.get("t") == "RawBlock":
            c = primary_block.get("c")
            fmt = str(c[0] or "").strip().lower() if isinstance(c, list) and len(c) == 2 else ""
            payload = str(c[1] or "") if fmt in {"markdown", "md"} else ""
        else:
            payload = inlines_to_card_text(primary_block.get("c"))

        comment_id, meta, kind = parse_comment_card_payload_text(payload, parent_hint=parent_hint)
        if not comment_id or kind not in {"COMMENT", "REPLY"}:
            return []

        detected_meta: dict[str, str] = {}
        detected_meta_id = ""
        for quote_block in quote_blocks:
            if not isinstance(quote_block, dict):
                continue
            qtype = quote_block.get("t")
            qdata = quote_block.get("c")
            marker_id = ""
            marker_meta: dict[str, str] = {}
            if qtype == "RawBlock" and isinstance(qdata, list) and len(qdata) == 2:
                fmt = str(qdata[0] or "").strip().lower()
                if fmt == "html":
                    marker_id, marker_meta = parse_card_meta_marker(str(qdata[1] or ""))
            elif qtype in {"Para", "Plain"} and isinstance(qdata, list):
                marker_id, marker_meta = parse_card_meta_marker(inlines_to_card_text(qdata))
            elif qtype == "Header" and isinstance(qdata, list) and len(qdata) >= 3 and isinstance(qdata[2], list):
                marker_id, marker_meta = parse_card_meta_marker(inlines_to_card_text(qdata[2]))
            if marker_id:
                if not detected_meta_id:
                    detected_meta_id = marker_id
                if marker_id == comment_id or not comment_id:
                    detected_meta.update(marker_meta)

        if detected_meta_id and detected_meta_id != comment_id:
            comment_id = detected_meta_id
        for key in [
            "author",
            "date",
            "parent",
            "state",
            "paraId",
            "durableId",
            "presenceProvider",
            "presenceUserId",
        ]:
            value = str(detected_meta.get(key) or "").strip()
            if value and not meta.get(key):
                meta[key] = value
        meta["state"] = normalize_state_token(meta.get("state"))
        if kind == "REPLY" and parent_hint and not meta.get("parent"):
            meta["parent"] = str(parent_hint).strip()

        parts = []
        body_text = normalize_comment_text(meta.get("text") or "")
        if body_text:
            parts.append(body_text)
        for idx, quote_block in enumerate(quote_blocks):
            if idx == primary_idx or not isinstance(quote_block, dict):
                continue
            if quote_block.get("t") == "BlockQuote":
                continue
            extra = extract_comment_card_text([quote_block])
            if extra:
                parts.append(extra)
        meta["text"] = normalize_comment_text("\n\n".join([part for part in parts if part]))

        entries = [(comment_id, meta)]
        for quote_block in quote_blocks:
            if not isinstance(quote_block, dict) or quote_block.get("t") != "BlockQuote":
                continue
            entries.extend(parse_comment_card_blockquote(quote_block, parent_hint=comment_id))
        return entries

    def ensure_id(comment_id: str) -> None:
        own_text_by_id.setdefault(comment_id, "")
        metadata_by_id.setdefault(comment_id, {})
        state_by_id.setdefault(comment_id, "active")

    def parse_attr(attr):
        if not (isinstance(attr, list) and len(attr) == 3):
            return None
        identifier = attr[0]
        classes = attr[1] or []
        kvs = attr[2] or []
        meta = {}
        if isinstance(kvs, list):
            for item in kvs:
                if isinstance(item, list) and len(item) == 2 and isinstance(item[0], str):
                    meta[item[0]] = item[1]
        return identifier, classes, meta

    def on_start(identifier: str, meta: dict, nested_inlines: list) -> None:
        nonlocal start_order
        comment_id = (identifier or meta.get("id") or "").strip()
        if not comment_id:
            return
        ensure_id(comment_id)
        card_meta = card_by_id.get(comment_id) or {}
        text = inlines_to_text(nested_inlines) or normalize_comment_text(card_meta.get("text") or "")
        existing = own_text_by_id.get(comment_id, "").strip()
        if text:
            if not existing:
                own_text_by_id[comment_id] = text
            elif text != existing and text not in existing:
                own_text_by_id[comment_id] = f"{existing}\n\n{text}"
        author = (meta.get("author") or card_meta.get("author") or "").strip()
        date = (meta.get("date") or card_meta.get("date") or "").strip()
        parent = (meta.get("parent") or card_meta.get("parent") or "").strip()
        state = normalize_state_token(meta.get("state") or card_meta.get("state") or "")
        if author and not metadata_by_id[comment_id].get("author"):
            metadata_by_id[comment_id]["author"] = author
        if date and not metadata_by_id[comment_id].get("date"):
            metadata_by_id[comment_id]["date"] = date
        if parent:
            parent_candidate_by_id[comment_id] = parent
        if "state" in meta or card_meta.get("state"):
            state_by_id[comment_id] = state
        starts.append(
            MarkdownCommentStart(
                id=comment_id,
                order=start_order,
                text=text,
                author=author,
                date=date,
                parent=parent,
                state=state,
            )
        )
        start_order += 1

    def on_end(identifier: str, meta: dict) -> None:
        comment_id = (meta.get("id") or identifier or "").strip()
        if comment_id:
            end_ids_order.append(comment_id)

    def walk_inlines(inlines) -> None:
        text_nodes = {"Str", "Space", "SoftBreak", "LineBreak"}
        i = 0
        while i < len(inlines or []):
            node = inlines[i]
            if not isinstance(node, dict):
                i += 1
                continue
            t = node.get("t")
            c = node.get("c")
            if t in text_nodes:
                j = i
                parts = []
                while j < len(inlines or []):
                    probe = inlines[j]
                    if not isinstance(probe, dict):
                        break
                    pt = probe.get("t")
                    if pt == "Str":
                        parts.append(probe.get("c") or "")
                    elif pt == "Space":
                        parts.append(" ")
                    elif pt in {"SoftBreak", "LineBreak"}:
                        parts.append("\n")
                    else:
                        break
                    j += 1
                chunk = "".join(parts)
                for match in MILESTONE_TOKEN_RE.finditer(chunk):
                    cid, edge = milestone_match_id_edge(match)
                    if edge == "s":
                        on_start(cid, {}, [])
                    elif edge == "e":
                        on_end(cid, {"id": cid})
                i = j
                continue
            if t == "Span" and isinstance(c, list) and len(c) == 2:
                parsed = parse_attr(c[0])
                nested = c[1] if isinstance(c[1], list) else []
                if parsed is not None:
                    identifier, classes, meta = parsed
                    if "comment-start" in classes:
                        on_start(identifier, meta, nested)
                        i += 1
                        continue
                    if "comment-end" in classes:
                        on_end(identifier, meta)
                        i += 1
                        continue
                walk_inlines(nested)
                i += 1
                continue
            if t == "Header" and isinstance(c, list) and len(c) >= 3 and isinstance(c[2], list):
                walk_inlines(c[2])
                i += 1
                continue
            if t in {"Link", "Image"} and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                walk_inlines(c[1])
                i += 1
                continue
            if isinstance(c, list):
                for item in c:
                    if isinstance(item, dict):
                        walk_inlines([item])
                    elif isinstance(item, list):
                        walk_inlines(item)
            i += 1

    def walk_blocks(blocks) -> None:
        for block in blocks or []:
            if not isinstance(block, dict):
                continue
            t = block.get("t")
            c = block.get("c")
            if t in {"Para", "Plain"} and isinstance(c, list):
                walk_inlines(c)
                continue
            if t == "Header" and isinstance(c, list) and len(c) >= 3 and isinstance(c[2], list):
                walk_inlines(c[2])
                continue
            if t == "BlockQuote" and isinstance(c, list):
                walk_blocks(c)
                continue
            if t == "Div" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                walk_blocks(c[1])
                continue
            if t in {"BulletList", "OrderedList"} and isinstance(c, list):
                items = c if t == "BulletList" else (c[1] if len(c) > 1 else [])
                for item in items:
                    walk_blocks(item)
                continue
            if isinstance(c, list):
                for item in c:
                    if isinstance(item, dict):
                        walk_blocks([item])
                    elif isinstance(item, list):
                        walk_blocks([x for x in item if isinstance(x, dict)])

    def collect_cards(blocks) -> None:
        for block in blocks or []:
            if not isinstance(block, dict):
                continue
            t = block.get("t")
            c = block.get("c")
            if t == "BlockQuote" and isinstance(c, list):
                entries = parse_comment_card_blockquote(block, parent_hint="")
                if entries:
                    for comment_id, meta in entries:
                        if not comment_id:
                            continue
                        card_by_id[comment_id] = {
                            "author": str(meta.get("author") or "").strip(),
                            "date": str(meta.get("date") or "").strip(),
                            "parent": str(meta.get("parent") or "").strip(),
                            "state": normalize_state_token(meta.get("state") or ""),
                            "text": normalize_comment_text(meta.get("text") or ""),
                        }
                    continue
                collect_cards(c)
                continue
            if t == "Div" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                collect_cards(c[1])
            elif t in {"BulletList", "OrderedList"} and isinstance(c, list):
                items = c if t == "BulletList" else (c[1] if len(c) > 1 else [])
                for item in items:
                    collect_cards(item)
            elif isinstance(c, list):
                for item in c:
                    if isinstance(item, list):
                        collect_cards([x for x in item if isinstance(x, dict)])
                    elif isinstance(item, dict):
                        collect_cards([item])

    collect_cards(doc.get("blocks", []))
    walk_blocks(doc.get("blocks", []))

    # Card-only comments (typically threaded replies) may not have milestone markers in prose.
    for comment_id, card_meta in card_by_id.items():
        if any(s.id == comment_id for s in starts):
            continue
        ensure_id(comment_id)
        text = normalize_comment_text(card_meta.get("text") or "")
        if text and not own_text_by_id.get(comment_id):
            own_text_by_id[comment_id] = text
        parent = (card_meta.get("parent") or "").strip()
        author = (card_meta.get("author") or "").strip()
        date = (card_meta.get("date") or "").strip()
        if author:
            metadata_by_id[comment_id]["author"] = author
        if date:
            metadata_by_id[comment_id]["date"] = date
        if parent:
            parent_candidate_by_id[comment_id] = parent
        state_by_id[comment_id] = normalize_state_token(card_meta.get("state") or "active")
        starts.append(
            MarkdownCommentStart(
                id=comment_id,
                order=start_order,
                text=text,
                author=author,
                date=date,
                parent=parent,
                state=state_by_id[comment_id],
            )
        )
        start_order += 1

    start_ids_order = [span.id for span in starts]
    started_ids = set(start_ids_order)
    parent_by_id: dict[str, str] = {}
    children_by_id: dict[str, list[str]] = {cid: [] for cid in started_ids}
    for child_id in start_ids_order:
        parent_id = (parent_candidate_by_id.get(child_id) or "").strip()
        if not parent_id or parent_id == child_id:
            continue
        if parent_id not in started_ids:
            continue
        parent_by_id[child_id] = parent_id
        children_by_id.setdefault(parent_id, [])
        if child_id not in children_by_id[parent_id]:
            children_by_id[parent_id].append(child_id)

    child_ids = sorted(parent_by_id.keys(), key=lambda cid: start_ids_order.index(cid))
    root_ids_order = [cid for cid in start_ids_order if cid not in parent_by_id]

    flattened_by_root: dict[str, str] = {}

    def flatten(comment_id: str, seen: set[str]) -> str:
        if comment_id in seen:
            return (own_text_by_id.get(comment_id) or "").strip()
        seen = set(seen)
        seen.add(comment_id)
        parts = []
        own_text = (own_text_by_id.get(comment_id) or "").strip()
        if own_text:
            parts.append(own_text)
        for child_id in children_by_id.get(comment_id, []):
            child_flat = flatten(child_id, seen)
            if child_flat:
                meta = metadata_by_id.get(child_id, {})
                parts.append(f"{reply_header(meta.get('author', ''), meta.get('date', ''))}\n{child_flat}")
        return "\n\n".join(parts).strip()

    for root_id in root_ids_order:
        flattened_by_root[root_id] = flatten(root_id, set())

    for cid in start_ids_order:
        state_by_id[cid] = normalize_state_token(state_by_id.get(cid, "active"))

    markdown_text = markdown_path.read_text(encoding="utf-8")
    placeholder_shape_match_count = 0
    for match in INLINE_IMAGE_RE.finditer(markdown_text):
        alt = (match.group("alt") or "").strip()
        title = (match.group("title") or "").strip().lower()
        if not alt and title == "shape":
            placeholder_shape_match_count += 1
    none_line_count = sum(1 for line in markdown_text.splitlines() if line.strip() == "None.")

    return MarkdownCommentSnapshot(
        path=str(markdown_path),
        starts=starts,
        start_ids_order=start_ids_order,
        end_ids_order=end_ids_order,
        parent_by_id=parent_by_id,
        child_ids=child_ids,
        root_ids_order=root_ids_order,
        own_text_by_id=own_text_by_id,
        flattened_by_root=flattened_by_root,
        state_by_id=state_by_id,
        placeholder_shape_match_count=placeholder_shape_match_count,
        none_line_count=none_line_count,
    )


def extract_comment_start_attrs(markdown_path: Path) -> dict[str, dict[str, str]]:
    text = Path(markdown_path).read_text(encoding="utf-8")
    attrs_by_id = {}
    for match in COMMENT_START_RE.finditer(text):
        attrs = {key: value for key, value in KV_ATTR_RE.findall(match.group("attrs"))}
        comment_id = (attrs.get("id") or "").strip()
        if comment_id:
            attrs_by_id[comment_id] = attrs
    return attrs_by_id
