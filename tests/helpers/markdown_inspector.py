from __future__ import annotations

import json
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path

COMMENT_START_RE = re.compile(r"\{\.comment-start(?P<attrs>[^}]*)\}", re.DOTALL)
KV_ATTR_RE = re.compile(r'([A-Za-z_:][-A-Za-z0-9_:.]*)="([^"]*)"')
INLINE_IMAGE_RE = re.compile(
    r'!\[(?P<alt>[^\]]*)\]\((?P<src>[^)\s]+)(?:\s+"(?P<title>[^"]*)")?\)\{(?P<attrs>[^}]*)\}',
    re.DOTALL,
)


@dataclass(frozen=True)
class MarkdownCommentStart:
    id: str
    order: int
    text: str
    author: str
    date: str
    parent: str

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "order": self.order,
            "text": self.text,
            "author": self.author,
            "date": self.date,
            "parent": self.parent,
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
            "placeholder_shape_match_count": self.placeholder_shape_match_count,
            "none_line_count": self.none_line_count,
        }


def normalize_comment_text(text: str) -> str:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\\+[ \t]*\n", "\n", text)
    text = re.sub(r"\\\\[ \t]+", "\n", text)
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
        elif t in {"Quoted", "Cite"}:
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


def reply_header(author: str, date: str) -> str:
    safe_author = (author or "Unknown").strip() or "Unknown"
    safe_date = (date or "").strip()
    if safe_date:
        return f"---\nReply from: {safe_author} ({safe_date})\n---"
    return f"---\nReply from: {safe_author}\n---"


def run_pandoc_json(markdown_path: Path) -> dict:
    cmd = ["pandoc", str(markdown_path), "-f", "markdown", "-t", "json"]
    out = subprocess.check_output(cmd, text=True)
    return json.loads(out)


def inspect_markdown_comments(markdown_path: Path) -> MarkdownCommentSnapshot:
    markdown_path = Path(markdown_path)
    doc = run_pandoc_json(markdown_path)

    starts: list[MarkdownCommentStart] = []
    end_ids_order: list[str] = []
    own_text_by_id: dict[str, str] = {}
    metadata_by_id: dict[str, dict[str, str]] = {}
    parent_candidate_by_id: dict[str, str] = {}
    start_order = 0

    def ensure_id(comment_id: str) -> None:
        own_text_by_id.setdefault(comment_id, "")
        metadata_by_id.setdefault(comment_id, {})

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
        comment_id = identifier or (meta.get("id") or "")
        comment_id = comment_id.strip()
        if not comment_id:
            return
        ensure_id(comment_id)
        text = inlines_to_text(nested_inlines)
        existing = own_text_by_id.get(comment_id, "").strip()
        if text:
            if not existing:
                own_text_by_id[comment_id] = text
            elif text != existing and text not in existing:
                own_text_by_id[comment_id] = f"{existing}\n\n{text}"
        author = (meta.get("author") or "").strip()
        date = (meta.get("date") or "").strip()
        parent = (meta.get("parent") or "").strip()
        if author and not metadata_by_id[comment_id].get("author"):
            metadata_by_id[comment_id]["author"] = author
        if date and not metadata_by_id[comment_id].get("date"):
            metadata_by_id[comment_id]["date"] = date
        if parent:
            parent_candidate_by_id[comment_id] = parent
        starts.append(
            MarkdownCommentStart(
                id=comment_id,
                order=start_order,
                text=text,
                author=author,
                date=date,
                parent=parent,
            )
        )
        start_order += 1

    def on_end(identifier: str, meta: dict) -> None:
        comment_id = (meta.get("id") or identifier or "").strip()
        if comment_id:
            end_ids_order.append(comment_id)

    def walk_inlines(inlines) -> None:
        for node in inlines or []:
            if not isinstance(node, dict):
                continue
            t = node.get("t")
            c = node.get("c")
            if t == "Span" and isinstance(c, list) and len(c) == 2:
                parsed = parse_attr(c[0])
                nested = c[1] if isinstance(c[1], list) else []
                if parsed is not None:
                    identifier, classes, meta = parsed
                    if "comment-start" in classes:
                        on_start(identifier, meta, nested)
                        continue
                    if "comment-end" in classes:
                        on_end(identifier, meta)
                        continue
                walk_inlines(nested)
                continue
            if t == "Header" and isinstance(c, list) and len(c) >= 3 and isinstance(c[2], list):
                walk_inlines(c[2])
                continue
            if t in {"Link", "Image"} and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                walk_inlines(c[1])
                continue
            if isinstance(c, list):
                for item in c:
                    if isinstance(item, dict):
                        walk_inlines([item])
                    elif isinstance(item, list):
                        walk_inlines(item)

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

    walk_blocks(doc.get("blocks", []))

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
