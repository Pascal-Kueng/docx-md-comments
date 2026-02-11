from __future__ import annotations

import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


@dataclass(frozen=True)
class CommentNode:
    id: str
    author: str
    date: str
    text: str
    order: int
    parent_id: str = ""
    para_id: str = ""

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "author": self.author,
            "date": self.date,
            "text": self.text,
            "order": self.order,
            "parent_id": self.parent_id,
            "para_id": self.para_id,
        }


@dataclass(frozen=True)
class DocxCommentSnapshot:
    path: str
    comments_by_id: dict[str, CommentNode]
    comment_ids_order: list[str]
    parent_map: dict[str, str]
    children_by_id: dict[str, list[str]]
    root_ids_order: list[str]
    anchor_ids_order: list[str]
    range_start_ids: list[str]
    range_end_ids: list[str]
    reference_ids: list[str]
    anchor_text_by_id: dict[str, str]
    range_start_count_by_id: dict[str, int]
    range_end_count_by_id: dict[str, int]
    reference_count_by_id: dict[str, int]
    resolved_by_id: dict[str, bool]
    has_comments_extended: bool
    has_comments_ids: bool

    def to_dict(self) -> dict:
        return {
            "path": self.path,
            "comments_by_id": {cid: node.to_dict() for cid, node in self.comments_by_id.items()},
            "comment_ids_order": self.comment_ids_order,
            "parent_map": self.parent_map,
            "children_by_id": self.children_by_id,
            "root_ids_order": self.root_ids_order,
            "anchor_ids_order": self.anchor_ids_order,
            "range_start_ids": self.range_start_ids,
            "range_end_ids": self.range_end_ids,
            "reference_ids": self.reference_ids,
            "anchor_text_by_id": self.anchor_text_by_id,
            "range_start_count_by_id": self.range_start_count_by_id,
            "range_end_count_by_id": self.range_end_count_by_id,
            "reference_count_by_id": self.reference_count_by_id,
            "resolved_by_id": self.resolved_by_id,
            "has_comments_extended": self.has_comments_extended,
            "has_comments_ids": self.has_comments_ids,
        }


@dataclass(frozen=True)
class FlattenExpectation:
    root_ids_order: list[str]
    child_ids: list[str]
    flattened_by_root: dict[str, str]

    def to_dict(self) -> dict:
        return {
            "root_ids_order": self.root_ids_order,
            "child_ids": self.child_ids,
            "flattened_by_root": self.flattened_by_root,
        }


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def get_attr_local(elem: ET.Element, attr_name: str) -> str | None:
    for key, value in elem.attrib.items():
        if key == attr_name or key.endswith("}" + attr_name):
            return value
    return None


def normalize_comment_text(text: str) -> str:
    normalized = (text or "").replace("\r\n", "\n").replace("\r", "\n").replace("\xa0", " ")
    normalized = unicodedata.normalize("NFKC", normalized)
    lines = []
    for raw_line in normalized.split("\n"):
        line = re.sub(r"[ \t]+", " ", raw_line).strip()
        if line:
            lines.append(line)
    return "\n".join(lines).strip()


def normalize_anchor_text(text: str) -> str:
    # Anchor extraction across run/paragraph boundaries can vary in whether
    # boundary whitespace is represented, while still preserving the same span.
    # Compare anchors whitespace-insensitively to avoid false positives.
    normalized = normalize_comment_text(text)
    return re.sub(r"\s+", "", normalized)


def extract_comment_text(comment_elem: ET.Element) -> str:
    paragraphs = []
    for p in comment_elem.iter(f"{{{W_NS}}}p"):
        parts = []
        for node in p.iter():
            lname = local_name(node.tag)
            if lname == "t" and node.text:
                parts.append(node.text)
            elif lname == "tab":
                parts.append("\t")
            elif lname in {"br", "cr"}:
                parts.append("\n")
        text = "".join(parts).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs).strip()


def unique_in_order(values: list[str]) -> list[str]:
    seen = set()
    ordered = []
    for value in values:
        if value not in seen:
            seen.add(value)
            ordered.append(value)
    return ordered


def story_xml_names(zip_file: zipfile.ZipFile) -> list[str]:
    out = []
    for name in zip_file.namelist():
        if not (name.startswith("word/") and name.endswith(".xml")):
            continue
        basename = Path(name).name
        if (
            basename == "document.xml"
            or basename in {"footnotes.xml", "endnotes.xml"}
            or re.match(r"^header[0-9]+\.xml$", basename)
            or re.match(r"^footer[0-9]+\.xml$", basename)
        ):
            out.append(name)
    return sorted(out)


def inspect_docx(docx_path: Path) -> DocxCommentSnapshot:
    docx_path = Path(docx_path)
    comments_by_id: dict[str, CommentNode] = {}
    comment_ids_order: list[str] = []
    parent_map: dict[str, str] = {}
    para_to_id: dict[str, str] = {}
    resolved_by_id: dict[str, bool] = {}

    with zipfile.ZipFile(docx_path, "r") as zip_file:
        names = set(zip_file.namelist())
        has_comments_extended = "word/commentsExtended.xml" in names
        has_comments_ids = "word/commentsIds.xml" in names

        if "word/comments.xml" in names:
            comments_root = ET.fromstring(zip_file.read("word/comments.xml"))
            for idx, comment in enumerate(comments_root.findall(f".//{{{W_NS}}}comment")):
                cid = get_attr_local(comment, "id")
                if cid is None:
                    continue
                node = CommentNode(
                    id=cid,
                    author=get_attr_local(comment, "author") or "",
                    date=get_attr_local(comment, "date") or "",
                    text=extract_comment_text(comment),
                    order=idx,
                    parent_id=get_attr_local(comment, "parentId") or "",
                    para_id=get_attr_local(comment, "paraId") or "",
                )
                if not node.para_id:
                    first_p = comment.find(f".//{{{W_NS}}}p")
                    if first_p is not None:
                        node = CommentNode(
                            id=node.id,
                            author=node.author,
                            date=node.date,
                            text=node.text,
                            order=node.order,
                            parent_id=node.parent_id,
                            para_id=get_attr_local(first_p, "paraId") or "",
                        )
                comments_by_id[cid] = node
                comment_ids_order.append(cid)
                if node.parent_id:
                    parent_map[cid] = node.parent_id
                if node.para_id:
                    para_to_id[node.para_id] = cid
                resolved_by_id[cid] = False

        if has_comments_ids and comment_ids_order:
            ids_root = ET.fromstring(zip_file.read("word/commentsIds.xml"))
            para_ids = []
            for elem in ids_root.iter():
                if local_name(elem.tag) != "commentId":
                    continue
                para_id = get_attr_local(elem, "paraId")
                if para_id:
                    para_ids.append(para_id)
            if len(para_ids) == len(comment_ids_order):
                for comment_id, para_id in zip(comment_ids_order, para_ids):
                    if para_id and para_id not in para_to_id:
                        para_to_id[para_id] = comment_id

        if has_comments_extended and para_to_id:
            ext_root = ET.fromstring(zip_file.read("word/commentsExtended.xml"))
            for elem in ext_root.iter():
                if local_name(elem.tag) != "commentEx":
                    continue
                child_para = get_attr_local(elem, "paraId")
                parent_para = get_attr_local(elem, "paraIdParent")
                done = get_attr_local(elem, "done")
                child_id = para_to_id.get(child_para or "")
                parent_id = para_to_id.get(parent_para or "")
                if child_id:
                    resolved_by_id[child_id] = str(done or "").strip() == "1"
                if child_id and parent_id and child_id not in parent_map:
                    parent_map[child_id] = parent_id

        parent_map = {
            child: parent
            for child, parent in parent_map.items()
            if child in comments_by_id and parent in comments_by_id and child != parent
        }

        children_by_id = {cid: [] for cid in comments_by_id}
        for child_id, parent_id in parent_map.items():
            children_by_id[parent_id].append(child_id)
        for siblings in children_by_id.values():
            siblings.sort(key=lambda cid: comments_by_id[cid].order)

        anchors = []
        range_start = []
        range_end = []
        references = []
        anchor_text_parts_by_id: dict[str, list[str]] = {}
        range_start_count_by_id: dict[str, int] = {}
        range_end_count_by_id: dict[str, int] = {}
        reference_count_by_id: dict[str, int] = {}
        for story_name in story_xml_names(zip_file):
            root = ET.fromstring(zip_file.read(story_name))
            active_comment_ids: list[str] = []
            for elem in root.iter():
                lname = local_name(elem.tag)
                if lname == "commentRangeStart":
                    cid = get_attr_local(elem, "id")
                    if cid is not None:
                        anchors.append(cid)
                        range_start.append(cid)
                        range_start_count_by_id[cid] = range_start_count_by_id.get(cid, 0) + 1
                        if cid not in active_comment_ids:
                            active_comment_ids.append(cid)
                elif lname == "commentRangeEnd":
                    cid = get_attr_local(elem, "id")
                    if cid is not None:
                        range_end.append(cid)
                        range_end_count_by_id[cid] = range_end_count_by_id.get(cid, 0) + 1
                        if cid in active_comment_ids:
                            active_comment_ids.remove(cid)
                elif lname == "commentReference":
                    cid = get_attr_local(elem, "id")
                    if cid is not None:
                        references.append(cid)
                        reference_count_by_id[cid] = reference_count_by_id.get(cid, 0) + 1
                elif lname == "t" and elem.text:
                    for cid in active_comment_ids:
                        anchor_text_parts_by_id.setdefault(cid, []).append(elem.text)
                elif lname == "tab":
                    for cid in active_comment_ids:
                        anchor_text_parts_by_id.setdefault(cid, []).append("\t")
                elif lname in {"br", "cr"}:
                    for cid in active_comment_ids:
                        anchor_text_parts_by_id.setdefault(cid, []).append("\n")

    root_ids_order = [cid for cid in comment_ids_order if cid not in parent_map]
    anchor_text_by_id = {
        cid: "".join(parts).strip()
        for cid, parts in anchor_text_parts_by_id.items()
    }

    return DocxCommentSnapshot(
        path=str(docx_path),
        comments_by_id=comments_by_id,
        comment_ids_order=comment_ids_order,
        parent_map=parent_map,
        children_by_id=children_by_id,
        root_ids_order=root_ids_order,
        anchor_ids_order=unique_in_order(anchors),
        range_start_ids=unique_in_order(range_start),
        range_end_ids=unique_in_order(range_end),
        reference_ids=unique_in_order(references),
        anchor_text_by_id=anchor_text_by_id,
        range_start_count_by_id=range_start_count_by_id,
        range_end_count_by_id=range_end_count_by_id,
        reference_count_by_id=reference_count_by_id,
        resolved_by_id=resolved_by_id,
        has_comments_extended=has_comments_extended,
        has_comments_ids=has_comments_ids,
    )


def thread_root(comment_id: str, snapshot: DocxCommentSnapshot) -> str:
    root_id = comment_id
    seen = set()
    while root_id in snapshot.parent_map and root_id not in seen:
        seen.add(root_id)
        parent_id = snapshot.parent_map[root_id]
        if parent_id not in snapshot.comments_by_id:
            break
        root_id = parent_id
    return root_id


def reply_header(comment: CommentNode) -> str:
    author = (comment.author or "Unknown").strip() or "Unknown"
    date = (comment.date or "").strip()
    if date:
        return f"---\nReply from: {author} ({date})\n---"
    return f"---\nReply from: {author}\n---"


def flatten_comment(comment_id: str, snapshot: DocxCommentSnapshot, seen: set[str]) -> str:
    if comment_id in seen:
        return (snapshot.comments_by_id.get(comment_id) or CommentNode("", "", "", "", 0)).text.strip()

    seen = set(seen)
    seen.add(comment_id)
    comment = snapshot.comments_by_id.get(comment_id)
    if comment is None:
        return ""

    parts = []
    own = (comment.text or "").strip()
    if own:
        parts.append(own)

    for child_id in snapshot.children_by_id.get(comment_id, []):
        child_flat = flatten_comment(child_id, snapshot, seen)
        if child_flat:
            child = snapshot.comments_by_id.get(child_id)
            if child is not None:
                parts.append(f"{reply_header(child)}\n{child_flat}")

    return "\n\n".join(parts).strip()


def build_flatten_expectation(snapshot: DocxCommentSnapshot) -> FlattenExpectation:
    root_ids_order = []
    seen_roots = set()
    for anchor_id in snapshot.anchor_ids_order:
        if anchor_id not in snapshot.comments_by_id:
            continue
        root_id = thread_root(anchor_id, snapshot)
        if root_id not in seen_roots:
            seen_roots.add(root_id)
            root_ids_order.append(root_id)
    for root_id in snapshot.root_ids_order:
        if root_id not in seen_roots:
            seen_roots.add(root_id)
            root_ids_order.append(root_id)

    child_ids = sorted(
        [cid for cid in snapshot.parent_map.keys() if cid in snapshot.comments_by_id],
        key=lambda cid: snapshot.comments_by_id[cid].order,
    )
    flattened_by_root = {
        root_id: flatten_comment(root_id, snapshot, set())
        for root_id in root_ids_order
    }
    return FlattenExpectation(
        root_ids_order=root_ids_order,
        child_ids=child_ids,
        flattened_by_root=flattened_by_root,
    )
