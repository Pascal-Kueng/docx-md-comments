from __future__ import annotations

import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
COMMENTS_EXT_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
COMMENTS_IDS_REL_TYPE = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
COMMENTS_EXTENSIBLE_REL_TYPE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
PEOPLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/people"
COMMENTS_EXT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
COMMENTS_IDS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
COMMENTS_EXTENSIBLE_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"
)
PEOPLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"


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
    has_comments_extensible: bool
    has_people: bool
    has_comments_extended_rel: bool
    has_comments_ids_rel: bool
    has_comments_extensible_rel: bool
    has_people_rel: bool
    has_comments_extended_content_type: bool
    has_comments_ids_content_type: bool
    has_comments_extensible_content_type: bool
    has_people_content_type: bool
    has_comments_xml_w15_ns: bool
    has_comments_xml_w14_ns: bool
    has_comments_xml_w16cid_ns: bool
    has_comments_xml_w16cex_ns: bool
    comments_xml_ignorable_tokens: list[str]
    comments_extended_ignorable_tokens: list[str]
    comments_ids_ignorable_tokens: list[str]
    comments_extensible_ignorable_tokens: list[str]
    has_settings_xml_w15_ns: bool
    has_settings_xml_w14_ns: bool
    settings_xml_ignorable_tokens: list[str]
    settings_compatibility_mode: str
    people_presence_provider_by_author: dict[str, str]
    people_presence_user_by_author: dict[str, str]
    first_paragraph_para_by_id: dict[str, str]
    last_paragraph_para_by_id: dict[str, str]
    paragraph_count_by_id: dict[str, int]
    annotation_ref_count_by_id: dict[str, int]
    comment_state_attr_by_id: dict[str, str]
    comment_parent_attr_by_id: dict[str, str]
    comment_para_attr_by_id: dict[str, str]
    comment_durable_attr_by_id: dict[str, str]
    comments_extended_para_ids: list[str]
    comments_extended_parent_para_ids: list[str]
    comments_ids_para_ids: list[str]
    comments_ids_durable_ids: list[str]
    comments_ids_durable_by_para: dict[str, str]
    comments_extensible_durable_ids: list[str]

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
            "has_comments_extensible": self.has_comments_extensible,
            "has_people": self.has_people,
            "has_comments_extended_rel": self.has_comments_extended_rel,
            "has_comments_ids_rel": self.has_comments_ids_rel,
            "has_comments_extensible_rel": self.has_comments_extensible_rel,
            "has_people_rel": self.has_people_rel,
            "has_comments_extended_content_type": self.has_comments_extended_content_type,
            "has_comments_ids_content_type": self.has_comments_ids_content_type,
            "has_comments_extensible_content_type": self.has_comments_extensible_content_type,
            "has_people_content_type": self.has_people_content_type,
            "has_comments_xml_w15_ns": self.has_comments_xml_w15_ns,
            "has_comments_xml_w14_ns": self.has_comments_xml_w14_ns,
            "has_comments_xml_w16cid_ns": self.has_comments_xml_w16cid_ns,
            "has_comments_xml_w16cex_ns": self.has_comments_xml_w16cex_ns,
            "comments_xml_ignorable_tokens": self.comments_xml_ignorable_tokens,
            "comments_extended_ignorable_tokens": self.comments_extended_ignorable_tokens,
            "comments_ids_ignorable_tokens": self.comments_ids_ignorable_tokens,
            "comments_extensible_ignorable_tokens": self.comments_extensible_ignorable_tokens,
            "has_settings_xml_w15_ns": self.has_settings_xml_w15_ns,
            "has_settings_xml_w14_ns": self.has_settings_xml_w14_ns,
            "settings_xml_ignorable_tokens": self.settings_xml_ignorable_tokens,
            "settings_compatibility_mode": self.settings_compatibility_mode,
            "people_presence_provider_by_author": self.people_presence_provider_by_author,
            "people_presence_user_by_author": self.people_presence_user_by_author,
            "first_paragraph_para_by_id": self.first_paragraph_para_by_id,
            "last_paragraph_para_by_id": self.last_paragraph_para_by_id,
            "paragraph_count_by_id": self.paragraph_count_by_id,
            "annotation_ref_count_by_id": self.annotation_ref_count_by_id,
            "comment_state_attr_by_id": self.comment_state_attr_by_id,
            "comment_parent_attr_by_id": self.comment_parent_attr_by_id,
            "comment_para_attr_by_id": self.comment_para_attr_by_id,
            "comment_durable_attr_by_id": self.comment_durable_attr_by_id,
            "comments_extended_para_ids": self.comments_extended_para_ids,
            "comments_extended_parent_para_ids": self.comments_extended_parent_para_ids,
            "comments_ids_para_ids": self.comments_ids_para_ids,
            "comments_ids_durable_ids": self.comments_ids_durable_ids,
            "comments_ids_durable_by_para": self.comments_ids_durable_by_para,
            "comments_extensible_durable_ids": self.comments_extensible_durable_ids,
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
        has_comments_extensible = "word/commentsExtensible.xml" in names
        has_people = "word/people.xml" in names
        has_comments_extended_rel = False
        has_comments_ids_rel = False
        has_comments_extensible_rel = False
        has_people_rel = False
        has_comments_extended_content_type = False
        has_comments_ids_content_type = False
        has_comments_extensible_content_type = False
        has_people_content_type = False
        comments_extended_para_ids: set[str] = set()
        comments_extended_parent_para_ids: set[str] = set()
        comments_ids_para_ids: set[str] = set()
        comments_ids_durable_ids: set[str] = set()
        comments_ids_durable_by_para: dict[str, str] = {}
        comments_extensible_durable_ids: set[str] = set()
        has_comments_xml_w15_ns = False
        has_comments_xml_w14_ns = False
        has_comments_xml_w16cid_ns = False
        has_comments_xml_w16cex_ns = False
        comments_xml_ignorable_tokens: list[str] = []
        comments_extended_ignorable_tokens: list[str] = []
        comments_ids_ignorable_tokens: list[str] = []
        comments_extensible_ignorable_tokens: list[str] = []
        has_settings_xml_w15_ns = False
        has_settings_xml_w14_ns = False
        settings_xml_ignorable_tokens: list[str] = []
        settings_compatibility_mode = ""
        people_presence_provider_by_author: dict[str, str] = {}
        people_presence_user_by_author: dict[str, str] = {}
        first_paragraph_para_by_id: dict[str, str] = {}
        last_paragraph_para_by_id: dict[str, str] = {}
        paragraph_count_by_id: dict[str, int] = {}
        annotation_ref_count_by_id: dict[str, int] = {}
        comment_state_attr_by_id: dict[str, str] = {}
        comment_parent_attr_by_id: dict[str, str] = {}
        comment_para_attr_by_id: dict[str, str] = {}
        comment_durable_attr_by_id: dict[str, str] = {}

        if "word/_rels/document.xml.rels" in names:
            rel_root = ET.fromstring(zip_file.read("word/_rels/document.xml.rels"))
            for rel in rel_root.iter():
                if local_name(rel.tag) != "Relationship":
                    continue
                rel_type = rel.attrib.get("Type", "")
                if rel_type == COMMENTS_EXT_REL_TYPE:
                    has_comments_extended_rel = True
                elif rel_type == COMMENTS_IDS_REL_TYPE:
                    has_comments_ids_rel = True
                elif rel_type == COMMENTS_EXTENSIBLE_REL_TYPE:
                    has_comments_extensible_rel = True
                elif rel_type == PEOPLE_REL_TYPE:
                    has_people_rel = True

        if "[Content_Types].xml" in names:
            content_root = ET.fromstring(zip_file.read("[Content_Types].xml"))
            override_by_part = {}
            for elem in content_root.iter():
                if local_name(elem.tag) != "Override":
                    continue
                part_name = elem.attrib.get("PartName", "")
                ctype = elem.attrib.get("ContentType", "")
                if part_name:
                    override_by_part[part_name] = ctype
            has_comments_extended_content_type = (
                override_by_part.get("/word/commentsExtended.xml") == COMMENTS_EXT_CONTENT_TYPE
            )
            has_comments_ids_content_type = (
                override_by_part.get("/word/commentsIds.xml") == COMMENTS_IDS_CONTENT_TYPE
            )
            has_comments_extensible_content_type = (
                override_by_part.get("/word/commentsExtensible.xml") == COMMENTS_EXTENSIBLE_CONTENT_TYPE
            )
            has_people_content_type = (
                override_by_part.get("/word/people.xml") == PEOPLE_CONTENT_TYPE
            )

        if "word/comments.xml" in names:
            comments_xml_raw = zip_file.read("word/comments.xml").decode("utf-8", errors="replace")
            has_comments_xml_w15_ns = 'xmlns:w15="' in comments_xml_raw
            has_comments_xml_w14_ns = 'xmlns:w14="' in comments_xml_raw
            has_comments_xml_w16cid_ns = 'xmlns:w16cid="' in comments_xml_raw
            has_comments_xml_w16cex_ns = 'xmlns:w16cex="' in comments_xml_raw
            ignorable_match = re.search(r'\b(?:mc:)?Ignorable="([^"]*)"', comments_xml_raw)
            if ignorable_match:
                comments_xml_ignorable_tokens = sorted(
                    [token for token in (ignorable_match.group(1) or "").split() if token]
                )

            comments_root = ET.fromstring(comments_xml_raw)
            for idx, comment in enumerate(comments_root.findall(f".//{{{W_NS}}}comment")):
                cid = get_attr_local(comment, "id")
                if cid is None:
                    continue
                comment_state_attr_by_id[cid] = get_attr_local(comment, "state") or ""
                comment_parent_attr_by_id[cid] = get_attr_local(comment, "parentId") or ""
                comment_para_attr_by_id[cid] = get_attr_local(comment, "paraId") or ""
                comment_durable_attr_by_id[cid] = get_attr_local(comment, "durableId") or ""
                paragraph_para_ids = []
                paragraphs = comment.findall(f"./{{{W_NS}}}p")
                paragraph_count_by_id[cid] = len(paragraphs)
                annotation_ref_count_by_id[cid] = len(list(comment.iter(f"{{{W_NS}}}annotationRef")))
                for paragraph in paragraphs:
                    paragraph_para_id = get_attr_local(paragraph, "paraId") or ""
                    if paragraph_para_id:
                        paragraph_para_ids.append(paragraph_para_id)
                        para_to_id[paragraph_para_id] = cid
                first_paragraph_para_by_id[cid] = paragraph_para_ids[0] if paragraph_para_ids else ""
                last_paragraph_para_by_id[cid] = paragraph_para_ids[-1] if paragraph_para_ids else ""
                thread_para_id = paragraph_para_ids[-1] if paragraph_para_ids else ""
                node = CommentNode(
                    id=cid,
                    author=get_attr_local(comment, "author") or "",
                    date=get_attr_local(comment, "date") or "",
                    text=extract_comment_text(comment),
                    order=idx,
                    parent_id=get_attr_local(comment, "parentId") or "",
                    para_id=thread_para_id or (get_attr_local(comment, "paraId") or ""),
                )
                comments_by_id[cid] = node
                comment_ids_order.append(cid)
                if node.parent_id:
                    parent_map[cid] = node.parent_id
                if node.para_id:
                    para_to_id[node.para_id] = cid
                resolved_by_id[cid] = False

        if "word/commentsExtended.xml" in names:
            comments_extended_raw = zip_file.read("word/commentsExtended.xml").decode("utf-8", errors="replace")
            ignorable_match = re.search(r'\b(?:mc:)?Ignorable="([^"]*)"', comments_extended_raw)
            if ignorable_match:
                comments_extended_ignorable_tokens = sorted(
                    [token for token in (ignorable_match.group(1) or "").split() if token]
                )

        if "word/commentsIds.xml" in names:
            comments_ids_raw = zip_file.read("word/commentsIds.xml").decode("utf-8", errors="replace")
            ignorable_match = re.search(r'\b(?:mc:)?Ignorable="([^"]*)"', comments_ids_raw)
            if ignorable_match:
                comments_ids_ignorable_tokens = sorted(
                    [token for token in (ignorable_match.group(1) or "").split() if token]
                )

        if "word/commentsExtensible.xml" in names:
            comments_extensible_raw = zip_file.read("word/commentsExtensible.xml").decode("utf-8", errors="replace")
            ignorable_match = re.search(r'\b(?:mc:)?Ignorable="([^"]*)"', comments_extensible_raw)
            if ignorable_match:
                comments_extensible_ignorable_tokens = sorted(
                    [token for token in (ignorable_match.group(1) or "").split() if token]
                )

        if "word/settings.xml" in names:
            settings_xml_raw = zip_file.read("word/settings.xml").decode("utf-8", errors="replace")
            has_settings_xml_w15_ns = 'xmlns:w15="' in settings_xml_raw
            has_settings_xml_w14_ns = 'xmlns:w14="' in settings_xml_raw
            ignorable_match = re.search(r'\b(?:mc:)?Ignorable="([^"]*)"', settings_xml_raw)
            if ignorable_match:
                settings_xml_ignorable_tokens = sorted(
                    [token for token in (ignorable_match.group(1) or "").split() if token]
                )
            settings_root = ET.fromstring(settings_xml_raw)
            for setting in settings_root.iter(f"{{{W_NS}}}compatSetting"):
                name = (get_attr_local(setting, "name") or "").strip()
                uri = (get_attr_local(setting, "uri") or "").strip()
                if name == "compatibilityMode" and uri == "http://schemas.microsoft.com/office/word":
                    settings_compatibility_mode = (get_attr_local(setting, "val") or "").strip()
                    break

        if has_comments_ids and comment_ids_order:
            ids_root = ET.fromstring(zip_file.read("word/commentsIds.xml"))
            para_ids = []
            for elem in ids_root.iter():
                if local_name(elem.tag) != "commentId":
                    continue
                para_id = get_attr_local(elem, "paraId")
                durable_id = get_attr_local(elem, "durableId")
                if para_id:
                    para_ids.append(para_id)
                    comments_ids_para_ids.add(para_id)
                if durable_id:
                    comments_ids_durable_ids.add(durable_id)
                if para_id and durable_id:
                    comments_ids_durable_by_para[para_id] = durable_id
            if len(para_ids) == len(comment_ids_order):
                for comment_id, para_id in zip(comment_ids_order, para_ids):
                    if para_id and para_id not in para_to_id:
                        para_to_id[para_id] = comment_id
                    if para_id and comment_id in comments_by_id:
                        node = comments_by_id[comment_id]
                        comments_by_id[comment_id] = CommentNode(
                            id=node.id,
                            author=node.author,
                            date=node.date,
                            text=node.text,
                            order=node.order,
                            parent_id=node.parent_id,
                            para_id=para_id,
                        )

        if has_comments_extended and para_to_id:
            ext_root = ET.fromstring(zip_file.read("word/commentsExtended.xml"))
            for elem in ext_root.iter():
                if local_name(elem.tag) != "commentEx":
                    continue
                child_para = get_attr_local(elem, "paraId")
                parent_para = get_attr_local(elem, "paraIdParent")
                done = get_attr_local(elem, "done")
                if child_para:
                    comments_extended_para_ids.add(child_para)
                if parent_para:
                    comments_extended_parent_para_ids.add(parent_para)
                child_id = para_to_id.get(child_para or "")
                parent_id = para_to_id.get(parent_para or "")
                if child_id:
                    resolved_by_id[child_id] = str(done or "").strip() == "1"
                    node = comments_by_id.get(child_id)
                    if node is not None and child_para:
                        comments_by_id[child_id] = CommentNode(
                            id=node.id,
                            author=node.author,
                            date=node.date,
                            text=node.text,
                            order=node.order,
                            parent_id=node.parent_id,
                            para_id=child_para,
                        )
                if child_id and parent_id and child_id not in parent_map:
                    parent_map[child_id] = parent_id

        if has_comments_extensible:
            extensible_root = ET.fromstring(zip_file.read("word/commentsExtensible.xml"))
            for elem in extensible_root.iter():
                if local_name(elem.tag) != "commentExtensible":
                    continue
                durable_id = get_attr_local(elem, "durableId")
                if durable_id:
                    comments_extensible_durable_ids.add(durable_id)

        if has_people:
            people_root = ET.fromstring(zip_file.read("word/people.xml"))
            for person in people_root.iter():
                if local_name(person.tag) != "person":
                    continue
                author = (get_attr_local(person, "author") or "").strip()
                if not author:
                    continue
                provider_id = ""
                user_id = ""
                for child in list(person):
                    if local_name(child.tag) != "presenceInfo":
                        continue
                    provider_id = (get_attr_local(child, "providerId") or "").strip()
                    user_id = (get_attr_local(child, "userId") or "").strip()
                    break
                people_presence_provider_by_author[author] = provider_id
                people_presence_user_by_author[author] = user_id

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
        has_comments_extensible=has_comments_extensible,
        has_people=has_people,
        has_comments_extended_rel=has_comments_extended_rel,
        has_comments_ids_rel=has_comments_ids_rel,
        has_comments_extensible_rel=has_comments_extensible_rel,
        has_people_rel=has_people_rel,
        has_comments_extended_content_type=has_comments_extended_content_type,
        has_comments_ids_content_type=has_comments_ids_content_type,
        has_comments_extensible_content_type=has_comments_extensible_content_type,
        has_people_content_type=has_people_content_type,
        has_comments_xml_w15_ns=has_comments_xml_w15_ns,
        has_comments_xml_w14_ns=has_comments_xml_w14_ns,
        has_comments_xml_w16cid_ns=has_comments_xml_w16cid_ns,
        has_comments_xml_w16cex_ns=has_comments_xml_w16cex_ns,
        comments_xml_ignorable_tokens=comments_xml_ignorable_tokens,
        comments_extended_ignorable_tokens=comments_extended_ignorable_tokens,
        comments_ids_ignorable_tokens=comments_ids_ignorable_tokens,
        comments_extensible_ignorable_tokens=comments_extensible_ignorable_tokens,
        has_settings_xml_w15_ns=has_settings_xml_w15_ns,
        has_settings_xml_w14_ns=has_settings_xml_w14_ns,
        settings_xml_ignorable_tokens=settings_xml_ignorable_tokens,
        settings_compatibility_mode=settings_compatibility_mode,
        people_presence_provider_by_author=people_presence_provider_by_author,
        people_presence_user_by_author=people_presence_user_by_author,
        first_paragraph_para_by_id=first_paragraph_para_by_id,
        last_paragraph_para_by_id=last_paragraph_para_by_id,
        paragraph_count_by_id=paragraph_count_by_id,
        annotation_ref_count_by_id=annotation_ref_count_by_id,
        comment_state_attr_by_id=comment_state_attr_by_id,
        comment_parent_attr_by_id=comment_parent_attr_by_id,
        comment_para_attr_by_id=comment_para_attr_by_id,
        comment_durable_attr_by_id=comment_durable_attr_by_id,
        comments_extended_para_ids=sorted(comments_extended_para_ids),
        comments_extended_parent_para_ids=sorted(comments_extended_parent_para_ids),
        comments_ids_para_ids=sorted(comments_ids_para_ids),
        comments_ids_durable_ids=sorted(comments_ids_durable_ids),
        comments_ids_durable_by_para=comments_ids_durable_by_para,
        comments_extensible_durable_ids=sorted(comments_extensible_durable_ids),
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
