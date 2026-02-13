import sys

import argparse
import json
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
import zlib
import xml.etree.ElementTree as ET
from pathlib import Path

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
W15_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
W16CEX_NS = "http://schemas.microsoft.com/office/word/2018/wordml/cex"
W16CID_NS = "http://schemas.microsoft.com/office/word/2016/wordml/cid"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
XML_NS = "http://www.w3.org/XML/1998/namespace"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
PKG_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
COMMENTS_EXT_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
COMMENTS_IDS_REL_TYPE = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"
COMMENTS_EXTENSIBLE_REL_TYPE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"
PEOPLE_REL_TYPE = "http://schemas.microsoft.com/office/2011/relationships/people"
COMMENTS_EXT_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
COMMENTS_IDS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"
COMMENTS_EXTENSIBLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"
PEOPLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"
COMMENTS_EXT_PART_NAME = "/word/commentsExtended.xml"
COMMENTS_IDS_PART_NAME = "/word/commentsIds.xml"
COMMENTS_EXTENSIBLE_PART_NAME = "/word/commentsExtensible.xml"
PEOPLE_PART_NAME = "/word/people.xml"

ET.register_namespace("w", W_NS)
ET.register_namespace("w14", W14_NS)
ET.register_namespace("w15", W15_NS)
ET.register_namespace("w16cex", W16CEX_NS)
ET.register_namespace("w16cid", W16CID_NS)
ET.register_namespace("mc", MC_NS)
COMMENT_START_ATTR_BLOCK_RE = re.compile(r"\{\.comment-start(?P<attrs>[^}]*)\}", re.DOTALL)
COMMENT_END_ATTR_BLOCK_RE = re.compile(r"\{\.comment-end(?P<attrs>[^}]*)\}", re.DOTALL)
NESTED_COMMENT_END_WRAPPER_RE = re.compile(
    r"\[(?P<inner>(?:\s*\[\]\{\.comment-end[^}]*\}\s*)+)\]\{\.comment-end(?P<attrs>[^}]*)\}",
    re.DOTALL,
)
KV_ATTR_RE = re.compile(r'([A-Za-z_:][-A-Za-z0-9_:.]*)="([^"]*)"')
CARD_META_INLINE_RE = re.compile(
    r"<!--\s*CARD_META\s*\{\s*#(?P<id>[A-Za-z0-9][A-Za-z0-9_-]*)\s*(?P<attrs>.*?)\}\s*-->",
    re.DOTALL,
)
CARD_HEADER_RE = re.compile(
    r"^\[!\s*(?P<kind>COMMENT|REPLY)\s+(?P<id>[A-Za-z0-9][A-Za-z0-9_-]*)\s*:\s*(?P<author>.+?)\s*\((?P<state>active|resolved)\)\s*\]$",
    re.IGNORECASE,
)
MILESTONE_TOKEN_RE = re.compile(
    r"(?:(?P<markeq>==)\s*)?"
    r"(?:/{3}\s*C(?P<id3c>[0-9][A-Za-z0-9_-]*)\s*\.\s*(?P<edge3c>[sSeE]|[Ss][Tt][Aa][Rr][Tt]|[Ee][Nn][Dd])\s*/{3}"
    r"|/{3}\s*(?P<id3>[A-Za-z0-9][A-Za-z0-9_-]*)\s*\.\s*(?P<edge3>[sSeE]|[Ss][Tt][Aa][Rr][Tt]|[Ee][Nn][Dd])\s*/{3})"
    r"(?(markeq)\s*==)"
)
MILESTONE_CORE_TOKEN_RE = re.compile(
    r"(?:/{3}\s*C(?P<id3c>[0-9][A-Za-z0-9_-]*)\s*\.\s*(?P<edge3c>[sSeE]|[Ss][Tt][Aa][Rr][Tt]|[Ee][Nn][Dd])\s*/{3})"
    r"|(?:/{3}\s*(?P<id3>[A-Za-z0-9][A-Za-z0-9_-]*)\s*\.\s*(?P<edge3>[sSeE]|[Ss][Tt][Aa][Rr][Tt]|[Ee][Nn][Dd])\s*/{3})"
)
INLINE_IMAGE_RE = re.compile(
    r'!\[(?P<alt>[^\]]*)\]\((?P<src>[^)\s]+)(?:\s+"(?P<title>[^"]*)")?\)\{(?P<attrs>[^}]*)\}',
    re.DOTALL,
)
IMAGE_LINK_RE = re.compile(r'!\[[^\]]*\]\((?P<src>[^)\s]+)(?:\s+"[^"]*")?\)', re.DOTALL)
MIN_PANDOC_VERSION = (2, 14)


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def get_attr_local(elem: ET.Element, attr_name: str):
    for key, value in elem.attrib.items():
        if key == attr_name or key.endswith("}" + attr_name):
            return value
    return None


def read_xml(xml_path: Path):
    tree = ET.parse(xml_path)
    return tree, tree.getroot()


def write_xml(tree: ET.ElementTree, xml_path: Path):
    tree.write(xml_path, encoding="utf-8", xml_declaration=True)


def extract_docx(docx_path: Path, target_dir: Path):
    with zipfile.ZipFile(docx_path, "r") as zf:
        zf.extractall(target_dir)


def pack_docx(source_dir: Path, output_docx: Path):
    with zipfile.ZipFile(output_docx, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(source_dir):
            for filename in sorted(files):
                full_path = Path(root) / filename
                arcname = full_path.relative_to(source_dir).as_posix()
                zf.write(full_path, arcname)


def run_pandoc(in_path: Path, out_path: Path, fmt_from=None, fmt_to=None, extra_args=None, cwd=None):
    cmd = ["pandoc", str(in_path)]
    if fmt_from:
        cmd.extend(["-f", fmt_from])
    if fmt_to:
        cmd.extend(["-t", fmt_to])
    if extra_args:
        cmd.extend(extra_args)
    cmd.extend(["-o", str(out_path)])
    subprocess.run(cmd, check=True, cwd=str(cwd) if cwd else None)


def temp_dir_root_for(path: Path):
    parent = path.parent
    if parent.exists() and os.access(parent, os.W_OK):
        return parent
    return None


def parse_state_token(state_value):
    token = str(state_value or "").strip().lower()
    if token == "resolved":
        return "resolved"
    return "active"


def resolve_pandoc_writer_format(extra_args, default_format="markdown"):
    args = list(extra_args or [])
    writer = default_format
    i = 0
    while i < len(args):
        arg = args[i]
        if arg in {"-t", "--to"} and i + 1 < len(args):
            writer = args[i + 1]
            i += 2
            continue
        if arg.startswith("--to="):
            writer = arg.split("=", 1)[1]
        i += 1
    return writer or default_format


def pandoc_args_for_json_markdown_render(extra_args):
    args = list(extra_args or [])
    out = []
    i = 0
    while i < len(args):
        arg = args[i]
        if arg in {"-f", "--from", "-t", "--to", "-o", "--output", "--extract-media", "--track-changes"}:
            i += 2
            continue
        if (
            arg.startswith("--from=")
            or arg.startswith("--to=")
            or arg.startswith("--output=")
            or arg.startswith("--extract-media=")
            or arg.startswith("--track-changes=")
        ):
            i += 1
            continue
        out.append(arg)
        i += 1
    return out


def ensure_attr_pair(kvs, key: str, value: str):
    for item in kvs:
        if isinstance(item, list) and len(item) == 2 and item[0] == key:
            if item[1]:
                return False
            item[1] = value
            return True
    kvs.append([key, value])
    return True


def remove_attr_pairs(kvs, keys):
    kept = []
    removed = False
    for item in kvs:
        if isinstance(item, list) and len(item) == 2 and item[0] in keys:
            removed = True
            continue
        kept.append(item)
    if removed:
        kvs[:] = kept
    return removed


def normalize_comment_span_id_attr(attr, classes, kvs):
    if "comment-start" not in classes and "comment-end" not in classes:
        return False
    identifier = str(attr[0] or "").strip()
    if not identifier:
        return False
    changed = False
    has_id = False
    for item in kvs:
        if isinstance(item, list) and len(item) == 2 and item[0] == "id":
            has_id = True
            if not item[1]:
                item[1] = identifier
                changed = True
            break
    if not has_id:
        kvs.insert(0, ["id", identifier])
        changed = True
    if attr[0] != "":
        attr[0] = ""
        changed = True
    return changed


def walk_pandoc_spans(doc, span_handler):
    changed = 0

    def walk_inlines(inlines):
        nonlocal changed
        for node in inlines or []:
            if not isinstance(node, dict):
                continue
            t = node.get("t")
            c = node.get("c")
            if t == "Span" and isinstance(c, list) and len(c) == 2:
                attr = c[0] if isinstance(c[0], list) else None
                if attr is not None and span_handler(attr):
                    changed += 1
                nested = c[1] if isinstance(c[1], list) else []
                walk_inlines(nested)
                continue
            if t in {"Para", "Plain"} and isinstance(c, list):
                walk_inlines(c)
                continue
            if t == "Header" and isinstance(c, list) and len(c) >= 3 and isinstance(c[2], list):
                walk_inlines(c[2])
                continue
            if t == "BlockQuote" and isinstance(c, list):
                walk_blocks(c)
                continue
            if t in {"BulletList", "OrderedList"} and isinstance(c, list):
                items = c if t == "BulletList" else (c[1] if len(c) > 1 else [])
                for item in items:
                    walk_blocks(item)
                continue
            if t == "DefinitionList" and isinstance(c, list):
                for term, defs in c:
                    walk_inlines(term)
                    for d in defs:
                        walk_blocks(d)
                continue
            if t == "Div" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                walk_blocks(c[1])
                continue
            if t in {"Link", "Image"} and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                walk_inlines(c[1])
                continue
            if isinstance(c, list):
                for item in c:
                    if isinstance(item, dict):
                        walk_blocks([item])
                    elif isinstance(item, list):
                        walk_inlines(item)

    def walk_blocks(blocks):
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
            if t in {"BulletList", "OrderedList"} and isinstance(c, list):
                items = c if t == "BulletList" else (c[1] if len(c) > 1 else [])
                for item in items:
                    walk_blocks(item)
                continue
            if t == "DefinitionList" and isinstance(c, list):
                for term, defs in c:
                    walk_inlines(term)
                    for d in defs:
                        walk_blocks(d)
                continue
            if t == "Div" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                walk_blocks(c[1])
                continue
            if t == "Table" and isinstance(c, list):
                for item in c:
                    if isinstance(item, list):
                        walk_blocks([x for x in item if isinstance(x, dict)])
                continue
            if isinstance(c, list):
                for item in c:
                    if isinstance(item, dict):
                        walk_blocks([item])
                    elif isinstance(item, list):
                        walk_blocks([x for x in item if isinstance(x, dict)])

    walk_blocks(doc.get("blocks", []))
    return changed


def text_to_pandoc_inlines(text: str):
    if not text:
        return []
    out = []
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch == "\n":
            out.append({"t": "SoftBreak"})
            i += 1
            continue
        if ch.isspace():
            while i < n and text[i].isspace() and text[i] != "\n":
                i += 1
            out.append({"t": "Space"})
            continue
        j = i
        while j < n and (not text[j].isspace()) and text[j] != "\n":
            j += 1
        out.append({"t": "Str", "c": text[i:j]})
        i = j
    return out


def milestone_marker_inline(comment_id: str, edge: str):
    cid = str(comment_id or "").strip()
    edge_token = normalize_milestone_edge(edge)
    if edge_token not in {"s", "e"}:
        edge_token = "s"
    edge_label = "START" if edge_token == "s" else "END"
    marker_id = f"C{cid}" if cid.isdigit() else cid
    return {"t": "Str", "c": f"==///{marker_id}.{edge_label}///=="}


def normalize_milestone_edge(edge_token: str) -> str:
    token = str(edge_token or "").strip().lower()
    if token in {"s", "start"}:
        return "s"
    if token in {"e", "end"}:
        return "e"
    return ""


def milestone_match_id_edge(match: re.Match):
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


def inlines_to_card_text(inlines):
    parts = []

    def emit(text):
        if text:
            parts.append(text)

    def walk_inline(node):
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
                    walk_inline(item)
        elif t in {"Emph", "Strong", "Strikeout", "Superscript", "Subscript", "SmallCaps", "Underline"}:
            if isinstance(c, list):
                for item in c:
                    walk_inline(item)
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
                    walk_inline(item)
                emit(close_quote)
        elif t == "Cite":
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk_inline(item)
        elif t in {"Link", "Image"}:
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk_inline(item)
        elif isinstance(c, list):
            for item in c:
                if isinstance(item, dict):
                    walk_inline(item)

    for item in inlines or []:
        walk_inline(item)
    text = "".join(parts)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    return text


def extract_comment_card_text_from_blocks(blocks):
    parts = []

    def normalize_card_line(text: str):
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
            line = normalize_card_line(inlines_to_card_text(c))
            if line:
                parts.append(line)
        elif t == "Header" and isinstance(c, list) and len(c) >= 3 and isinstance(c[2], list):
            line = inlines_to_card_text(c[2])
            if line:
                parts.append(line)
        elif t == "RawBlock" and isinstance(c, list) and len(c) == 2:
            raw = str(c[1] or "")
            if CARD_META_INLINE_RE.search(raw):
                continue
            line = normalize_card_line(raw)
            if line:
                parts.append(line)
        elif t == "Div" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
            nested = extract_comment_card_text_from_blocks(c[1])
            if nested:
                parts.append(nested)
    return "\n\n".join([p for p in parts if p]).strip()


def parse_card_meta_marker(raw_html: str):
    match = CARD_META_INLINE_RE.search(raw_html or "")
    if not match:
        return "", {}
    comment_id = str(match.group("id") or "").strip()
    attrs_raw = str(match.group("attrs") or "").strip()
    meta = {}
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


def build_card_meta_marker(comment_id: str, meta: dict):
    attrs = []
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
        value = str((meta or {}).get(key) or "").strip()
        if value:
            attrs.append(f"{json.dumps(key)}:{json.dumps(value, ensure_ascii=False)}")
    attrs_part = (" " + ",".join(attrs)) if attrs else ""
    return f"<!--CARD_META{{#{comment_id}{attrs_part}}}-->"


def parse_comment_callout_header(line: str):
    match = CARD_HEADER_RE.match(str(line or "").strip())
    if not match:
        return {}
    return {
        "kind": str(match.group("kind") or "").strip().upper(),
        "id": str(match.group("id") or "").strip(),
        "author": str(match.group("author") or "").strip(),
        "state": parse_state_token(match.group("state")),
    }


def parse_comment_card_payload_text(payload_text: str, parent_hint=""):
    lines = str(payload_text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")
    header_idx = None
    header = {}
    for idx, line in enumerate(lines):
        if not str(line or "").strip():
            continue
        parsed = parse_comment_callout_header(line)
        if parsed:
            header_idx = idx
            header = parsed
            break
    if header_idx is None:
        return "", {}, "", ""

    meta_line_idx = None
    meta_comment_id = ""
    meta = {}
    for idx, line in enumerate(lines):
        candidate_id, candidate_meta = parse_card_meta_marker(line)
        if candidate_id:
            meta_line_idx = idx
            meta_comment_id = candidate_id
            meta = candidate_meta
            break

    comment_id = str(meta_comment_id or header.get("id") or "").strip()
    if not comment_id:
        return "", {}, "", ""

    out_meta = {}
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
        value = str(meta.get(key) or "").strip()
        if value:
            out_meta[key] = value

    out_meta.setdefault("author", str(header.get("author") or "").strip())
    out_meta["state"] = parse_state_token(out_meta.get("state") or header.get("state"))
    if header.get("kind") == "REPLY" and parent_hint and not out_meta.get("parent"):
        out_meta["parent"] = str(parent_hint).strip()

    body_lines = []
    for idx, line in enumerate(lines):
        if idx == header_idx or idx == meta_line_idx:
            continue
        body_lines.append(line)
    body_text = normalize_markdown_comment_text("\n".join(body_lines))
    return comment_id, out_meta, str(header.get("kind") or ""), body_text


def parse_comment_card_blockquote(block, parent_hint=""):
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

    comment_id, meta, kind, body_text = parse_comment_card_payload_text(payload, parent_hint=parent_hint)
    if not comment_id or kind not in {"COMMENT", "REPLY"}:
        return []

    detected_meta = {}
    detected_meta_id = ""
    for quote_block in quote_blocks:
        if not isinstance(quote_block, dict):
            continue
        qtype = quote_block.get("t")
        qdata = quote_block.get("c")
        marker_id = ""
        marker_meta = {}
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
    meta["state"] = parse_state_token(meta.get("state"))
    if kind == "REPLY" and parent_hint and not meta.get("parent"):
        meta["parent"] = str(parent_hint).strip()

    parts = []
    if body_text:
        parts.append(body_text)
    for idx, quote_block in enumerate(quote_blocks):
        if idx == primary_idx or not isinstance(quote_block, dict):
            continue
        if quote_block.get("t") == "BlockQuote":
            continue
        extra = extract_comment_card_text_from_blocks([quote_block])
        if extra:
            parts.append(extra)
    meta["text"] = "\n\n".join([part for part in parts if part]).strip()

    entries = [(comment_id, meta)]
    for quote_block in quote_blocks:
        if not isinstance(quote_block, dict) or quote_block.get("t") != "BlockQuote":
            continue
        entries.extend(parse_comment_card_blockquote(quote_block, parent_hint=comment_id))
    return entries


def normalize_card_layout_text(text: str) -> str:
    out = str(text or "")
    # Keep callout headers directly above metadata lines.
    out = re.sub(
        r"(^>+\s*\[!(?:COMMENT|REPLY)[^\n]*\])\n>+[ \t]*\n(?=>+[ \t]*<!--CARD_META)",
        r"\1\n",
        out,
        flags=re.MULTILINE,
    )
    # Keep metadata lines directly above body lines.
    out = re.sub(
        r"(^>+[ \t]*<!--CARD_META\{#[^\n]*\}-->)\n>+[ \t]*\n(?=>+[ \t]*\S)",
        r"\1\n",
        out,
        flags=re.MULTILINE,
    )
    return out


def parse_comment_cards_from_doc(doc):
    card_by_id = {}
    removed = 0

    def process_blocks(blocks):
        nonlocal removed
        kept = []
        idx = 0
        while idx < len(blocks or []):
            block = blocks[idx]
            if not isinstance(block, dict):
                kept.append(block)
                idx += 1
                continue
            t = block.get("t")
            c = block.get("c")
            if t == "BlockQuote":
                entries = parse_comment_card_blockquote(block, parent_hint="")
                if entries:
                    for comment_id, meta in entries:
                        if not comment_id:
                            continue
                        card_by_id[comment_id] = {
                            "author": str(meta.get("author") or "").strip(),
                            "date": str(meta.get("date") or "").strip(),
                            "parent": str(meta.get("parent") or "").strip(),
                            "state": parse_state_token(meta.get("state")),
                            "paraId": str(meta.get("paraId") or "").strip(),
                            "durableId": str(meta.get("durableId") or "").strip(),
                            "presenceProvider": str(meta.get("presenceProvider") or "").strip(),
                            "presenceUserId": str(meta.get("presenceUserId") or "").strip(),
                            "anchor": str(meta.get("anchor") or "").strip(),
                            "text": normalize_markdown_comment_text(meta.get("text") or ""),
                        }
                        removed += 1
                    idx += 1
                    continue
                if isinstance(c, list):
                    block["c"] = process_blocks(c)
            elif t == "Div" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                block["c"][1] = process_blocks(c[1])
            elif t in {"BulletList", "OrderedList"} and isinstance(c, list):
                items = c if t == "BulletList" else (c[1] if len(c) > 1 else [])
                for item in items:
                    if isinstance(item, list):
                        item[:] = process_blocks(item)
            elif t == "DefinitionList" and isinstance(c, list):
                for _, defs in c:
                    for d in defs:
                        if isinstance(d, list):
                            d[:] = process_blocks(d)
            elif t == "Table" and isinstance(c, list):
                for item in c:
                    if isinstance(item, list):
                        for maybe_block in item:
                            if isinstance(maybe_block, dict) and maybe_block.get("t") == "BlockQuote":
                                maybe_entries = parse_comment_card_blockquote(maybe_block, parent_hint="")
                                if maybe_entries:
                                    for comment_id, meta in maybe_entries:
                                        if not comment_id:
                                            continue
                                        card_by_id[comment_id] = {
                                            "author": str(meta.get("author") or "").strip(),
                                            "date": str(meta.get("date") or "").strip(),
                                            "parent": str(meta.get("parent") or "").strip(),
                                            "state": parse_state_token(meta.get("state")),
                                            "paraId": str(meta.get("paraId") or "").strip(),
                                            "durableId": str(meta.get("durableId") or "").strip(),
                                            "presenceProvider": str(meta.get("presenceProvider") or "").strip(),
                                            "presenceUserId": str(meta.get("presenceUserId") or "").strip(),
                                            "anchor": str(meta.get("anchor") or "").strip(),
                                            "text": normalize_markdown_comment_text(meta.get("text") or ""),
                                        }
                                        removed += 1
            kept.append(block)
            idx += 1
        return kept

    if doc.get("blocks"):
        doc["blocks"] = process_blocks(doc.get("blocks", []))
    return card_by_id, removed


def make_comment_span_inline(comment_id: str, edge: str, card_by_id=None):
    if edge == "e":
        return {"t": "Span", "c": [["", ["comment-end"], [["id", comment_id]]], []]}
    card = (card_by_id or {}).get(comment_id) or {}
    attrs = [["id", comment_id]]
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
        value = str(card.get(key) or "").strip()
        if value:
            attrs.append([key, value])
    anchor = str(card.get("anchor") or "")
    if (
        "\n" in anchor
        or "\r" in anchor
        or "---" in anchor
        or len(anchor) > 120
    ):
        anchor = ""
    nested = text_to_pandoc_inlines(anchor) if anchor else []
    return {"t": "Span", "c": [["", ["comment-start"], attrs], nested]}


def expand_milestone_tokens_in_text(text: str, card_by_id=None):
    if not text:
        return None, 0
    matches = list(MILESTONE_TOKEN_RE.finditer(text))
    if not matches:
        return None, 0
    out = []
    cursor = 0
    for match in matches:
        if match.start() > cursor:
            out.extend(text_to_pandoc_inlines(text[cursor : match.start()]))
        comment_id, edge = milestone_match_id_edge(match)
        if comment_id and edge in {"s", "e"}:
            out.append(make_comment_span_inline(comment_id, edge, card_by_id))
            next_cursor = match.end()
            if edge == "s":
                card = (card_by_id or {}).get(comment_id) or {}
                anchor = str(card.get("anchor") or "")
                if (
                    "\n" in anchor
                    or "\r" in anchor
                    or "---" in anchor
                    or len(anchor) > 120
                ):
                    anchor = ""
                if anchor and text[next_cursor:].startswith(anchor):
                    next_cursor += len(anchor)
        else:
            out.extend(text_to_pandoc_inlines(match.group(0)))
            next_cursor = match.end()
        cursor = next_cursor
    if cursor < len(text):
        out.extend(text_to_pandoc_inlines(text[cursor:]))
    return out, len(matches)


def rewrite_milestone_tokens_in_inlines(inlines, card_by_id=None):
    changed = 0
    out = []
    i = 0
    text_nodes = {"Str", "Space", "SoftBreak", "LineBreak"}

    def node_text(node):
        t = node.get("t")
        if t == "Str":
            return node.get("c") or ""
        if t == "Space":
            return " "
        if t in {"SoftBreak", "LineBreak"}:
            return "\n"
        return ""

    while i < len(inlines):
        node = inlines[i]
        if isinstance(node, dict) and node.get("t") in text_nodes:
            start = i
            text_parts = []
            while i < len(inlines):
                probe = inlines[i]
                if not isinstance(probe, dict) or probe.get("t") not in text_nodes:
                    break
                text_parts.append(node_text(probe))
                i += 1
            segment = "".join(text_parts)
            replacement, replaced_count = expand_milestone_tokens_in_text(segment, card_by_id=card_by_id)
            if replaced_count:
                out.extend(replacement)
                changed += replaced_count
            else:
                out.extend(inlines[start:i])
            continue
        out.append(node)
        i += 1
    if changed:
        inlines[:] = out
    return changed


def rewrite_milestone_tokens_in_doc(doc, card_by_id=None):
    changed = 0

    def walk_inlines(inlines):
        nonlocal changed
        if not isinstance(inlines, list):
            return
        changed += rewrite_milestone_tokens_in_inlines(inlines, card_by_id=card_by_id)
        for node in inlines:
            if not isinstance(node, dict):
                continue
            t = node.get("t")
            c = node.get("c")
            if t == "Span" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                walk_inlines(c[1])
                continue
            if t in {"Emph", "Strong", "Strikeout", "Superscript", "Subscript", "SmallCaps", "Underline"}:
                if isinstance(c, list):
                    walk_inlines(c)
                continue
            if t == "Quoted" and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                walk_inlines(c[1])
                continue
            if t == "Cite" and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                walk_inlines(c[1])
                continue
            if t in {"Link", "Image"} and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                walk_inlines(c[1])
                continue

    def walk_blocks(blocks):
        if not isinstance(blocks, list):
            return
        for block in blocks:
            if not isinstance(block, dict):
                continue
            t = block.get("t")
            c = block.get("c")
            if t in {"Para", "Plain"} and isinstance(c, list):
                walk_inlines(c)
                continue
            if t == "Header" and isinstance(c, list) and len(c) >= 3 and isinstance(c[2], list):
                walk_inlines(c[2])
                if isinstance(c[1], list) and len(c[1]) >= 1:
                    header_id = str(c[1][0] or "").strip().lower()
                    if "dc_comment" in header_id:
                        c[1][0] = ""
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
            if t == "DefinitionList" and isinstance(c, list):
                for term, defs in c:
                    walk_inlines(term)
                    for d in defs:
                        walk_blocks(d)
                continue
            if t == "Table" and isinstance(c, list):
                for item in c:
                    if isinstance(item, list):
                        walk_blocks([x for x in item if isinstance(x, dict)])
                continue

    walk_blocks(doc.get("blocks", []))
    return changed


def build_comment_card_blockquote(comment_id: str, meta: dict, children=None):
    local_meta = dict(meta or {})
    author = str(local_meta.get("author") or "").strip() or "Unknown"
    state = parse_state_token(local_meta.get("state"))
    kind = "REPLY" if str(local_meta.get("parent") or "").strip() else "COMMENT"
    header_text = f"[!{kind} {comment_id}: {author} ({state})]"
    marker_text = build_card_meta_marker(comment_id, local_meta)
    body_text = normalize_markdown_comment_text(local_meta.get("text") or "")

    quote_blocks = [
        {"t": "RawBlock", "c": ["markdown", header_text]},
        {"t": "RawBlock", "c": ["html", marker_text]},
    ]
    if body_text:
        quote_blocks.append({"t": "RawBlock", "c": ["markdown", body_text]})
    for child in children or []:
        if isinstance(child, dict):
            quote_blocks.append(child)
    return {"t": "BlockQuote", "c": quote_blocks}


def rewrite_comment_spans_to_milestones_in_doc(doc, child_ids=None):
    child_id_set = set(str(cid) for cid in (child_ids or set()))
    start_order = []
    seen_starts = set()
    anchor_by_id = {}
    changed = 0

    def comment_id_from_attr(attr):
        if not (isinstance(attr, list) and len(attr) == 3):
            return ""
        identifier = str(attr[0] or "").strip()
        if identifier:
            return identifier
        kvs = attr[2] if isinstance(attr[2], list) else []
        for item in kvs:
            if isinstance(item, list) and len(item) == 2 and item[0] == "id":
                return str(item[1] or "").strip()
        return ""

    def walk_inlines(inlines):
        nonlocal changed
        if not isinstance(inlines, list):
            return
        out = []
        for node in inlines:
            if not isinstance(node, dict):
                out.append(node)
                continue
            t = node.get("t")
            c = node.get("c")
            if t == "Span" and isinstance(c, list) and len(c) == 2 and isinstance(c[0], list):
                attr = c[0]
                nested = c[1] if isinstance(c[1], list) else []
                classes = attr[1] if isinstance(attr[1], list) else []
                cid = comment_id_from_attr(attr)
                if cid and "comment-start" in classes:
                    # Root anchors stay in prose as milestones; replies are carried by cards only.
                    if cid not in child_id_set:
                        out.append(milestone_marker_inline(cid, "s"))
                        if cid not in seen_starts:
                            seen_starts.add(cid)
                            start_order.append(cid)
                    changed += 1
                    continue
                if cid and "comment-end" in classes:
                    if cid not in child_id_set:
                        out.append(milestone_marker_inline(cid, "e"))
                    changed += 1
                    continue
                walk_inlines(nested)
                out.append({"t": "Span", "c": [attr, nested]})
                continue
            if t in {"Emph", "Strong", "Strikeout", "Superscript", "Subscript", "SmallCaps", "Underline"}:
                if isinstance(c, list):
                    walk_inlines(c)
                out.append(node)
                continue
            if t == "Quoted" and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                walk_inlines(c[1])
                out.append(node)
                continue
            if t == "Cite" and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                walk_inlines(c[1])
                out.append(node)
                continue
            if t in {"Link", "Image"} and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                walk_inlines(c[1])
                out.append(node)
                continue
            out.append(node)
        inlines[:] = out

    def walk_blocks(blocks):
        if not isinstance(blocks, list):
            return
        for block in blocks:
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
            if t == "DefinitionList" and isinstance(c, list):
                for term, defs in c:
                    walk_inlines(term)
                    for d in defs:
                        walk_blocks(d)
                continue
            if t == "Table" and isinstance(c, list):
                for item in c:
                    if isinstance(item, list):
                        walk_blocks([x for x in item if isinstance(x, dict)])
                continue

    walk_blocks(doc.get("blocks", []))
    return changed, start_order, anchor_by_id


def emit_milestones_and_cards_ast(
    md_path: Path,
    comment_cards_by_id,
    child_ids,
    pandoc_extra_args=None,
    writer_format="markdown",
    cwd=None,
):
    doc = run_pandoc_json(md_path, fmt_from="markdown", extra_args=pandoc_extra_args)
    changed, start_order, anchor_by_id = rewrite_comment_spans_to_milestones_in_doc(doc, child_ids=child_ids)
    if start_order:
        cards_meta_by_id = {str(cid): dict(meta or {}) for cid, meta in (comment_cards_by_id or {}).items()}
        order_index = {cid: idx for idx, cid in enumerate(cards_meta_by_id.keys())}
        children_by_parent = {}
        for cid, meta in cards_meta_by_id.items():
            parent = str((meta or {}).get("parent") or "").strip()
            if parent and parent != cid and parent in cards_meta_by_id:
                children_by_parent.setdefault(parent, []).append(cid)
        for parent, children in children_by_parent.items():
            children.sort(key=lambda x: order_index.get(x, 10**9))

        def build_thread_blockquote(comment_id, seen=None):
            if seen is None:
                seen = set()
            if comment_id in seen:
                return None
            seen.add(comment_id)
            meta = dict(cards_meta_by_id.get(comment_id) or {})
            children = []
            for child_id in children_by_parent.get(comment_id, []):
                child_block = build_thread_blockquote(child_id, seen=seen)
                if child_block is not None:
                    children.append(child_block)
            return build_comment_card_blockquote(comment_id, meta, children=children)

        cards_by_root_id = {}
        for root_id in start_order:
            root_block = build_thread_blockquote(root_id, seen=set())
            cards_by_root_id[root_id] = [root_block] if root_block is not None else []

        marker_order = []

        def push_marker_ids_from_text(text, start_ids, end_ids):
            for match in MILESTONE_TOKEN_RE.finditer(text or ""):
                mid, edge = milestone_match_id_edge(match)
                if not mid or edge not in {"s", "e"}:
                    continue
                if edge == "s":
                    if mid not in start_ids:
                        start_ids.append(mid)
                    if mid not in marker_order:
                        marker_order.append(mid)
                else:
                    if mid not in end_ids:
                        end_ids.append(mid)

        def scan_inlines_for_markers(inlines, start_ids, end_ids):
            if not isinstance(inlines, list):
                return
            for node in inlines:
                if not isinstance(node, dict):
                    continue
                t = node.get("t")
                c = node.get("c")
                if t == "Str":
                    push_marker_ids_from_text(str(c or ""), start_ids, end_ids)
                    continue
                if t in {
                    "Emph",
                    "Strong",
                    "Strikeout",
                    "Superscript",
                    "Subscript",
                    "SmallCaps",
                    "Underline",
                } and isinstance(c, list):
                    scan_inlines_for_markers(c, start_ids, end_ids)
                    continue
                if t == "Span" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                    scan_inlines_for_markers(c[1], start_ids, end_ids)
                    continue
                if t == "Quoted" and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                    scan_inlines_for_markers(c[1], start_ids, end_ids)
                    continue
                if t == "Cite" and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                    scan_inlines_for_markers(c[1], start_ids, end_ids)
                    continue
                if t in {"Link", "Image"} and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                    scan_inlines_for_markers(c[1], start_ids, end_ids)
                    continue

        def collect_block_markers(block, start_ids, end_ids):
            if not isinstance(block, dict):
                return
            t = block.get("t")
            c = block.get("c")
            if t in {"Para", "Plain"} and isinstance(c, list):
                scan_inlines_for_markers(c, start_ids, end_ids)
            elif t == "Header" and isinstance(c, list) and len(c) >= 3 and isinstance(c[2], list):
                scan_inlines_for_markers(c[2], start_ids, end_ids)

        def block_markers_in_subtree(block):
            starts = []
            ends = []

            def walk_block(node):
                if not isinstance(node, dict):
                    return
                collect_block_markers(node, starts, ends)
                t = node.get("t")
                c = node.get("c")
                if t == "BlockQuote" and isinstance(c, list):
                    for child in c:
                        walk_block(child)
                    return
                if t == "Div" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                    for child in c[1]:
                        walk_block(child)
                    return
                if t == "BulletList" and isinstance(c, list):
                    for item in c:
                        if isinstance(item, list):
                            for child in item:
                                walk_block(child)
                    return
                if t == "OrderedList" and isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                    for item in c[1]:
                        if isinstance(item, list):
                            for child in item:
                                walk_block(child)
                    return
                if t == "DefinitionList" and isinstance(c, list):
                    for term, defs in c:
                        scan_inlines_for_markers(term, starts, ends)
                        for d in defs:
                            if isinstance(d, list):
                                for child in d:
                                    walk_block(child)
                    return
                if t == "Table" and isinstance(c, list):
                    for item in c:
                        if isinstance(item, list):
                            for maybe_block in item:
                                if isinstance(maybe_block, dict):
                                    walk_block(maybe_block)

            walk_block(block)
            return starts, ends

        top_blocks = doc.get("blocks", [])
        first_start_index = {}
        last_end_index = {}
        root_set = set(start_order)
        for idx, block in enumerate(top_blocks):
            starts, ends = block_markers_in_subtree(block)
            for cid in starts:
                if cid in root_set and cid not in first_start_index:
                    first_start_index[cid] = idx
            for cid in ends:
                if cid in root_set:
                    last_end_index[cid] = idx

        cards_after_index = {}
        pending_append = []
        for cid in start_order:
            card_blocks = cards_by_root_id.get(cid) or []
            if not card_blocks:
                continue
            insert_after = last_end_index.get(cid, first_start_index.get(cid))
            if insert_after is None:
                pending_append.extend(card_blocks)
                continue
            cards_after_index.setdefault(insert_after, []).extend(card_blocks)

        out_blocks = []
        for idx, block in enumerate(top_blocks):
            out_blocks.append(block)
            out_blocks.extend(cards_after_index.get(idx, []))
        out_blocks.extend(pending_append)
        doc["blocks"] = out_blocks
    if changed or start_order:
        render_pandoc_json_to_markdown(
            doc,
            md_path,
            writer_format=writer_format,
            extra_args=pandoc_extra_args,
            cwd=cwd,
        )
        normalized = normalize_card_layout_text(md_path.read_text(encoding="utf-8"))
        md_path.write_text(normalized, encoding="utf-8")
    return changed, len(start_order)


def normalize_milestone_tokens_ast(
    md_path: Path,
    out_md_path: Path,
    pandoc_extra_args=None,
    writer_format="markdown",
    cwd=None,
):
    doc = run_pandoc_json(md_path, fmt_from="markdown", extra_args=pandoc_extra_args)
    card_by_id, removed_cards = parse_comment_cards_from_doc(doc)
    changed = rewrite_milestone_tokens_in_doc(doc, card_by_id=card_by_id)
    if changed or removed_cards or md_path != out_md_path:
        render_pandoc_json_to_markdown(
            doc,
            out_md_path,
            writer_format=writer_format,
            extra_args=pandoc_extra_args,
            cwd=cwd,
        )
    return changed, card_by_id


def line_col_for_offset(text: str, offset: int):
    safe_offset = max(0, min(len(text or ""), int(offset)))
    line = (text or "").count("\n", 0, safe_offset) + 1
    last_newline = (text or "").rfind("\n", 0, safe_offset)
    col = safe_offset - (last_newline + 1) + 1
    return line, col


def line_excerpt(text: str, line_no: int, max_len: int = 140):
    lines = (text or "").splitlines()
    if line_no < 1 or line_no > len(lines):
        return ""
    excerpt = lines[line_no - 1].strip()
    if len(excerpt) > max_len:
        excerpt = excerpt[: max_len - 3] + "..."
    return excerpt


def collect_span_marker_positions(markdown_text: str):
    starts_by_id = {}
    ends_by_id = {}

    for match in COMMENT_START_ATTR_BLOCK_RE.finditer(markdown_text or ""):
        attrs = {k: v for k, v in KV_ATTR_RE.findall(match.group("attrs"))}
        comment_id = str(attrs.get("id") or "").strip()
        if not comment_id:
            continue
        line_no, col_no = line_col_for_offset(markdown_text, match.start())
        starts_by_id.setdefault(comment_id, []).append(
            {"offset": match.start(), "line": line_no, "col": col_no}
        )

    for match in COMMENT_END_ATTR_BLOCK_RE.finditer(markdown_text or ""):
        attrs = {k: v for k, v in KV_ATTR_RE.findall(match.group("attrs"))}
        comment_id = str(attrs.get("id") or "").strip()
        if not comment_id:
            continue
        line_no, col_no = line_col_for_offset(markdown_text, match.start())
        ends_by_id.setdefault(comment_id, []).append(
            {"offset": match.start(), "line": line_no, "col": col_no}
        )

    return starts_by_id, ends_by_id


def collect_root_card_lines(markdown_text: str):
    root_line_by_id = {}
    for match in CARD_META_INLINE_RE.finditer(markdown_text or ""):
        comment_id, meta = parse_card_meta_marker(match.group(0) or "")
        if not comment_id:
            continue
        parent_id = str((meta or {}).get("parent") or "").strip()
        if parent_id:
            continue
        if comment_id in root_line_by_id:
            continue
        line_no, _ = line_col_for_offset(markdown_text, match.start())
        root_line_by_id[comment_id] = line_no
    return root_line_by_id


def format_marker_locations(locations):
    if not locations:
        return "(none)"
    return ", ".join([f"{loc['line']}:{loc['col']}" for loc in locations])


def collect_one_sided_wrapper_issues(markdown_text: str):
    text = markdown_text or ""
    issues = []
    for match in MILESTONE_CORE_TOKEN_RE.finditer(text):
        comment_id, edge = milestone_match_id_edge(match)
        if not comment_id or edge not in {"s", "e"}:
            continue

        left_idx = match.start()
        while left_idx > 0 and text[left_idx - 1] in {" ", "\t"}:
            left_idx -= 1
        has_left_wrapper = left_idx >= 2 and text[left_idx - 2:left_idx] == "=="

        right_idx = match.end()
        text_len = len(text)
        while right_idx < text_len and text[right_idx] in {" ", "\t"}:
            right_idx += 1
        has_right_wrapper = right_idx + 1 < text_len and text[right_idx:right_idx + 2] == "=="

        if has_left_wrapper == has_right_wrapper:
            continue

        line_no, col_no = line_col_for_offset(text, match.start())
        excerpt = line_excerpt(text, line_no)
        token = "START" if edge == "s" else "END"
        issue = (
            f"Comment {comment_id} has one-sided highlight wrapper around {token} marker at "
            f"{line_no}:{col_no}. Use either plain `///C{comment_id}.{token}///` or fully wrapped "
            f"`==///C{comment_id}.{token}///==`."
        )
        if excerpt:
            issue += f" Line {line_no}: `{excerpt}`"
        issues.append(issue)
    return issues


def validate_comment_marker_integrity(source_text: str, normalized_text: str, card_by_id=None, source_label=""):
    starts_by_id, ends_by_id = collect_span_marker_positions(normalized_text)
    root_line_by_id = collect_root_card_lines(source_text)
    for comment_id, card in (card_by_id or {}).items():
        cid = str(comment_id or "").strip()
        if not cid:
            continue
        if str((card or {}).get("parent") or "").strip():
            continue
        root_line_by_id.setdefault(cid, 0)

    issues = []
    all_ids = set(starts_by_id.keys()) | set(ends_by_id.keys()) | set(root_line_by_id.keys())
    for comment_id in sorted(all_ids, key=lambda value: (len(value), value)):
        starts = starts_by_id.get(comment_id, [])
        ends = ends_by_id.get(comment_id, [])
        start_count = len(starts)
        end_count = len(ends)
        is_root_card = comment_id in root_line_by_id

        if is_root_card and (start_count != 1 or end_count != 1):
            card_line = int(root_line_by_id.get(comment_id) or 0)
            card_ref = f"CARD_META line {card_line}" if card_line else "CARD_META line unknown"
            issues.append(
                f"Root comment {comment_id} ({card_ref}) must have exactly one START and one END marker; "
                f"found START={start_count}, END={end_count}."
            )
        elif start_count != end_count:
            issues.append(
                f"Comment {comment_id} has unbalanced markers; START={start_count} at {format_marker_locations(starts)}, "
                f"END={end_count} at {format_marker_locations(ends)}."
            )

        if start_count > 1 or end_count > 1:
            issues.append(
                f"Comment {comment_id} has duplicate markers; START lines {format_marker_locations(starts)}, "
                f"END lines {format_marker_locations(ends)}."
            )

        if starts and ends and starts[0]["offset"] > ends[-1]["offset"]:
            issues.append(
                f"Comment {comment_id} has END before START; first START at {starts[0]['line']}:{starts[0]['col']}, "
                f"last END at {ends[-1]['line']}:{ends[-1]['col']}."
            )

        if start_count and not end_count:
            first_start_line = starts[0]["line"]
            excerpt = line_excerpt(normalized_text, first_start_line)
            if excerpt:
                issues.append(
                    f"Comment {comment_id} START marker line {first_start_line}: `{excerpt}`"
                )
        if end_count and not start_count:
            first_end_line = ends[0]["line"]
            excerpt = line_excerpt(normalized_text, first_end_line)
            if excerpt:
                issues.append(
                    f"Comment {comment_id} END marker line {first_end_line}: `{excerpt}`"
                )

    one_sided_wrapper_issues = collect_one_sided_wrapper_issues(source_text)
    if one_sided_wrapper_issues:
        issues.extend(one_sided_wrapper_issues)

    if issues:
        source_hint = source_label or "<markdown input>"
        guidance = (
            "Fix markers by keeping one exact pair per root anchor in prose: "
            "`///C<ID>.START/// ... ///C<ID>.END///` (optional wrapper: `==///C<ID>.START///==`). "
            "Ensure each root `CARD_META` entry has a matching marker pair with the same ID."
        )
        raise ValueError(
            "Comment marker validation failed before md->docx conversion.\n"
            f"Source: {source_hint}\n"
            + "\n".join([f"- {issue}" for issue in issues])
            + "\n- "
            + guidance
        )


def render_pandoc_json_to_markdown(
    doc,
    out_path: Path,
    writer_format="markdown",
    extra_args=None,
    cwd=None,
):
    with tempfile.TemporaryDirectory(prefix=".docx-comments-json-", dir=temp_dir_root_for(out_path)) as tmp:
        tmp_dir = Path(tmp)
        json_in = tmp_dir / "ast.json"
        json_in.write_text(json.dumps(doc, ensure_ascii=False), encoding="utf-8")
        render_args = pandoc_args_for_json_markdown_render(extra_args)
        run_pandoc(
            json_in,
            out_path,
            fmt_from="json",
            fmt_to=writer_format or "markdown",
            extra_args=render_args,
            cwd=cwd,
        )


def annotate_markdown_comment_attrs(
    md_path: Path,
    parent_map,
    state_by_id,
    para_by_id=None,
    durable_by_id=None,
    presence_provider_by_id=None,
    presence_user_by_id=None,
    pandoc_extra_args=None,
    writer_format="markdown",
    cwd=None,
):
    if (
        not parent_map
        and not state_by_id
        and not para_by_id
        and not durable_by_id
        and not presence_provider_by_id
        and not presence_user_by_id
    ):
        return 0
    doc = run_pandoc_json(md_path, fmt_from="markdown", extra_args=pandoc_extra_args)

    def on_span(attr):
        if not (isinstance(attr, list) and len(attr) == 3):
            return False
        classes = attr[1] if isinstance(attr[1], list) else []
        kvs = attr[2] if isinstance(attr[2], list) else []
        changed_here = normalize_comment_span_id_attr(attr, classes, kvs)
        identifier = str(attr[0] or "").strip()
        if "comment-start" not in classes:
            return changed_here
        kv = {}
        for item in kvs:
            if isinstance(item, list) and len(item) == 2:
                kv[item[0]] = item[1]
        cid = identifier or str(kv.get("id") or "").strip()
        if not cid:
            return changed_here

        pid = (parent_map or {}).get(cid)
        if pid and not kv.get("parent"):
            changed_here = ensure_attr_pair(kvs, "parent", str(pid)) or changed_here
        if not kv.get("state"):
            state_token = parse_state_token((state_by_id or {}).get(cid))
            changed_here = ensure_attr_pair(kvs, "state", state_token) or changed_here
        para_id = (para_by_id or {}).get(cid)
        if para_id and not kv.get("paraId"):
            changed_here = ensure_attr_pair(kvs, "paraId", str(para_id)) or changed_here
        durable_id = (durable_by_id or {}).get(cid)
        if durable_id and not kv.get("durableId"):
            changed_here = ensure_attr_pair(kvs, "durableId", str(durable_id)) or changed_here
        presence_provider = (presence_provider_by_id or {}).get(cid)
        if presence_provider and not kv.get("presenceProvider"):
            changed_here = ensure_attr_pair(kvs, "presenceProvider", str(presence_provider)) or changed_here
        presence_user = (presence_user_by_id or {}).get(cid)
        if presence_user and not kv.get("presenceUserId"):
            changed_here = ensure_attr_pair(kvs, "presenceUserId", str(presence_user)) or changed_here
        return changed_here

    changed = walk_pandoc_spans(doc, on_span)
    if changed:
        render_pandoc_json_to_markdown(
            doc,
            md_path,
            writer_format=writer_format,
            extra_args=pandoc_extra_args,
            cwd=cwd,
        )
    return changed


def strip_comment_transport_attrs_ast(
    md_path: Path,
    out_md_path: Path,
    pandoc_extra_args=None,
    writer_format="markdown",
    cwd=None,
):
    doc = run_pandoc_json(md_path, fmt_from="markdown", extra_args=pandoc_extra_args)

    def on_span(attr):
        if not (isinstance(attr, list) and len(attr) == 3):
            return False
        classes = attr[1] if isinstance(attr[1], list) else []
        kvs = attr[2] if isinstance(attr[2], list) else []
        changed_here = normalize_comment_span_id_attr(attr, classes, kvs)
        if "comment-start" not in classes:
            return changed_here
        return remove_attr_pairs(kvs, {"paraId", "durableId", "presenceProvider", "presenceUserId"}) or changed_here

    changed = walk_pandoc_spans(doc, on_span)
    render_pandoc_json_to_markdown(
        doc,
        out_md_path,
        writer_format=writer_format,
        extra_args=pandoc_extra_args,
        cwd=cwd,
    )
    return changed


def repair_unbalanced_comment_markers(markdown_text: str):
    starts = []
    end_positions_by_id = {}

    for m in COMMENT_START_ATTR_BLOCK_RE.finditer(markdown_text):
        attrs = {k: v for k, v in KV_ATTR_RE.findall(m.group("attrs"))}
        cid = attrs.get("id")
        if not cid:
            continue
        starts.append(
            {
                "id": cid,
                "parent": attrs.get("parent", ""),
                "pos_end": m.end(),
            }
        )

    for m in COMMENT_END_ATTR_BLOCK_RE.finditer(markdown_text):
        attrs = {k: v for k, v in KV_ATTR_RE.findall(m.group("attrs"))}
        cid = attrs.get("id")
        if not cid:
            continue
        end_positions_by_id.setdefault(cid, []).append(m.end())

    start_ids = [s["id"] for s in starts]
    if not start_ids:
        return markdown_text, 0
    missing_ids = [cid for cid in start_ids if cid not in end_positions_by_id]
    if not missing_ids:
        return markdown_text, 0

    starts_by_id = {s["id"]: s for s in starts}
    children = {}
    for s in starts:
        pid = (s.get("parent") or "").strip()
        if pid:
            children.setdefault(pid, []).append(s["id"])

    def descendant_end_positions(cid, seen):
        if cid in seen:
            return []
        seen.add(cid)
        out = []
        for child_id in children.get(cid, []):
            out.extend(end_positions_by_id.get(child_id, []))
            out.extend(descendant_end_positions(child_id, seen))
        return out

    insertions = []
    for cid in missing_ids:
        base = starts_by_id[cid]["pos_end"]
        desc_ends = descendant_end_positions(cid, set())
        pos = max(desc_ends) if desc_ends else base
        insertions.append((pos, f'[]{{.comment-end id="{cid}"}}'))

    updated = markdown_text
    for pos, token in sorted(insertions, key=lambda x: x[0], reverse=True):
        updated = updated[:pos] + token + updated[pos:]

    return updated, len(insertions)


def strip_comment_transport_attrs(markdown_text: str):
    removed = 0

    def repl(match):
        nonlocal removed
        block = match.group(0)
        updated = re.sub(r'\s+paraId="[^"]*"', "", block)
        updated = re.sub(r'\s+durableId="[^"]*"', "", updated)
        updated = re.sub(r'\s+presenceProvider="[^"]*"', "", updated)
        updated = re.sub(r'\s+presenceUserId="[^"]*"', "", updated)
        if updated != block:
            removed += 1
        return updated

    out = COMMENT_START_ATTR_BLOCK_RE.sub(repl, markdown_text)
    return out, removed


def normalize_nested_comment_end_markers(markdown_text: str):
    # Pandoc can emit nested wrappers like:
    # [[]{.comment-end id="a"}]{.comment-end id="b"}
    # or [[]{.comment-end id="a"}[]{.comment-end id="b"}]{.comment-end id="c"}
    # Flatten these to sibling markers so md->docx preserves all ends/references.
    updated = markdown_text
    changed = 0
    while True:
        updated, count = NESTED_COMMENT_END_WRAPPER_RE.subn(
            lambda m: f"{m.group('inner')}[]{{.comment-end{m.group('attrs')}}}",
            updated,
        )
        if not count:
            break
        changed += count
    return updated, changed


def normalize_markdown_comment_text(text: str) -> str:
    text = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    # Pandoc emits hard line breaks in comment brackets as backslash-newline.
    text = re.sub(r"\\+[ \t]*\n", "\n", text)
    # Handle wrapped hard-break output forms like "\\ " conservatively.
    text = re.sub(r"\\\\[ \t]+", "\n", text)
    text = (
        text.replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
    )
    text = re.sub(r"(?m)^[\u2014\u2015]\s*$", "---", text)
    return text.strip()


def run_pandoc_json(in_path: Path, fmt_from=None, extra_args=None):
    cmd = ["pandoc", str(in_path)]
    if fmt_from:
        cmd.extend(["-f", fmt_from])
    if extra_args:
        cmd.extend(extra_args)
    cmd.extend(["-t", "json"])
    # Pandoc emits UTF-8 JSON; force decoding so Windows locale codecs do not break.
    out = subprocess.check_output(cmd, text=True, encoding="utf-8")
    return json.loads(out)


def has_extract_media_arg(args):
    for arg in args or []:
        if arg == "--extract-media" or arg.startswith("--extract-media="):
            return True
    return False


def parse_length_to_inches(length_value: str):
    if not length_value:
        return None
    m = re.match(r"^\s*([0-9]*\.?[0-9]+(?:e[-+]?[0-9]+)?)\s*([a-zA-Z]*)\s*$", length_value)
    if not m:
        return None
    value = float(m.group(1))
    unit = (m.group(2) or "in").lower()
    if unit == "in":
        return value
    if unit == "cm":
        return value / 2.54
    if unit == "mm":
        return value / 25.4
    if unit == "pt":
        return value / 72.0
    if unit == "px":
        return value / 96.0
    return None


def should_strip_placeholder_image(alt, src, title, attrs):
    title = (title or "").strip().lower()
    alt = (alt or "").strip()
    if alt:
        return False
    if title != "shape":
        return False

    src_norm = (src or "").strip().lower()
    if "/media/" not in f"/{src_norm}" and not src_norm.startswith("./media/"):
        return False

    basename = Path(src_norm).name
    if not re.match(r"^image[0-9]+\.(png|jpg|jpeg|gif|bmp|emf|wmf|svg)$", basename):
        return False

    kv = {k: v for k, v in KV_ATTR_RE.findall(attrs or "")}
    width_in = parse_length_to_inches(kv.get("width", ""))
    height_in = parse_length_to_inches(kv.get("height", ""))
    if width_in is None or height_in is None:
        return False
    return width_in <= 0.03 and height_in <= 0.03


def strip_placeholder_shape_images(markdown_text: str):
    removed = []

    def repl(match):
        alt = match.group("alt")
        src = match.group("src")
        title = match.group("title")
        attrs = match.group("attrs")
        if should_strip_placeholder_image(alt, src, title, attrs):
            removed.append(src)
            return ""
        return match.group(0)

    updated = INLINE_IMAGE_RE.sub(repl, markdown_text)
    # Normalize excess whitespace left behind by removals.
    updated = re.sub(r"\n{3,}", "\n\n", updated).strip() + "\n"
    return updated, removed


def list_files_relative(root_dir: Path):
    if not root_dir.exists():
        return set()
    out = set()
    for p in root_dir.rglob("*"):
        if p.is_file():
            out.add(p.relative_to(root_dir).as_posix())
    return out


def extract_media_refs_from_markdown(markdown_text: str):
    refs = set()
    for m in IMAGE_LINK_RE.finditer(markdown_text):
        src = (m.group("src") or "").strip()
        if not src:
            continue
        if src.startswith("./"):
            src = src[2:]
        if src.startswith("media/"):
            refs.add(src[len("media/") :])
    return refs


def prune_unreferenced_new_media(media_dir: Path, files_before, markdown_text: str):
    if not media_dir.exists():
        return 0
    refs = extract_media_refs_from_markdown(markdown_text)
    files_after = list_files_relative(media_dir)
    created = files_after - files_before
    removed = 0
    for rel in created:
        if rel not in refs:
            p = media_dir / rel
            if p.exists() and p.is_file():
                p.unlink()
                removed += 1
    # Clean up empty directories left behind.
    for d in sorted(media_dir.rglob("*"), reverse=True):
        if d.is_dir():
            try:
                d.rmdir()
            except OSError:
                pass
    if media_dir.exists():
        try:
            media_dir.rmdir()
        except OSError:
            pass
    return removed


def inlines_to_text(inlines):
    parts = []

    def emit(text):
        if text:
            parts.append(text)

    def walk_inline(node):
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
            # Convert line break markers in markdown comments to real newlines.
            if parts and parts[-1].endswith("\\"):
                parts[-1] = parts[-1].rstrip("\\")
            emit("\n")
        elif t in {"Code", "Math"}:
            if isinstance(c, list) and c:
                emit(c[-1] or "")
            elif isinstance(c, str):
                emit(c)
        elif t == "Span":
            # Span c = [attr, inlines]
            if isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk_inline(item)
        elif t in {"Emph", "Strong", "Strikeout", "Superscript", "Subscript", "SmallCaps", "Underline"}:
            # Emph/Strong/etc. c = inlines
            if isinstance(c, list):
                for item in c:
                    walk_inline(item)
        elif t == "Quoted":
            # Quoted c = [quoteType, inlines]
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk_inline(item)
        elif t == "Cite":
            # Cite c = [citations, inlines]
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk_inline(item)
        elif t in {"Link", "Image"}:
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], list):
                for item in c[1]:
                    walk_inline(item)
        elif t == "RawInline":
            # Keep plain textual payload where possible.
            if isinstance(c, list) and len(c) >= 2 and isinstance(c[1], str):
                emit(c[1])
        elif isinstance(c, list):
            for item in c:
                if isinstance(item, dict):
                    walk_inline(item)

    for item in inlines or []:
        walk_inline(item)
    text = "".join(parts)
    text = re.sub(r"[ \t]+\n", "\n", text)
    return normalize_markdown_comment_text(text)


def extract_comment_texts_from_markdown(md_path: Path, pandoc_extra_args, card_by_id=None):
    doc = run_pandoc_json(md_path, fmt_from="markdown", extra_args=pandoc_extra_args)
    own_text_by_id = {}
    children_by_id = {}
    meta_by_id = {}
    parent_by_id = {}
    state_by_id = {}
    para_by_id = {}
    durable_by_id = {}
    presence_by_author = {}
    seen_order = []
    started_ids = set()

    def ensure_comment_id(comment_id: str):
        if comment_id not in children_by_id:
            children_by_id[comment_id] = []
        if comment_id not in own_text_by_id:
            own_text_by_id[comment_id] = ""
        if comment_id not in meta_by_id:
            meta_by_id[comment_id] = {}
        if comment_id not in state_by_id:
            state_by_id[comment_id] = "active"
        if comment_id not in seen_order:
            seen_order.append(comment_id)

    def add_child(parent_id: str, child_id: str):
        ensure_comment_id(parent_id)
        ensure_comment_id(child_id)
        if child_id not in children_by_id[parent_id]:
            children_by_id[parent_id].append(child_id)

    def on_comment_start(comment_id: str, text: str, meta):
        ensure_comment_id(comment_id)
        started_ids.add(comment_id)
        text = (text or "").strip()
        if text:
            existing = (own_text_by_id.get(comment_id) or "").strip()
            if not existing:
                own_text_by_id[comment_id] = text
            elif text != existing and text not in existing:
                own_text_by_id[comment_id] = f"{existing}\n\n{text}"
        if meta:
            if meta.get("author") and not meta_by_id[comment_id].get("author"):
                meta_by_id[comment_id]["author"] = meta["author"]
            if meta.get("date") and not meta_by_id[comment_id].get("date"):
                meta_by_id[comment_id]["date"] = meta["date"]
            if meta.get("presenceProvider") and not meta_by_id[comment_id].get("presenceProvider"):
                meta_by_id[comment_id]["presenceProvider"] = meta["presenceProvider"]
            if meta.get("presenceUserId") and not meta_by_id[comment_id].get("presenceUserId"):
                meta_by_id[comment_id]["presenceUserId"] = meta["presenceUserId"]
            parent_id = (meta.get("parent") or "").strip()
            if parent_id:
                parent_by_id[comment_id] = parent_id
            if "state" in meta:
                state_by_id[comment_id] = parse_state_token(meta.get("state"))
            para_id = (meta.get("paraId") or "").strip()
            if para_id:
                para_by_id[comment_id] = para_id
            durable_id = (meta.get("durableId") or "").strip()
            if durable_id:
                durable_by_id[comment_id] = durable_id

    def parse_span_meta(attr):
        if not (isinstance(attr, list) and len(attr) == 3):
            return None
        identifier = attr[0]
        classes = attr[1] or []
        kvs = attr[2] or []
        meta = {}
        if isinstance(kvs, list):
            for item in kvs:
                if isinstance(item, list) and len(item) == 2:
                    key, value = item
                    if key in {
                        "author",
                        "date",
                        "parent",
                        "state",
                        "paraId",
                        "durableId",
                        "presenceProvider",
                        "presenceUserId",
                    }:
                        meta[key] = value
        return identifier, classes, meta

    def walk_inlines(inlines):
        for node in inlines or []:
            if not isinstance(node, dict):
                continue
            t = node.get("t")
            c = node.get("c")

            if t == "Span" and isinstance(c, list) and len(c) == 2:
                attr, nested_inlines = c
                parsed = parse_span_meta(attr)
                if parsed is not None:
                    identifier, classes, meta = parsed
                    comment_id = (identifier or meta.get("id") or "").strip()
                    if comment_id and "comment-start" in classes:
                        on_comment_start(comment_id, inlines_to_text(nested_inlines), meta)
                        continue
                walk_inlines(nested_inlines)
                continue

            if t == "Str":
                continue

            if t in {
                "Emph",
                "Strong",
                "Strikeout",
                "Superscript",
                "Subscript",
                "SmallCaps",
                "Underline",
                "Quoted",
                "Cite",
            } and isinstance(c, list):
                tail = c[-1] if c else []
                if isinstance(tail, list):
                    walk_inlines(tail)
                continue

            if t in {"Link", "Image"} and isinstance(c, list) and len(c) >= 2:
                if isinstance(c[1], list):
                    walk_inlines(c[1])
                continue

            if isinstance(c, list):
                # Fallback for rarely used inline constructors.
                walk_inlines([item for item in c if isinstance(item, dict)])

    def walk_blocks(blocks):
        for block in blocks or []:
            if not isinstance(block, dict):
                continue
            t = block.get("t")
            c = block.get("c")

            if t in {"Para", "Plain", "Header"} and isinstance(c, list):
                if t == "Header" and len(c) >= 3 and isinstance(c[2], list):
                    walk_inlines(c[2])
                else:
                    walk_inlines(c)
                continue

            if t == "BlockQuote" and isinstance(c, list):
                walk_blocks(c)
                continue

            if t in {"BulletList", "OrderedList"} and isinstance(c, list):
                items = c if t == "BulletList" else (c[1] if len(c) > 1 else [])
                for item in items:
                    walk_blocks(item)
                continue

            if t == "DefinitionList" and isinstance(c, list):
                for term, defs in c:
                    walk_inlines(term)
                    for d in defs:
                        walk_blocks(d)
                continue

            if t == "Div" and isinstance(c, list) and len(c) == 2 and isinstance(c[1], list):
                walk_blocks(c[1])
                continue

            if t == "Table" and isinstance(c, list):
                # Traverse nested table content for completeness.
                for item in c:
                    if isinstance(item, list):
                        walk_blocks([x for x in item if isinstance(x, dict)])
                continue

    walk_blocks(doc.get("blocks", []))
    def reply_header(comment_id: str):
        meta = meta_by_id.get(comment_id, {})
        author = (meta.get("author") or "Unknown").strip() or "Unknown"
        date = (meta.get("date") or "").strip()
        if date:
            return f"---\nReply from: {author} ({date})\n---"
        return f"---\nReply from: {author}\n---"

    def flatten_comment(comment_id: str, seen):
        if comment_id in seen:
            return own_text_by_id.get(comment_id, "").strip()
        seen.add(comment_id)

        parts = []
        own = own_text_by_id.get(comment_id, "").strip()
        if own:
            parts.append(own)
        for child_id in children_by_id.get(comment_id, []):
            child_flat = flatten_comment(child_id, seen)
            if child_flat:
                parts.append(f"{reply_header(child_id)}\n{child_flat}")
        return "\n\n".join(parts).strip()

    if card_by_id:
        for cid, card in card_by_id.items():
            ensure_comment_id(cid)
            started_ids.add(cid)
            card_text = normalize_markdown_comment_text(card.get("text") or "")
            if card_text:
                own_text_by_id[cid] = card_text
            author = (card.get("author") or "").strip()
            date = (card.get("date") or "").strip()
            parent_id = (card.get("parent") or "").strip()
            if author and not meta_by_id[cid].get("author"):
                meta_by_id[cid]["author"] = author
            if date and not meta_by_id[cid].get("date"):
                meta_by_id[cid]["date"] = date
            if parent_id:
                parent_by_id[cid] = parent_id
            state_by_id[cid] = parse_state_token(card.get("state"))
            para_id = (card.get("paraId") or "").strip()
            durable_id = (card.get("durableId") or "").strip()
            if para_id:
                para_by_id[cid] = para_id
            if durable_id:
                durable_by_id[cid] = durable_id
            provider = (card.get("presenceProvider") or "").strip()
            user = (card.get("presenceUserId") or "").strip()
            if provider or user:
                meta_by_id[cid]["presenceProvider"] = provider
                meta_by_id[cid]["presenceUserId"] = user

    ordered_started_ids = [cid for cid in seen_order if cid in started_ids]
    if not ordered_started_ids:
        ordered_started_ids = list(seen_order)

    invalid_parent_issues = []
    valid_parent_by_id = {}
    started_id_set = set(ordered_started_ids)
    for child_id in ordered_started_ids:
        parent_id = (parent_by_id.get(child_id) or "").strip()
        # Accept threading when IDs are present in extracted comment set,
        # including card-only reply comments that have no inline markers.
        if not parent_id:
            continue
        if parent_id == child_id:
            invalid_parent_issues.append(f"comment {child_id} cannot be its own parent")
            continue
        if parent_id not in started_id_set:
            invalid_parent_issues.append(
                f"comment {child_id} references unknown parent {parent_id}"
            )
            continue
        valid_parent_by_id[child_id] = parent_id
        add_child(parent_id, child_id)

    visit_state = {}
    cycle_issues = []

    def detect_cycle(comment_id, stack):
        state = visit_state.get(comment_id, 0)
        if state == 1:
            if comment_id in stack:
                idx = stack.index(comment_id)
                cycle = stack[idx:] + [comment_id]
            else:
                cycle = stack + [comment_id]
            cycle_issues.append(" -> ".join(cycle))
            return
        if state == 2:
            return

        visit_state[comment_id] = 1
        stack.append(comment_id)
        parent_id = valid_parent_by_id.get(comment_id)
        if parent_id:
            detect_cycle(parent_id, stack)
        stack.pop()
        visit_state[comment_id] = 2

    for cid in ordered_started_ids:
        if cid in valid_parent_by_id and visit_state.get(cid, 0) == 0:
            detect_cycle(cid, [])

    if invalid_parent_issues or cycle_issues:
        issues = []
        issues.extend(invalid_parent_issues)
        for cycle in cycle_issues:
            issues.append(f"thread cycle detected: {cycle}")
        raise ValueError(
            "Invalid comment thread relationships in markdown cards/spans.\n"
            + "\n".join([f"- {issue}" for issue in issues])
            + "\n- Fix parent IDs so every reply points to an existing comment and no cycles exist."
        )

    flattened_by_id = {}
    for cid in ordered_started_ids:
        flattened_by_id[cid] = flatten_comment(cid, set())

    child_ids = set(valid_parent_by_id.keys())
    root_ids = [cid for cid in ordered_started_ids if cid not in child_ids]
    if not root_ids:
        root_ids = list(ordered_started_ids)

    for cid in ordered_started_ids:
        state_by_id[cid] = parse_state_token(state_by_id.get(cid))
        meta = meta_by_id.get(cid, {})
        author = (meta.get("author") or "").strip()
        if author:
            provider_id = (meta.get("presenceProvider") or "").strip()
            user_id = (meta.get("presenceUserId") or "").strip()
            if provider_id or user_id:
                presence_by_author[author] = {
                    "provider_id": provider_id,
                    "user_id": user_id,
                }

    text_by_id = {}
    author_by_id = {}
    date_by_id = {}
    for cid in ordered_started_ids:
        meta = meta_by_id.get(cid, {})
        text_by_id[cid] = normalize_markdown_comment_text(own_text_by_id.get(cid) or "")
        author_by_id[cid] = (meta.get("author") or "").strip()
        date_by_id[cid] = (meta.get("date") or "").strip()

    return {
        "ordered_ids": ordered_started_ids,
        "text_by_id": text_by_id,
        "author_by_id": author_by_id,
        "date_by_id": date_by_id,
        "flattened_by_id": flattened_by_id,
        "parent_by_id": valid_parent_by_id,
        "child_ids": child_ids,
        "root_ids": root_ids,
        "state_by_id": state_by_id,
        "para_by_id": para_by_id,
        "durable_by_id": durable_by_id,
        "presence_by_author": presence_by_author,
    }


def append_comment_paragraph(
    comment_elem: ET.Element,
    text: str,
    with_annotation_ref=False,
    paragraph_attrs=None,
):
    p = ET.SubElement(comment_elem, f"{{{W_NS}}}p")
    if paragraph_attrs:
        for key, value in paragraph_attrs.items():
            p.set(key, value)
    ppr = ET.SubElement(p, f"{{{W_NS}}}pPr")
    pstyle = ET.SubElement(ppr, f"{{{W_NS}}}pStyle")
    pstyle.set(f"{{{W_NS}}}val", "CommentText")

    if with_annotation_ref:
        ref_r = ET.SubElement(p, f"{{{W_NS}}}r")
        ref_rpr = ET.SubElement(ref_r, f"{{{W_NS}}}rPr")
        ref_style = ET.SubElement(ref_rpr, f"{{{W_NS}}}rStyle")
        ref_style.set(f"{{{W_NS}}}val", "CommentReference")
        ET.SubElement(ref_r, f"{{{W_NS}}}annotationRef")

    r = ET.SubElement(p, f"{{{W_NS}}}r")
    t = ET.SubElement(r, f"{{{W_NS}}}t")
    if text[:1].isspace() or text[-1:].isspace():
        t.set(f"{{{XML_NS}}}space", "preserve")
    t.text = text


def topological_comment_order(ordered_ids, parent_by_id):
    pending = [str(cid) for cid in ordered_ids or [] if str(cid)]
    pending_set = set(pending)
    emitted = []
    while pending_set:
        progressed = False
        for cid in pending:
            if cid not in pending_set:
                continue
            parent_id = str((parent_by_id or {}).get(cid) or "").strip()
            if parent_id and parent_id in pending_set:
                continue
            emitted.append(cid)
            pending_set.remove(cid)
            progressed = True
        if progressed:
            continue
        # Should be unreachable after cycle validation, but keep deterministic fallback.
        for cid in pending:
            if cid in pending_set:
                emitted.append(cid)
                pending_set.remove(cid)
    return emitted


def rewrite_comments_from_markdown_threaded(
    docx_dir: Path,
    ordered_ids,
    text_by_id,
    author_by_id,
    date_by_id,
    parent_by_id,
):
    comments_path = docx_dir / "word" / "comments.xml"
    if not comments_path.exists() or not ordered_ids:
        return 0

    tree, root = read_xml(comments_path)
    changed = False

    for child in list(root):
        if local_name(child.tag) == "comment":
            root.remove(child)
            changed = True

    ordered = topological_comment_order(ordered_ids, parent_by_id)
    for cid in ordered:
        comment = ET.SubElement(root, f"{{{W_NS}}}comment")
        comment.set(f"{{{W_NS}}}id", str(cid))

        author = str((author_by_id or {}).get(cid) or "").strip()
        if author:
            comment.set(f"{{{W_NS}}}author", author)

        date = str((date_by_id or {}).get(cid) or "").strip()
        if date:
            comment.set(f"{{{W_NS}}}date", date)

        comment.set(f"{{{W_NS}}}initials", "DC")

        raw = normalize_markdown_comment_text((text_by_id or {}).get(cid) or "")
        lines = raw.split("\n") if raw else [""]

        for line in lines:
            append_comment_paragraph(comment, line, with_annotation_ref=False)

    write_xml(tree, comments_path)
    return 1 if (changed or ordered) else 0


def generate_unique_para_id(seed: str, used_para_ids):
    counter = 0
    while True:
        candidate = f"{zlib.crc32(f'{seed}:{counter}'.encode('utf-8')) & 0xFFFFFFFF:08X}"
        if candidate != "00000000" and candidate not in used_para_ids:
            used_para_ids.add(candidate)
            return candidate
        counter += 1


def generate_unique_durable_id(seed: str, used_durable_ids):
    counter = 0
    while True:
        candidate = f"{zlib.crc32(f'{seed}:{counter}'.encode('utf-8')) & 0xFFFFFFFF:08X}"
        if candidate != "00000000" and candidate not in used_durable_ids:
            used_durable_ids.add(candidate)
            return candidate
        counter += 1


def ensure_word_relationship(docx_dir: Path, rel_type: str, target: str):
    rels_path = docx_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return False

    tree, root = read_xml(rels_path)
    changed = False
    max_rid = 0
    has_ext_rel = False

    for rel in root:
        if local_name(rel.tag) != "Relationship":
            continue
        rid = rel.attrib.get("Id", "")
        m = re.match(r"^rId([0-9]+)$", rid)
        if m:
            max_rid = max(max_rid, int(m.group(1)))
        current_type = rel.attrib.get("Type", "")
        current_target = rel.attrib.get("Target", "")
        if current_type == rel_type:
            has_ext_rel = True
            if current_target != target:
                rel.attrib["Target"] = target
                changed = True

    if not has_ext_rel:
        rel_ns = root.tag.split("}", 1)[0][1:] if root.tag.startswith("{") else PKG_REL_NS
        rel = ET.SubElement(root, f"{{{rel_ns}}}Relationship")
        rel.set("Id", f"rId{max_rid + 1}")
        rel.set("Type", rel_type)
        rel.set("Target", target)
        changed = True

    if changed:
        write_xml(tree, rels_path)
    return changed


def ensure_word_content_type_override(docx_dir: Path, part_name: str, content_type: str):
    content_types_path = docx_dir / "[Content_Types].xml"
    if not content_types_path.exists():
        return False

    tree, root = read_xml(content_types_path)
    changed = False
    has_override = False
    ct_ns = root.tag.split("}", 1)[0][1:] if root.tag.startswith("{") else PKG_CT_NS

    for elem in root:
        if local_name(elem.tag) != "Override":
            continue
        current_part_name = elem.attrib.get("PartName", "")
        if current_part_name == part_name:
            has_override = True
            if elem.attrib.get("ContentType", "") != content_type:
                elem.attrib["ContentType"] = content_type
                changed = True
            break

    if not has_override:
        override = ET.SubElement(root, f"{{{ct_ns}}}Override")
        override.set("PartName", part_name)
        override.set("ContentType", content_type)
        changed = True

    if changed:
        write_xml(tree, content_types_path)
    return changed


def ensure_comments_extended_relationship(docx_dir: Path):
    return ensure_word_relationship(docx_dir, COMMENTS_EXT_REL_TYPE, "commentsExtended.xml")


def ensure_comments_ids_relationship(docx_dir: Path):
    return ensure_word_relationship(docx_dir, COMMENTS_IDS_REL_TYPE, "commentsIds.xml")


def ensure_comments_extensible_relationship(docx_dir: Path):
    return ensure_word_relationship(docx_dir, COMMENTS_EXTENSIBLE_REL_TYPE, "commentsExtensible.xml")


def ensure_comments_extended_content_type(docx_dir: Path):
    return ensure_word_content_type_override(docx_dir, COMMENTS_EXT_PART_NAME, COMMENTS_EXT_CONTENT_TYPE)


def ensure_comments_ids_content_type(docx_dir: Path):
    return ensure_word_content_type_override(docx_dir, COMMENTS_IDS_PART_NAME, COMMENTS_IDS_CONTENT_TYPE)


def ensure_comments_extensible_content_type(docx_dir: Path):
    return ensure_word_content_type_override(
        docx_dir,
        COMMENTS_EXTENSIBLE_PART_NAME,
        COMMENTS_EXTENSIBLE_CONTENT_TYPE,
    )


def ensure_people_relationship(docx_dir: Path):
    return ensure_word_relationship(docx_dir, PEOPLE_REL_TYPE, "people.xml")


def ensure_people_content_type(docx_dir: Path):
    return ensure_word_content_type_override(docx_dir, PEOPLE_PART_NAME, PEOPLE_CONTENT_TYPE)


def rewrite_people_part(docx_dir: Path, authors, presence_by_author=None):
    people_root = ET.Element(f"{{{W15_NS}}}people")
    for author in sorted({str(a).strip() for a in (authors or []) if str(a).strip()}):
        person = ET.SubElement(people_root, f"{{{W15_NS}}}person")
        person.set(f"{{{W15_NS}}}author", author)
        presence = (presence_by_author or {}).get(author) or {}
        provider_id = str(presence.get("provider_id") or "").strip()
        user_id = str(presence.get("user_id") or "").strip()
        if provider_id or user_id:
            presence_info = ET.SubElement(person, f"{{{W15_NS}}}presenceInfo")
            if provider_id:
                presence_info.set(f"{{{W15_NS}}}providerId", provider_id)
            if user_id:
                presence_info.set(f"{{{W15_NS}}}userId", user_id)
    people_path = docx_dir / "word" / "people.xml"
    write_xml(ET.ElementTree(people_root), people_path)
    return True


def ensure_comments_xml_state_compatibility(root: ET.Element):
    changed = False
    current_ignorable = root.attrib.get(f"{{{MC_NS}}}Ignorable")
    if current_ignorable is not None:
        normalized = " ".join([tok for tok in str(current_ignorable).split() if tok])
        if normalized != current_ignorable:
            root.set(f"{{{MC_NS}}}Ignorable", normalized)
            changed = True

    return changed


def load_comments_ids_durable_map(docx_dir: Path):
    comments_ids_path = docx_dir / "word" / "commentsIds.xml"
    para_to_durable = {}
    used_durable_ids = set()
    if not comments_ids_path.exists():
        return para_to_durable, used_durable_ids

    _, root = read_xml(comments_ids_path)
    for elem in root.iter():
        if local_name(elem.tag) != "commentId":
            continue
        para_id = get_attr_local(elem, "paraId")
        durable_id = get_attr_local(elem, "durableId")
        if durable_id:
            used_durable_ids.add(durable_id)
        if para_id and durable_id:
            para_to_durable[para_id] = durable_id
    return para_to_durable, used_durable_ids


def rewrite_comments_extended_state(
    docx_dir: Path,
    ordered_ids,
    parent_by_id,
    state_by_id,
    para_by_id=None,
    durable_by_id=None,
    presence_by_author=None,
):
    comments_path = docx_dir / "word" / "comments.xml"
    if not comments_path.exists() or not ordered_ids:
        return 0

    tree, root = read_xml(comments_path)
    changed_comments_xml = False
    authors = set()

    if ensure_comments_xml_state_compatibility(root):
        changed_comments_xml = True

    used_para_ids = set()
    comments_by_id = {}
    for comment in root.findall(f".//{{{W_NS}}}comment"):
        cid = get_attr_local(comment, "id")
        if cid is None:
            continue
        comments_by_id[cid] = comment
        for attr_name in list(comment.attrib.keys()):
            if attr_name == "state" or attr_name.endswith("}state"):
                del comment.attrib[attr_name]
                changed_comments_xml = True
        para_id = get_attr_local(comment, "paraId")
        if para_id:
            used_para_ids.add(para_id)
        for p in comment.findall(f"./{{{W_NS}}}p"):
            p_para_id = get_attr_local(p, "paraId")
            if p_para_id:
                used_para_ids.add(p_para_id)

    existing_durable_by_para, used_durable_ids = load_comments_ids_durable_map(docx_dir)
    ordered = [cid for cid in topological_comment_order(ordered_ids, parent_by_id) if cid in comments_by_id]
    if not ordered:
        return 0

    comment_meta_by_id = {}
    for cid in ordered:
        comment = comments_by_id[cid]
        paragraphs = comment.findall(f"./{{{W_NS}}}p")
        if not paragraphs:
            append_comment_paragraph(comment, "", with_annotation_ref=True)
            paragraphs = comment.findall(f"./{{{W_NS}}}p")
            changed_comments_xml = True
        thread_p = paragraphs[-1]

        preferred_para_id = str((para_by_id or {}).get(cid) or "").strip()
        para_id = get_attr_local(thread_p, "paraId") or get_attr_local(comment, "paraId")
        if preferred_para_id:
            if para_id != preferred_para_id:
                thread_p.set(f"{{{W14_NS}}}paraId", preferred_para_id)
                changed_comments_xml = True
            para_id = preferred_para_id
        if not para_id:
            para_id = generate_unique_para_id(f"comment-{cid}", used_para_ids)
            thread_p.set(f"{{{W14_NS}}}paraId", para_id)
            changed_comments_xml = True
        else:
            used_para_ids.add(para_id)

        durable_id = str((durable_by_id or {}).get(cid) or "").strip() or existing_durable_by_para.get(para_id)
        if not durable_id:
            durable_id = generate_unique_durable_id(f"durable-{para_id}", used_durable_ids)
        else:
            used_durable_ids.add(durable_id)

        author = (get_attr_local(comment, "author") or "").strip()
        if author:
            authors.add(author)

        comment_meta_by_id[cid] = {
            "para_id": para_id,
            "durable_id": durable_id,
            "date": get_attr_local(comment, "date") or "",
        }

    if changed_comments_xml:
        write_xml(tree, comments_path)

    comments_ext_root = ET.Element(f"{{{W15_NS}}}commentsEx")
    for cid in ordered:
        meta = comment_meta_by_id.get(cid) or {}
        para_id = meta.get("para_id")
        if not para_id:
            continue
        state_token = parse_state_token((state_by_id or {}).get(cid))
        entry = ET.SubElement(comments_ext_root, f"{{{W15_NS}}}commentEx")
        entry.set(f"{{{W15_NS}}}paraId", para_id)
        entry.set(f"{{{W15_NS}}}done", "1" if state_token == "resolved" else "0")
        parent_id = str((parent_by_id or {}).get(cid) or "").strip()
        if parent_id:
            parent_meta = comment_meta_by_id.get(parent_id) or {}
            parent_para_id = str(parent_meta.get("para_id") or "").strip()
            if parent_para_id:
                entry.set(f"{{{W15_NS}}}paraIdParent", parent_para_id)

    comments_ids_root = ET.Element(f"{{{W16CID_NS}}}commentsIds")
    comments_extensible_root = ET.Element(f"{{{W16CEX_NS}}}commentsExtensible")
    for cid in ordered:
        meta = comment_meta_by_id.get(cid) or {}
        para_id = meta.get("para_id")
        durable_id = meta.get("durable_id")
        date_utc = meta.get("date") or ""
        if not para_id or not durable_id:
            continue
        id_entry = ET.SubElement(comments_ids_root, f"{{{W16CID_NS}}}commentId")
        id_entry.set(f"{{{W16CID_NS}}}paraId", para_id)
        id_entry.set(f"{{{W16CID_NS}}}durableId", durable_id)
        ext_entry = ET.SubElement(comments_extensible_root, f"{{{W16CEX_NS}}}commentExtensible")
        ext_entry.set(f"{{{W16CEX_NS}}}durableId", durable_id)
        if date_utc:
            ext_entry.set(f"{{{W16CEX_NS}}}dateUtc", date_utc)

    comments_ext_path = docx_dir / "word" / "commentsExtended.xml"
    comments_ids_path = docx_dir / "word" / "commentsIds.xml"
    comments_extensible_path = docx_dir / "word" / "commentsExtensible.xml"
    rewrite_people_part(docx_dir, authors, presence_by_author)
    write_xml(ET.ElementTree(comments_ext_root), comments_ext_path)
    write_xml(ET.ElementTree(comments_ids_root), comments_ids_path)
    write_xml(ET.ElementTree(comments_extensible_root), comments_extensible_path)

    rel_changed = False
    rel_changed |= ensure_comments_extended_relationship(docx_dir)
    rel_changed |= ensure_comments_ids_relationship(docx_dir)
    rel_changed |= ensure_comments_extensible_relationship(docx_dir)
    rel_changed |= ensure_people_relationship(docx_dir)

    ct_changed = False
    ct_changed |= ensure_comments_extended_content_type(docx_dir)
    ct_changed |= ensure_comments_ids_content_type(docx_dir)
    ct_changed |= ensure_comments_extensible_content_type(docx_dir)
    ct_changed |= ensure_people_content_type(docx_dir)

    return 1 if (changed_comments_xml or rel_changed or ct_changed or ordered) else 0


def word_story_xml_candidates(docx_dir: Path):
    word_dir = docx_dir / "word"
    candidates = [
        word_dir / "document.xml",
        word_dir / "footnotes.xml",
        word_dir / "endnotes.xml",
    ]
    candidates.extend(sorted(word_dir.glob("header*.xml")))
    candidates.extend(sorted(word_dir.glob("footer*.xml")))
    return [p for p in candidates if p.exists()]


def collect_story_marker_counts(docx_dir: Path):
    counts = {
        "start": {},
        "end": {},
        "ref": {},
    }
    marker_to_bucket = {
        "commentRangeStart": "start",
        "commentRangeEnd": "end",
        "commentReference": "ref",
    }
    for xml_path in word_story_xml_candidates(docx_dir):
        _, root = read_xml(xml_path)
        for elem in root.iter():
            bucket = marker_to_bucket.get(local_name(elem.tag))
            if not bucket:
                continue
            cid = (get_attr_local(elem, "id") or "").strip()
            if not cid:
                continue
            counts[bucket][cid] = int(counts[bucket].get(cid, 0)) + 1
    return counts


def synthesize_child_markers_in_story(
    root: ET.Element,
    parent_id: str,
    child_id: str,
    need_start: bool,
    need_end: bool,
    need_ref: bool,
):
    inserted = {"start": 0, "end": 0, "ref": 0}
    parent_id = str(parent_id)
    child_id = str(child_id)

    def is_marker(elem: ET.Element, marker_name: str, cid: str) -> bool:
        if local_name(elem.tag) != marker_name:
            return False
        return (get_attr_local(elem, "id") or "").strip() == cid

    def make_marker(local_tag: str) -> ET.Element:
        marker = ET.Element(f"{{{W_NS}}}{local_tag}")
        marker.set(f"{{{W_NS}}}id", child_id)
        return marker

    for container in root.iter():
        children = list(container)
        if not children:
            continue
        idx = 0
        while idx < len(children):
            elem = children[idx]

            if need_start and inserted["start"] == 0 and is_marker(elem, "commentRangeStart", parent_id):
                insert_at = idx + 1
                while insert_at < len(children) and local_name(children[insert_at].tag) == "commentRangeStart":
                    insert_at += 1
                container.insert(insert_at, make_marker("commentRangeStart"))
                inserted["start"] += 1
                children = list(container)
                idx = insert_at + 1
                continue

            if need_end and inserted["end"] == 0 and is_marker(elem, "commentRangeEnd", parent_id):
                insert_at = idx
                while insert_at > 0 and local_name(children[insert_at - 1].tag) == "commentRangeEnd":
                    insert_at -= 1
                container.insert(insert_at, make_marker("commentRangeEnd"))
                inserted["end"] += 1
                children = list(container)
                idx += 1
                continue

            if need_ref and inserted["ref"] == 0 and is_marker(elem, "commentReference", parent_id):
                insert_at = idx + 1
                while insert_at < len(children) and local_name(children[insert_at].tag) == "commentReference":
                    insert_at += 1
                container.insert(insert_at, make_marker("commentReference"))
                inserted["ref"] += 1
                children = list(container)
                idx = insert_at + 1
                continue

            idx += 1

            if (
                (not need_start or inserted["start"] > 0)
                and (not need_end or inserted["end"] > 0)
                and (not need_ref or inserted["ref"] > 0)
            ):
                return inserted

    return inserted


def ensure_thread_reply_anchors(docx_dir: Path, ordered_ids, parent_by_id):
    child_ids = [str(cid) for cid in (ordered_ids or []) if str((parent_by_id or {}).get(str(cid)) or "").strip()]
    if not child_ids:
        return 0

    marker_counts = collect_story_marker_counts(docx_dir)
    changed_files = 0
    unresolved = []

    for child_id in topological_comment_order(ordered_ids or [], parent_by_id or {}):
        parent_id = str((parent_by_id or {}).get(str(child_id)) or "").strip()
        if not parent_id:
            continue
        child_id = str(child_id)
        has_start = int(marker_counts["start"].get(child_id, 0)) > 0
        has_end = int(marker_counts["end"].get(child_id, 0)) > 0
        has_ref = int(marker_counts["ref"].get(child_id, 0)) > 0
        if has_start and has_end and has_ref:
            continue

        if (
            int(marker_counts["start"].get(parent_id, 0)) < 1
            or int(marker_counts["end"].get(parent_id, 0)) < 1
            or int(marker_counts["ref"].get(parent_id, 0)) < 1
        ):
            unresolved.append(
                f"cannot synthesize anchors for reply {child_id}: parent {parent_id} has incomplete anchors"
            )
            continue

        need_start = not has_start
        need_end = not has_end
        need_ref = not has_ref

        for xml_path in word_story_xml_candidates(docx_dir):
            if not (need_start or need_end or need_ref):
                break
            tree, root = read_xml(xml_path)
            inserted = synthesize_child_markers_in_story(
                root,
                parent_id=parent_id,
                child_id=child_id,
                need_start=need_start,
                need_end=need_end,
                need_ref=need_ref,
            )
            if inserted["start"] or inserted["end"] or inserted["ref"]:
                write_xml(tree, xml_path)
                changed_files += 1
                marker_counts["start"][child_id] = int(marker_counts["start"].get(child_id, 0)) + int(
                    inserted["start"]
                )
                marker_counts["end"][child_id] = int(marker_counts["end"].get(child_id, 0)) + int(inserted["end"])
                marker_counts["ref"][child_id] = int(marker_counts["ref"].get(child_id, 0)) + int(inserted["ref"])
                need_start = int(marker_counts["start"].get(child_id, 0)) < 1
                need_end = int(marker_counts["end"].get(child_id, 0)) < 1
                need_ref = int(marker_counts["ref"].get(child_id, 0)) < 1

        if need_start or need_end or need_ref:
            missing = []
            if need_start:
                missing.append("commentRangeStart")
            if need_end:
                missing.append("commentRangeEnd")
            if need_ref:
                missing.append("commentReference")
            unresolved.append(
                f"reply {child_id} still missing: {', '.join(missing)} (parent {parent_id})"
            )

    if unresolved:
        raise ValueError(
            "Unable to restore reply anchors required for threaded Word comments.\n"
            + "\n".join([f"- {issue}" for issue in unresolved])
            + "\n- Fix malformed or missing parent anchors in markdown/converted DOCX and retry."
        )
    return changed_files


def prune_child_comment_artifacts(docx_dir: Path, child_ids):
    child_set = {str(cid) for cid in (child_ids or []) if str(cid)}
    if not child_set:
        return 0

    removed = 0

    # Remove child anchors/references in all word "story" XMLs.
    for xml_path in word_story_xml_candidates(docx_dir):
        tree, root = read_xml(xml_path)
        changed = False
        for parent in root.iter():
            for child in list(parent):
                lname = local_name(child.tag)
                if lname not in {"commentRangeStart", "commentRangeEnd", "commentReference"}:
                    continue
                cid = get_attr_local(child, "id")
                if cid in child_set:
                    parent.remove(child)
                    removed += 1
                    changed = True
        if changed:
            write_xml(tree, xml_path)

    # Remove child comment nodes from comments.xml and capture paraIds.
    child_para_ids = set()
    comments_path = docx_dir / "word" / "comments.xml"
    if comments_path.exists():
        tree, root = read_xml(comments_path)
        changed = False
        for parent in root.iter():
            for child in list(parent):
                if local_name(child.tag) != "comment":
                    continue
                cid = get_attr_local(child, "id")
                if cid not in child_set:
                    continue
                para_id = get_attr_local(child, "paraId")
                if not para_id:
                    for node in child.iter():
                        if local_name(node.tag) == "p":
                            para_id = get_attr_local(node, "paraId")
                            if para_id:
                                break
                if para_id:
                    child_para_ids.add(para_id)
                parent.remove(child)
                removed += 1
                changed = True
        if changed:
            write_xml(tree, comments_path)

    # Keep package internals consistent: prune child entries in extension/id files.
    if child_para_ids:
        child_durable_ids = set()
        comments_ext_path = docx_dir / "word" / "commentsExtended.xml"
        if comments_ext_path.exists():
            tree, root = read_xml(comments_ext_path)
            changed = False
            for parent in root.iter():
                for child in list(parent):
                    if local_name(child.tag) != "commentEx":
                        continue
                    para_id = get_attr_local(child, "paraId")
                    if para_id in child_para_ids:
                        parent.remove(child)
                        removed += 1
                        changed = True
            if changed:
                write_xml(tree, comments_ext_path)

        comments_ids_path = docx_dir / "word" / "commentsIds.xml"
        if comments_ids_path.exists():
            tree, root = read_xml(comments_ids_path)
            changed = False
            for parent in root.iter():
                for child in list(parent):
                    if local_name(child.tag) != "commentId":
                        continue
                    para_id = get_attr_local(child, "paraId")
                    if para_id in child_para_ids:
                        durable_id = get_attr_local(child, "durableId")
                        if durable_id:
                            child_durable_ids.add(durable_id)
                        parent.remove(child)
                        removed += 1
                        changed = True
            if changed:
                write_xml(tree, comments_ids_path)

        comments_extensible_path = docx_dir / "word" / "commentsExtensible.xml"
        if comments_extensible_path.exists() and child_durable_ids:
            tree, root = read_xml(comments_extensible_path)
            changed = False
            for parent in root.iter():
                for child in list(parent):
                    if local_name(child.tag) != "commentExtensible":
                        continue
                    durable_id = get_attr_local(child, "durableId")
                    if durable_id in child_durable_ids:
                        parent.remove(child)
                        removed += 1
                        changed = True
            if changed:
                write_xml(tree, comments_extensible_path)

    return removed


def extract_comment_text(comment_elem: ET.Element) -> str:
    paragraphs = []
    for p in comment_elem.iter(f"{{{W_NS}}}p"):
        pieces = []
        for node in p.iter():
            lname = local_name(node.tag)
            if lname == "t" and node.text:
                pieces.append(node.text)
            elif lname == "tab":
                pieces.append("\t")
            elif lname in {"br", "cr"}:
                pieces.append("\n")
        text = "".join(pieces).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs).strip()


def comment_paragraph_para_ids(comment_elem: ET.Element):
    para_ids = []
    for p in comment_elem.findall(f"./{{{W_NS}}}p"):
        para_id = get_attr_local(p, "paraId")
        if para_id:
            para_ids.append(para_id)
    return para_ids


def comment_thread_para_id(comment_elem: ET.Element):
    # Word's commentEx/commentsIds mapping uses the thread paragraph ID,
    # which for multi-paragraph comments is the last paragraph's paraId.
    para_ids = comment_paragraph_para_ids(comment_elem)
    if para_ids:
        return para_ids[-1]
    return get_attr_local(comment_elem, "paraId") or ""


def parse_docx_comments(docx_dir: Path):
    comments_path = docx_dir / "word" / "comments.xml"
    comments_ext_path = docx_dir / "word" / "commentsExtended.xml"
    comments_ids_path = docx_dir / "word" / "commentsIds.xml"

    comments = {}
    parent_map = {}
    para_to_id = {}
    durable_by_para = {}
    ordered_comment_ids = []
    resolved_by_id = {}

    if comments_path.exists():
        _, root = read_xml(comments_path)
        for idx, comment in enumerate(root.findall(f".//{{{W_NS}}}comment")):
            cid = get_attr_local(comment, "id")
            if cid is None:
                continue
            author = get_attr_local(comment, "author") or ""
            date = get_attr_local(comment, "date") or ""
            text = extract_comment_text(comment)
            para_id = comment_thread_para_id(comment)
            for p_para_id in comment_paragraph_para_ids(comment):
                para_to_id[p_para_id] = cid
            parent = get_attr_local(comment, "parentId")
            if para_id:
                para_to_id[para_id] = cid
            if parent:
                parent_map[cid] = parent
            ordered_comment_ids.append(cid)
            resolved_by_id[cid] = False
            comments[cid] = {
                "author": author,
                "date": date,
                "text": text,
                "order": idx,
                "para_id": para_id or "",
                "durable_id": "",
                "resolved": False,
            }

    # Some Word versions store paraId only in commentsIds.xml.
    # In these files, comments.xml and commentsIds.xml are aligned by order.
    if comments_ids_path.exists() and ordered_comment_ids:
        _, cid_root = read_xml(comments_ids_path)
        para_ids_in_order = []
        for elem in cid_root.iter():
            if local_name(elem.tag) != "commentId":
                continue
            para_id = get_attr_local(elem, "paraId")
            durable_id = get_attr_local(elem, "durableId")
            if para_id:
                para_ids_in_order.append(para_id)
            if para_id and durable_id:
                durable_by_para[para_id] = durable_id
        if len(para_ids_in_order) == len(ordered_comment_ids):
            for comment_id, para_id in zip(ordered_comment_ids, para_ids_in_order):
                if para_id and para_id not in para_to_id:
                    para_to_id[para_id] = comment_id
                if para_id and comment_id in comments:
                    comments[comment_id]["para_id"] = para_id

    for cid, meta in comments.items():
        para_id = (meta.get("para_id") or "").strip()
        if para_id and para_id in durable_by_para:
            comments[cid]["durable_id"] = durable_by_para[para_id]

    if comments_ext_path.exists() and para_to_id:
        _, root = read_xml(comments_ext_path)
        for elem in root.iter():
            if local_name(elem.tag) != "commentEx":
                continue
            para_id = get_attr_local(elem, "paraId")
            parent_para_id = get_attr_local(elem, "paraIdParent")
            done = get_attr_local(elem, "done")
            child_id = para_to_id.get(para_id) if para_id else None
            if child_id:
                is_resolved = str(done or "").strip() == "1"
                resolved_by_id[child_id] = is_resolved
                if child_id in comments:
                    comments[child_id]["resolved"] = is_resolved
                    if para_id:
                        comments[child_id]["para_id"] = para_id
            if not para_id or not parent_para_id:
                continue
            parent_id = para_to_id.get(parent_para_id)
            if child_id and parent_id and child_id not in parent_map:
                parent_map[child_id] = parent_id

    children = {cid: [] for cid in comments}
    for child_id, parent_id in parent_map.items():
        if child_id in comments and parent_id in comments:
            children[parent_id].append(child_id)
    for sibling_ids in children.values():
        sibling_ids.sort(key=lambda cid: comments[cid]["order"])

    for cid in comments:
        comments[cid]["resolved"] = bool(resolved_by_id.get(cid, False))

    return comments, parent_map, children


def parse_docx_people_presence(docx_dir: Path):
    people_path = docx_dir / "word" / "people.xml"
    presence_by_author = {}
    if not people_path.exists():
        return presence_by_author

    _, root = read_xml(people_path)
    for person in root.iter():
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
        if provider_id or user_id:
            presence_by_author[author] = {
                "provider_id": provider_id,
                "user_id": user_id,
            }

    return presence_by_author


def collect_anchors_from_xml(xml_path: Path):
    _, root = read_xml(xml_path)
    anchors = []
    for elem in root.iter():
        if local_name(elem.tag) == "commentRangeStart":
            cid = get_attr_local(elem, "id")
            if cid is not None:
                anchors.append(cid)
    return anchors


def get_anchor_comment_ids(docx_dir: Path):
    anchors = []
    for xml_path in word_story_xml_candidates(docx_dir):
        anchors.extend(collect_anchors_from_xml(xml_path))

    seen = set()
    ordered = []
    for cid in anchors:
        if cid not in seen:
            seen.add(cid)
            ordered.append(cid)
    return ordered


def thread_root(comment_id: str, comments, parent_map) -> str:
    root_id = comment_id
    seen = set()
    while root_id in parent_map and root_id not in seen:
        seen.add(root_id)
        parent_id = parent_map[root_id]
        if parent_id not in comments:
            break
        root_id = parent_id
    return root_id


def flatten_thread(anchor_id: str, comments, parent_map, children):
    if anchor_id not in comments:
        return anchor_id, ""

    root_id = thread_root(anchor_id, comments, parent_map)
    ordered_ids = []
    seen_ids = set()

    def walk(cid):
        if cid in seen_ids:
            return
        seen_ids.add(cid)
        ordered_ids.append(cid)
        for child_id in children.get(cid, []):
            walk(child_id)

    walk(root_id)

    if len(ordered_ids) == 1:
        text = comments[root_id]["text"] or "(empty)"
        return root_id, text

    blocks = []
    for cid in ordered_ids:
        entry = comments[cid]
        who = entry["author"] or "Unknown"
        when = entry["date"]
        head = f"{who} ({when})" if when else who
        body = entry["text"] or "(empty)"
        blocks.append(f"{head}: {body}")
    return root_id, "\n\n".join(blocks)


def make_comment_element(comment_id: str, author: str, date: str, text: str) -> ET.Element:
    comment = ET.Element(f"{{{W_NS}}}comment")
    comment.set(f"{{{W_NS}}}id", comment_id)
    if author:
        comment.set(f"{{{W_NS}}}author", author)
    if date:
        comment.set(f"{{{W_NS}}}date", date)
    comment.set(f"{{{W_NS}}}initials", "DC")

    lines = text.splitlines() if text else [""]
    for line in lines:
        p = ET.SubElement(comment, f"{{{W_NS}}}p")
        r = ET.SubElement(p, f"{{{W_NS}}}r")
        t = ET.SubElement(r, f"{{{W_NS}}}t")
        if line[:1].isspace() or line[-1:].isspace():
            t.set(f"{{{XML_NS}}}space", "preserve")
        t.text = line
    return comment


def rewrite_comments_with_flattened_threads(docx_dir: Path):
    comments_path = docx_dir / "word" / "comments.xml"
    if not comments_path.exists() or not (docx_dir / "word" / "document.xml").exists():
        return 0

    comments, parent_map, children = parse_docx_comments(docx_dir)
    if not comments:
        return 0

    anchors = get_anchor_comment_ids(docx_dir)
    if not anchors:
        return 0

    new_root = ET.Element(f"{{{W_NS}}}comments")
    count = 0

    for anchor_id in anchors:
        root_id, flat_text = flatten_thread(anchor_id, comments, parent_map, children)
        meta = comments.get(root_id) or comments.get(anchor_id) or {}
        author = meta.get("author", "")
        date = meta.get("date", "")
        if not flat_text:
            flat_text = "(comment content unavailable)"
        new_root.append(make_comment_element(anchor_id, author, date, flat_text))
        count += 1

    write_xml(ET.ElementTree(new_root), comments_path)

    # Remove thread-specific extras so the temporary package is internally consistent.
    for rel_path in [
        "word/commentsExtended.xml",
        "word/commentsIds.xml",
        "word/people.xml",
        "word/_rels/comments.xml.rels",
    ]:
        p = docx_dir / rel_path
        if p.exists():
            p.unlink()

    return count


def convert_docx_to_md(in_docx: Path, out_md: Path, pandoc_extra_args):
    with tempfile.TemporaryDirectory(prefix=".docx-comments-", dir=temp_dir_root_for(in_docx)) as tmp:
        tmp_dir = Path(tmp)
        src_dir = tmp_dir / "src"
        src_dir.mkdir(parents=True, exist_ok=True)
        out_md.parent.mkdir(parents=True, exist_ok=True)
        media_dir = out_md.parent / "media"
        media_before = list_files_relative(media_dir)

        extract_docx(in_docx, src_dir)
        comments, parent_map, _ = parse_docx_comments(src_dir)
        people_presence_by_author = parse_docx_people_presence(src_dir)
        state_by_id = {}
        para_by_id = {}
        durable_by_id = {}
        presence_provider_by_id = {}
        presence_user_by_id = {}
        for cid, meta in comments.items():
            state_by_id[cid] = "resolved" if meta.get("resolved") else "active"
            para_id = (meta.get("para_id") or "").strip()
            if para_id:
                para_by_id[cid] = para_id
            durable_id = (meta.get("durable_id") or "").strip()
            if durable_id:
                durable_by_id[cid] = durable_id
            author = (meta.get("author") or "").strip()
            if author:
                presence = people_presence_by_author.get(author) or {}
                provider_id = str(presence.get("provider_id") or "").strip()
                user_id = str(presence.get("user_id") or "").strip()
                if provider_id:
                    presence_provider_by_id[cid] = provider_id
                if user_id:
                    presence_user_by_id[cid] = user_id

        args = ["--track-changes=all"]
        if pandoc_extra_args:
            args.extend(pandoc_extra_args)
        if not has_extract_media_arg(args):
            args.append("--extract-media=.")
        run_pandoc(
            in_docx,
            out_md,
            fmt_to="markdown",
            extra_args=args,
            cwd=out_md.parent,
        )
        md_writer = resolve_pandoc_writer_format(pandoc_extra_args, default_format="markdown")
        annotate_markdown_comment_attrs(
            out_md,
            parent_map,
            state_by_id,
            para_by_id,
            durable_by_id,
            presence_provider_by_id,
            presence_user_by_id,
            pandoc_extra_args=pandoc_extra_args,
            writer_format=md_writer,
            cwd=out_md.parent,
        )
        text = out_md.read_text(encoding="utf-8")
        text, _ = repair_unbalanced_comment_markers(text)
        text, _ = normalize_nested_comment_end_markers(text)
        cleaned, _ = strip_placeholder_shape_images(text)
        out_md.write_text(cleaned, encoding="utf-8")
        comment_cards_by_id = {}
        for cid, meta in comments.items():
            comment_cards_by_id[cid] = {
                "author": (meta.get("author") or "").strip(),
                "date": (meta.get("date") or "").strip(),
                "parent": (parent_map.get(cid) or "").strip(),
                "state": "resolved" if meta.get("resolved") else "active",
                "paraId": (meta.get("para_id") or "").strip(),
                "durableId": (meta.get("durable_id") or "").strip(),
                "presenceProvider": (presence_provider_by_id.get(cid) or "").strip(),
                "presenceUserId": (presence_user_by_id.get(cid) or "").strip(),
                "text": normalize_markdown_comment_text(meta.get("text") or ""),
            }
        emit_milestones_and_cards_ast(
            out_md,
            comment_cards_by_id,
            set(parent_map.keys()),
            pandoc_extra_args=pandoc_extra_args,
            writer_format=md_writer,
            cwd=out_md.parent,
        )
        cleaned = out_md.read_text(encoding="utf-8")
        prune_unreferenced_new_media(media_dir, media_before, cleaned)


def convert_md_to_docx(in_md: Path, out_docx: Path, pandoc_extra_args):
    with tempfile.TemporaryDirectory(prefix=".docx-comments-mdinput-", dir=temp_dir_root_for(in_md)) as tmp:
        tmp_dir = Path(tmp)
        sanitized_md = tmp_dir / "input.md"
        normalized_md = tmp_dir / "normalized.md"
        pandoc_input_md = tmp_dir / "pandoc-input.md"
        text = in_md.read_text(encoding="utf-8")
        cleaned, _ = strip_placeholder_shape_images(text)
        sanitized_md.write_text(cleaned, encoding="utf-8")
        _, card_by_id = normalize_milestone_tokens_ast(
            sanitized_md,
            normalized_md,
            pandoc_extra_args=pandoc_extra_args,
            writer_format="markdown",
            cwd=in_md.parent,
        )
        normalized_text = normalized_md.read_text(encoding="utf-8")
        normalized_text, _ = normalize_nested_comment_end_markers(normalized_text)
        normalized_md.write_text(normalized_text, encoding="utf-8")
        validate_comment_marker_integrity(
            cleaned,
            normalized_text,
            card_by_id=card_by_id,
            source_label=str(in_md),
        )

        comment_data = extract_comment_texts_from_markdown(
            normalized_md,
            pandoc_extra_args,
            card_by_id=card_by_id,
        )
        strip_comment_transport_attrs_ast(
            normalized_md,
            pandoc_input_md,
            pandoc_extra_args=pandoc_extra_args,
            writer_format="markdown",
            cwd=in_md.parent,
        )
        pandoc_text = pandoc_input_md.read_text(encoding="utf-8")
        pandoc_text, _ = normalize_nested_comment_end_markers(pandoc_text)
        pandoc_input_md.write_text(pandoc_text, encoding="utf-8")

        run_pandoc(
            pandoc_input_md,
            out_docx,
            fmt_from="markdown",
            extra_args=pandoc_extra_args,
            cwd=in_md.parent,
        )
    if not comment_data or not comment_data.get("ordered_ids"):
        return

    with tempfile.TemporaryDirectory(prefix=".docx-comments-md2docx-", dir=temp_dir_root_for(in_md)) as tmp:
        tmp_dir = Path(tmp)
        unpacked = tmp_dir / "docx"
        unpacked.mkdir(parents=True, exist_ok=True)
        extract_docx(out_docx, unpacked)
        changed = rewrite_comments_from_markdown_threaded(
            unpacked,
            comment_data.get("ordered_ids", []),
            comment_data.get("text_by_id", {}),
            comment_data.get("author_by_id", {}),
            comment_data.get("date_by_id", {}),
            comment_data.get("parent_by_id", {}),
        )
        anchor_changed = ensure_thread_reply_anchors(
            unpacked,
            comment_data.get("ordered_ids", []),
            comment_data.get("parent_by_id", {}),
        )
        state_updated = rewrite_comments_extended_state(
            unpacked,
            comment_data.get("ordered_ids", []),
            comment_data.get("parent_by_id", {}),
            comment_data.get("state_by_id", {}),
            comment_data.get("para_by_id", {}),
            comment_data.get("durable_by_id", {}),
            comment_data.get("presence_by_author", {}),
        )
        if not changed and not state_updated and not anchor_changed:
            return
        patched_docx = tmp_dir / "patched.docx"
        pack_docx(unpacked, patched_docx)
        shutil.copyfile(patched_docx, out_docx)


def default_out_path(in_path: Path, suffix: str):
    return in_path.with_suffix(suffix)


def detect_mode_from_path(input_path: Path) -> str:
    suffix = input_path.suffix.lower()
    if suffix == ".docx":
        return "docx2md"
    if suffix in {".md", ".markdown", ".mdown", ".mkd"}:
        return "md2docx"
    raise ValueError(
        f"Cannot infer conversion mode from '{input_path.name}'. "
        "Use .docx, .md, .markdown, .mdown, or .mkd, or pass --mode explicitly."
    )


def normalize_argv(argv):
    # Backward compatibility: docx-comments docx2md input.docx -o out.md
    # becomes docx-comments --mode docx2md input.docx -o out.md
    if argv and argv[0] in {"docx2md", "md2docx"}:
        return ["--mode", argv[0], *argv[1:]]
    return argv


def build_parser(prog_name=None):
    parser = argparse.ArgumentParser(
        prog=prog_name or os.environ.get("DOCX_COMMENTS_PROG") or (Path(sys.argv[0]).name or "docx-comments"),
        description=(
            "Convert .docx <-> markdown while preserving comments. "
            "Threaded Word comments are reconstructed as native reply threads in .docx output."
        ),
    )
    parser.add_argument("input", type=Path, help="Input file (.docx or markdown)")
    parser.add_argument("-o", "--output", type=Path, help="Output file path")
    parser.add_argument(
        "--mode",
        choices=["auto", "docx2md", "md2docx"],
        default="auto",
        help="Conversion mode; default auto infers from input extension",
    )
    return parser


def parse_pandoc_version(version_output: str):
    lines = (version_output or "").splitlines()
    first_line = lines[0] if lines else ""
    m = re.search(r"\b(\d+)\.(\d+)(?:\.(\d+))?", first_line)
    if not m:
        return None
    major = int(m.group(1))
    minor = int(m.group(2))
    patch = int(m.group(3) or 0)
    return (major, minor, patch)


def check_prerequisites():
    pandoc_bin = shutil.which("pandoc")
    if pandoc_bin is None:
        raise RuntimeError("pandoc is not installed or not on PATH.")

    proc = subprocess.run([pandoc_bin, "--version"], capture_output=True, text=True, check=False)
    if proc.returncode != 0:
        raise RuntimeError(f"failed to run pandoc --version (exit {proc.returncode}).")

    version = parse_pandoc_version(proc.stdout)
    if version is None:
        raise RuntimeError("could not parse pandoc version output.")

    minimum = MIN_PANDOC_VERSION + (0,)
    if version < minimum:
        current = ".".join(str(x) for x in version)
        required = ".".join(str(x) for x in minimum[:2])
        raise RuntimeError(
            f"pandoc version {current} is too old; require at least {required}."
        )


def run_conversion(mode: str, input_path: Path, output_path: Path | None, pandoc_extra_args):
    check_prerequisites()
    resolved_mode = mode
    if resolved_mode == "auto":
        resolved_mode = detect_mode_from_path(input_path)

    if resolved_mode == "docx2md":
        out_md = output_path or default_out_path(input_path, ".md")
        convert_docx_to_md(input_path, out_md, pandoc_extra_args)
        return 0
    if resolved_mode == "md2docx":
        out_docx = output_path or default_out_path(input_path, ".docx")
        convert_md_to_docx(input_path, out_docx, pandoc_extra_args)
        return 0
    raise ValueError(f"Unknown mode: {resolved_mode}")


def legacy_main(argv=None, prog_name=None):
    parser = build_parser(prog_name=prog_name)
    parsed_argv = sys.argv[1:] if argv is None else list(argv)
    args, pandoc_extra_args = parser.parse_known_args(normalize_argv(parsed_argv))

    try:
        return run_conversion(args.mode, args.input, args.output, pandoc_extra_args)
    except subprocess.CalledProcessError as exc:
        print(f"pandoc failed (exit {exc.returncode}): {' '.join(exc.cmd)}", file=sys.stderr)
        return 2
    except ValueError as exc:
        parser.error(str(exc))
    except Exception as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1


def main():
    return legacy_main()


if __name__ == "__main__":
    sys.exit(main())
