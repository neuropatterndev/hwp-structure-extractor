#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hwp_full_parser_v28.py

HWP(.hwp) structured parser + built-in verification web UI.
Version: v16 rollback-no-conversion, exact binary picture-id mapping and GSO grouping build.

Purpose
-------
This file is designed for document-AI verification workflows where you need to
check whether text, tables, images, and captions were extracted correctly from
HWP files.

It intentionally does NOT depend on LibreOffice. The left panel of the web UI
shows the HWP internal PrvImage first-page preview when available. It no longer
shows duplicated PrvText or parser-reconstructed blocks. Exact full-page HWP
rendering still requires an HWP rendering engine.

Main extraction paths
---------------------
1. pyhwp XML path, preferred when hwp5proc is available
   - Runs: hwp5proc xml --embedbin --output intermediate.xml input.hwp
   - Parses the XML generically and extracts paragraphs, tables, images, captions.

2. OLE/Binary fallback path
   - Reads HWP OLE/CFB streams directly with olefile.
   - Extracts BodyText/Section* paragraph text records.
   - Extracts BinData/* image/binary streams robustly.
   - Emits best-effort table/image placeholders from HWP records.

3. Web verification UI
   - Upload new HWP files repeatedly.
   - Left: HWP internal first-page preview from PrvImage when available.
   - Right: block-by-block extraction visualization.
   - Tables rendered as HTML tables.
   - Images rendered as images.
   - Captions shown with method, position, and safe attachment diagnostics.
   - Unlinked media gallery shown so image extraction failures are visible.

Recommended install
-------------------
    pip install olefile lxml pyhwp

Usage
-----
    python hwp_full_parser_v12.py --web
    python hwp_full_parser_v12.py sample.hwp --web
    python hwp_full_parser_v12.py sample.hwp -o parsed --print-summary

Notes on completeness
---------------------
HWP is a complex binary document format. This script is a practical extraction
and verification parser, not a full HWP layout renderer. It is intentionally
built to preserve raw attributes and diagnostics so that project-specific HWP
variants can be extended without losing evidence.
"""

from __future__ import annotations

import argparse
import base64
import binascii
import hashlib
import html
import json
import mimetypes
import os
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import threading
import uuid
import webbrowser
import zlib
from dataclasses import dataclass, field, asdict
from email import policy as email_policy
from email.parser import BytesParser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Tuple, Union
from urllib.parse import parse_qs, quote, unquote, urlparse

try:
    import olefile  # type: ignore
except Exception:  # pragma: no cover
    olefile = None

try:
    from lxml import etree as LET  # type: ignore
except Exception:  # pragma: no cover
    LET = None
    import xml.etree.ElementTree as ET  # type: ignore


# =============================================================================
# HWP record constants
# =============================================================================

HWPTAG_DOCUMENT_PROPERTIES = 16
HWPTAG_ID_MAPPINGS = 17
HWPTAG_BIN_DATA = 18
HWPTAG_FACE_NAME = 19
HWPTAG_BORDER_FILL = 21
HWPTAG_CHAR_SHAPE = 22
HWPTAG_TAB_DEF = 23
HWPTAG_NUMBERING = 24
HWPTAG_BULLET = 25
HWPTAG_PARA_SHAPE = 26
HWPTAG_STYLE = 27
HWPTAG_MEMO_SHAPE = 28
HWPTAG_TRACK_CHANGE = 29
HWPTAG_TRACK_CHANGE_AUTHOR = 30

HWPTAG_PARA_HEADER = 66
HWPTAG_PARA_TEXT = 67
HWPTAG_PARA_CHAR_SHAPE = 68
HWPTAG_PARA_LINE_SEG = 69
HWPTAG_PARA_RANGE_TAG = 70
HWPTAG_CTRL_HEADER = 71
HWPTAG_LIST_HEADER = 72
HWPTAG_PAGE_DEF = 73
HWPTAG_FOOTNOTE_SHAPE = 74
HWPTAG_PAGE_BORDER_FILL = 75
HWPTAG_SHAPE_COMPONENT = 76
HWPTAG_TABLE = 77
HWPTAG_SHAPE_COMPONENT_LINE = 78
HWPTAG_SHAPE_COMPONENT_RECTANGLE = 79
HWPTAG_SHAPE_COMPONENT_ELLIPSE = 80
HWPTAG_SHAPE_COMPONENT_ARC = 81
HWPTAG_SHAPE_COMPONENT_POLYGON = 82
HWPTAG_SHAPE_COMPONENT_CURVE = 83
HWPTAG_SHAPE_COMPONENT_OLE = 84
HWPTAG_SHAPE_COMPONENT_PICTURE = 85
HWPTAG_SHAPE_COMPONENT_CONTAINER = 86
HWPTAG_CTRL_DATA = 87
HWPTAG_EQEDIT = 88
HWPTAG_SHAPE_COMPONENT_TEXTART = 89
HWPTAG_FORM_OBJECT = 90
HWPTAG_MEMO_LIST = 91
HWPTAG_CHART_DATA = 92
HWPTAG_VIDEO_DATA = 93

TAG_NAMES = {
    HWPTAG_DOCUMENT_PROPERTIES: "DOCUMENT_PROPERTIES",
    HWPTAG_ID_MAPPINGS: "ID_MAPPINGS",
    HWPTAG_BIN_DATA: "BIN_DATA",
    HWPTAG_FACE_NAME: "FACE_NAME",
    HWPTAG_BORDER_FILL: "BORDER_FILL",
    HWPTAG_CHAR_SHAPE: "CHAR_SHAPE",
    HWPTAG_TAB_DEF: "TAB_DEF",
    HWPTAG_NUMBERING: "NUMBERING",
    HWPTAG_BULLET: "BULLET",
    HWPTAG_PARA_SHAPE: "PARA_SHAPE",
    HWPTAG_STYLE: "STYLE",
    HWPTAG_PARA_HEADER: "PARA_HEADER",
    HWPTAG_PARA_TEXT: "PARA_TEXT",
    HWPTAG_PARA_CHAR_SHAPE: "PARA_CHAR_SHAPE",
    HWPTAG_PARA_LINE_SEG: "PARA_LINE_SEG",
    HWPTAG_PARA_RANGE_TAG: "PARA_RANGE_TAG",
    HWPTAG_CTRL_HEADER: "CTRL_HEADER",
    HWPTAG_LIST_HEADER: "LIST_HEADER",
    HWPTAG_PAGE_DEF: "PAGE_DEF",
    HWPTAG_FOOTNOTE_SHAPE: "FOOTNOTE_SHAPE",
    HWPTAG_PAGE_BORDER_FILL: "PAGE_BORDER_FILL",
    HWPTAG_SHAPE_COMPONENT: "SHAPE_COMPONENT",
    HWPTAG_TABLE: "TABLE",
    HWPTAG_SHAPE_COMPONENT_PICTURE: "SHAPE_COMPONENT_PICTURE",
    HWPTAG_CTRL_DATA: "CTRL_DATA",
    HWPTAG_EQEDIT: "EQEDIT",
    HWPTAG_CHART_DATA: "CHART_DATA",
    HWPTAG_VIDEO_DATA: "VIDEO_DATA",
}

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".wmf", ".emf", ".svg", ".webp"}

# XML structural heuristics. The parser is permissive because pyhwp/HWPML-like
# XML may use different naming conventions.
TABLE_ROOT_TAGS = {
    "table", "tbl", "tablecontrol", "tablectrl", "hwp-table", "hwp-table-control",
}
TABLE_BODY_TAGS = {
    "tablebody", "tbody", "table-body", "table-bodytype", "tablecontent", "tablecontents",
    "rowlist", "celllist", "tablecelllist", "cells", "rows",
}
ROW_TAGS = {"tr", "row", "tablerow", "table-row", "rowcontrol", "tablerowcontrol"}
CELL_TAGS = {"tc", "cell", "tablecell", "table-cell", "cellcontrol", "tablecellcontrol", "cellzone"}
PARAGRAPH_TAGS = {"p", "para", "paragraph", "hwp-paragraph"}
TEXT_TAGS = {"t", "text", "paratext", "para-text", "char", "run"}
IMAGE_TAGS = {
    "pic", "picture", "image", "img", "shapecomponentpicture", "picturecontrol",
    "imagecontrol", "bindatapicture", "hwp-picture",
}
CAPTION_TAGS = {"caption", "cap", "tablecaption", "imagecaption", "picturecaption", "figcaption"}
# v5: HWP built-in captions are often emitted by pyhwp/HWPML-like XML as
# standalone control/list nodes rather than as direct descendants of the table
# or picture object. Keep these broad on purpose. False positives are filtered
# in the structural resolver by proximity to table/image/equation blocks.
CAPTION_CONTROL_TAGS = {
    "caption", "cap", "captioncontrol", "captionctrl", "capcontrol",
    "tablecaption", "imagecaption", "picturecaption", "figcaption",
    "captionlist", "caption-list", "captionbody", "caption-body",
    "captionparagraph", "caption-para", "captionpara", "caplist",
}
CAPTION_ATTR_KEYS = {
    "caption", "captionpos", "caption-pos", "captionposition", "caption-position",
    "captiontype", "caption-type", "captiondirection", "caption-direction",
    "captionwidth", "caption-width",
}

CAPTION_PATTERN = re.compile(
    r"^\s*(?:\[?\(?\s*)?"
    r"(?P<label>(표|그림|사진|도|수식|Table|TABLE|Fig\.?|FIG\.?|Figure|FIGURE|Equation|EQUATION)"
    r"\s*[-–—]??\s*(?P<num>[0-9０-９]+|[IVXivx]+|[A-Za-z가-힣]+)?\s*"
    r"(?:[.)．:：]|[-–—])?)\s*(?P<body>.*)$"
)

BASE64_CHARS = set("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=\r\n\t ")


# =============================================================================
# Data model
# =============================================================================

@dataclass
class Caption:
    text: str
    method: str = "unknown"
    position: Optional[str] = None
    raw: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return _drop_empty(asdict(self))


@dataclass
class TableCell:
    text: str = ""
    row: Optional[int] = None
    col: Optional[int] = None
    row_span: int = 1
    col_span: int = 1
    attrs: Dict[str, Any] = field(default_factory=dict)
    paragraphs: List[str] = field(default_factory=list)
    content: List[Dict[str, Any]] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return _drop_empty(asdict(self))


@dataclass
class Block:
    type: str
    id: str
    order: int
    section: Optional[int] = None
    text: Optional[str] = None
    caption: Optional[Caption] = None
    rows: Optional[List[List[TableCell]]] = None
    image_path: Optional[str] = None
    media_type: Optional[str] = None
    geometry: Dict[str, Any] = field(default_factory=dict)
    attrs: Dict[str, Any] = field(default_factory=dict)
    raw: Dict[str, Any] = field(default_factory=dict)
    children: List["Block"] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        result: Dict[str, Any] = {"type": self.type, "id": self.id, "order": self.order}
        if self.section is not None:
            result["section"] = self.section
        if self.text not in (None, ""):
            result["text"] = self.text
        if self.caption is not None:
            result["caption"] = self.caption.to_dict()
        if self.rows is not None:
            result["rows"] = [[cell.to_dict() for cell in row] for row in self.rows]
        if self.image_path:
            result["image_path"] = self.image_path
        if self.media_type:
            result["media_type"] = self.media_type
        if self.geometry:
            result["geometry"] = self.geometry
        if self.attrs:
            result["attrs"] = self.attrs
        if self.raw:
            result["raw"] = self.raw
        if self.children:
            result["children"] = [c.to_dict() for c in self.children]
        return result


@dataclass
class MediaItem:
    path: str
    name: str
    media_type: str
    source: str = "unknown"
    refs: List[str] = field(default_factory=list)
    sha1: str = ""
    linked: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return _drop_empty(asdict(self))


@dataclass
class ParsedDocument:
    source_path: str
    method: str
    blocks: List[Block]
    media_items: List[MediaItem] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    @property
    def media_files(self) -> List[str]:
        return [m.path for m in self.media_items]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "source_path": self.source_path,
            "method": self.method,
            "metadata": self.metadata,
            "warnings": self.warnings,
            "errors": self.errors,
            "media_files": [m.to_dict() for m in self.media_items],
            "blocks": [b.to_dict() for b in self.blocks],
        }


@dataclass
class HwpRecord:
    tag_id: int
    level: int
    size: int
    payload: bytes
    offset: int
    section: Optional[int] = None

    @property
    def tag_name(self) -> str:
        return TAG_NAMES.get(self.tag_id, f"TAG_{self.tag_id}")

    def to_meta(self) -> Dict[str, Any]:
        return {
            "tag_id": self.tag_id,
            "tag_name": self.tag_name,
            "level": self.level,
            "size": self.size,
            "offset": self.offset,
            "section": self.section,
        }


# =============================================================================
# Utilities
# =============================================================================

def _drop_empty(d: Dict[str, Any]) -> Dict[str, Any]:
    return {k: v for k, v in d.items() if v not in (None, {}, [], "")}


def unique_id(prefix: str) -> str:
    return f"{prefix}-{uuid.uuid4().hex[:10]}"


def ensure_dir(path: Union[str, Path]) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def local_name(tag: Any) -> str:
    s = str(tag)
    if "}" in s:
        s = s.rsplit("}", 1)[-1]
    if ":" in s:
        s = s.rsplit(":", 1)[-1]
    return s


def norm_name(tag: Any) -> str:
    return local_name(tag).strip().lower().replace("_", "-")


def normalize_space(text: str) -> str:
    text = text.replace("\x00", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def safe_int(value: Any, default: Optional[int] = 0) -> Optional[int]:
    if value is None:
        return default
    try:
        if isinstance(value, str):
            value = value.strip()
            if not value:
                return default
            if value.lower().startswith("0x"):
                return int(value, 16)
            if re.fullmatch(r"[0-9a-fA-F]+", value) and any(ch in "abcdefABCDEF" for ch in value):
                return int(value, 16)
        return int(value)
    except Exception:
        return default


def attrs_to_dict(elem: Any) -> Dict[str, Any]:
    try:
        return {local_name(k): v for k, v in dict(elem.attrib).items()}
    except Exception:
        return {}


def is_probable_base64(text: str) -> bool:
    stripped = "".join(str(text).split())
    if len(stripped) < 80 or len(stripped) % 4 != 0:
        return False
    return all(ch in BASE64_CHARS for ch in stripped[:4096])


def decode_base64_maybe(text: str) -> Optional[bytes]:
    if not is_probable_base64(text):
        return None
    compact = "".join(text.split())
    try:
        return base64.b64decode(compact, validate=False)
    except (binascii.Error, ValueError):
        return None


def detect_extension(data: bytes, fallback: str = ".bin") -> Tuple[str, str]:
    head = data[:256]
    if head.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png", "image/png"
    if head.startswith(b"\xff\xd8\xff"):
        return ".jpg", "image/jpeg"
    if head.startswith(b"GIF87a") or head.startswith(b"GIF89a"):
        return ".gif", "image/gif"
    if head.startswith(b"BM"):
        return ".bmp", "image/bmp"
    if head.startswith(b"II*\x00") or head.startswith(b"MM\x00*"):
        return ".tif", "image/tiff"
    if head.startswith(b"RIFF") and b"WEBP" in head[:16]:
        return ".webp", "image/webp"
    if head.startswith(b"<svg") or b"<svg" in head[:128].lower():
        return ".svg", "image/svg+xml"
    # EMF usually contains this magic at offset 40, but accept early occurrence.
    if b"\xd7\xcd\xc6\x9a" in head[:160]:
        return ".emf", "image/emf"
    # WMF placeable/metafile starts vary; these are best-effort.
    if head.startswith(b"\x01\x00\x09\x00") or head.startswith(b"\xd7\xcd\xc6\x9a"):
        return ".wmf", "image/wmf"
    if head.startswith(b"%PDF"):
        return ".pdf", "application/pdf"
    if head.startswith(b"PK\x03\x04"):
        return ".zip", "application/zip"
    return fallback, "application/octet-stream"


def sha1_bytes(data: bytes) -> str:
    return hashlib.sha1(data).hexdigest()


def write_unique_file(directory: Path, filename: str, data: bytes) -> Path:
    directory.mkdir(parents=True, exist_ok=True)
    filename = re.sub(r"[^0-9A-Za-z가-힣._()\- ]+", "_", filename).strip() or f"file_{uuid.uuid4().hex[:8]}.bin"
    target = directory / filename
    if not target.exists():
        target.write_bytes(data)
        return target
    stem, suffix = target.stem, target.suffix
    for i in range(1, 10000):
        candidate = directory / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            candidate.write_bytes(data)
            return candidate
    raise RuntimeError("could not allocate a unique filename")


def choose_filename(name: str, data: bytes, default_prefix: str = "BIN") -> str:
    base = Path(name).name.strip() if name else ""
    magic_ext, _ = detect_extension(data)
    if base and Path(base).suffix.lower() in IMAGE_EXTENSIONS | {".bin", ".pdf", ".zip"}:
        suffix = Path(base).suffix.lower()
        # If stream has .bin but data is clearly image, use detected extension.
        if suffix == ".bin" and magic_ext != ".bin":
            return Path(base).stem + magic_ext
        return base
    stem = Path(base).stem if base else f"{default_prefix}_{uuid.uuid4().hex[:8]}"
    return stem + magic_ext


def maybe_decompress_variants(data: bytes) -> List[Tuple[str, bytes]]:
    variants: List[Tuple[str, bytes]] = [("raw", data)]
    seen = {sha1_bytes(data)}
    for label, wbits in [("deflate_raw", -15), ("zlib", zlib.MAX_WBITS), ("gzip", 16 + zlib.MAX_WBITS)]:
        try:
            out = zlib.decompress(data, wbits)
            h = sha1_bytes(out)
            if out and h not in seen:
                variants.append((label, out))
                seen.add(h)
        except Exception:
            pass
    return variants


def carve_image_payload(data: bytes) -> Optional[bytes]:
    """Find a likely image payload inside a larger BinData stream.

    Some HWP BinData streams contain wrapper bytes or compressed payloads. This
    function searches for common image signatures and trims obvious tails for
    PNG/JPEG/GIF where an end marker is available.
    """
    signatures = [
        (b"\x89PNG\r\n\x1a\n", "png"),
        (b"\xff\xd8\xff", "jpg"),
        (b"GIF87a", "gif"),
        (b"GIF89a", "gif"),
        (b"BM", "bmp"),
        (b"II*\x00", "tif"),
        (b"MM\x00*", "tif"),
        (b"RIFF", "riff"),
    ]
    best: Optional[Tuple[int, str, bytes]] = None
    for sig, kind in signatures:
        idx = data.find(sig)
        if idx >= 0:
            best = (idx, kind, sig)
            break
    if best is None:
        # EMF magic can occur after EMR header fields.
        idx = data.find(b"\xd7\xcd\xc6\x9a")
        if idx >= 0:
            # EMF usually begins 40 bytes before magic, but avoid negative.
            start = max(0, idx - 40)
            return data[start:]
        return None

    idx, kind, _ = best
    payload = data[idx:]
    if kind == "png":
        end = payload.find(b"IEND")
        if end >= 0 and end + 8 <= len(payload):
            return payload[:end + 8]
    if kind == "jpg":
        end = payload.rfind(b"\xff\xd9")
        if end >= 0:
            return payload[:end + 2]
    if kind == "gif":
        end = payload.rfind(b"\x3b")
        if end >= 0:
            return payload[:end + 1]
    return payload


def strip_caption_marker_prefix(line: str) -> Tuple[str, bool]:
    """Remove parser/UI caption markers such as 'caption 그림 ...'.

    pyhwp sometimes serializes Hancom built-in captions as text beginning with
    a literal 'caption' token. The semantic caption starts after that token.
    """
    s = (line or "").strip()
    changed = False
    m = re.match(r"^(?:caption|캡션)\s*[:：\-–—]?\s*(.+)$", s, flags=re.IGNORECASE)
    if m:
        s = m.group(1).strip()
        changed = True
    return s, changed


def caption_lookup_key(text: str) -> str:
    text = normalize_space(str(text or ""))
    if not text:
        return ""
    first = text.splitlines()[0].strip()
    first, _ = strip_caption_marker_prefix(first)
    return re.sub(r"\s+", " ", first).strip().lower()


def caption_from_text(text: str, object_type: Optional[str] = None) -> Optional[Caption]:
    """Extract a typed caption from text, preserving soft-line-broken captions.

    HWP captions created with Shift+Enter can be serialized as one paragraph
    containing multiple lines. Older versions used only the first physical line,
    which truncated captions such as "그림. 트랜스포머-오토인코더 기반 이상탐지 모델의\n서버 운영 ...".

    This function now recognizes the caption label from the first line, then
    joins subsequent non-empty lines as part of the same caption. The join is
    deliberately bounded to avoid swallowing ordinary body paragraphs when a
    malformed document starts a long paragraph with a caption-looking token.
    """
    text = normalize_space(text)
    if not text:
        return None

    raw_lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    if not raw_lines:
        return None

    first_line, stripped_marker = strip_caption_marker_prefix(raw_lines[0])
    if len(first_line) > 220:
        return None
    m = CAPTION_PATTERN.match(first_line)
    if not m:
        return None

    label = (m.group("label") or "").lower()
    if object_type == "table" and not (label.startswith("표") or label.startswith("table")):
        return None
    if object_type == "image" and not (
        label.startswith("그림") or label.startswith("사진") or label.startswith("도") or label.startswith("fig") or label.startswith("figure")
    ):
        return None
    if object_type == "equation" and not (label.startswith("수식") or label.startswith("equation")):
        return None

    caption_lines = [first_line]
    # Preserve Shift+Enter / soft line break continuations. Stop if the next
    # line itself starts a different caption or a clearly new numbered section.
    for extra in raw_lines[1:]:
        extra_stripped, extra_marker = strip_caption_marker_prefix(extra)
        if not extra_stripped:
            continue
        # Another caption starts; do not merge independent captions.
        if CAPTION_PATTERN.match(extra_stripped):
            break
        # A new major numbered/checkbox section is very unlikely to be part of
        # a caption. Korean captions in this corpus are short, so this is safe.
        if re.match(r"^(?:[0-9]+[.)]|[①-⑳]|□|<\s*서식|[가-힣A-Za-z]+\s*[:：])", extra_stripped) and len(" ".join(caption_lines)) > 20:
            break
        # Bound the caption to avoid swallowing body paragraphs.
        candidate = normalize_space(" ".join(caption_lines + [extra_stripped]))
        if len(candidate) > 500:
            break
        caption_lines.append(extra_stripped)

    caption_text = normalize_space(" ".join(caption_lines))
    return Caption(text=caption_text, method="pattern", raw={"label": m.group("label"), "body": m.group("body"), "stripped_caption_marker": stripped_marker, "soft_line_joined_v19": len(caption_lines) > 1})


def infer_caption_target_type(text: str = "", attrs: Optional[Dict[str, Any]] = None, tag: str = "") -> Optional[str]:
    """Infer the object type that a caption is intended to describe.

    HWP built-in captions may be emitted without a visible "표 1"/"그림 1"
    label. In that case pyhwp often preserves hints such as number-category,
    chid, control-name, or caption-related attributes. This function is
    intentionally conservative: it returns None when no signal is present, and
    the resolver then uses proximity.
    """
    hay = " ".join([str(tag or ""), str(text or "")] + [f"{k}={v}" for k, v in (attrs or {}).items()]).lower()
    if re.search(r"(table|tbl|표|number-category\s*=\s*table|numbercategory\s*=\s*table)", hay):
        return "table"
    if re.search(r"(figure|fig\.?|image|img|picture|pic|photo|gso|그림|사진|도|number-category\s*=\s*(figure|image|picture)|numbercategory\s*=\s*(figure|image|picture))", hay):
        return "image"
    if re.search(r"(equation|eqedit|eqed|수식)", hay):
        return "equation"
    if caption_from_text(text, "table"):
        return "table"
    if caption_from_text(text, "image"):
        return "image"
    if caption_from_text(text, "equation"):
        return "equation"
    return None




def xml_path_tuple(value: Any) -> Tuple[int, ...]:
    if value is None:
        return tuple()
    if isinstance(value, tuple):
        return tuple(int(x) for x in value if isinstance(x, int) or str(x).lstrip('-').isdigit())
    if isinstance(value, list):
        return tuple(int(x) for x in value if isinstance(x, int) or str(x).lstrip('-').isdigit())
    if isinstance(value, str):
        nums = re.findall(r"-?\d+", value)
        return tuple(int(x) for x in nums)
    return tuple()


def common_prefix_len(a: Sequence[int], b: Sequence[int]) -> int:
    n = min(len(a), len(b))
    i = 0
    while i < n and a[i] == b[i]:
        i += 1
    return i


def xml_path_gap(a: Sequence[int], b: Sequence[int]) -> int:
    """A small structural distance metric between two XML element paths."""
    lca = common_prefix_len(a, b)
    gap = (len(a) - lca) + (len(b) - lca)
    if lca < min(len(a), len(b)):
        try:
            gap += min(abs(a[lca] - b[lca]), 25)
        except Exception:
            pass
    return gap


def normalize_caption_position(value: Any) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip().lower()
    if not s:
        return None
    if re.search(r"(bottom|below|under|after|down|하단|아래|밑)", s):
        return "after"
    if re.search(r"(top|above|before|up|상단|위)", s):
        return "before"
    if re.search(r"(left|왼쪽|좌)", s):
        return "left"
    if re.search(r"(right|오른쪽|우)", s):
        return "right"
    if s in {"0", "bottom", "1:bottom"}:
        return "after"
    if s in {"1", "top", "0:top"}:
        return "before"
    return None


def caption_position_from_attrs(attrs: Optional[Dict[str, Any]]) -> Optional[str]:
    if not attrs:
        return None
    for key, value in attrs.items():
        lk = str(key).lower()
        if "caption" in lk or lk in {"pos", "position", "placement", "side", "align", "valign"}:
            pos = normalize_caption_position(value)
            if pos:
                return pos
    return None


def extract_object_keys(attrs: Optional[Dict[str, Any]], tag: str = "", text: str = "") -> List[str]:
    """Return normalized object/link keys useful for table/image-caption matching."""
    if not attrs:
        attrs = {}
    key_names = {
        "id", "object-id", "objectid", "instance-id", "instanceid", "instid",
        "table-id", "tableid", "shape-id", "shapeid", "control-id", "controlid",
        "ctrl-id", "ctrlid", "ref", "href", "target", "target-id", "targetid",
        "owner-id", "ownerid", "parent-id", "parentid", "caption-of", "captionof",
        "bindata", "bin-id", "binid", "stream", "stream-id", "streamid",
    }
    values: List[str] = []
    for k, v in attrs.items():
        lk = str(k).strip().lower()
        sv = str(v).strip()
        if not sv:
            continue
        if lk in key_names or any(tok in lk for tok in ("instance", "object", "shape", "control", "target", "owner", "parent", "caption", "table-id", "bin")):
            values.append(f"{lk}={sv}")
            values.append(sv)
    hay = " ".join([tag or "", text or ""] + [f"{k}={v}" for k, v in attrs.items()])
    for m in re.finditer(r"(?:BIN|bin)\s*0*([0-9]+)", hay):
        n = int(m.group(1))
        values.extend([f"BIN{n:04d}", f"BIN{n}", str(n)])
    out: List[str] = []
    seen = set()
    for v in values:
        nv = normalize_space(str(v)).lower()
        if not nv or nv in {"0", "1", "true", "false", "none"}:
            continue
        if nv not in seen:
            seen.add(nv)
            out.append(nv)
    return out[:40]

def first_nonempty_line(text: str, max_len: int = 400) -> str:
    for line in normalize_space(text).splitlines():
        line = line.strip()
        if line:
            return line[:max_len]
    return ""


def natural_bin_sort_key(name: str) -> Tuple[int, str]:
    """Sort BIN0001, BIN000A, embedded_0011 in a stable document-like order."""
    s = str(name or "")
    m = re.search(r"BIN([0-9A-Fa-f]+)", s)
    if m:
        token = m.group(1)
        try:
            return (int(token, 16), s.lower())
        except Exception:
            pass
    m = re.search(r"(\d+)", s)
    if m:
        return (int(m.group(1)), s.lower())
    return (10**9, s.lower())


def decode_hwp_preview_text(data: bytes) -> str:
    for enc in ("utf-16le", "utf-16", "cp949", "utf-8"):
        try:
            text = data.decode(enc, errors="ignore")
            text = normalize_space(text)
            if text:
                return text
        except Exception:
            pass
    return ""


# =============================================================================
# Media store and OLE extraction
# =============================================================================

class MediaStore:
    def __init__(self, media_dir: Union[str, Path]):
        self.media_dir = ensure_dir(media_dir)
        self.items: List[MediaItem] = []
        self.ref_map: Dict[str, str] = {}
        self._sha_seen: Dict[str, str] = {}
        # Explicit picture-control references recovered from BodyText records.
        # Used only as a safer fallback for real picture controls, not for
        # caption-only text.
        self.picture_ref_order: List[str] = []
        self._picture_ref_cursor = 0
        # caption text -> media path hints recovered from the binary BodyText
        # sequence: picture record immediately before a caption paragraph.
        self.caption_image_hints: Dict[str, str] = {}
        self.caption_image_hint_order: List[Tuple[str, str]] = []

    def add_bytes(self, data: bytes, name: str, source: str, refs: Optional[Sequence[str]] = None) -> Optional[MediaItem]:
        if not data or len(data) < 4:
            return None
        # v14: prefer a complete decompressed image stream over a carved inner
        # signature. Several HWP BinData BMP streams contain JPEG-like byte
        # patterns inside the bitmap payload; older versions carved those and
        # produced invalid/wrong images. We first accept whole valid image-like
        # variants, and only carve when no whole variant is recognized.
        variants = maybe_decompress_variants(data)
        best_label, best_data = variants[0]
        best_ext, best_type = detect_extension(best_data)
        whole_hits: List[Tuple[str, bytes, str]] = []
        carved_hits: List[Tuple[str, bytes, str]] = []
        for label, variant in variants:
            ext, mtype = detect_extension(variant)
            if mtype.startswith("image/"):
                whole_hits.append((label, variant, mtype))
                continue
            carved = carve_image_payload(variant)
            if carved and carved != variant:
                cext, cmtype = detect_extension(carved)
                if cmtype.startswith("image/"):
                    carved_hits.append((label + "+carved", carved, cmtype))
        if whole_hits:
            # Prefer non-raw decompressed whole payloads over raw compressed bytes.
            whole_hits.sort(key=lambda x: (0 if x[0] != "raw" else 1, len(x[1])))
            best_label, best_data, best_type = whole_hits[0]
        elif carved_hits:
            carved_hits.sort(key=lambda x: (0 if x[0] != "raw+carved" else 1, len(x[1])))
            best_label, best_data, best_type = carved_hits[0]

        h = sha1_bytes(best_data)
        if h in self._sha_seen:
            existing_path = self._sha_seen[h]
            self._register_refs(existing_path, refs or [], extra_names=[name])
            for item in self.items:
                if item.path == existing_path:
                    return item
            return None

        filename = choose_filename(name, best_data, default_prefix="BIN")
        path = write_unique_file(self.media_dir, filename, best_data)
        item = MediaItem(
            path=str(path),
            name=path.name,
            media_type=best_type,
            source=source + (f":{best_label}" if best_label != "raw" else ""),
            refs=[],
            sha1=h,
        )
        self.items.append(item)
        self._sha_seen[h] = str(path)
        self._register_refs(str(path), refs or [], extra_names=[name, path.name, path.stem])
        item.refs = sorted({r for r, p in self.ref_map.items() if p == str(path) and r})[:50]
        return item

    def _register_refs(self, path: str, refs: Sequence[str], extra_names: Sequence[str] = ()) -> None:
        all_refs = set()
        for ref in list(refs) + list(extra_names):
            if ref is None:
                continue
            s = str(ref).strip()
            if not s:
                continue
            all_refs.add(s)
            all_refs.add(s.lower())
            all_refs.add(s.upper())
            all_refs.add(Path(s).name)
            all_refs.add(Path(s).stem)
            m = re.search(r"BIN\s*0*([0-9]+)", s, re.IGNORECASE)
            if m:
                n = int(m.group(1))
                for width in (1, 2, 3, 4, 5):
                    all_refs.add(f"BIN{n:0{width}d}")
                    all_refs.add(f"BIN{n:0{width}d}".lower())
            # hwp XML may use numeric bindata id.
            if s.isdigit():
                n = int(s)
                for width in (1, 2, 3, 4, 5):
                    all_refs.add(f"BIN{n:0{width}d}")
        for r in all_refs:
            if r:
                self.ref_map[r] = path

    def resolve(self, ref: Optional[str]) -> Optional[str]:
        if not ref:
            return None
        candidates = [ref, ref.strip("#"), Path(ref).name, Path(ref).stem, ref.lower(), ref.upper()]
        m = re.search(r"BIN\s*0*([0-9]+)", ref, re.IGNORECASE)
        if m:
            n = int(m.group(1))
            for width in (1, 2, 3, 4, 5):
                candidates.append(f"BIN{n:0{width}d}")
                candidates.append(f"BIN{n:0{width}d}".lower())
        if str(ref).isdigit():
            n = int(ref)
            for width in (1, 2, 3, 4, 5):
                candidates.append(f"BIN{n:0{width}d}")
        for c in candidates:
            if c in self.ref_map:
                return self.ref_map[c]
        ref_l = Path(ref).name.lower()
        for k, v in self.ref_map.items():
            if ref_l and (Path(k).name.lower() == ref_l or ref_l in Path(k).name.lower()):
                return v
        return None

    def add_caption_image_hint(self, caption_text: str, image_path: str) -> None:
        key = caption_lookup_key(caption_text)
        if key and image_path:
            if key not in self.caption_image_hints:
                self.caption_image_hints[key] = image_path
                self.caption_image_hint_order.append((key, image_path))

    def resolve_caption_image(self, caption_text: str) -> Optional[str]:
        key = caption_lookup_key(caption_text)
        if not key:
            return None
        if key in self.caption_image_hints:
            path = self.caption_image_hints[key]
            # Do not reuse an already linked image in a different location.
            for item in self.items:
                try:
                    same = item.path == path or str(Path(item.path).resolve()) == str(Path(path).resolve())
                except Exception:
                    same = item.path == path
                if same:
                    if item.linked:
                        return None
                    if item.media_type.startswith("image/") and not Path(item.path).name.lower().startswith("prvimage"):
                        item.linked = True
                        return item.path
            return path
        # Fuzzy fallback for pyhwp variants that slightly change caption text
        # between BodyText and XML serialization, e.g. dropping a prefix or
        # normalizing punctuation. This never invents a new image; it only uses
        # hints already recovered from nearby binary picture records.
        for hint_key, path in list(self.caption_image_hint_order):
            if not hint_key or not path:
                continue
            if key == hint_key or key in hint_key or hint_key in key:
                for item in self.items:
                    try:
                        same = item.path == path or str(Path(item.path).resolve()) == str(Path(path).resolve())
                    except Exception:
                        same = item.path == path
                    if same:
                        if item.linked:
                            return None
                        if item.media_type.startswith("image/") and not Path(item.path).name.lower().startswith("prvimage"):
                            item.linked = True
                            return item.path
                return path
        return None

    def add_picture_ref_order(self, refs: Sequence[str]) -> None:
        for ref in refs:
            s = str(ref or "").strip()
            if s and s not in self.picture_ref_order:
                self.picture_ref_order.append(s)

    def next_picture_order_image(self) -> Optional[str]:
        while self._picture_ref_cursor < len(self.picture_ref_order):
            ref = self.picture_ref_order[self._picture_ref_cursor]
            self._picture_ref_cursor += 1
            path = self.resolve(ref)
            if not path:
                continue
            item_obj = None
            for item in self.items:
                try:
                    same = item.path == path or str(Path(item.path).resolve()) == str(Path(path).resolve())
                except Exception:
                    same = item.path == path
                if same:
                    item_obj = item
                    break
            if item_obj is None or item_obj.linked or not item_obj.media_type.startswith("image/"):
                continue
            if Path(item_obj.path).name.lower().startswith("prvimage") or item_obj.source.lower().startswith("prvimage"):
                continue
            item_obj.linked = True
            return item_obj.path
        return None

    def mark_linked(self, path: Optional[str]) -> None:
        if not path:
            return
        for item in self.items:
            if str(Path(item.path).resolve()) == str(Path(path).resolve()) or item.path == path:
                item.linked = True

    def next_unlinked_image(self, include_preview: bool = False, prefer_ole_bindata: bool = True) -> Optional[str]:
        """Return the next unlinked real document image.

        PrvImage is only a first-page preview and must not be consumed as a
        figure/screenshot. v9 therefore skips PrvImage by default. OLE BinData
        items are preferred because they are the actual embedded HWP media;
        pyhwp embedded-base64 duplicates are used only as fallback.
        """
        def usable(item: MediaItem) -> bool:
            if item.linked or not item.media_type.startswith("image/"):
                return False
            name = Path(item.path).name.lower()
            if not include_preview and (name.startswith("prvimage") or item.source.lower().startswith("prvimage")):
                return False
            return True

        candidates = [item for item in self.items if usable(item)]
        if prefer_ole_bindata:
            candidates.sort(key=lambda it: (0 if "OLE:BinData" in it.source else 1, natural_bin_sort_key(it.name)))
        else:
            candidates.sort(key=lambda it: natural_bin_sort_key(it.name))
        for item in candidates:
            item.linked = True
            return item.path
        return None

    def figure_candidate_items(self) -> List[MediaItem]:
        """Return real document image candidates in deterministic BinData order.

        PrvImage is excluded because it is only the HWP first-page preview. OLE
        BinData items are preferred over pyhwp embedded-base64 duplicates. This
        list is used only together with BodyText picture-record order, not as a
        blind caption fallback.
        """
        def usable(item: MediaItem) -> bool:
            if not item.media_type.startswith("image/"):
                return False
            name = Path(item.path).name.lower()
            if name.startswith("prvimage") or item.source.lower().startswith("prvimage"):
                return False
            return True
        items = [item for item in self.items if usable(item)]
        items.sort(key=lambda it: (0 if "OLE:BinData" in it.source else 1, natural_bin_sort_key(it.name)))
        return items

    def resolve_picture_event_image(self, refs: Sequence[str], event_index: int) -> Optional[str]:
        """Resolve one HWP picture record to a media path.

        First use explicit BinData references decoded from the picture record.
        If the record does not expose a usable reference, fall back to the image
        at the same ordinal among real document BinData candidates. This is safer
        than global caption-order fallback because the ordinal comes from actual
        HWPTAG_SHAPE_COMPONENT_PICTURE records in BodyText order.
        """
        for ref in refs or []:
            path = self.resolve(str(ref))
            if not path:
                continue
            for item in self.items:
                try:
                    same = item.path == path or str(Path(item.path).resolve()) == str(Path(path).resolve())
                except Exception:
                    same = item.path == path
                if same and item.media_type.startswith("image/") and not Path(item.path).name.lower().startswith("prvimage"):
                    return item.path
        candidates = self.figure_candidate_items()
        if 0 <= event_index < len(candidates):
            return candidates[event_index].path
        return None


def _hwp_file_is_compressed_from_ole(ole: Any) -> bool:
    try:
        data = ole.openstream("FileHeader").read()
        return len(data) > 36 and bool(data[36] & 0x01)
    except Exception:
        return False


def _iter_hwp_records_from_bytes(data: bytes, section: Optional[int] = None) -> Iterator[HwpRecord]:
    offset = 0
    total = len(data)
    while offset + 4 <= total:
        rec_offset = offset
        try:
            header = struct.unpack_from("<I", data, offset)[0]
        except Exception:
            break
        offset += 4
        tag_id = header & 0x3FF
        level = (header >> 10) & 0x3FF
        size = (header >> 20) & 0xFFF
        if size == 0xFFF:
            if offset + 4 > total:
                break
            size = struct.unpack_from("<I", data, offset)[0]
            offset += 4
        if size < 0 or offset + size > total:
            break
        payload = data[offset:offset + size]
        offset += size
        yield HwpRecord(tag_id, level, size, payload, rec_offset, section=section)


def _read_hwp_section_streams_from_ole(ole: Any) -> List[Tuple[int, bytes]]:
    sections: List[Tuple[int, bytes]] = []
    compressed = _hwp_file_is_compressed_from_ole(ole)
    paths: List[List[str]] = []
    for entry in ole.listdir(streams=True, storages=False):
        if len(entry) == 2 and entry[0] == "BodyText" and entry[1].startswith("Section"):
            paths.append(entry)
    paths.sort(key=lambda p: int(re.search(r"(\d+)$", p[-1]).group(1)) if re.search(r"(\d+)$", p[-1]) else 10**9)
    for idx, path in enumerate(paths):
        raw = ole.openstream(path).read()
        data = raw
        if compressed:
            for wbits in (-15, zlib.MAX_WBITS):
                try:
                    data = zlib.decompress(raw, wbits)
                    break
                except Exception:
                    pass
        sections.append((idx, data))
    return sections


def _bin_refs_from_picture_payload(payload: bytes, media_store: MediaStore) -> List[str]:
    refs: List[str] = []
    text = payload[:512].decode("latin1", errors="ignore")
    for m in re.finditer(r"BIN[0-9A-Fa-f]{1,6}|BinData[/\\][^\x00\s]+", text, re.I):
        refs.append(m.group(0))
    max_id = 0
    for item in media_store.items:
        k = natural_bin_sort_key(item.name)[0]
        if isinstance(k, int) and k < 10**8:
            max_id = max(max_id, k)
    if max_id > 0:
        priority_offsets = [22, 24, 26, 28, 30, 32, 34, 36, 38, 40, 44, 48, 52, 56, 60, 64]
        seen_vals = set()
        for off in priority_offsets + list(range(0, min(len(payload) - 1, 160), 2)):
            if off < 0 or off + 2 > len(payload):
                continue
            v = int.from_bytes(payload[off:off+2], "little", signed=False)
            if 1 <= v <= max_id and v not in seen_vals:
                seen_vals.add(v)
                refs.extend([str(v), f"BIN{v:04X}", f"BIN{v:04d}"])
    resolved: List[str] = []
    for ref in refs:
        if media_store.resolve(ref) and ref not in resolved:
            resolved.append(ref)
    return resolved


def collect_picture_ref_order_from_ole(ole: Any, media_store: MediaStore) -> List[str]:
    refs: List[str] = []
    try:
        for section_idx, data in _read_hwp_section_streams_from_ole(ole):
            for rec in _iter_hwp_records_from_bytes(data, section=section_idx):
                if rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
                    for ref in _bin_refs_from_picture_payload(rec.payload, media_store):
                        if ref not in refs:
                            refs.append(ref)
    except Exception:
        return refs
    return refs


def _decode_hwp_para_text_for_hints(payload: bytes) -> str:
    out: List[str] = []
    i = 0
    n = len(payload)
    while i + 2 <= n:
        code = int.from_bytes(payload[i:i+2], "little")
        i += 2
        if code >= 32:
            try:
                out.append(chr(code))
            except Exception:
                pass
        elif code == 0x0009:
            out.append("\t")
        elif code in {0x000a, 0x000d}:
            out.append("\n")
        else:
            # Skip extended control payloads when they look like 14-byte HWP controls.
            if i + 14 <= n:
                raw = payload[i:i+14]
                if any(all(32 <= b <= 126 for b in cand) for cand in (raw[:4], raw[:4][::-1])):
                    i += 14
    return normalize_space("".join(out))


def collect_caption_image_hints_from_ole(ole: Any, media_store: MediaStore, max_gap_paragraphs: int = 12) -> List[Tuple[str, str]]:
    """Map image captions to preceding picture records in binary BodyText order.

    v12 improves v11 by pairing captions with *picture-record events*, not with
    arbitrary global BinData order. If a picture record's BinData reference is
    not explicitly decodable, the event ordinal is mapped to the same ordinal in
    real document BinData candidates. Stale picture events are dropped when too
    many non-empty paragraphs occur before the caption, preventing cover/title
    images from being attached to later captions.
    """
    hints: List[Tuple[str, str]] = []
    pending: List[Dict[str, Any]] = []
    used_paths: set[str] = set()
    picture_event_index = 0
    try:
        for section_idx, data in _read_hwp_section_streams_from_ole(ole):
            pending.clear()
            for rec in _iter_hwp_records_from_bytes(data, section=section_idx):
                if rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
                    refs = _bin_refs_from_picture_payload(rec.payload, media_store)
                    path = media_store.resolve_picture_event_image(refs, picture_event_index)
                    event = {
                        "event_index": picture_event_index,
                        "path": path,
                        "refs": refs,
                        "gap": 0,
                        "section": section_idx,
                        "offset": rec.offset,
                    }
                    picture_event_index += 1
                    if path and not Path(path).name.lower().startswith("prvimage"):
                        pending.append(event)
                        pending = pending[-16:]
                    continue

                if rec.tag_id != HWPTAG_PARA_TEXT:
                    continue
                text = _decode_hwp_para_text_for_hints(rec.payload)
                if not text:
                    continue

                pending = [dict(ev, gap=int(ev.get("gap", 0)) + 1) for ev in pending if int(ev.get("gap", 0)) + 1 <= max_gap_paragraphs]
                for line in text.splitlines():
                    cap = caption_from_text(line, "image")
                    if not cap or not pending:
                        continue
                    # nearest preceding picture event that still has an unused image
                    chosen_idx = None
                    chosen_path = None
                    for idx in range(len(pending) - 1, -1, -1):
                        path = pending[idx].get("path")
                        if path and path not in used_paths:
                            chosen_idx = idx
                            chosen_path = path
                            break
                    if chosen_idx is None or not chosen_path:
                        continue
                    pending.pop(chosen_idx)
                    used_paths.add(chosen_path)
                    hints.append((cap.text, chosen_path))
                    media_store.add_caption_image_hint(cap.text, chosen_path)
                    break
    except Exception:
        return hints
    return hints



def extract_ole_media(input_path: Union[str, Path], media_store: MediaStore, warnings: List[str], errors: List[str]) -> Dict[str, Any]:
    preview: Dict[str, Any] = {"preview_image": None, "preview_text": None, "notes": []}
    if olefile is None:
        warnings.append("olefile is not installed; OLE BinData extraction skipped. Install with: pip install olefile")
        return preview
    input_path = Path(input_path)
    if not olefile.isOleFile(str(input_path)):
        warnings.append("Input is not an OLE/CFB HWP file; OLE media extraction skipped.")
        return preview
    try:
        with olefile.OleFileIO(str(input_path)) as ole:
            entries = ole.listdir(streams=True, storages=False)
            for entry in entries:
                try:
                    if entry == ["PrvImage"]:
                        data = ole.openstream(entry).read()
                        item = media_store.add_bytes(data, "PrvImage", source="PrvImage", refs=["PrvImage"])
                        if item:
                            preview["preview_image"] = item.path
                    elif entry == ["PrvText"]:
                        data = ole.openstream(entry).read()
                        text = decode_hwp_preview_text(data)
                        if text:
                            preview["preview_text"] = text
                    elif entry and entry[0] == "BinData":
                        data = ole.openstream(entry).read()
                        stream_name = entry[-1]
                        refs = [stream_name, "/".join(entry), Path(stream_name).stem]
                        # Also register BIN number from stream name.
                        media_store.add_bytes(data, stream_name, source="OLE:BinData", refs=refs)
                except Exception as e:
                    errors.append(f"Failed to extract OLE stream {'/'.join(entry)}: {e}")
            try:
                picture_refs = collect_picture_ref_order_from_ole(ole, media_store)
                if picture_refs:
                    media_store.add_picture_ref_order(picture_refs)
                    preview["picture_ref_order"] = picture_refs
                    preview.setdefault("notes", []).append(f"Collected {len(picture_refs)} picture-control media references from BodyText records.")
                caption_hints = collect_caption_image_hints_from_ole(ole, media_store)
                if caption_hints:
                    preview["caption_image_hints"] = [{"caption": c, "image_path": p} for c, p in caption_hints]
                    preview.setdefault("notes", []).append(f"Recovered {len(caption_hints)} caption-to-image hints from BodyText record order.")
            except Exception as e:
                warnings.append(f"Failed to collect picture-control references/caption-image hints: {e}")
    except Exception as e:
        errors.append(f"OLE media extraction failed: {e}")
    return preview


# =============================================================================
# PyHWP XML parser
# =============================================================================

class PyHwpXmlParser:
    def __init__(self, xml_path: Union[str, Path], output_dir: Union[str, Path], media_store: Optional[MediaStore] = None):
        self.xml_path = Path(xml_path)
        self.output_dir = ensure_dir(output_dir)
        self.media_store = media_store or MediaStore(self.output_dir / "media")
        self.blocks: List[Block] = []
        self.warnings: List[str] = []
        self.errors: List[str] = []
        self._order = 0
        # v11: map pyhwp --embedbin base64 XML nodes back to extracted media.
        # This preserves the real DOM position of images inside table cells
        # instead of guessing by global BinData order.
        self._embedded_image_by_node_id: Dict[int, str] = {}

    def parse(self) -> ParsedDocument:
        root = self._read_xml_root()
        self._extract_embedded_binaries(root)
        self._walk(root, context={"section": None, "inside_para": False, "inside_table": False, "xml_path": ()})
        self._postprocess_blocks()
        return ParsedDocument(
            source_path=str(self.xml_path),
            method="pyhwp_xml",
            blocks=self.blocks,
            media_items=self.media_store.items,
            metadata={"xml_path": str(self.xml_path), "media_count": len(self.media_store.items)},
            warnings=self.warnings,
            errors=self.errors,
        )

    def _read_xml_root(self) -> Any:
        if LET is not None:
            parser = LET.XMLParser(recover=True, huge_tree=True, remove_blank_text=False)
            return LET.parse(str(self.xml_path), parser).getroot()
        return ET.parse(str(self.xml_path)).getroot()  # type: ignore[name-defined]

    def _next_order(self) -> int:
        self._order += 1
        return self._order

    def _children(self, elem: Any) -> List[Any]:
        try:
            return list(elem)
        except Exception:
            return []

    def _descendants(self, elem: Any) -> Iterator[Any]:
        for child in self._children(elem):
            yield child
            yield from self._descendants(child)

    def _is_table_body(self, elem: Any) -> bool:
        return norm_name(getattr(elem, "tag", "")) in TABLE_BODY_TAGS

    def _is_table(self, elem: Any) -> bool:
        name = norm_name(getattr(elem, "tag", ""))
        attrs = {k.lower(): str(v).lower() for k, v in attrs_to_dict(elem).items()}
        values = " ".join(attrs.values())
        # v5: built-in caption controls may contain number-category=table. Do
        # not misclassify such caption controls as tables.
        if name in CAPTION_CONTROL_TAGS or attrs.get("chid", "").strip().lower() in {"cap", "cap "}:
            return False
        if "caption" in name or re.search(r"\bcaption\b|\bcap\b", values):
            if not (name in TABLE_ROOT_TAGS or attrs.get("chid", "").strip().lower() == "tbl"):
                return False
        if name in TABLE_BODY_TAGS:
            return False
        if name in TABLE_ROOT_TAGS:
            return True
        if name in {"control", "ctrl", "object", "shape", "component"} and re.search(r"\b(tbl|table)\b", values):
            return True
        if attrs.get("chid", "").strip().lower() == "tbl":
            return True
        return False

    def _is_row(self, elem: Any) -> bool:
        name = norm_name(getattr(elem, "tag", ""))
        return name in ROW_TAGS or ("row" in name and "arrow" not in name and "border" not in name)

    def _is_cell(self, elem: Any) -> bool:
        name = norm_name(getattr(elem, "tag", ""))
        return name in CELL_TAGS or ("cell" in name and "border" not in name and "spacing" not in name)

    def _is_para(self, elem: Any) -> bool:
        name = norm_name(getattr(elem, "tag", ""))
        if name in PARAGRAPH_TAGS:
            return True
        # Do not treat every generic Text as a top-level paragraph. It causes duplicates.
        return name in {"paratext", "para-text"}

    def _is_image(self, elem: Any) -> bool:
        name = norm_name(getattr(elem, "tag", ""))
        if name in IMAGE_TAGS:
            return True
        attrs = {k.lower(): str(v).lower() for k, v in attrs_to_dict(elem).items()}
        values = " ".join(attrs.values())
        chid = attrs.get("chid", "").strip().lower().replace("$", "")
        number_category = attrs.get("number-category", attrs.get("numbercategory", "")).strip().lower()
        # pyhwp/HWPML variants may emit picture controls as generic Control,
        # ShapeComponent, GShapeObjectControl, or as a node with chid="$pic".
        if re.search(r"(pic|picture|image|img|photo|gso|shapecomponentpicture)", name):
            return True
        if name in {"control", "ctrl", "object", "shape", "component", "gshapeobjectcontrol", "gshapeobject", "shapecomponent"} and re.search(r"\b(pic|picture|image|img|photo)\b", values):
            return True
        if chid and ("pic" in chid or chid in {"gso", "pict", "picture"}):
            return True
        if number_category in {"figure", "image", "picture"}:
            return True
        return False

    def _is_caption_node(self, elem: Any) -> bool:
        name = norm_name(getattr(elem, "tag", ""))
        if name in CAPTION_TAGS or name in CAPTION_CONTROL_TAGS or "caption" in name:
            return True
        attrs = {k.lower(): str(v).lower() for k, v in attrs_to_dict(elem).items()}
        if any(k in CAPTION_ATTR_KEYS or "caption" in k for k in attrs):
            return True
        if attrs.get("chid", "").strip().lower() in {"cap", "cap "}:
            return True
        return any("caption" in v or re.fullmatch(r"cap\s*", v) for v in attrs.values())

    def _is_caption_control(self, elem: Any) -> bool:
        """Return True for standalone HWP caption controls/lists.

        This is stricter than _is_caption_node because it is used during tree
        walking to create a separate caption block. The aim is to catch Hancom's
        'insert caption' structures without misclassifying ordinary paragraphs.
        """
        name = norm_name(getattr(elem, "tag", ""))
        attrs = {k.lower(): str(v).lower() for k, v in attrs_to_dict(elem).items()}
        values = " ".join(attrs.values())
        if name in CAPTION_CONTROL_TAGS:
            return True
        if "caption" in name and name not in TABLE_ROOT_TAGS:
            return True
        if attrs.get("chid", "").strip().lower() in {"cap", "cap "}:
            return True
        if ("list" in name or "header" in name) and attrs.get("number-category", attrs.get("numbercategory", "")).strip().lower() in {"table", "figure", "image", "picture", "equation"}:
            return True
        if any(k in CAPTION_ATTR_KEYS or "caption" in k for k in attrs):
            # Avoid classifying TableControl/PictureControl themselves as
            # captions when they merely carry caption-related geometry.
            if not (self._is_table(elem) or self._is_image(elem)):
                return True
        if re.search(r"\bcaption\b|\bcap\b", values) and not (self._is_table(elem) or self._is_image(elem)):
            return True
        return False

    def _text_of(self, elem: Any, exclude_objects: bool = False) -> str:
        parts: List[str] = []

        def rec(node: Any) -> None:
            if exclude_objects and node is not elem and (self._is_table(node) or self._is_image(node) or self._is_caption_control(node)):
                return
            text = getattr(node, "text", None)
            if text and not is_probable_base64(text):
                parts.append(str(text))
            for child in self._children(node):
                rec(child)
                tail = getattr(child, "tail", None)
                if tail and not is_probable_base64(tail):
                    parts.append(str(tail))
            if self._is_para(node):
                parts.append("\n")

        rec(elem)
        return normalize_space("".join(parts))

    def _xml_meta(self, elem: Any, context: Dict[str, Any]) -> Dict[str, Any]:
        attrs = attrs_to_dict(elem)
        tag = local_name(getattr(elem, "tag", ""))
        path = xml_path_tuple(context.get("xml_path", ()))
        parent_path = path[:-1]
        keys = extract_object_keys(attrs, tag=tag, text="")
        return {
            "xml_tag": tag,
            "xml_path": list(path),
            "xml_parent_path": list(parent_path),
            "xml_depth": len(path),
            "xml_sibling_index": path[-1] if path else 0,
            "object_keys": keys,
        }

    def _walk(self, elem: Any, context: Dict[str, Any]) -> None:
        path = xml_path_tuple(context.get("xml_path", ()))
        if self._is_table(elem):
            self.blocks.append(self._parse_table(elem, context.get("section"), context))
            return
        if self._is_image(elem):
            self.blocks.append(self._parse_image(elem, context.get("section"), context))
            return
        if self._is_caption_control(elem):
            cap_block = self._parse_caption_control(elem, context.get("section"), context)
            if cap_block is not None:
                self.blocks.append(cap_block)
            return

        section = context.get("section")
        name = norm_name(getattr(elem, "tag", ""))
        attrs = attrs_to_dict(elem)
        if "section" in name or "sect" in name:
            m = re.search(r"(\d+)$", name)
            if m:
                section = int(m.group(1))
            else:
                for key in ("id", "index", "number", "section"):
                    if key in attrs:
                        section = safe_int(attrs[key], section)
                        break

        is_para = self._is_para(elem)
        if is_para and not context.get("inside_para") and not context.get("inside_table"):
            text = self._text_of(elem, exclude_objects=True)
            if self._valid_text_block(text):
                para_raw = self._xml_meta(elem, context)
                self.blocks.append(Block("paragraph", unique_id("p"), self._next_order(), section=section, text=text, attrs=attrs, raw=para_raw))

        child_context = dict(context)
        child_context["section"] = section
        if is_para:
            child_context["inside_para"] = True
        for child_index, child in enumerate(self._children(elem)):
            child_context["xml_path"] = path + (child_index,)
            self._walk(child, child_context)

    def _valid_text_block(self, text: str) -> bool:
        if not text:
            return False
        if is_probable_base64(text):
            return False
        if len(text) < 2 and not re.search(r"[가-힣A-Za-z0-9]", text):
            return False
        if text.count("�") > 5:
            return False
        return True

    def _parse_caption_control(self, elem: Any, section: Optional[int], context: Optional[Dict[str, Any]] = None) -> Optional[Block]:
        context = context or {}
        attrs = attrs_to_dict(elem)
        xml_meta = self._xml_meta(elem, context)
        # Do not exclude nested caption body/list nodes here; they are where
        # pyhwp commonly places the real caption text.
        text = self._text_of(elem, exclude_objects=False)
        text = first_nonempty_line(text)
        if not text:
            # Some XML variants put caption text into attributes.
            for k, v in attrs.items():
                if "caption" in str(k).lower() and str(v).strip() and str(v).lower() not in {"true", "false", "1", "0"}:
                    text = first_nonempty_line(str(v))
                    break
        if not self._valid_text_block(text):
            return None
        target_type = infer_caption_target_type(text, attrs, local_name(getattr(elem, "tag", "")))
        pos_hint = caption_position_from_attrs(attrs)
        cap = Caption(
            text=text,
            method="structural-control",
            position=pos_hint,
            raw={
                **xml_meta,
                "target_type_hint": target_type,
                "attrs": attrs,
            },
        )
        return Block(
            "caption",
            unique_id("cap"),
            self._next_order(),
            section=section,
            text=text,
            caption=cap,
            attrs=attrs,
            raw={**xml_meta, "target_type_hint": target_type, "caption_position_hint": pos_hint},
        )

    def _parse_table(self, elem: Any, section: Optional[int], context: Optional[Dict[str, Any]] = None) -> Block:
        context = context or {}
        attrs = attrs_to_dict(elem)
        xml_meta = self._xml_meta(elem, context)
        rows: List[List[TableCell]] = []
        row_elems = self._find_rows(elem)
        for r_idx, row_elem in enumerate(row_elems):
            cells = [self._cell_from_element(c, r_idx, c_idx, section) for c_idx, c in enumerate(self._find_cells(row_elem))]
            if cells:
                rows.append(cells)
        if not rows:
            cells = self._find_cells(elem)
            if cells:
                rows = self._group_cells(cells, section)

        if not rows:
            inline_content = self._extract_inline_content(elem, section=section)
            if inline_content:
                cell_text = self._content_plain_text(inline_content)
                rows = [[TableCell(text=cell_text, row=0, col=0, paragraphs=[i.get("text", "") for i in inline_content if i.get("type") == "paragraph"], content=inline_content)]]

        caption = self._extract_caption(elem, "table") or caption_from_text(self._text_of(elem, exclude_objects=True), "table")
        raw: Dict[str, Any] = {
            **xml_meta,
            "row_count": len(rows),
            "cell_count": sum(len(r) for r in rows),
        }
        if not rows:
            raw["child_tags"] = [local_name(getattr(c, "tag", "")) for c in self._children(elem)[:50]]
            raw["descendant_tag_sample"] = [local_name(getattr(d, "tag", "")) for _, d in zip(range(80), self._descendants(elem))]
            raw["recovery"] = "empty table will be passed to postprocess follower-attachment heuristic"
        return Block("table", unique_id("tbl"), self._next_order(), section=section, caption=caption, rows=rows, attrs=attrs, raw=raw)

    def _find_rows(self, table_elem: Any) -> List[Any]:
        out: List[Any] = []

        def rec(node: Any) -> None:
            for child in self._children(node):
                if child is not table_elem and self._is_table(child) and not self._is_table_body(child):
                    continue
                if self._is_row(child):
                    out.append(child)
                    continue
                rec(child)
        rec(table_elem)
        return out

    def _find_cells(self, elem: Any) -> List[Any]:
        out: List[Any] = []

        def rec(node: Any) -> None:
            for child in self._children(node):
                if child is not elem and self._is_table(child) and not self._is_table_body(child):
                    continue
                if self._is_cell(child):
                    out.append(child)
                    continue
                rec(child)
        rec(elem)
        # If elem itself is a cell, include it only if no descendants were found.
        if not out and self._is_cell(elem):
            out.append(elem)
        return _unique_by_id(out)

    def _cell_from_element(self, cell_elem: Any, row: int, col: int, section: Optional[int] = None) -> TableCell:
        attrs = attrs_to_dict(cell_elem)
        content = self._extract_inline_content(cell_elem, section=section)
        paras = [str(item.get("text", "")) for item in content if item.get("type") == "paragraph" and str(item.get("text", "")).strip()]
        if not paras:
            paras = self._paragraphs_inside(cell_elem)
        text = self._content_plain_text(content)
        if not text:
            text = normalize_space("\n".join(paras)) if paras else self._text_of(cell_elem, exclude_objects=True)
        return TableCell(
            text=text,
            row=row,
            col=col,
            row_span=max(1, self._first_int(attrs, ["rowspan", "rowSpan", "row-span", "mergedRows", "rowSpanValue"], 1)),
            col_span=max(1, self._first_int(attrs, ["colspan", "colSpan", "col-span", "mergedCols", "colSpanValue"], 1)),
            attrs=attrs,
            paragraphs=paras,
            content=content,
        )

    def _content_plain_text(self, content: Sequence[Dict[str, Any]]) -> str:
        parts: List[str] = []
        for item in content or []:
            typ = str(item.get("type", ""))
            if typ == "paragraph" and item.get("text"):
                parts.append(str(item.get("text")))
            elif typ == "caption" and item.get("text"):
                parts.append(str(item.get("text")))
            elif typ == "image":
                cap = item.get("caption") or {}
                if isinstance(cap, dict) and cap.get("text"):
                    parts.append(str(cap.get("text")))
            elif typ == "table" and item.get("text"):
                parts.append(str(item.get("text")))
        return normalize_space("\n".join(parts))

    def _extract_inline_content(self, elem: Any, section: Optional[int] = None) -> List[Dict[str, Any]]:
        """Extract mixed paragraph/image/caption content inside a table cell.

        v11 keeps the DOM order of paragraphs, picture controls, embedded BinData
        payloads, and caption controls within each table cell. It no longer
        attaches arbitrary unlinked BinData images to captions; images are
        inserted only when a picture control or embedded binary node provides
        a structural position.
        This fixes the common HWP pattern where
        a large bordered table acts as a visual box and contains sections,
        screenshots, and captions inside one cell.
        """
        content: List[Dict[str, Any]] = []
        seen_nodes: set[int] = set()

        def add_paragraph(text: str, raw: Optional[Dict[str, Any]] = None) -> None:
            """Add text while preserving caption lines as separate inline items.

            pyhwp frequently flattens a whole table cell into a single paragraph
            containing lines like 'caption 그림. ...'. If we keep that as plain
            text, the figure cannot be paired with its caption. v9 splits only
            caption-looking lines and leaves ordinary text grouped.
            """
            text = normalize_space(text)
            if not self._valid_text_block(text):
                return
            buf: List[str] = []

            def flush_para() -> None:
                para = normalize_space("\n".join(buf))
                buf.clear()
                if not self._valid_text_block(para):
                    return
                if content and content[-1].get("type") == "paragraph" and normalize_space(str(content[-1].get("text", ""))) == para:
                    return
                content.append({"type": "paragraph", "text": para, "raw": raw or {}})

            for line in text.splitlines():
                line = normalize_space(line)
                if not line:
                    continue
                cap = caption_from_text(line)
                if cap is not None:
                    flush_para()
                    hint = infer_caption_target_type(cap.text, {}, "")
                    content.append({
                        "type": "caption",
                        "text": cap.text,
                        "caption": cap.to_dict(),
                        "target_type_hint": hint,
                        "raw": {**(raw or {}), "split_from_multiline_paragraph": True},
                    })
                else:
                    buf.append(line)
            flush_para()

        def add_caption(item: Optional[Dict[str, Any]]) -> None:
            if not item:
                return
            text = normalize_space(str(item.get("text", "")))
            if not text:
                return
            if content and content[-1].get("type") == "caption" and normalize_space(str(content[-1].get("text", ""))) == text:
                return
            content.append(item)

        def scan(node: Any, allow_para: bool = True) -> None:
            for child in self._children(node):
                ident = id(child)
                if ident in seen_nodes:
                    continue
                seen_nodes.add(ident)
                embedded_path = self._embedded_image_by_node_id.get(id(child))
                if embedded_path:
                    self.media_store.mark_linked(embedded_path)
                    content.append(self._image_content_from_media(embedded_path, caption_item=None, section=section, reason="embedded-base64-dom-position"))
                    continue
                if child is not elem and self._is_table(child) and not self._is_table_body(child):
                    try:
                        nested = self._parse_table(child, section, {"xml_path": ()})
                        nested_content = _flatten_table_content(nested.rows or [])
                        content.append({
                            "type": "table",
                            "text": self._content_plain_text(nested_content),
                            "rows": [[c.to_dict() for c in row] for row in (nested.rows or [])],
                            "caption": nested.caption.to_dict() if nested.caption else None,
                            "raw": nested.raw,
                        })
                    except Exception:
                        txt = self._text_of(child, exclude_objects=True)
                        add_paragraph(txt, {"note": "nested table fallback text"})
                    continue
                if self._is_image(child):
                    content.append(self._image_content_from_element(child, section=section))
                    continue
                if self._is_caption_control(child):
                    add_caption(self._caption_content_from_element(child))
                    continue
                if allow_para and self._is_para(child):
                    txt = self._text_of(child, exclude_objects=True)
                    add_paragraph(txt, {"xml_tag": local_name(getattr(child, "tag", "")), "attrs": attrs_to_dict(child)})
                    scan(child, allow_para=False)
                    continue
                scan(child, allow_para=allow_para)

        scan(elem, allow_para=True)
        if not content:
            txt = self._text_of(elem, exclude_objects=True)
            if self._valid_text_block(txt):
                add_paragraph(txt, {"fallback": "cell_text_of"})
        content = self._attach_inline_captions(content)
        content = self._materialize_missing_images_for_captions(content, section=section)
        return content

    def _image_content_from_element(self, elem: Any, section: Optional[int] = None) -> Dict[str, Any]:
        attrs = attrs_to_dict(elem)
        ref = self._find_binary_ref(elem)
        image_path = self.media_store.resolve(ref) if ref else None
        if not image_path:
            # Prefer an embedded base64 payload physically inside this picture
            # control. This is structural and therefore safe.
            image_path = self._embedded_image_by_node_id.get(id(elem))
            if not image_path:
                for d in self._descendants(elem):
                    image_path = self._embedded_image_by_node_id.get(id(d))
                    if image_path:
                        break
        if not image_path:
            # Last resort for an explicit picture control only. This is still
            # less risky than attaching images to caption-only text, because
            # the DOM told us a picture exists at this exact location.
            image_path = self.media_store.next_picture_order_image()
        if image_path:
            self.media_store.mark_linked(image_path)
        media_type = None
        if image_path:
            try:
                _, media_type = detect_extension(Path(image_path).read_bytes())
            except Exception:
                media_type = mimetypes.guess_type(image_path)[0]
        cap = self._extract_caption(elem, "image")
        return {
            "type": "image",
            "id": unique_id("imgin"),
            "image_path": image_path,
            "media_type": media_type,
            "caption": cap.to_dict() if cap else None,
            "geometry": self._extract_geometry(attrs),
            "attrs": attrs,
            "raw": {"xml_tag": local_name(getattr(elem, "tag", "")), "binary_ref": ref, "linked": bool(image_path), "inline_in_table_cell": True},
        }

    def _image_content_from_media(self, image_path: str, caption_item: Optional[Dict[str, Any]] = None, section: Optional[int] = None, reason: str = "caption-only-media-fallback") -> Dict[str, Any]:
        media_type = None
        if image_path:
            try:
                _, media_type = detect_extension(Path(image_path).read_bytes())
            except Exception:
                media_type = mimetypes.guess_type(image_path)[0]
        cap_obj = None
        if caption_item:
            cap_obj = caption_item.get("caption") or {"text": caption_item.get("text", ""), "method": "caption-only"}
            if isinstance(cap_obj, Caption):
                cap_obj = cap_obj.to_dict()
            if isinstance(cap_obj, dict):
                cap_obj.setdefault("position", "after")
                cap_obj.setdefault("method", "caption-only-media-fallback")
                cap_obj.setdefault("raw", {})
                if isinstance(cap_obj.get("raw"), dict):
                    cap_obj["raw"]["attach_method"] = reason
        return {
            "type": "image",
            "id": unique_id("imgcap"),
            "image_path": image_path,
            "media_type": media_type,
            "caption": cap_obj,
            "geometry": {},
            "attrs": {},
            "raw": {"linked": bool(image_path), "link_method": reason, "created_from_caption_without_picture_control": True},
        }

    def _materialize_missing_images_for_captions(self, content: List[Dict[str, Any]], section: Optional[int] = None) -> List[Dict[str, Any]]:
        """Do not attach arbitrary BinData images to caption-only text.

        Earlier versions inserted the next unlinked image whenever a line looked
        like an image caption. That placed cover-page/title images under later
        captions when pyhwp did not expose picture controls in the same cell.

        v11 keeps caption-only items as captions unless an image item was
        structurally present in the DOM and was already inserted by
        _extract_inline_content(). This avoids showing the wrong picture as if
        it had been correctly extracted.
        """
        out: List[Dict[str, Any]] = []
        for item in content:
            if not isinstance(item, dict):
                continue
            typ = item.get("type")
            if typ == "table":
                for row in item.get("rows", []) or []:
                    for cell in row or []:
                        if isinstance(cell, dict) and cell.get("content"):
                            cell["content"] = self._materialize_missing_images_for_captions(cell.get("content") or [], section=section)
                out.append(item)
                continue
            if typ == "caption":
                text = normalize_space(str(item.get("text", "")))
                hint = item.get("target_type_hint") or infer_caption_target_type(text, {}, "")
                if hint == "image" and not (out and out[-1].get("type") == "image"):
                    hinted_path = self.media_store.resolve_caption_image(text)
                    if hinted_path:
                        out.append(self._image_content_from_media(hinted_path, caption_item=item, section=section, reason="binary-record-caption-image-hint"))
                        continue
                    item.setdefault("raw", {})
                    if isinstance(item.get("raw"), dict):
                        item["raw"]["unresolved_image_reason"] = (
                            "image caption was detected, but no picture control, embedded binary node, or reliable BodyText picture-record hint was present at this location; "
                            "arbitrary BinData fallback is disabled to avoid linking the wrong image"
                        )
                out.append(item)
                continue
            out.append(item)
        return out

    def _caption_content_from_element(self, elem: Any) -> Optional[Dict[str, Any]]:
        attrs = attrs_to_dict(elem)
        text = first_nonempty_line(self._text_of(elem, exclude_objects=False))
        if not text:
            for k, v in attrs.items():
                if "caption" in str(k).lower() and str(v).strip() and str(v).lower() not in {"true", "false", "1", "0"}:
                    text = first_nonempty_line(str(v))
                    break
        if not self._valid_text_block(text):
            return None
        hint = infer_caption_target_type(text, attrs, local_name(getattr(elem, "tag", "")))
        cap = caption_from_text(text, hint) if hint in {"table", "image", "equation"} else caption_from_text(text)
        if cap is None:
            cap = Caption(text=text, method="structural-inline-control")
        else:
            cap.method = "structural-inline-control"
        cap.position = cap.position or caption_position_from_attrs(attrs)
        cap.raw.update({"attrs": attrs, "target_type_hint": hint, "xml_tag": local_name(getattr(elem, "tag", ""))})
        return {"type": "caption", "text": cap.text, "caption": cap.to_dict(), "target_type_hint": hint, "raw": cap.raw}

    def _attach_inline_captions(self, content: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        def compatible(obj: Dict[str, Any], hint: Optional[str]) -> bool:
            if not hint:
                return obj.get("type") in {"image", "table", "equation"}
            if hint == "image":
                return obj.get("type") == "image"
            return obj.get("type") == hint

        out: List[Dict[str, Any]] = []
        i = 0
        while i < len(content):
            item = content[i]
            if item.get("type") == "caption":
                text = normalize_space(str(item.get("text", "")))
                hint = item.get("target_type_hint") or infer_caption_target_type(text, {}, "")
                if out and compatible(out[-1], hint) and not out[-1].get("caption"):
                    cap = item.get("caption") or {"text": text, "method": "inline-caption"}
                    if isinstance(cap, dict):
                        cap.setdefault("position", "after")
                        cap.setdefault("method", "inline-caption")
                        cap.setdefault("raw", {})
                        if isinstance(cap.get("raw"), dict):
                            cap["raw"]["attach_method"] = "inline-cell-previous-object-v8"
                    out[-1]["caption"] = cap
                    i += 1
                    continue
                if i + 1 < len(content) and compatible(content[i + 1], hint) and not content[i + 1].get("caption"):
                    cap = item.get("caption") or {"text": text, "method": "inline-caption"}
                    if isinstance(cap, dict):
                        cap.setdefault("position", "before")
                        cap.setdefault("method", "inline-caption")
                        cap.setdefault("raw", {})
                        if isinstance(cap.get("raw"), dict):
                            cap["raw"]["attach_method"] = "inline-cell-next-object-v8"
                    content[i + 1]["caption"] = cap
                    i += 1
                    continue
            out.append(item)
            i += 1
        return out

    def _paragraphs_inside(self, elem: Any) -> List[str]:
        paras: List[str] = []

        def rec(node: Any) -> None:
            if node is not elem and (self._is_table(node) or self._is_image(node)):
                return
            if node is not elem and self._is_para(node):
                text = self._text_of(node, exclude_objects=True)
                if self._valid_text_block(text):
                    paras.append(text)
                return
            for child in self._children(node):
                rec(child)
        rec(elem)
        if not paras:
            text = self._text_of(elem, exclude_objects=True)
            if self._valid_text_block(text):
                paras.append(text)
        return _dedupe_adjacent_strings(paras)

    def _group_cells(self, cells: List[Any], section: Optional[int] = None) -> List[List[TableCell]]:
        coords: List[Tuple[int, int, Any]] = []
        has_coord = False
        for i, c in enumerate(cells):
            attrs = attrs_to_dict(c)
            r = self._first_int(attrs, ["row", "rowIndex", "row-index", "rowAddr", "row-addr", "rowNo", "rownum"], -1)
            col = self._first_int(attrs, ["col", "colIndex", "col-index", "colAddr", "col-addr", "colNo", "colnum", "column"], i)
            if r >= 0:
                has_coord = True
            coords.append((r, col, c))
        if has_coord:
            explicit_rows = [r for r, _, _ in coords if r >= 0]
            r_base = 1 if explicit_rows and min(explicit_rows) == 1 else 0
            bucket: Dict[int, List[Tuple[int, Any]]] = {}
            for seq, (r, c, cell) in enumerate(coords):
                rr = (r - r_base) if r >= 0 else 0
                bucket.setdefault(rr, []).append((c, cell))
            result: List[List[TableCell]] = []
            for rr in sorted(bucket):
                row: List[TableCell] = []
                for _, cell in sorted(bucket[rr], key=lambda x: x[0]):
                    row.append(self._cell_from_element(cell, len(result), len(row), section))
                result.append(row)
            return result
        # Preserve all text if no geometry exists.
        return [[self._cell_from_element(cell, 0, i, section) for i, cell in enumerate(cells)]]

    def _first_int(self, attrs: Dict[str, Any], keys: Sequence[str], default: int) -> int:
        lower = {k.lower(): v for k, v in attrs.items()}
        for key in keys:
            if key in attrs:
                return int(safe_int(attrs[key], default) or default)
            if key.lower() in lower:
                return int(safe_int(lower[key.lower()], default) or default)
        return default

    def _parse_image(self, elem: Any, section: Optional[int], context: Optional[Dict[str, Any]] = None) -> Block:
        context = context or {}
        attrs = attrs_to_dict(elem)
        xml_meta = self._xml_meta(elem, context)
        caption = self._extract_caption(elem, "image")
        ref = self._find_binary_ref(elem)
        image_path = self.media_store.resolve(ref) if ref else None
        if image_path:
            self.media_store.mark_linked(image_path)
        geometry = self._extract_geometry(attrs)
        if not geometry:
            for d in self._descendants(elem):
                geometry = self._extract_geometry(attrs_to_dict(d))
                if geometry:
                    break
        media_type = None
        if image_path:
            try:
                _, media_type = detect_extension(Path(image_path).read_bytes())
            except Exception:
                media_type = mimetypes.guess_type(image_path)[0]
        return Block(
            "image",
            unique_id("img"),
            self._next_order(),
            section=section,
            caption=caption,
            image_path=image_path,
            media_type=media_type,
            geometry=geometry,
            attrs=attrs,
            raw={**xml_meta, "binary_ref": ref, "linked": bool(image_path)},
        )

    def _extract_geometry(self, attrs: Dict[str, Any]) -> Dict[str, Any]:
        keys = ["width", "height", "x", "y", "left", "top", "right", "bottom", "szWidth", "szHeight", "posX", "posY", "offsetX", "offsetY"]
        out: Dict[str, Any] = {}
        lower = {k.lower(): k for k in attrs}
        for key in keys:
            if key in attrs:
                out[key] = attrs[key]
            elif key.lower() in lower:
                out[key] = attrs[lower[key.lower()]]
        return out

    def _extract_caption(self, elem: Any, object_type: str) -> Optional[Caption]:
        """Extract a caption nested directly in an object, with type safety.

        v8 change: when parsing an outer visual table, do not steal image
        captions that live inside the table. A caption whose label/hint says
        image must remain with the image control, not the surrounding table.
        """
        def is_compatible(text: str, attrs: Dict[str, Any], tag: str) -> Tuple[bool, Optional[str], Optional[Caption]]:
            target_hint = infer_caption_target_type(text, attrs, tag)
            typed_cap = caption_from_text(text, object_type)
            if typed_cap is not None:
                return True, target_hint, typed_cap
            # Explicit contradictory signal: skip. This is the critical guard
            # against attaching "그림..." captions to a table block.
            if target_hint is not None and target_hint != object_type:
                return False, target_hint, None
            # A visible caption pattern for another object type is also a skip.
            if object_type != "table" and caption_from_text(text, "table"):
                return False, "table", None
            if object_type != "image" and caption_from_text(text, "image"):
                return False, "image", None
            if object_type != "equation" and caption_from_text(text, "equation"):
                return False, "equation", None
            return True, target_hint, None

        candidates = [elem] + list(self._descendants(elem))
        for d in candidates:
            if self._is_caption_node(d) or self._is_caption_control(d):
                text = first_nonempty_line(self._text_of(d, exclude_objects=False))
                attrs = attrs_to_dict(d)
                if not text:
                    for k, v in attrs.items():
                        if "caption" in str(k).lower() and str(v).strip() and str(v).lower() not in {"true", "false", "1", "0"}:
                            text = first_nonempty_line(str(v))
                            break
                if text:
                    ok, target_hint, typed_cap = is_compatible(text, attrs, local_name(getattr(d, "tag", "")))
                    if not ok:
                        continue
                    cap = typed_cap or Caption(text=text, method="structural")
                    cap.method = "structural"
                    cap.position = cap.position or caption_position_from_attrs(attrs)
                    cap.raw.update({
                        "tag": local_name(getattr(d, "tag", "")),
                        "attrs": attrs,
                        "target_type_hint": target_hint,
                        "position_hint": cap.position,
                        "type_safe_v8": True,
                    })
                    return cap
        for d in candidates:
            attrs = attrs_to_dict(d)
            lower_attrs = {str(k).lower(): str(v) for k, v in attrs.items()}
            if any("caption" in k or k in CAPTION_ATTR_KEYS for k in lower_attrs) or any("caption" in v.lower() for v in lower_attrs.values()):
                text = first_nonempty_line(self._text_of(d, exclude_objects=False))
                if not text:
                    for v in lower_attrs.values():
                        if v.strip() and v.lower() not in {"true", "false", "1", "0"}:
                            text = first_nonempty_line(v)
                            break
                if text:
                    ok, target_hint, typed_cap = is_compatible(text, attrs, local_name(getattr(d, "tag", "")))
                    if not ok:
                        continue
                    return typed_cap or Caption(text=text, method="structural-attr", position=caption_position_from_attrs(attrs), raw={"attrs": attrs, "tag": local_name(getattr(d, "tag", "")), "position_hint": caption_position_from_attrs(attrs), "target_type_hint": target_hint, "type_safe_v8": True})
        return None

    def _find_binary_ref(self, elem: Any) -> Optional[str]:
        # v11: do not treat generic object ids such as instance-id/table-id as
        # image references. They can accidentally resolve to BIN0001-like media
        # and produce a visually plausible but wrong image.
        strong_key_pat = re.compile(r"(href|src|path|file|bin|bindata|binary|image|pic|picture|stream)", re.I)
        val_pat = re.compile(r"(BIN\s*[0-9A-Fa-f]+|BinData[/\\][^\x00\s]+|\.(png|jpg|jpeg|gif|bmp|tif|tiff|wmf|emf|webp)\b)", re.I)

        def check(node: Any) -> Optional[str]:
            attrs = attrs_to_dict(node)
            # Value evidence is strongest regardless of key.
            for k, v in attrs.items():
                sv = str(v).strip()
                if sv and val_pat.search(sv):
                    return sv
            # Otherwise require an image/binary-specific key.
            for k, v in attrs.items():
                sv = str(v).strip()
                if not sv:
                    continue
                if strong_key_pat.search(str(k)):
                    # Avoid boolean/layout flags.
                    if sv.lower() in {"true", "false", "0", "1", "inline", "block"}:
                        continue
                    return sv
            return None
        ref = check(elem)
        if ref:
            return ref
        for d in self._descendants(elem):
            ref = check(d)
            if ref:
                return ref
        return None

    def _extract_embedded_binaries(self, root: Any) -> None:
        count = 0
        for node in self._descendants(root):
            text = getattr(node, "text", None)
            if not text:
                continue
            data = decode_base64_maybe(text)
            if not data or len(data) < 16:
                continue
            attrs = attrs_to_dict(node)
            filename = self._filename_from_attrs(attrs) or f"embedded_{count:04d}"
            refs = list(attrs.values()) + [filename, local_name(getattr(node, "tag", ""))]
            item = self.media_store.add_bytes(data, filename, source="pyhwp:embedded-base64", refs=refs)
            if item and item.media_type.startswith("image/"):
                self._embedded_image_by_node_id[id(node)] = item.path
            count += 1

    def _filename_from_attrs(self, attrs: Dict[str, Any]) -> Optional[str]:
        lower = {k.lower(): v for k, v in attrs.items()}
        for key in ("filename", "file", "href", "path", "name", "id", "stream", "streamname"):
            if key in lower and str(lower[key]).strip():
                return Path(str(lower[key])).name
        return None

    def _postprocess_blocks(self) -> None:
        # v7: keep visual verification order as close as possible to the XML
        # document order. Do this before any caption attachment so that
        # paragraph captions immediately before/after objects remain adjacent.
        self._stabilize_block_order()
        self._dedupe_adjacent_paragraphs()
        self._attach_following_paragraphs_to_empty_tables()
        self._resolve_structural_caption_blocks()
        self._resolve_nearby_captions()
        self._assign_unlinked_images_by_order()
        self._materialize_caption_images_in_table_cells()
        # Do NOT append unresolved BinData images to the block stream. Appending
        # media at the end made real captions and objects appear far apart.
        # Unresolved media are still visible in the media gallery.
        self._stabilize_block_order()
        for i, block in enumerate(self.blocks, start=1):
            block.order = i

    def _stabilize_block_order(self) -> None:
        def key(block: Block) -> Tuple[int, ...]:
            path = xml_path_tuple((block.raw or {}).get("xml_path", []))
            if path:
                return tuple([0] + list(path) + [block.order])
            return (1, block.order)
        self.blocks = sorted(self.blocks, key=key)

    def _dedupe_adjacent_paragraphs(self) -> None:
        out: List[Block] = []
        for b in self.blocks:
            if b.type == "paragraph" and out and out[-1].type == "paragraph":
                if normalize_space(out[-1].text or "") == normalize_space(b.text or ""):
                    continue
            out.append(b)
        self.blocks = out

    def _attach_following_paragraphs_to_empty_tables(self) -> None:
        """Recover table text when pyhwp emits TableControl as an empty marker.

        Some pyhwp XMLs expose <TableControl ...> with no descendants, while the
        paragraph records that belong to the table appear immediately after the
        marker. This heuristic does not pretend to know exact cell geometry; it
        preserves the text as single-column table rows and marks the method.
        """
        out: List[Block] = []
        i = 0
        while i < len(self.blocks):
            b = self.blocks[i]
            if b.type == "table" and not b.rows:
                followers: List[Block] = []
                j = i + 1
                while j < len(self.blocks) and self.blocks[j].type == "paragraph":
                    txt = normalize_space(self.blocks[j].text or "")
                    if not txt:
                        j += 1
                        continue
                    # Stop on clear caption; nearby caption resolver will handle it.
                    if caption_from_text(txt, "table") or caption_from_text(txt, "image"):
                        break
                    followers.append(self.blocks[j])
                    j += 1
                    # Keep heuristic bounded to avoid swallowing entire documents.
                    if len(followers) >= 30:
                        break
                if followers:
                    texts = _dedupe_adjacent_strings([f.text or "" for f in followers])
                    b.rows = [[TableCell(text=t, row=r, col=0, paragraphs=[t])] for r, t in enumerate(texts) if t]
                    b.raw["row_count"] = len(b.rows)
                    b.raw["cell_count"] = len(b.rows)
                    b.raw["recovery"] = "attached following paragraph records as single-column table rows; exact HWP cell geometry was not exposed by XML"
                    out.append(b)
                    i = i + 1 + len(followers)
                    continue
            out.append(b)
            i += 1
        self.blocks = out

    def _resolve_structural_caption_blocks(self) -> None:
        """Attach standalone CaptionControl blocks only with strong DOM evidence.

        v7 change:
        - v5/v6 could still attach a caption to a distant table/image when a
          weak type hint matched. This made captions appear beside the wrong
          object.
        - This resolver treats XML parent/sibling relation and explicit object
          keys as authoritative. Pure flattened block distance is no longer used
          for long-distance matching.
        - Captions without a structurally safe target remain as caption blocks at
          their own document position instead of corrupting object-caption pairs.
        """
        blocks = self.blocks
        consumed: set[int] = set()

        def compatible(obj: Block, target_hint: Optional[str]) -> bool:
            if obj.type not in {"table", "image", "equation"}:
                return False
            if obj.caption is not None:
                return False
            if target_hint is None:
                return True
            if target_hint == "image":
                return obj.type == "image"
            return obj.type == target_hint

        def as_path(block: Block, key: str) -> Tuple[int, ...]:
            return xml_path_tuple((block.raw or {}).get(key, []))

        def key_overlap(a: Block, b: Block) -> List[str]:
            ak = set((a.raw or {}).get("object_keys") or [])
            bk = set((b.raw or {}).get("object_keys") or [])
            ak.update(extract_object_keys(a.attrs, tag=(a.raw or {}).get("xml_tag", ""), text=a.text or ""))
            bk.update(extract_object_keys(b.attrs, tag=(b.raw or {}).get("xml_tag", ""), text=b.text or ""))
            noisy = {"0", "1", "true", "false", "table", "image", "picture", "figure"}
            return sorted((ak & bk) - noisy)

        def sibling_gap(a: Sequence[int], b: Sequence[int]) -> int:
            if not a or not b:
                return 999
            lca = common_prefix_len(a, b)
            if lca >= min(len(a), len(b)):
                return 0
            try:
                return abs(int(a[lca]) - int(b[lca]))
            except Exception:
                return 999

        for idx, block in enumerate(blocks):
            if block.type != "caption" or idx in consumed:
                continue

            cap_text = normalize_space(block.text or (block.caption.text if block.caption else ""))
            target_hint = (block.raw or {}).get("target_type_hint")
            if target_hint is None:
                target_hint = infer_caption_target_type(cap_text, block.attrs, (block.raw or {}).get("xml_tag", ""))

            cap = block.caption or Caption(text=cap_text, method="structural-control")
            cap.text = cap_text or cap.text
            cap.method = cap.method or "structural-control"
            cap_path = as_path(block, "xml_path")
            cap_parent = as_path(block, "xml_parent_path")
            cap_grandparent = cap_parent[:-1] if cap_parent else tuple()
            cap_pos_hint = (block.raw or {}).get("caption_position_hint") or cap.position

            candidates: List[Tuple[float, int, str, Block, Dict[str, Any]]] = []
            for obj_idx, obj in enumerate(blocks):
                if obj_idx == idx or not compatible(obj, target_hint):
                    continue

                obj_path = as_path(obj, "xml_path")
                obj_parent = as_path(obj, "xml_parent_path")
                obj_grandparent = obj_parent[:-1] if obj_parent else tuple()
                block_distance = abs(obj_idx - idx)
                overlap = key_overlap(block, obj)
                same_parent = bool(cap_parent and obj_parent and cap_parent == obj_parent)
                same_grandparent = bool(cap_grandparent and obj_grandparent and cap_grandparent == obj_grandparent)
                prefix_related = bool(cap_path and obj_path and (cap_path[:len(obj_path)] == obj_path or obj_path[:len(cap_path)] == cap_path))
                lca = common_prefix_len(cap_path, obj_path)
                gap = xml_path_gap(cap_path, obj_path) if cap_path and obj_path else 999
                sib_gap = sibling_gap(cap_path, obj_path)

                if cap_pos_hint in {"before", "after", "left", "right"}:
                    position = "after" if cap_pos_hint in {"after", "right"} else "before"
                elif cap_path and obj_path:
                    position = "after" if cap_path > obj_path else "before"
                else:
                    position = "after" if idx > obj_idx else "before"

                evidence: Dict[str, Any] = {
                    "caption_index": idx,
                    "object_index": obj_idx,
                    "block_distance": block_distance,
                    "xml_gap": gap,
                    "xml_lca_depth": lca,
                    "xml_sibling_gap": sib_gap,
                    "same_parent": same_parent,
                    "same_grandparent": same_grandparent,
                    "prefix_related": prefix_related,
                    "key_overlap": overlap,
                    "position_hint": cap_pos_hint,
                    "target_type_hint": target_hint,
                }

                has_key_evidence = bool(overlap)
                has_parent_evidence = same_parent and sib_gap <= 4
                has_grandparent_evidence = same_grandparent and sib_gap <= 6 and lca >= 2
                has_prefix_evidence = prefix_related and gap <= 10

                if not (has_key_evidence or has_parent_evidence or has_grandparent_evidence or has_prefix_evidence):
                    if not (block_distance <= 1 and target_hint is not None):
                        continue

                if cap_pos_hint in {"before", "after"} and not (has_key_evidence or has_prefix_evidence):
                    if cap_path and obj_path:
                        if cap_pos_hint == "before" and not (cap_path < obj_path):
                            continue
                        if cap_pos_hint == "after" and not (cap_path > obj_path):
                            continue
                    else:
                        if cap_pos_hint == "before" and not (idx < obj_idx):
                            continue
                        if cap_pos_hint == "after" and not (idx > obj_idx):
                            continue

                score = 0.0
                if has_key_evidence:
                    score += 200 + min(len(overlap), 5) * 15
                if has_prefix_evidence:
                    score += 140
                if has_parent_evidence:
                    score += 120 - sib_gap * 5
                if has_grandparent_evidence:
                    score += 80 - sib_gap * 4
                if lca >= 2:
                    score += min(lca * 6, 36)
                if gap < 999:
                    score += max(0, 30 - gap * 2)
                if block_distance <= 1:
                    score += 20
                elif block_distance > 6 and not has_key_evidence:
                    score -= min(block_distance * 3, 80)
                if target_hint is not None:
                    score += 15

                evidence["score"] = round(score, 2)
                candidates.append((score, sib_gap, position, obj, evidence))

            if not candidates:
                block.raw = block.raw or {}
                block.raw["unresolved_caption_reason"] = "no DOM-local or key-based target found; kept in original position to avoid wrong attachment"
                continue

            candidates.sort(key=lambda x: (-x[0], x[1]))
            score, _, position, obj, evidence = candidates[0]
            if score < 55:
                block.raw = block.raw or {}
                block.raw["unresolved_caption_reason"] = f"best DOM target score too low: {score:.1f}"
                continue

            cap.position = position
            if not cap.raw:
                cap.raw = {}
            cap.raw.update({
                "attached_from_block_id": block.id,
                "attach_method": "standalone-caption-control-dom-anchored-v7",
                **evidence,
            })
            obj.caption = cap
            consumed.add(idx)

        if consumed:
            self.blocks = [b for i, b in enumerate(blocks) if i not in consumed]

    def _resolve_nearby_captions(self) -> None:
        """Attach only immediate paragraph captions.

        This is deliberately conservative. If a caption paragraph is not
        immediately adjacent to a table/image/equation in the stabilized block
        stream, it remains in its own document position instead of being
        attached to a distant object.
        """
        out: List[Block] = []
        i = 0
        while i < len(self.blocks):
            b = self.blocks[i]
            if b.type in {"table", "image", "equation"} and b.caption is None:
                obj_type = "image" if b.type == "image" else b.type
                if out and out[-1].type == "paragraph":
                    cap = caption_from_text(out[-1].text or "", obj_type)
                    if cap:
                        cap.method = "adjacent-previous-paragraph"
                        cap.position = "before"
                        if not cap.raw:
                            cap.raw = {}
                        cap.raw["attach_method"] = "immediate-adjacent-only-v7"
                        b.caption = cap
                        out.pop()
                        out.append(b)
                        i += 1
                        continue
                if i + 1 < len(self.blocks) and self.blocks[i + 1].type == "paragraph":
                    cap = caption_from_text(self.blocks[i + 1].text or "", obj_type)
                    if cap:
                        cap.method = "adjacent-next-paragraph"
                        cap.position = "after"
                        if not cap.raw:
                            cap.raw = {}
                        cap.raw["attach_method"] = "immediate-adjacent-only-v7"
                        b.caption = cap
                        out.append(b)
                        i += 2
                        continue
            out.append(b)
            i += 1
        self.blocks = out

    def _materialize_caption_images_in_table_cells(self) -> None:
        """Second-pass recovery for tables already flattened before v9 splitting.

        When pyhwp emits a bordered HWP table as one giant cell, the cell text
        may contain caption lines but no PictureControl nodes. This pass converts
        cell.text into mixed content and inserts sequential BinData images at
        image-caption locations.
        """
        for block in self.blocks:
            if block.type != "table" or not block.rows:
                continue
            for row in block.rows or []:
                for cell in row or []:
                    content = getattr(cell, "content", None) or []
                    if not content and getattr(cell, "text", None):
                        temp: List[Dict[str, Any]] = []
                        buf: List[str] = []

                        def flush_para() -> None:
                            para = normalize_space("\n".join(buf))
                            buf.clear()
                            if self._valid_text_block(para):
                                temp.append({"type": "paragraph", "text": para, "raw": {"split_from_cell_text_second_pass": True}})

                        for line in normalize_space(cell.text or "").splitlines():
                            line = normalize_space(line)
                            if not line:
                                continue
                            cap = caption_from_text(line)
                            if cap is not None:
                                flush_para()
                                hint = infer_caption_target_type(cap.text, {}, "")
                                temp.append({"type": "caption", "text": cap.text, "caption": cap.to_dict(), "target_type_hint": hint, "raw": {"split_from_cell_text_second_pass": True}})
                            else:
                                buf.append(line)
                        flush_para()
                        content = temp
                    if content:
                        content = self._attach_inline_captions(content)
                        content = self._materialize_missing_images_for_captions(content, section=block.section)
                        cell.content = content
                        cell.text = self._content_plain_text(content)
                        cell.paragraphs = [str(i.get("text", "")) for i in content if isinstance(i, dict) and i.get("type") == "paragraph" and i.get("text")]

    def _assign_unlinked_images_by_order(self) -> None:
        for b in self.blocks:
            if b.type == "image" and not b.image_path:
                path = self.media_store.next_picture_order_image()
                if path:
                    b.image_path = path
                    try:
                        _, b.media_type = detect_extension(Path(path).read_bytes())
                    except Exception:
                        b.media_type = mimetypes.guess_type(path)[0]
                    b.raw["linked"] = True
                    b.raw["link_method"] = "sequential-unlinked-media-fallback"
            elif b.type == "image" and b.image_path:
                self.media_store.mark_linked(b.image_path)

    def _add_unplaced_media_blocks(self) -> None:
        for item in self.media_store.items:
            if item.media_type.startswith("image/") and not item.linked and Path(item.path).name.lower() != "prvimage":
                self.blocks.append(Block(
                    "image",
                    unique_id("img"),
                    self._next_order(),
                    image_path=item.path,
                    media_type=item.media_type,
                    raw={"unplaced_media": True, "source": item.source, "note": "image extracted from BinData but no corresponding picture control was resolved"},
                ))
                item.linked = True



def _flatten_table_content(rows: List[List[TableCell]]) -> List[Dict[str, Any]]:
    content: List[Dict[str, Any]] = []
    for row in rows or []:
        for cell in row or []:
            if getattr(cell, "content", None):
                content.extend(cell.content)
            elif getattr(cell, "text", None):
                content.append({"type": "paragraph", "text": cell.text})
    return content
def _unique_by_id(items: List[Any]) -> List[Any]:
    seen = set()
    out = []
    for item in items:
        ident = id(item)
        if ident not in seen:
            seen.add(ident)
            out.append(item)
    return out


def _dedupe_adjacent_strings(items: Sequence[str]) -> List[str]:
    out: List[str] = []
    for s in items:
        ns = normalize_space(s)
        if ns and (not out or out[-1] != ns):
            out.append(ns)
    return out


# =============================================================================
# Binary fallback parser
# =============================================================================

class BinaryHwpParser:
    """Structured HWP binary fallback parser.

    v14 changes the fallback from a record sniffer to a table-aware stream
    parser. It reconstructs HWP table cells from TABLE/LIST_HEADER/PARA_TEXT
    record groups, which is the path needed when pyhwp conversion times out on
    large or form-heavy HWP files.
    """

    def __init__(self, input_path: Union[str, Path], output_dir: Union[str, Path], media_store: Optional[MediaStore] = None, preview: Optional[Dict[str, Any]] = None):
        self.input_path = Path(input_path)
        self.output_dir = ensure_dir(output_dir)
        self.media_store = media_store or MediaStore(self.output_dir / "media")
        self.preview = preview or {}
        self.blocks: List[Block] = []
        self.warnings: List[str] = []
        self.errors: List[str] = []
        self.metadata: Dict[str, Any] = {}
        self._order = 0
        self._picture_event_index = 0

    def parse(self) -> ParsedDocument:
        if olefile is None:
            raise RuntimeError("olefile is required for binary parsing. Install: pip install olefile")
        if not olefile.isOleFile(str(self.input_path)):
            raise RuntimeError("not an OLE/CFB HWP file")
        with olefile.OleFileIO(str(self.input_path)) as ole:
            self.ole = ole
            self._parse_header()
            self._parse_sections()
        self._postprocess()
        return ParsedDocument(
            source_path=str(self.input_path),
            method="binary_structured_fallback_v14",
            blocks=self.blocks,
            media_items=self.media_store.items,
            metadata=self.metadata,
            warnings=self.warnings,
            errors=self.errors,
        )

    def _next_order(self) -> int:
        self._order += 1
        return self._order

    def _parse_header(self) -> None:
        try:
            data = self.ole.openstream("FileHeader").read()
            self.metadata["file_header_size"] = len(data)
            if len(data) > 40:
                prop = data[36]
                self.metadata["compressed"] = bool(prop & 0x01)
                self.metadata["encrypted"] = bool(prop & 0x02)
                self.metadata["distributable"] = bool(prop & 0x04)
                if prop & 0x02:
                    self.warnings.append("encrypted HWP flag is set; content extraction may fail")
                if prop & 0x04:
                    self.warnings.append("distributable HWP flag is set; some content may be protected")
        except Exception as e:
            self.warnings.append(f"failed to read FileHeader: {e}")

    def _section_paths(self) -> List[List[str]]:
        paths = []
        for entry in self.ole.listdir(streams=True, storages=False):
            if len(entry) == 2 and entry[0] == "BodyText" and entry[1].startswith("Section"):
                paths.append(entry)
        return sorted(paths, key=lambda p: int(re.search(r"(\d+)$", p[-1]).group(1)) if re.search(r"(\d+)$", p[-1]) else 10**9)

    def _parse_sections(self) -> None:
        for s_idx, path in enumerate(self._section_paths()):
            try:
                raw = self.ole.openstream(path).read()
                data = raw
                if self.metadata.get("compressed"):
                    for wbits in (-15, zlib.MAX_WBITS):
                        try:
                            data = zlib.decompress(raw, wbits)
                            break
                        except Exception:
                            pass
                records = list(self.iter_records(data, section=s_idx))
                self._parse_record_list(records, section=s_idx)
            except Exception as e:
                self.errors.append(f"failed to parse {'/'.join(path)}: {e}")

    @staticmethod
    def iter_records(data: bytes, section: Optional[int] = None) -> Iterator[HwpRecord]:
        offset = 0
        total = len(data)
        while offset + 4 <= total:
            rec_offset = offset
            header = struct.unpack_from("<I", data, offset)[0]
            offset += 4
            tag_id = header & 0x3FF
            level = (header >> 10) & 0x3FF
            size = (header >> 20) & 0xFFF
            if size == 0xFFF:
                if offset + 4 > total:
                    break
                size = struct.unpack_from("<I", data, offset)[0]
                offset += 4
            if size < 0 or offset + size > total:
                break
            payload = data[offset:offset + size]
            offset += size
            yield HwpRecord(tag_id, level, size, payload, rec_offset, section=section)

    def _parse_record_list(self, records: List[HwpRecord], section: Optional[int]) -> None:
        i = 0
        n = len(records)
        while i < n:
            rec = records[i]
            if rec.tag_id == HWPTAG_TABLE:
                block, next_i = self._parse_table_group(records, i, section)
                self.blocks.append(block)
                i = max(next_i, i + 1)
                continue
            if rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
                block = self._image_block_from_picture(rec, section)
                self.blocks.append(block)
                i += 1
                continue
            if rec.tag_id == HWPTAG_EQEDIT:
                txt = self._decode_bytes_text(rec.payload)
                if normalize_space(txt):
                    self.blocks.append(Block("equation", unique_id("eq"), self._next_order(), section=section, text=normalize_space(txt), raw={"record": rec.to_meta()}))
                i += 1
                continue
            if rec.tag_id == HWPTAG_PARA_TEXT:
                text, ctrls = self.decode_para_text(rec.payload)
                text = normalize_space(text)
                if text and not self._is_control_only_text(text):
                    self.blocks.append(Block("paragraph", unique_id("p"), self._next_order(), section=section, text=text, raw={"record": rec.to_meta(), "controls": ctrls[:20]}))
            i += 1

    def _parse_table_group(self, records: List[HwpRecord], start: int, section: Optional[int]) -> Tuple[Block, int]:
        table_rec = records[start]
        row_count, col_count = self._table_dimensions(table_rec.payload)
        rows: Dict[int, Dict[int, TableCell]] = {}
        mixed: List[Dict[str, Any]] = []
        raw_cells: List[Dict[str, Any]] = []

        j = start + 1
        cell_no = 0
        last_text_cell: Optional[TableCell] = None
        while j < len(records):
            r = records[j]
            if r.level < table_rec.level:
                break
            # A sibling TABLE at the same level after at least one cell normally means next object.
            if j != start + 1 and r.tag_id == HWPTAG_TABLE and r.level <= table_rec.level:
                break
            if r.tag_id == HWPTAG_LIST_HEADER:
                cell_meta = self._parse_list_header_as_cell(r.payload, cell_no)
                cell_no += 1
                k = j + 1
                cell_records: List[HwpRecord] = []
                while k < len(records):
                    nr = records[k]
                    if nr.level < table_rec.level:
                        break
                    if k != j + 1 and nr.tag_id == HWPTAG_LIST_HEADER and nr.level == r.level:
                        break
                    if k != j + 1 and nr.tag_id == HWPTAG_TABLE and nr.level <= table_rec.level:
                        break
                    cell_records.append(nr)
                    k += 1
                cell, cell_mixed = self._parse_cell_records(cell_records, cell_meta, section)
                rr = int(cell.row if cell.row is not None else 0)
                cc = int(cell.col if cell.col is not None else 0)
                rows.setdefault(rr, {})[cc] = cell
                raw_cells.append({"row": rr, "col": cc, "row_span": cell.row_span, "col_span": cell.col_span, "text_preview": (cell.text or "")[:120]})
                for item in cell_mixed:
                    item.setdefault("row", rr)
                    item.setdefault("col", cc)
                    mixed.append(item)
                last_text_cell = cell if cell.text else last_text_cell
                j = k
                continue
            j += 1

        table_rows = self._materialize_rows(rows, row_count, col_count)
        caption = self._extract_table_caption_from_mixed(mixed)
        block = Block(
            "table",
            unique_id("tbl"),
            self._next_order(),
            section=section,
            caption=caption,
            rows=table_rows,
            raw={
                "record": table_rec.to_meta(),
                "binary_table_parser": "v14",
                "row_count_declared": row_count,
                "col_count_declared": col_count,
                "cell_count": sum(len(r) for r in table_rows),
                "cells": raw_cells[:200],
                "mixed_content_count": len(mixed),
                "payload_prefix_hex": table_rec.payload[:96].hex(),
            },
        )
        if mixed:
            block.raw["mixed_content"] = mixed[:400]
        return block, j

    def _table_dimensions(self, payload: bytes) -> Tuple[Optional[int], Optional[int]]:
        # HWP TABLE payload commonly stores row_count/col_count at offsets 4/6.
        if len(payload) >= 8:
            r = int.from_bytes(payload[4:6], "little", signed=False)
            c = int.from_bytes(payload[6:8], "little", signed=False)
            if 1 <= r <= 1000 and 1 <= c <= 200:
                return r, c
        guesses = self._guess_table_dimensions(payload)
        if guesses:
            return guesses[0].get("row_count_guess"), guesses[0].get("col_count_guess")
        return None, None

    def _parse_list_header_as_cell(self, payload: bytes, seq: int) -> Dict[str, int]:
        # For table cells, LIST_HEADER stores col,row,colSpan,rowSpan as u16 at 8,10,12,14.
        def u16(off: int, default: int = 0) -> int:
            if off + 2 <= len(payload):
                return int.from_bytes(payload[off:off+2], "little", signed=False)
            return default
        col = u16(8, seq)
        row = u16(10, 0)
        col_span = max(1, u16(12, 1))
        row_span = max(1, u16(14, 1))
        return {"row": row, "col": col, "row_span": row_span, "col_span": col_span, "seq": seq, "payload_hex": payload[:64].hex()}

    def _parse_cell_records(self, cell_records: List[HwpRecord], cell_meta: Dict[str, int], section: Optional[int]) -> Tuple[TableCell, List[Dict[str, Any]]]:
        parts: List[str] = []
        paras: List[str] = []
        mixed: List[Dict[str, Any]] = []
        pending_caption: Optional[Caption] = None
        for rec in cell_records:
            if rec.tag_id == HWPTAG_PARA_TEXT:
                text, ctrls = self.decode_para_text(rec.payload)
                text = normalize_space(text)
                if not text or self._is_control_only_text(text):
                    continue
                cap_img = caption_from_text(text, "image")
                cap_tbl = caption_from_text(text, "table")
                if cap_img:
                    # If a preceding image in the same cell has no caption, attach it.
                    attached = False
                    for item in reversed(mixed):
                        if item.get("type") == "image" and not item.get("caption"):
                            item["caption"] = cap_img.to_dict()
                            attached = True
                            break
                    if not attached:
                        mixed.append({"type": "caption", "text": cap_img.text, "caption": cap_img.to_dict(), "target_type": "image"})
                    parts.append(cap_img.text)
                    paras.append(cap_img.text)
                    continue
                if cap_tbl:
                    mixed.append({"type": "caption", "text": cap_tbl.text, "caption": cap_tbl.to_dict(), "target_type": "table"})
                    parts.append(cap_tbl.text)
                    paras.append(cap_tbl.text)
                    continue
                parts.append(text)
                paras.append(text)
                mixed.append({"type": "paragraph", "text": text})
            elif rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
                image_block = self._image_block_from_picture(rec, section)
                item = {
                    "type": "image",
                    "image_path": image_block.image_path,
                    "media_type": image_block.media_type,
                    "raw": image_block.raw,
                }
                mixed.append(item)
            elif rec.tag_id == HWPTAG_TABLE:
                # Nested table marker inside a cell. We keep a placeholder in mixed content.
                mixed.append({"type": "nested_table_marker", "record": rec.to_meta()})
        cell = TableCell(
            text=normalize_space("\n".join(parts)),
            row=cell_meta.get("row", 0),
            col=cell_meta.get("col", 0),
            row_span=max(1, int(cell_meta.get("row_span", 1))),
            col_span=max(1, int(cell_meta.get("col_span", 1))),
            attrs={"binary_cell_meta": cell_meta},
            paragraphs=_dedupe_adjacent_strings(paras),
        )
        return cell, mixed

    def _materialize_rows(self, rows: Dict[int, Dict[int, TableCell]], row_count: Optional[int], col_count: Optional[int]) -> List[List[TableCell]]:
        if not rows:
            return []
        max_r = max(rows.keys())
        max_c = max((max(cols.keys()) if cols else 0) for cols in rows.values())
        if row_count:
            max_r = max(max_r, row_count - 1)
        if col_count:
            max_c = max(max_c, col_count - 1)
        out: List[List[TableCell]] = []
        for r in range(max_r + 1):
            row: List[TableCell] = []
            for c in sorted(rows.get(r, {}).keys()):
                row.append(rows[r][c])
            if row:
                out.append(row)
        return out

    def _extract_table_caption_from_mixed(self, mixed: List[Dict[str, Any]]) -> Optional[Caption]:
        for item in mixed:
            if item.get("type") == "caption" and item.get("target_type") == "table":
                text = item.get("text") or ""
                if text:
                    return Caption(text=text, method="binary-cell-caption", position="inside-table")
        return None

    def _image_block_from_picture(self, rec: HwpRecord, section: Optional[int]) -> Block:
        refs = _bin_refs_from_picture_payload(rec.payload, self.media_store)
        path = self.media_store.resolve_picture_event_image(refs, self._picture_event_index)
        self._picture_event_index += 1
        media_type = None
        if path:
            self.media_store.mark_linked(path)
            try:
                _, media_type = detect_extension(Path(path).read_bytes())
            except Exception:
                media_type = mimetypes.guess_type(path)[0]
        return Block(
            "image",
            unique_id("img"),
            self._next_order(),
            section=section,
            image_path=path,
            media_type=media_type,
            raw={"record": rec.to_meta(), "picture_refs": refs, "picture_event_index": self._picture_event_index - 1, "link_method": "binary-picture-event-order-v14" if path else "unresolved-binary-picture"},
        )

    @staticmethod
    def decode_para_text(payload: bytes) -> Tuple[str, List[Dict[str, Any]]]:
        out: List[str] = []
        controls: List[Dict[str, Any]] = []
        i = 0
        n = len(payload)
        while i + 2 <= n:
            code = int.from_bytes(payload[i:i+2], "little")
            i += 2
            if code >= 32:
                try:
                    out.append(chr(code))
                except Exception:
                    pass
                continue
            if code == 0x0009:
                out.append("\t")
                continue
            if code in {0x000a, 0x000d}:
                out.append("\n")
                continue
            info = {"code": code, "offset": i - 2}
            if i + 14 <= n:
                raw = payload[i:i+14]
                consumed = False
                for cand in (raw[:4], raw[:4][::-1]):
                    if all(32 <= b <= 126 for b in cand):
                        ctrl_id = cand.decode("ascii", errors="ignore").strip()
                        if ctrl_id:
                            info["ctrl_id"] = ctrl_id
                            # Keep semantic placeholder only for non-layout controls.
                            if ctrl_id not in {"lbt", "osg", "gso", "onta", "dces", "dloc", "pngp"}:
                                out.append(f"⟦CTRL:{ctrl_id}⟧")
                            i += 14
                            consumed = True
                            break
                if not consumed:
                    pass
            controls.append(info)
        return "".join(out), controls

    @staticmethod
    def _is_control_only_text(text: str) -> bool:
        s = normalize_space(text)
        if not s:
            return True
        stripped = re.sub(r"[\s\[\]⟦⟧:A-Za-z0-9_\-]+", "", s)
        return not re.search(r"[가-힣A-Za-z0-9]", stripped) and "CTRL" in s

    def _guess_table_dimensions(self, payload: bytes) -> List[Dict[str, int]]:
        guesses = []
        for off in range(0, min(96, max(0, len(payload) - 4)), 2):
            r = int.from_bytes(payload[off:off+2], "little")
            c = int.from_bytes(payload[off+2:off+4], "little")
            if 1 <= r <= 300 and 1 <= c <= 100:
                guesses.append({"offset": off, "row_count_guess": r, "col_count_guess": c})
        return guesses[:8]

    def _decode_bytes_text(self, data: bytes) -> str:
        for enc in ("utf-16le", "utf-8", "cp949", "latin1"):
            try:
                s = data.decode(enc, errors="ignore")
                if normalize_space(s):
                    return s
            except Exception:
                pass
        return ""

    def _postprocess(self) -> None:
        self._dedupe_adjacent_paragraphs()
        self._resolve_nearby_captions()
        # v14 does not add unplaced media into the document stream; they remain
        # visible in the media gallery. This prevents layout corruption.
        for i, b in enumerate(self.blocks, 1):
            b.order = i

    def _dedupe_adjacent_paragraphs(self) -> None:
        out: List[Block] = []
        for b in self.blocks:
            if b.type == "paragraph" and out and out[-1].type == "paragraph" and normalize_space(out[-1].text or "") == normalize_space(b.text or ""):
                continue
            out.append(b)
        self.blocks = out

    def _resolve_nearby_captions(self) -> None:
        out: List[Block] = []
        i = 0
        while i < len(self.blocks):
            b = self.blocks[i]
            if b.type in {"table", "image", "equation"} and b.caption is None:
                typ = "image" if b.type == "image" else b.type
                if out and out[-1].type == "paragraph":
                    cap = caption_from_text(out[-1].text or "", typ)
                    if cap:
                        cap.method = "binary-nearby-previous"
                        cap.position = "before"
                        b.caption = cap
                        out.pop()
                        out.append(b)
                        i += 1
                        continue
                if i + 1 < len(self.blocks) and self.blocks[i+1].type == "paragraph":
                    cap = caption_from_text(self.blocks[i+1].text or "", typ)
                    if cap:
                        cap.method = "binary-nearby-next"
                        cap.position = "after"
                        b.caption = cap
                        out.append(b)
                        i += 2
                        continue
            out.append(b)
            i += 1
        self.blocks = out


# =============================================================================
# Orchestrator
# =============================================================================

class FullHwpParser:
    def __init__(self, input_path: Union[str, Path], output_dir: Union[str, Path], mode: str = "auto", keep_intermediate: bool = False, hwp5proc_path: str = "hwp5proc"):
        self.input_path = Path(input_path)
        self.output_dir = ensure_dir(output_dir)
        self.mode = mode
        self.keep_intermediate = keep_intermediate
        self.hwp5proc_path = hwp5proc_path
        self.warnings: List[str] = []
        self.errors: List[str] = []
        self.preview: Dict[str, Any] = {}
        self.media_store = MediaStore(self.output_dir / "media")

    def parse(self) -> ParsedDocument:
        if not self.input_path.exists():
            raise FileNotFoundError(str(self.input_path))
        if self.mode not in {"auto", "pyhwp", "binary", "xml"}:
            raise ValueError("mode must be one of: auto, pyhwp, binary, xml")

        # Always extract OLE BinData first when possible. This fixes many image
        # linking failures because pyhwp XML may reference BIN IDs without
        # embedding data in a directly resolvable node.
        if self.input_path.suffix.lower() == ".hwp":
            self.preview = extract_ole_media(self.input_path, self.media_store, self.warnings, self.errors)

        if self.mode == "xml" or self.input_path.suffix.lower() == ".xml":
            doc = PyHwpXmlParser(self.input_path, self.output_dir, self.media_store).parse()
            doc.warnings = self.warnings + doc.warnings
            doc.errors = self.errors + doc.errors
            doc.metadata["preview"] = self.preview
            return doc

        if self.mode in {"auto", "pyhwp"}:
            doc = self._try_pyhwp()
            if doc is not None:
                doc.warnings = self.warnings + doc.warnings
                doc.errors = self.errors + doc.errors
                doc.metadata["preview"] = self.preview
                return doc
            if self.mode == "pyhwp":
                raise RuntimeError("pyhwp XML conversion failed; use --mode binary or inspect errors")

        doc = BinaryHwpParser(self.input_path, self.output_dir, self.media_store, self.preview).parse()
        doc.warnings = self.warnings + doc.warnings
        doc.errors = self.errors + doc.errors
        doc.metadata["preview"] = self.preview
        return doc

    def _try_pyhwp(self) -> Optional[ParsedDocument]:
        # v16: avoid known pyhwp hangs on large/form-heavy HWP files.
        # Such files are handled by the improved binary parser instead.
        try:
            if self.input_path.suffix.lower() == ".hwp" and self.input_path.stat().st_size > 12 * 1024 * 1024:
                self.warnings.append("Large/form-heavy HWP detected; skipping pyhwp XML conversion and using binary parser directly.")
                return None
        except Exception:
            pass
        exe = shutil.which(self.hwp5proc_path) or (self.hwp5proc_path if Path(self.hwp5proc_path).exists() else None)
        if not exe:
            self.warnings.append("hwp5proc was not found. Falling back to binary parser. Install pyhwp for better table/image/caption extraction: pip install pyhwp")
            return None
        if self.keep_intermediate:
            xml_path = self.output_dir / "intermediate.xml"
            tmp = None
        else:
            tmp = tempfile.TemporaryDirectory(prefix="hwp_xml_")
            xml_path = Path(tmp.name) / "intermediate.xml"
        try:
            cmd = [exe, "xml", "--embedbin", "--output", str(xml_path), str(self.input_path)]
            proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=45)
            if proc.returncode != 0 or not xml_path.exists():
                self.errors.append(f"pyhwp conversion failed: returncode={proc.returncode}; stderr={(proc.stderr or proc.stdout)[:1000]}")
                return None
            doc = PyHwpXmlParser(xml_path, self.output_dir, self.media_store).parse()
            doc.source_path = str(self.input_path)
            doc.method = "pyhwp_xml+ole_media"
            if self.keep_intermediate:
                doc.metadata["intermediate_xml"] = str(xml_path)
            return doc
        except Exception as e:
            self.errors.append(f"pyhwp conversion exception: {e}")
            return None
        finally:
            if tmp is not None:
                tmp.cleanup()


# =============================================================================
# JSON and summary
# =============================================================================

def write_json(doc: ParsedDocument, path: Union[str, Path], pretty: bool = True) -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(doc.to_dict(), ensure_ascii=False, indent=2 if pretty else None, separators=None if pretty else (",", ":"))
    p.write_text(text, encoding="utf-8")
    return p


def summarize(doc: ParsedDocument) -> Dict[str, Any]:
    counts: Dict[str, int] = {}
    table_nonempty = 0
    image_linked = 0
    captions = 0
    inline_images = 0
    inline_captions = 0

    def walk_inline(items: Sequence[Dict[str, Any]]) -> None:
        nonlocal inline_images, inline_captions, image_linked, captions
        for item in items or []:
            if not isinstance(item, dict):
                continue
            if item.get("type") == "image":
                inline_images += 1
                if item.get("image_path"):
                    image_linked += 1
            if item.get("caption"):
                inline_captions += 1
                captions += 1
            for row in item.get("rows", []) or []:
                for cell in row or []:
                    if isinstance(cell, dict):
                        walk_inline(cell.get("content", []) or [])

    for b in doc.blocks:
        counts[b.type] = counts.get(b.type, 0) + 1
        if b.type == "table" and b.rows and sum(len(r) for r in b.rows) > 0:
            table_nonempty += 1
            for row in b.rows or []:
                for cell in row or []:
                    walk_inline(getattr(cell, "content", []) or [])
        if b.type == "image" and b.image_path:
            image_linked += 1
        if b.caption:
            captions += 1
    return {
        "source_path": doc.source_path,
        "method": doc.method,
        "block_counts": counts,
        "nonempty_tables": table_nonempty,
        "linked_image_blocks": image_linked,
        "inline_image_blocks": inline_images,
        "caption_count": captions,
        "inline_caption_count": inline_captions,
        "media_file_count": len(doc.media_items),
        "warning_count": len(doc.warnings),
        "error_count": len(doc.errors),
    }

# =============================================================================
# Web UI
# =============================================================================

VIEWER_HTML = r'''<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>HWP Parser Verification UI v19</title>
<style>
:root{--bg:#f5f7fb;--panel:#fff;--text:#172033;--muted:#667085;--line:#d9e0ec;--accent:#273b7a;--soft:#eef2ff;--warn:#8a5200;--err:#9d1c1c;}
*{box-sizing:border-box} body{margin:0;font-family:Segoe UI,Arial,'Malgun Gothic',sans-serif;background:var(--bg);color:var(--text)}
header{height:64px;display:flex;align-items:center;justify-content:space-between;padding:0 18px;background:#111827;color:#fff;border-bottom:1px solid #000}
.title{font-size:18px;font-weight:700}.subtitle{font-size:12px;color:#cbd5e1;margin-top:4px}.toolbar{display:flex;gap:8px;align-items:center}.button,button{border:1px solid var(--line);background:#fff;color:#172033;border-radius:10px;padding:8px 12px;text-decoration:none;cursor:pointer;font-size:13px}.button.primary,button.primary{background:#2b3f86;color:#fff;border-color:#2b3f86}
main{display:grid;grid-template-columns:1fr 1fr;gap:12px;height:calc(100vh - 64px);padding:12px}@media(max-width:1100px){main{grid-template-columns:1fr;height:auto}.pane{height:80vh}}
.pane{background:var(--panel);border:1px solid var(--line);border-radius:16px;box-shadow:0 8px 22px rgba(15,23,42,.06);overflow:hidden;display:flex;flex-direction:column}.pane-head{display:flex;justify-content:space-between;align-items:center;padding:12px 14px;border-bottom:1px solid var(--line);background:#fbfcff}.pane-title{font-weight:700}.badge{font-size:11px;border:1px solid #cdd6e3;border-radius:999px;padding:3px 8px;color:#475467;background:#f8fafc}.pane-body{padding:14px;overflow:auto;flex:1}
.notice{padding:10px 12px;border-radius:12px;margin:8px 0;border:1px solid #d6dae5;background:#f8fafc;color:#344054}.notice.warn{background:#fff8e6;color:var(--warn);border-color:#ffe1a6}.notice.err{background:#fff1f1;color:var(--err);border-color:#ffc9c9}.notice.info{background:#eef5ff;color:#244b7a;border-color:#cfe3ff}
.meta-grid{display:grid;grid-template-columns:repeat(5,minmax(90px,1fr));gap:8px;margin-bottom:12px}.metric{background:#f8fafc;border:1px solid var(--line);border-radius:12px;padding:10px}.metric .num{font-size:22px;font-weight:800}.metric .label{font-size:12px;color:var(--muted)}
.controls{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0}select,input{border:1px solid var(--line);border-radius:10px;padding:8px 10px;background:#fff;color:#172033}
.block{border:1px solid var(--line);border-radius:14px;background:#fff;margin:10px 0;padding:12px}.block-head{display:flex;align-items:center;justify-content:space-between;gap:8px;color:#475467;font-size:12px;margin-bottom:8px}.block-title{font-weight:700;color:#111827}.para{white-space:pre-wrap;line-height:1.65}.caption{background:#eef2ff;border:1px solid #d7defd;border-radius:10px;padding:8px;margin:8px 0;font-size:13px}.caption b{color:#273b7a}
.table-wrap{overflow:auto;border:1px solid var(--line);border-radius:12px;margin-top:8px}table.extracted{border-collapse:collapse;width:100%;font-size:13px}table.extracted td,table.extracted th{border:1px solid #cfd7e6;padding:8px;vertical-align:top;white-space:pre-wrap;min-width:80px}table.extracted td.empty{color:#98a2b3;background:#f8fafc}.image-box{border:1px solid var(--line);border-radius:12px;padding:8px;background:#f8fafc}.image-box img{max-width:100%;display:block;margin:auto;border-radius:8px}.raw{margin-top:8px}.raw pre{white-space:pre-wrap;background:#0f172a;color:#dbeafe;border-radius:12px;padding:10px;overflow:auto;font-size:12px;max-height:360px}.empty{color:#667085;padding:16px;border:1px dashed #cbd5e1;border-radius:12px;background:#f8fafc}
.upload-wrap{min-height:100vh;display:grid;place-items:center;padding:20px}.upload-card{max-width:720px;width:100%;background:#fff;border:1px solid var(--line);border-radius:20px;padding:24px;box-shadow:0 20px 40px rgba(15,23,42,.08)}.upload-card h1{margin:0 0 8px}.dropzone{border:2px dashed #b7c4d6;border-radius:16px;padding:24px;margin:18px 0;background:#f8fafc;text-align:center}
.source-doc{background:#fdfdfd;border:1px solid #d7dce8;border-radius:18px;padding:20px;max-width:850px;margin:0 auto 16px auto;box-shadow:0 10px 22px rgba(15,23,42,.05)}.source-block{margin:12px 0}.source-table{margin:12px 0}.source-caption{font-size:13px;text-align:center;color:#475467;margin:6px 0}.source-image img{max-width:100%;display:block;margin:auto}.preview-image{max-width:100%;border:1px solid var(--line);border-radius:12px}.media-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(160px,1fr));gap:10px}.media-card{border:1px solid var(--line);border-radius:12px;padding:8px;background:#fff}.media-card img{width:100%;height:120px;object-fit:contain;background:#f8fafc;border-radius:8px}.small{font-size:12px;color:#667085;word-break:break-all}
</style>
</head>
<body><div id="app"></div>
<script>
const APP_STATE = __APP_STATE__;
function esc(s){return String(s ?? '').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));}
function pretty(o){return esc(JSON.stringify(o,null,2));}
function counts(blocks){const c={}; for(const b of blocks||[]) c[b.type]=(c[b.type]||0)+1; return c;}
function walkItems(items, cb){
  for(const item of items||[]){
    if(!item) continue;
    cb(item);
    for(const ch of item.children||[]) walkItems([ch],cb);
    for(const row of item.rows||[]) for(const cell of row||[]) walkItems(cell.content||[],cb);
  }
}
function countImagesAll(blocks){let n=0; walkItems(blocks,(x)=>{if(x.type==='image')n++;}); return n;}
function countLinkedImages(blocks){let n=0; walkItems(blocks,(x)=>{if(x.type==='image'){if(x.image_paths&&x.image_paths.length)n+=x.image_paths.length; else if(x.image_path)n++;}}); return n;}
function countCaptionsAll(blocks){let n=0; walkItems(blocks,(x)=>{if(x.caption)n++;}); return n;}
function captionHtml(b){return b.caption ? `<div class="caption"><b>caption</b> [${esc(b.caption.method||'')}${b.caption.position?' · '+esc(b.caption.position):''}] ${esc(b.caption.text||'')}</div>` : '';}
function inlineItemHtml(item){
  if(!item) return '';
  const typ=item.type||'';
  if(typ==='paragraph') return `<div class="para">${esc(item.text||'')}</div>`;
  if(typ==='caption') return `<div class="caption"><b>caption</b> ${esc(item.text||'')}</div>`;
  if(typ==='image') return withCaption(item,imageHtml(item));
  if(typ==='table') return withCaption(item,tableHtml(item.rows||[]));
  return `<div class="para">${esc(item.text||'')}</div>`;
}
function cellHtml(c){
  const content=c.content||[];
  if(content.length) return content.map(inlineItemHtml).join('');
  return esc(c.text||'');
}
function tableHtml(rows){
  if(!rows || !rows.length) return '<div class="empty">표 블록은 감지됐지만 행/열 구조가 비어 있습니다. raw 진단 정보를 확인하세요.</div>';
  return `<div class="table-wrap"><table class="extracted"><tbody>${rows.map(r=>`<tr>${(r||[]).map(c=>`<td rowspan="${esc(c.row_span||1)}" colspan="${esc(c.col_span||1)}" class="${(c.text||(c.content||[]).length)?'':'empty'}">${cellHtml(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}
function imageHtml(b){ const urls=(b.image_urls&&b.image_urls.length)?b.image_urls:(b.image_url?[b.image_url]:[]); const paths=(b.image_paths&&b.image_paths.length)?b.image_paths:(b.image_path?[b.image_path]:[]); if(!urls.length) return '<div class="empty">이미지 객체는 감지됐지만 연결된 이미지 파일이 없습니다. 아래 media gallery를 확인하세요.</div>'; return `<div class="image-box">${urls.map((u,i)=>`<img src="${esc(u)}" alt="${esc(b.id||'image')}"><div class="small">${esc(paths[i]||'')}</div>`).join('')}</div>`;}
function blockText(b){return [b.text, b.caption&&b.caption.text, b.image_path, JSON.stringify(b.attrs||{})].filter(Boolean).join(' ')}
function withCaption(b, content){const cap=captionHtml(b); if(!cap) return content; const typ=String((b&&b.type)||'').toLowerCase(); const pos=String((b.caption&&b.caption.position)||'').toLowerCase(); if(['before','top','above','left'].includes(pos)) return cap+content; if(['after','bottom','below','right'].includes(pos) || (!pos && ['table','image'].includes(typ))) return content+cap; return cap+content;}
function blockHtml(b){const raw={attrs:b.attrs||{},raw:b.raw||{},geometry:b.geometry||{}}; let body=''; if(b.type==='paragraph') body=`<div class="para">${esc(b.text||'')}</div>`; else if(b.type==='table') body=withCaption(b,tableHtml(b.rows)); else if(b.type==='image') body=withCaption(b,imageHtml(b)); else if(b.type==='equation') body=withCaption(b,`<div class="para">${esc(b.text||'')}</div>`); else if(b.type==='caption') body=`<div class="caption"><b>unresolved caption</b> ${esc(b.text||'')}</div>`; else body=`<div class="para">${esc(b.text||'')}</div>`; return `<div class="block" data-type="${esc(b.type)}" data-text="${esc(blockText(b))}"><div class="block-head"><span><span class="block-title">${esc(b.type)}</span> order=${esc(b.order)} · id=${esc(b.id)}</span><span class="badge">${esc(b.section!==undefined?'section '+b.section:'')}</span></div>${body}<details class="raw"><summary>raw / attrs 보기</summary><pre>${pretty(raw)}</pre></details></div>`;}
function mediaGallery(doc){const items=doc.media_files||[]; if(!items.length) return '<div class="empty">추출된 media 파일이 없습니다.</div>'; return `<div class="media-grid">${items.map(m=>`<div class="media-card">${m.media_type&&m.media_type.startsWith('image/')?`<img src="${esc(m.url)}">`:'<div class="empty">binary</div>'}<div class="small"><b>${esc(m.name)}</b><br>${esc(m.media_type||'')}<br>${esc(m.source||'')}<br>${m.linked?'linked':'unlinked'}</div></div>`).join('')}</div>`;}
function originalPane(doc, assets){if(assets&&assets.preview_image_url){return `<div class="notice info">왼쪽은 HWP 내부 PrvImage 기반 첫 페이지 원본 미리보기만 표시합니다. PrvText와 파서 재구성 블록은 중복 검증을 피하기 위해 제거했습니다.</div><img class="preview-image" src="${esc(assets.preview_image_url)}">`;} return `<div class="notice warn">이 HWP에는 표시 가능한 PrvImage 첫 페이지 미리보기가 없습니다. 원본 전체 페이지 렌더링은 한컴 렌더링 엔진이 필요하며, 이 파일은 추출 결과 패널에서만 검증할 수 있습니다.</div>`;}
function renderViewer(){const doc=APP_STATE.doc; const assets=APP_STATE.original_assets||{}; const c=counts(doc.blocks||[]); const media=doc.media_files||[]; const app=document.getElementById('app'); app.innerHTML=`<header><div><div class="title">HWP Parser Verification UI v19</div><div class="subtitle">${esc(doc.source_path)} · method=${esc(doc.method)}</div></div><div class="toolbar"><a class="button" href="/download/json">result.json</a><a class="button primary" href="/new">새 파일 업로드</a></div></header><main><section class="pane"><div class="pane-head"><div class="pane-title">왼쪽: HWP 내부 첫 페이지 미리보기</div><span class="badge">PrvImage</span></div><div class="pane-body">${originalPane(doc,assets)}</div></section><section class="pane"><div class="pane-head"><div class="pane-title">오른쪽: 추출 결과 시각화</div><span class="badge">parsed</span></div><div class="pane-body" id="rightBody"></div></section></main>`; const nonemptyTables=(doc.blocks||[]).filter(b=>b.type==='table'&&b.rows&&b.rows.length).length; const allImages=countImagesAll(doc.blocks||[]); const linkedImages=countLinkedImages(doc.blocks||[]); const captions=countCaptionsAll(doc.blocks||[]); const right=document.getElementById('rightBody'); const metric=`<div class="meta-grid"><div class="metric"><div class="num">${doc.blocks.length}</div><div class="label">blocks</div></div><div class="metric"><div class="num">${c.table||0}/${nonemptyTables}</div><div class="label">tables / nonempty</div></div><div class="metric"><div class="num">${allImages}/${linkedImages}</div><div class="label">images / linked</div></div><div class="metric"><div class="num">${captions}</div><div class="label">captions</div></div><div class="metric"><div class="num">${media.length}</div><div class="label">media</div></div></div>`; const warnings=(doc.warnings||[]).map(w=>`<div class="notice warn">${esc(w)}</div>`).join(''); const errors=(doc.errors||[]).map(e=>`<div class="notice err">${esc(e)}</div>`).join(''); const controls=`<div class="controls"><select id="typeFilter"><option value="all">전체 블록</option><option value="paragraph">문단</option><option value="table">표</option><option value="image">이미지</option><option value="caption">미해결 캡션</option><option value="equation">수식</option></select><input id="searchBox" type="text" placeholder="텍스트/캡션 검색" style="flex:1;min-width:180px"><button onclick="expandRaw(false)">raw 접기</button><button onclick="expandRaw(true)">raw 펼치기</button></div>`; right.innerHTML=metric+warnings+errors+controls+`<div id="blocks">${(doc.blocks||[]).map(blockHtml).join('')||'<div class="empty">추출된 블록이 없습니다.</div>'}</div><h3>추출 media gallery</h3>${mediaGallery(doc)}`; document.getElementById('typeFilter').addEventListener('change',applyFilter); document.getElementById('searchBox').addEventListener('input',applyFilter);}
function applyFilter(){const t=document.getElementById('typeFilter').value; const q=document.getElementById('searchBox').value.toLowerCase().trim(); for(const el of document.querySelectorAll('.block')){const okType=t==='all'||el.dataset.type===t; const okText=!q||(el.dataset.text||'').toLowerCase().includes(q)||el.innerText.toLowerCase().includes(q); el.style.display=okType&&okText?'':'none';}}
function expandRaw(open){for(const d of document.querySelectorAll('details.raw')) d.open=open;}
function renderUpload(){document.getElementById('app').innerHTML=`<div class="upload-wrap"><form class="upload-card" method="post" action="/upload" enctype="multipart/form-data"><h1>HWP Parser Verification UI v19</h1><p>HWP 파일을 업로드하면 왼쪽에는 HWP 내부 첫 페이지 미리보기, 오른쪽에는 추출된 문단·표·이미지·캡션을 시각화합니다.</p><div class="dropzone"><b>.hwp 또는 intermediate .xml 파일 선택</b><br><input type="file" name="hwp_file" accept=".hwp,.xml" required></div><div class="controls"><select name="mode"><option value="auto">auto: pyhwp XML 우선 + binary fallback</option><option value="pyhwp">pyhwp XML only</option><option value="binary">binary fallback only</option><option value="xml">intermediate XML 직접 파싱</option></select><button class="primary" type="submit">파싱 및 검증 UI 열기</button></div><div class="notice">권장 설치: <code>pip install olefile lxml pyhwp</code>. LibreOffice는 사용하지 않습니다.</div>${APP_STATE.last_error?`<div class="notice err">${esc(APP_STATE.last_error)}</div>`:''}</form></div>`;}
if(APP_STATE && APP_STATE.doc) renderViewer(); else renderUpload();
</script>
</body>
</html>'''


def _safe_json_for_script(data: Dict[str, Any]) -> str:
    return json.dumps(data, ensure_ascii=False).replace("</", "<\\/")


def _file_url(path: Optional[Union[str, Path]]) -> Optional[str]:
    if not path:
        return None
    return "/file?path=" + quote(str(Path(path).resolve()))


def _is_under(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except Exception:
        return False


def _augment_doc_for_web(doc: ParsedDocument) -> Dict[str, Any]:
    d = doc.to_dict()

    def augment_item(item: Any) -> None:
        if not isinstance(item, dict):
            return
        if item.get("image_path"):
            item["image_url"] = _file_url(item.get("image_path"))
        for child in item.get("children", []) or []:
            augment_item(child)
        for row in item.get("rows", []) or []:
            for cell in row or []:
                if isinstance(cell, dict):
                    for inline in cell.get("content", []) or []:
                        augment_item(inline)
                    for nested_row in cell.get("rows", []) or []:
                        for nested_cell in nested_row or []:
                            if isinstance(nested_cell, dict):
                                for inline in nested_cell.get("content", []) or []:
                                    augment_item(inline)

    for b in d.get("blocks", []):
        augment_item(b)
    media_out = []
    for item in doc.media_items:
        x = item.to_dict()
        x["url"] = _file_url(item.path)
        media_out.append(x)
    d["media_files"] = media_out
    return d
def _augment_assets_for_web(preview: Dict[str, Any]) -> Dict[str, Any]:
    out = dict(preview or {})
    if out.get("preview_image"):
        out["preview_image_url"] = _file_url(out["preview_image"])
    return out


class HwpWebState:
    def __init__(self, base_output_dir: Union[str, Path], hwp5proc_path: str = "hwp5proc"):
        self.base_output_dir = ensure_dir(base_output_dir)
        self.upload_dir = ensure_dir(self.base_output_dir / "uploads")
        self.hwp5proc_path = hwp5proc_path
        self.current_doc: Optional[ParsedDocument] = None
        self.current_output_dir: Optional[Path] = None
        self.current_input: Optional[Path] = None
        self.current_preview: Dict[str, Any] = {}
        self.last_error: Optional[str] = None

    def parse_input(self, input_path: Union[str, Path], mode: str = "auto", keep_intermediate: bool = True) -> ParsedDocument:
        input_path = Path(input_path)
        out_dir = ensure_dir(self.base_output_dir / f"{input_path.stem}_{uuid.uuid4().hex[:8]}")
        parser = FullHwpParser(input_path, out_dir, mode=mode, keep_intermediate=keep_intermediate, hwp5proc_path=self.hwp5proc_path)
        doc = parser.parse()
        write_json(doc, out_dir / "result.json", pretty=True)
        self.current_doc = doc
        self.current_output_dir = out_dir
        self.current_input = input_path
        self.current_preview = parser.preview
        self.last_error = None
        return doc

    def clear(self) -> None:
        self.current_doc = None
        self.current_output_dir = None
        self.current_input = None
        self.current_preview = {}
        self.last_error = None

    def allowed_roots(self) -> List[Path]:
        roots = [self.base_output_dir]
        if self.current_output_dir:
            roots.append(self.current_output_dir)
        if self.current_input:
            roots.append(self.current_input.parent)
        return [r.resolve() for r in roots]

    def app_state(self) -> Dict[str, Any]:
        if not self.current_doc:
            return {"doc": None, "original_assets": None, "last_error": self.last_error}
        return {"doc": _augment_doc_for_web(self.current_doc), "original_assets": _augment_assets_for_web(self.current_preview), "last_error": self.last_error}


class HwpViewerHandler(BaseHTTPRequestHandler):
    server_version = "HwpParserWeb/2.0"

    @property
    def state(self) -> HwpWebState:
        return self.server.state  # type: ignore[attr-defined]

    def log_message(self, fmt: str, *args: Any) -> None:
        sys.stderr.write("[web] " + fmt % args + "\n")

    def _send(self, status: int, body: bytes, content_type: str = "text/html; charset=utf-8") -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _redirect(self, loc: str) -> None:
        self.send_response(303)
        self.send_header("Location", loc)
        self.end_headers()

    def do_GET(self) -> None:  # noqa: N802
        parsed = urlparse(self.path)
        if parsed.path == "/new":
            self.state.clear()
            self._send(200, self._html())
            return
        if parsed.path == "/":
            self._send(200, self._html())
            return
        if parsed.path == "/api/doc":
            self._send(200, json.dumps(self.state.app_state(), ensure_ascii=False, indent=2).encode("utf-8"), "application/json; charset=utf-8")
            return
        if parsed.path == "/download/json":
            if not self.state.current_doc:
                self._send(404, b"No parsed document", "text/plain; charset=utf-8")
                return
            body = json.dumps(self.state.current_doc.to_dict(), ensure_ascii=False, indent=2).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Disposition", 'attachment; filename="result.json"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if parsed.path == "/file":
            qs = parse_qs(parsed.query)
            raw = qs.get("path", [""])[0]
            try:
                target = Path(unquote(raw)).resolve()
                if not any(_is_under(target, root) for root in self.state.allowed_roots()):
                    self._send(403, b"Forbidden", "text/plain; charset=utf-8")
                    return
                if not target.exists() or not target.is_file():
                    self._send(404, b"Not found", "text/plain; charset=utf-8")
                    return
                ctype = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
                self._send(200, target.read_bytes(), ctype)
            except Exception as e:
                self._send(500, str(e).encode("utf-8"), "text/plain; charset=utf-8")
            return
        self._send(404, b"Not found", "text/plain; charset=utf-8")

    def _html(self) -> bytes:
        return VIEWER_HTML.replace("__APP_STATE__", _safe_json_for_script(self.state.app_state())).encode("utf-8")

    def _read_multipart(self) -> Tuple[Dict[str, str], Dict[str, Dict[str, Any]]]:
        ctype = self.headers.get("Content-Type", "")
        clen = int(self.headers.get("Content-Length", "0") or "0")
        body = self.rfile.read(clen)
        header = f"Content-Type: {ctype}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
        msg = BytesParser(policy=email_policy.default).parsebytes(header + body)
        fields: Dict[str, str] = {}
        files: Dict[str, Dict[str, Any]] = {}
        if not msg.is_multipart():
            return fields, files
        for part in msg.iter_parts():
            if part.get_content_disposition() != "form-data":
                continue
            name = part.get_param("name", header="content-disposition")
            if not name:
                continue
            payload = part.get_payload(decode=True) or b""
            filename = part.get_filename()
            if filename:
                files[name] = {"filename": filename, "data": payload}
            else:
                charset = part.get_content_charset() or "utf-8"
                fields[name] = payload.decode(charset, errors="replace")
        return fields, files

    def do_POST(self) -> None:  # noqa: N802
        if urlparse(self.path).path != "/upload":
            self._send(404, b"Not found", "text/plain; charset=utf-8")
            return
        try:
            fields, files = self._read_multipart()
            f = files.get("hwp_file")
            if not f:
                self._send(400, b"No file uploaded", "text/plain; charset=utf-8")
                return
            filename = Path(str(f["filename"])).name
            target = self.state.upload_dir / f"{uuid.uuid4().hex[:8]}_{filename}"
            target.write_bytes(f.get("data", b""))
            try:
                self.state.parse_input(target, mode=fields.get("mode", "auto"), keep_intermediate=True)
            except Exception as e:
                self.state.last_error = str(e)
            self._redirect("/")
        except Exception as e:
            self._send(500, str(e).encode("utf-8"), "text/plain; charset=utf-8")


def start_web_ui(input_path: Optional[Union[str, Path]], output_dir: Union[str, Path], mode: str, keep_intermediate: bool, hwp5proc_path: str, host: str, port: int, open_browser: bool) -> int:
    state = HwpWebState(output_dir, hwp5proc_path=hwp5proc_path)
    if input_path:
        state.parse_input(input_path, mode=mode, keep_intermediate=keep_intermediate)
    server = ThreadingHTTPServer((host, port), HwpViewerHandler)
    server.state = state  # type: ignore[attr-defined]
    url = f"http://{host}:{server.server_port}/"
    print(f"HWP Parser Verification UI v19: {url}")
    print("LibreOffice is not used. Stop with Ctrl+C.")
    if open_browser:
        threading.Timer(0.8, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopping web server.")
    finally:
        server.server_close()
    return 0




# =============================================================================
# v13 targeted fixes for real HWP picture placement
# =============================================================================

# The uploaded verification file exposed two concrete failure modes:
# 1) HWP PictureInfo uses decimal bindata-id values, while OLE BinData stream
#    names switch to hexadecimal suffixes after BIN0009, e.g. bindata-id="10"
#    corresponds to BIN000A. Older ref resolution missed these images.
# 2) In visual-box tables, TableCell direct children are Paragraph elements; the
#    actual screenshot is a GShapeObjectControl nested inside a paragraph, and
#    its caption is nested inside the same GShapeObjectControl. Generic recursive
#    scanning could skip some picture controls or flatten caption paragraphs.
#    v13 therefore linearizes TableCell content paragraph-by-paragraph and emits
#    image controls at their exact paragraph position.

_HEX_DIGITS_V13 = "0123456789ABCDEF"


def _v13_bin_keys_from_number(n: int) -> List[str]:
    keys: List[str] = []
    if n < 0:
        return keys
    # decimal id forms
    keys.extend([str(n), f"BIN{n}", f"BIN{n:04d}", f"BIN{n:05d}"])
    # HWP OLE stream forms commonly use hex after 9: BIN000A, BIN000B, ...
    hx = format(n, "X")
    keys.extend([f"BIN{hx}", f"BIN{hx:0>4}", f"BIN{hx:0>5}"])
    keys.extend([k.lower() for k in list(keys)])
    return keys


def _v13_parse_bin_number(s: str) -> Optional[int]:
    s = str(s or "").strip()
    if not s:
        return None
    m = re.search(r"BIN\s*([0-9A-Fa-f]+)", s, re.IGNORECASE)
    if m:
        token = m.group(1)
        try:
            # If A-F appears, it is certainly hexadecimal. If only digits,
            # keep the old decimal interpretation because earlier BIN0008 etc.
            # worked that way, but v13 also registers hex aliases below.
            return int(token, 16) if re.search(r"[A-Fa-f]", token) else int(token, 10)
        except Exception:
            return None
    if re.fullmatch(r"[0-9]+", s):
        try:
            return int(s, 10)
        except Exception:
            return None
    if re.fullmatch(r"[0-9A-Fa-f]+", s) and re.search(r"[A-Fa-f]", s):
        try:
            return int(s, 16)
        except Exception:
            return None
    return None


def _v13_register_refs(self: MediaStore, path: str, refs: Sequence[str], extra_names: Sequence[str] = ()) -> None:
    all_refs = set()
    for ref in list(refs) + list(extra_names):
        if ref is None:
            continue
        s = str(ref).strip()
        if not s:
            continue
        variants = {s, s.lower(), s.upper(), Path(s).name, Path(s).stem}
        for v in list(variants):
            if v:
                all_refs.add(v)
                all_refs.add(v.lower())
                all_refs.add(v.upper())
        n = _v13_parse_bin_number(s)
        if n is not None:
            for k in _v13_bin_keys_from_number(n):
                all_refs.add(k)
        # Also recover from filenames such as BIN000A.png.
        stem = Path(s).stem
        n2 = _v13_parse_bin_number(stem)
        if n2 is not None:
            for k in _v13_bin_keys_from_number(n2):
                all_refs.add(k)
    for r in all_refs:
        if r:
            self.ref_map[r] = path


def _v13_resolve(self: MediaStore, ref: Optional[str]) -> Optional[str]:
    if not ref:
        return None
    candidates: List[str] = []
    s = str(ref).strip()
    candidates.extend([s, s.strip("#"), Path(s).name, Path(s).stem, s.lower(), s.upper()])
    n = _v13_parse_bin_number(s)
    if n is not None:
        candidates.extend(_v13_bin_keys_from_number(n))
    # If a value is decimal 10, explicitly try hex stream names BIN000A.
    if s.isdigit():
        try:
            n2 = int(s, 10)
            candidates.extend(_v13_bin_keys_from_number(n2))
        except Exception:
            pass
    seen = set()
    for c in candidates:
        if not c or c in seen:
            continue
        seen.add(c)
        if c in self.ref_map:
            return self.ref_map[c]
    ref_l = Path(s).name.lower()
    for k, v in self.ref_map.items():
        if ref_l and (Path(k).name.lower() == ref_l or ref_l in Path(k).name.lower()):
            return v
    return None


MediaStore._register_refs = _v13_register_refs  # type: ignore[method-assign]
MediaStore.resolve = _v13_resolve  # type: ignore[method-assign]


def _v13_picture_refs_from_element(parser: PyHwpXmlParser, elem: Any) -> List[str]:
    refs: List[str] = []
    candidates = [elem] + list(parser._descendants(elem))
    for d in candidates:
        attrs = attrs_to_dict(d)
        tag = norm_name(getattr(d, "tag", ""))
        # Exact HWPML/pyhwp picture reference.
        if tag == "pictureinfo" and attrs.get("bindata-id") not in (None, ""):
            refs.append(str(attrs.get("bindata-id")))
        for k, v in attrs.items():
            lk = str(k).lower()
            sv = str(v).strip()
            if not sv:
                continue
            if lk in {"bindata-id", "bindataid", "bin-data-id", "binary-id", "binaryid"}:
                refs.append(sv)
            elif re.search(r"(bindata|bin-data|binary|picture|image|src|href|file|path)", lk, re.I) and re.search(r"BIN\s*[0-9A-Fa-f]+|^[0-9]+$|\.(png|jpg|jpeg|gif|bmp|tif|tiff|webp|wmf|emf)\b", sv, re.I):
                refs.append(sv)
    # unique preserving order
    out: List[str] = []
    seen = set()
    for r in refs:
        if r not in seen:
            seen.add(r)
            out.append(r)
    return out


def _v13_is_real_picture_control(parser: PyHwpXmlParser, elem: Any) -> bool:
    tag = norm_name(getattr(elem, "tag", ""))
    attrs = {k.lower(): str(v).lower() for k, v in attrs_to_dict(elem).items()}
    if tag in {"gshapeobjectcontrol", "picturecontrol", "imagecontrol", "pic", "picture", "image"}:
        if attrs.get("number-category") in {"figure", "image", "picture"}:
            return True
        if attrs.get("chid", "").strip().replace("$", "") in {"gso", "pic", "picture", "pict"}:
            return True
        if _v13_picture_refs_from_element(parser, elem):
            return True
    # ShapeComponent with chid=$pic is not used as a top-level content item;
    # it belongs to the enclosing GShapeObjectControl.
    return False


def _v13_top_picture_controls(parser: PyHwpXmlParser, elem: Any) -> List[Any]:
    out: List[Any] = []
    def rec(node: Any) -> None:
        for child in parser._children(node):
            if _v13_is_real_picture_control(parser, child):
                out.append(child)
                continue
            # Do not traverse into nested tables as part of this paragraph.
            if child is not elem and parser._is_table(child) and not parser._is_table_body(child):
                continue
            rec(child)
    rec(elem)
    return _unique_by_id(out)


def _v13_image_content_from_element(self: PyHwpXmlParser, elem: Any, section: Optional[int] = None) -> Dict[str, Any]:
    attrs = attrs_to_dict(elem)
    refs = _v13_picture_refs_from_element(self, elem)
    paths: List[str] = []
    for ref in refs:
        path = self.media_store.resolve(ref)
        if path and path not in paths:
            paths.append(path)
    # Structural embedded base64 fallback only, not arbitrary BinData fallback.
    if not paths:
        embedded_path = self._embedded_image_by_node_id.get(id(elem))
        if not embedded_path:
            for d in self._descendants(elem):
                embedded_path = self._embedded_image_by_node_id.get(id(d))
                if embedded_path:
                    break
        if embedded_path:
            paths.append(embedded_path)
    for path in paths:
        self.media_store.mark_linked(path)
    image_path = paths[0] if paths else None
    media_type = None
    if image_path:
        try:
            _, media_type = detect_extension(Path(image_path).read_bytes())
        except Exception:
            media_type = mimetypes.guess_type(image_path)[0]
    cap = self._extract_caption(elem, "image")
    item: Dict[str, Any] = {
        "type": "image",
        "id": unique_id("imgin"),
        "image_path": image_path,
        "media_type": media_type,
        "caption": cap.to_dict() if cap else None,
        "geometry": self._extract_geometry(attrs),
        "attrs": attrs,
        "raw": {
            "xml_tag": local_name(getattr(elem, "tag", "")),
            "binary_refs": refs,
            "binary_ref": refs[0] if refs else None,
            "linked": bool(image_path),
            "inline_in_table_cell": True,
            "v13_structural_picture_control": True,
        },
    }
    if len(paths) > 1:
        item["image_paths"] = paths
    return item


def _v13_extract_inline_content(self: PyHwpXmlParser, elem: Any, section: Optional[int] = None) -> List[Dict[str, Any]]:
    content: List[Dict[str, Any]] = []

    def add_para(text: str, raw: Optional[Dict[str, Any]] = None) -> None:
        text = normalize_space(text)
        if not self._valid_text_block(text):
            return
        if content and content[-1].get("type") == "paragraph" and normalize_space(str(content[-1].get("text", ""))) == text:
            return
        content.append({"type": "paragraph", "text": text, "raw": raw or {}})

    children = self._children(elem)
    # v13 fast path: HWP table cells usually contain Paragraph children directly.
    if children and any(self._is_para(ch) for ch in children):
        for child in children:
            if self._is_para(child):
                # Ordinary paragraph text, excluding object bodies/captions.
                txt = self._text_of(child, exclude_objects=True)
                add_para(txt, {"xml_tag": local_name(getattr(child, "tag", "")), "attrs": attrs_to_dict(child), "v13_direct_cell_paragraph": True})
                # Insert picture controls at the exact paragraph position.
                for pic in _v13_top_picture_controls(self, child):
                    content.append(_v13_image_content_from_element(self, pic, section=section))
                continue
            if child is not elem and self._is_table(child) and not self._is_table_body(child):
                try:
                    nested = self._parse_table(child, section, {"xml_path": ()})
                    content.append({
                        "type": "table",
                        "text": self._content_plain_text(_flatten_table_content(nested.rows or [])),
                        "rows": [[c.to_dict() for c in row] for row in (nested.rows or [])],
                        "caption": nested.caption.to_dict() if nested.caption else None,
                        "raw": nested.raw,
                    })
                except Exception:
                    add_para(self._text_of(child, exclude_objects=True), {"note": "nested table fallback text", "v13": True})
                continue
            if _v13_is_real_picture_control(self, child):
                content.append(_v13_image_content_from_element(self, child, section=section))
                continue
            if self._is_caption_control(child):
                cap_item = self._caption_content_from_element(child)
                if cap_item:
                    content.append(cap_item)
                continue
            # Fallback for unusual non-paragraph container children.
            sub = _v13_extract_inline_content(self, child, section=section)
            if sub:
                content.extend(sub)
            else:
                add_para(self._text_of(child, exclude_objects=True), {"v13_fallback_child": local_name(getattr(child, "tag", ""))})
    else:
        # Generic fallback: find top-level picture controls and normal text.
        pics = _v13_top_picture_controls(self, elem)
        if pics:
            for pic in pics:
                content.append(_v13_image_content_from_element(self, pic, section=section))
        else:
            txt = self._text_of(elem, exclude_objects=True)
            add_para(txt, {"fallback": "v13_text_of"})

    # Attach caption controls that are adjacent to images, but do not invent
    # images from captions. The real image controls above carry their captions.
    content = self._attach_inline_captions(content)
    return content


PyHwpXmlParser._image_content_from_element = _v13_image_content_from_element  # type: ignore[method-assign]
PyHwpXmlParser._extract_inline_content = _v13_extract_inline_content  # type: ignore[method-assign]


# Web output augmentation for image groups.
_old_augment_doc_for_web_v13 = _augment_doc_for_web

def _augment_doc_for_web(doc: ParsedDocument) -> Dict[str, Any]:  # type: ignore[no-redef]
    d = _old_augment_doc_for_web_v13(doc)
    def augment_item(item: Any) -> None:
        if not isinstance(item, dict):
            return
        if item.get("image_paths"):
            item["image_urls"] = [_file_url(p) for p in item.get("image_paths") or [] if p]
        for child in item.get("children", []) or []:
            augment_item(child)
        for row in item.get("rows", []) or []:
            for cell in row or []:
                if isinstance(cell, dict):
                    for inline in cell.get("content", []) or []:
                        augment_item(inline)
                    for nested_row in cell.get("rows", []) or []:
                        for nested_cell in nested_row or []:
                            if isinstance(nested_cell, dict):
                                for inline in nested_cell.get("content", []) or []:
                                    augment_item(inline)
    for b in d.get("blocks", []) or []:
        augment_item(b)
    return d



# =============================================================================
# v16 rollback and binary picture placement fixes
# =============================================================================
# v16 intentionally does NOT use Hancom COM conversion or HWPX conversion.


def _v16_bin_refs_from_picture_payload(payload: bytes, media_store: MediaStore) -> List[str]:
    refs: List[str] = []
    text = payload[:512].decode("latin1", errors="ignore")
    for m in re.finditer(r"BIN[0-9A-Fa-f]{1,6}|BinData[/\\][^\x00\s]+", text, re.I):
        refs.append(m.group(0))
    max_id = 0
    for item in media_store.items:
        k = natural_bin_sort_key(item.name)[0]
        if isinstance(k, int) and k < 10**8:
            max_id = max(max_id, k)
    for off in (70, 72):
        if off + 2 <= len(payload):
            v = int.from_bytes(payload[off:off + 2], "little", signed=False)
            candidates: List[int] = []
            if v >= 256 and (v & 0xFF) == 0:
                candidates.append(v >> 8)
            if 1 <= v <= max_id:
                candidates.append(v)
            for n in candidates:
                if 1 <= n <= max_id:
                    # HWP OLE BinData stream suffixes are hexadecimal.
                    # 10 -> BIN000A, 16 -> BIN0010, 17 -> BIN0011.
                    # Use only the exact hexadecimal OLE stream stem.
                    # Do not include decimal aliases such as "10", because the
                    # media store also has BIN0010 and alias expansion can
                    # attach the wrong image.
                    exact = [f"BIN{n:04X}", f"BIN{n:04X}".lower()]
                    refs.extend(exact)
                    break
            if refs:
                break
    out: List[str] = []
    seen = set()
    for ref in refs:
        if ref in seen:
            continue
        seen.add(ref)
        if media_store.resolve(ref):
            out.append(ref)
    return out


_bin_refs_from_picture_payload = _v16_bin_refs_from_picture_payload  # type: ignore[assignment]


# Exact filename resolution must beat alias expansion. Without this, aliases such
# as decimal 10 -> BIN000A can be overwritten by later BIN0010 registration.
def _v16_resolve(self: MediaStore, ref: Optional[str]) -> Optional[str]:
    if not ref:
        return None
    s = str(ref).strip()
    if not s:
        return None
    names = {s, s.strip('#'), Path(s).name, Path(s).stem}
    names |= {x.lower() for x in list(names)}
    for item in self.items:
        item_name = Path(item.path).name
        item_stem = Path(item.path).stem
        variants = {item.name, item_name, item_stem, item.name.lower(), item_name.lower(), item_stem.lower()}
        if names & variants:
            return item.path
    try:
        return _v13_resolve(self, ref)
    except Exception:
        return None


MediaStore.resolve = _v16_resolve  # type: ignore[method-assign]


def _v16_ctrl_id(payload: bytes) -> str:
    if len(payload) < 4:
        return ""
    raw = payload[:4]
    for b in (raw[::-1], raw):
        try:
            s = b.decode("ascii", errors="ignore").strip().replace("$", "")
        except Exception:
            s = ""
        if s:
            return s
    return ""


def _v16_caption_for_image_from_records(parser: "BinaryHwpParser", records: Sequence[HwpRecord]) -> Optional[Caption]:
    for rec in records:
        if rec.tag_id == HWPTAG_PARA_TEXT:
            text, _ = parser.decode_para_text(rec.payload)
            cap = caption_from_text(normalize_space(text), "image")
            if cap:
                cap.method = "binary-gso-structural"
                return cap
    return None


def _v16_mark_path_linked(media_store: MediaStore, path: str) -> None:
    media_store.mark_linked(path)


def _v16_image_group_block_from_picture_records(parser: "BinaryHwpParser", pic_records: Sequence[HwpRecord], section: Optional[int], caption: Optional[Caption] = None, raw_extra: Optional[Dict[str, Any]] = None) -> Block:
    paths: List[str] = []
    refs_all: List[str] = []
    first_rec = pic_records[0] if pic_records else None
    for rec in pic_records:
        refs = _v16_bin_refs_from_picture_payload(rec.payload, parser.media_store)
        refs_all.extend(refs)
        resolved_any = False
        for ref in refs:
            path = parser.media_store.resolve(ref)
            if not path or path in paths:
                continue
            if Path(path).name.lower().startswith("prvimage"):
                continue
            paths.append(path)
            _v16_mark_path_linked(parser.media_store, path)
            resolved_any = True
        if not resolved_any and not refs:
            candidates = [it for it in parser.media_store.figure_candidate_items() if not it.linked]
            if 0 <= parser._picture_event_index < len(candidates):
                path = candidates[parser._picture_event_index].path
            elif candidates:
                path = candidates[0].path
            else:
                path = None
            if path and path not in paths:
                paths.append(path)
                _v16_mark_path_linked(parser.media_store, path)
        parser._picture_event_index += 1
    media_type = None
    if paths:
        try:
            _, media_type = detect_extension(Path(paths[0]).read_bytes())
        except Exception:
            media_type = mimetypes.guess_type(paths[0])[0]
    block = Block("image", unique_id("img"), parser._next_order(), section=section, caption=caption, image_path=paths[0] if paths else None, media_type=media_type, raw={"record": first_rec.to_meta() if first_rec else {}, "picture_refs": refs_all, "grouped_picture_count": len(pic_records), "link_method": "binary-picture-explicit-bin-id-v16" if paths else "unresolved-binary-picture-v16", **(raw_extra or {})})
    if len(paths) > 1:
        setattr(block, "image_paths", paths)
    return block


def _v16_block_to_dict(self: Block) -> Dict[str, Any]:
    result: Dict[str, Any] = {"type": self.type, "id": self.id, "order": self.order}
    if self.section is not None: result["section"] = self.section
    if self.text not in (None, ""): result["text"] = self.text
    if self.caption is not None: result["caption"] = self.caption.to_dict()
    if self.rows is not None: result["rows"] = [[cell.to_dict() for cell in row] for row in self.rows]
    if self.image_path: result["image_path"] = self.image_path
    if hasattr(self, "image_paths"):
        paths = getattr(self, "image_paths") or []
        if paths: result["image_paths"] = paths
    if self.media_type: result["media_type"] = self.media_type
    if self.geometry: result["geometry"] = self.geometry
    if self.attrs: result["attrs"] = self.attrs
    if self.raw: result["raw"] = self.raw
    if self.children: result["children"] = [c.to_dict() for c in self.children]
    return result


Block.to_dict = _v16_block_to_dict  # type: ignore[method-assign]


def _v16_parse_gso_group(self: "BinaryHwpParser", records: List[HwpRecord], start: int, section: Optional[int]) -> Tuple[Optional[Block], int]:
    ctrl = records[start]
    group: List[HwpRecord] = [ctrl]
    j = start + 1
    while j < len(records):
        r = records[j]
        if r.level <= ctrl.level and j > start + 1:
            break
        group.append(r)
        j += 1
    pics = [r for r in group if r.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE]
    if not pics:
        return None, j
    caption = _v16_caption_for_image_from_records(self, group)
    block = _v16_image_group_block_from_picture_records(self, pics, section, caption, raw_extra={"gso_group_start": ctrl.to_meta(), "binary_gso_group_v16": True})
    return block, j


def _v16_parse_record_list(self: "BinaryHwpParser", records: List[HwpRecord], section: Optional[int]) -> None:
    i = 0
    n = len(records)
    while i < n:
        rec = records[i]
        if rec.tag_id == HWPTAG_TABLE:
            block, next_i = self._parse_table_group(records, i, section)
            self.blocks.append(block)
            i = max(next_i, i + 1); continue
        if rec.tag_id == HWPTAG_CTRL_HEADER and _v16_ctrl_id(rec.payload) == "gso":
            block, next_i = self._parse_gso_group(records, i, section)
            if block:
                self.blocks.append(block)
                i = max(next_i, i + 1); continue
        if rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
            self.blocks.append(self._image_block_from_picture(rec, section))
            i += 1; continue
        if rec.tag_id == HWPTAG_EQEDIT:
            txt = self._decode_bytes_text(rec.payload)
            if normalize_space(txt):
                self.blocks.append(Block("equation", unique_id("eq"), self._next_order(), section=section, text=normalize_space(txt), raw={"record": rec.to_meta()}))
            i += 1; continue
        if rec.tag_id == HWPTAG_PARA_TEXT:
            text, ctrls = self.decode_para_text(rec.payload)
            text = normalize_space(text)
            if text and not self._is_control_only_text(text):
                self.blocks.append(Block("paragraph", unique_id("p"), self._next_order(), section=section, text=text, raw={"record": rec.to_meta(), "controls": ctrls[:20]}))
        i += 1


def _v16_parse_cell_records(self: "BinaryHwpParser", cell_records: List[HwpRecord], cell_meta: Dict[str, int], section: Optional[int]) -> Tuple[TableCell, List[Dict[str, Any]]]:
    parts: List[str] = []
    paras: List[str] = []
    mixed: List[Dict[str, Any]] = []
    i = 0
    while i < len(cell_records):
        rec = cell_records[i]
        if rec.tag_id == HWPTAG_CTRL_HEADER and _v16_ctrl_id(rec.payload) == "gso":
            block, next_i = self._parse_gso_group(cell_records, i, section)
            if block:
                item: Dict[str, Any] = {"type": "image", "image_path": block.image_path, "media_type": block.media_type, "caption": block.caption.to_dict() if block.caption else None, "raw": block.raw}
                if hasattr(block, "image_paths"): item["image_paths"] = getattr(block, "image_paths")
                mixed.append(item)
                i = max(next_i, i + 1); continue
        if rec.tag_id == HWPTAG_PARA_TEXT:
            text, ctrls = self.decode_para_text(rec.payload)
            text = normalize_space(text)
            if not text or self._is_control_only_text(text): i += 1; continue
            cap_img = caption_from_text(text, "image")
            cap_tbl = caption_from_text(text, "table")
            if cap_img:
                attached = False
                for item in reversed(mixed):
                    if item.get("type") == "image" and not item.get("caption"):
                        item["caption"] = cap_img.to_dict(); attached = True; break
                if not attached: mixed.append({"type": "caption", "text": cap_img.text, "caption": cap_img.to_dict(), "target_type": "image"})
                parts.append(cap_img.text); paras.append(cap_img.text); i += 1; continue
            if cap_tbl:
                mixed.append({"type": "caption", "text": cap_tbl.text, "caption": cap_tbl.to_dict(), "target_type": "table"})
                parts.append(cap_tbl.text); paras.append(cap_tbl.text); i += 1; continue
            parts.append(text); paras.append(text); mixed.append({"type": "paragraph", "text": text})
        elif rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
            image_block = self._image_block_from_picture(rec, section)
            item = {"type": "image", "image_path": image_block.image_path, "media_type": image_block.media_type, "raw": image_block.raw}
            if hasattr(image_block, "image_paths"): item["image_paths"] = getattr(image_block, "image_paths")
            mixed.append(item)
        elif rec.tag_id == HWPTAG_TABLE:
            mixed.append({"type": "nested_table_marker", "record": rec.to_meta()})
        i += 1
    cell = TableCell(text=normalize_space("\n".join(parts)), row=cell_meta.get("row", 0), col=cell_meta.get("col", 0), row_span=max(1, int(cell_meta.get("row_span", 1))), col_span=max(1, int(cell_meta.get("col_span", 1))), attrs={"binary_cell_meta": cell_meta}, paragraphs=_dedupe_adjacent_strings(paras), content=mixed)
    return cell, mixed


def _v16_image_block_from_picture(self: "BinaryHwpParser", rec: HwpRecord, section: Optional[int]) -> Block:
    return _v16_image_group_block_from_picture_records(self, [rec], section, None)


BinaryHwpParser._parse_record_list = _v16_parse_record_list  # type: ignore[method-assign]
BinaryHwpParser._parse_gso_group = _v16_parse_gso_group  # type: ignore[attr-defined]
BinaryHwpParser._parse_cell_records = _v16_parse_cell_records  # type: ignore[method-assign]
BinaryHwpParser._image_block_from_picture = _v16_image_block_from_picture  # type: ignore[method-assign]



# =============================================================================
# v17 nested table support for binary fallback
# =============================================================================
# The v16 binary fallback could reconstruct flat HWP tables, but when a table
# cell contained another TABLE record it emitted only a nested_table_marker and
# then allowed the nested table's cell text to leak into the parent cell as flat
# paragraphs. v17 recursively parses TABLE records encountered inside cell
# record spans and stores them as inline table content so that table-in-table
# layouts are rendered as nested HTML tables in the verification UI.


def _v17_cell_text_from_rows(rows):
    parts = []
    for row in rows or []:
        row_parts = []
        for cell in row or []:
            if getattr(cell, "content", None):
                sub_parts = []
                for item in cell.content or []:
                    if not isinstance(item, dict):
                        continue
                    typ = item.get("type")
                    if typ in {"paragraph", "caption"} and item.get("text"):
                        sub_parts.append(str(item.get("text")))
                    elif typ == "image":
                        cap = item.get("caption") or {}
                        if isinstance(cap, dict) and cap.get("text"):
                            sub_parts.append(str(cap.get("text")))
                    elif typ == "table":
                        sub_parts.append(_v17_cell_text_from_rows(_v17_rows_from_inline_table_item(item)))
                txt = normalize_space(" ".join(sub_parts))
            else:
                txt = normalize_space(getattr(cell, "text", "") or "")
            if txt:
                row_parts.append(txt)
        if row_parts:
            parts.append(" | ".join(row_parts))
    return normalize_space("\n".join(parts))


def _v17_rows_from_inline_table_item(item):
    rows = []
    for r in item.get("rows") or []:
        row = []
        for c in r or []:
            if isinstance(c, TableCell):
                row.append(c)
            elif isinstance(c, dict):
                row.append(TableCell(
                    text=str(c.get("text", "") or ""),
                    row=safe_int(c.get("row"), None),
                    col=safe_int(c.get("col"), None),
                    row_span=max(1, int(safe_int(c.get("row_span"), 1) or 1)),
                    col_span=max(1, int(safe_int(c.get("col_span"), 1) or 1)),
                    attrs=c.get("attrs") or {},
                    paragraphs=c.get("paragraphs") or [],
                    content=c.get("content") or [],
                ))
        rows.append(row)
    return rows


def _v17_inline_table_item_from_block(block):
    rows_dict = [[cell.to_dict() for cell in row] for row in (block.rows or [])]
    text = _v17_cell_text_from_rows(block.rows or [])
    item = {
        "type": "table",
        "id": block.id,
        "text": text,
        "rows": rows_dict,
        "caption": block.caption.to_dict() if block.caption else None,
        "attrs": block.attrs,
        "raw": {**(block.raw or {}), "inline_nested_table_v18": True},
    }
    return _drop_empty(item)


def _v17_parse_cell_records(self, cell_records, cell_meta, section):
    parts = []
    paras = []
    mixed = []
    i = 0
    while i < len(cell_records):
        rec = cell_records[i]

        if rec.tag_id == HWPTAG_TABLE:
            try:
                nested_block, next_i = self._parse_table_group(cell_records, i, section)
                nested_item = _v17_inline_table_item_from_block(nested_block)
                mixed.append(nested_item)
                nested_text = normalize_space(nested_item.get("text", ""))
                if nested_text:
                    parts.append(nested_text)
                    paras.append(nested_text)
                i = max(next_i, i + 1)
                continue
            except Exception as e:
                mixed.append({"type": "table", "rows": [], "raw": {"record": rec.to_meta(), "nested_table_parse_error_v17": str(e)}})
                i += 1
                continue

        if rec.tag_id == HWPTAG_CTRL_HEADER and _v16_ctrl_id(rec.payload) == "gso":
            block, next_i = self._parse_gso_group(cell_records, i, section)
            if block:
                item = {
                    "type": "image",
                    "image_path": block.image_path,
                    "media_type": block.media_type,
                    "caption": block.caption.to_dict() if block.caption else None,
                    "raw": block.raw,
                }
                if hasattr(block, "image_paths"):
                    item["image_paths"] = getattr(block, "image_paths")
                mixed.append(item)
                cap_txt = (item.get("caption") or {}).get("text") if isinstance(item.get("caption"), dict) else None
                if cap_txt:
                    parts.append(str(cap_txt)); paras.append(str(cap_txt))
                i = max(next_i, i + 1)
                continue

        if rec.tag_id == HWPTAG_PARA_TEXT:
            text, ctrls = self.decode_para_text(rec.payload)
            text = normalize_space(text)
            if not text or self._is_control_only_text(text):
                i += 1
                continue
            cap_img = caption_from_text(text, "image")
            cap_tbl = caption_from_text(text, "table")
            if cap_img:
                attached = False
                for item in reversed(mixed):
                    if item.get("type") == "image" and not item.get("caption"):
                        item["caption"] = cap_img.to_dict(); attached = True; break
                if not attached:
                    mixed.append({"type": "caption", "text": cap_img.text, "caption": cap_img.to_dict(), "target_type": "image"})
                parts.append(cap_img.text); paras.append(cap_img.text); i += 1; continue
            if cap_tbl:
                attached = False
                for item in reversed(mixed):
                    if item.get("type") == "table" and not item.get("caption"):
                        item["caption"] = cap_tbl.to_dict(); attached = True; break
                    if item.get("type") not in {"paragraph", "caption"}:
                        break
                if not attached:
                    mixed.append({"type": "caption", "text": cap_tbl.text, "caption": cap_tbl.to_dict(), "target_type": "table"})
                parts.append(cap_tbl.text); paras.append(cap_tbl.text); i += 1; continue
            parts.append(text); paras.append(text); mixed.append({"type": "paragraph", "text": text})
            i += 1
            continue

        if rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
            image_block = self._image_block_from_picture(rec, section)
            item = {"type": "image", "image_path": image_block.image_path, "media_type": image_block.media_type, "raw": image_block.raw}
            if hasattr(image_block, "image_paths"):
                item["image_paths"] = getattr(image_block, "image_paths")
            mixed.append(item)
            i += 1
            continue

        i += 1

    cell = TableCell(
        text=normalize_space("\n".join(parts)),
        row=cell_meta.get("row", 0),
        col=cell_meta.get("col", 0),
        row_span=max(1, int(cell_meta.get("row_span", 1))),
        col_span=max(1, int(cell_meta.get("col_span", 1))),
        attrs={"binary_cell_meta": cell_meta, "nested_table_aware_v18": True},
        paragraphs=_dedupe_adjacent_strings(paras),
        content=mixed,
    )
    return cell, mixed


BinaryHwpParser._parse_cell_records = _v17_parse_cell_records  # type: ignore[method-assign]

try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v17", "HWP Parser Verification UI v19")
except Exception:
    pass


# =============================================================================
# v19 caption continuation and caption/table relocation fixes
# =============================================================================
# Two additional HWP layout cases are handled here:
# 1) Captions written with Shift+Enter are kept as multi-line captions instead
#    of being truncated at the first physical line.
# 2) In binary fallback, some table captions appear in the record stream before
#    the nested table they visually describe. v19 relocates standalone table
#    captions to the nearest structurally plausible following/previous table,
#    including tables nested inside table cells.


def _v19_caption_dict_from_text(text: str, target_type: Optional[str] = None) -> Optional[Dict[str, Any]]:
    cap = caption_from_text(text or "", target_type)
    if not cap:
        return None
    d = cap.to_dict()
    if target_type:
        d.setdefault("target_type_hint", target_type)
    return d


def _v19_item_caption_text(item: Dict[str, Any]) -> str:
    if not isinstance(item, dict):
        return ""
    if item.get("type") == "caption":
        cap = item.get("caption") or {}
        if isinstance(cap, dict) and cap.get("text"):
            return normalize_space(str(cap.get("text")))
        return normalize_space(str(item.get("text") or ""))
    if item.get("type") == "paragraph":
        return normalize_space(str(item.get("text") or ""))
    return ""


def _v19_caption_target_from_item(item: Dict[str, Any]) -> Optional[str]:
    text = _v19_item_caption_text(item)
    if not text:
        return None
    hint = item.get("target_type") or item.get("target_type_hint")
    if hint in {"table", "image", "equation"}:
        return str(hint)
    return infer_caption_target_type(text, {}, "")


def _v19_caption_obj_from_item(item: Dict[str, Any], target_type: str, method: str, position: str = "after") -> Optional[Dict[str, Any]]:
    text = _v19_item_caption_text(item)
    if not text:
        return None
    cap = caption_from_text(text, target_type)
    if not cap:
        return None
    cap.method = method
    cap.position = position
    d = cap.to_dict()
    raw = d.setdefault("raw", {})
    raw["relocated_v19"] = True
    return d


def _v19_make_caption_item_from_paragraph(item: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not isinstance(item, dict) or item.get("type") != "paragraph":
        return None
    text = normalize_space(str(item.get("text") or ""))
    if not text:
        return None
    hint = infer_caption_target_type(text, {}, "")
    if hint not in {"table", "image", "equation"}:
        return None
    cap = caption_from_text(text, hint)
    if not cap:
        return None
    out = dict(item)
    out["type"] = "caption"
    out["caption"] = cap.to_dict()
    out["text"] = cap.text
    out["target_type"] = hint
    out.setdefault("raw", {})["converted_from_paragraph_caption_v19"] = True
    return out


def _v19_relocate_captions_in_content(content: List[Dict[str, Any]], max_lookahead: int = 14, max_text_chars: int = 9000) -> List[Dict[str, Any]]:
    """Attach standalone captions to inline images/tables when record order is skewed.

    The function is conservative for images, but deliberately more permissive
    for tables because HWP binary fallback often emits the table caption before
    the nested table contents. Attached captions are removed from their original
    position so the UI no longer shows a detached caption far above the table.
    """
    if not content:
        return content

    # First recurse into nested tables and normalize paragraph captions.
    normalized: List[Dict[str, Any]] = []
    for item in content:
        if not isinstance(item, dict):
            continue
        item = dict(item)
        if item.get("type") == "table":
            _v19_relocate_captions_in_table_item(item)
        cap_item = _v19_make_caption_item_from_paragraph(item)
        normalized.append(cap_item if cap_item is not None else item)

    consumed = set()

    def attach_to_item(target_idx: int, cap_idx: int, target_type: str, method: str, position: str = "after") -> bool:
        if target_idx < 0 or target_idx >= len(normalized) or cap_idx in consumed:
            return False
        target = normalized[target_idx]
        if not isinstance(target, dict) or target.get("type") != target_type:
            return False
        if target.get("caption"):
            return False
        cap_obj = _v19_caption_obj_from_item(normalized[cap_idx], target_type, method, position)
        if not cap_obj:
            return False
        target["caption"] = cap_obj
        target.setdefault("raw", {})["caption_attached_from_item_index_v19"] = cap_idx
        consumed.add(cap_idx)
        return True

    for i, item in enumerate(normalized):
        if not isinstance(item, dict) or item.get("type") != "caption":
            continue
        target_type = _v19_caption_target_from_item(item)
        if target_type not in {"table", "image", "equation"}:
            continue

        # Immediate previous object is the normal case for captions rendered below.
        for j in range(i - 1, max(-1, i - 4), -1):
            if j < 0:
                break
            prev = normalized[j]
            if not isinstance(prev, dict):
                continue
            if prev.get("type") == target_type and not prev.get("caption"):
                if attach_to_item(j, i, target_type, "binary-inline-caption-previous-v19", "after"):
                    break
            # A non-empty paragraph between object and caption weakens previous matching.
            if prev.get("type") == "paragraph" and normalize_space(str(prev.get("text") or "")):
                break
        if i in consumed:
            continue

        # Look ahead for table captions that binary record order placed before
        # the actual nested table. Allow paragraph text between them, but cap the
        # distance so unrelated captions are not stolen.
        if target_type == "table":
            text_budget = 0
            for j in range(i + 1, min(len(normalized), i + 1 + max_lookahead)):
                nxt = normalized[j]
                if not isinstance(nxt, dict):
                    continue
                if nxt.get("type") == "caption" and _v19_caption_target_from_item(nxt) == "table":
                    break
                if nxt.get("type") == "paragraph":
                    text_budget += len(str(nxt.get("text") or ""))
                    if text_budget > max_text_chars:
                        break
                    continue
                if nxt.get("type") == "table" and not nxt.get("caption"):
                    attach_to_item(j, i, "table", "binary-table-caption-lookahead-v19", "after")
                    break
                # Images/equations between caption and table usually indicate a
                # different visual object; stop.
                if nxt.get("type") in {"image", "equation"}:
                    break

    return [item for idx, item in enumerate(normalized) if idx not in consumed]


def _v19_relocate_captions_in_table_item(table_item: Dict[str, Any]) -> None:
    rows = table_item.get("rows") or []
    for row in rows:
        if not isinstance(row, list):
            continue
        for cell in row:
            if not isinstance(cell, dict):
                continue
            content = cell.get("content") or []
            if isinstance(content, list) and content:
                cell["content"] = _v19_relocate_captions_in_content(content)


def _v19_relocate_captions_in_block(block: Block) -> None:
    if block.type != "table" or not block.rows:
        return
    for row in block.rows or []:
        for cell in row or []:
            content = getattr(cell, "content", None) or []
            if content:
                cell.content = _v19_relocate_captions_in_content(content)


def _v19_relocate_top_level_table_captions(blocks: List[Block], max_lookahead: int = 10) -> List[Block]:
    if not blocks:
        return blocks
    consumed = set()
    for i, b in enumerate(blocks):
        if b.type not in {"caption", "paragraph"}:
            continue
        text = normalize_space(b.text or (b.caption.text if b.caption else ""))
        cap = caption_from_text(text, "table")
        if not cap:
            continue
        # Prefer a following table because binary fallback may emit caption
        # records before the table object even when visual caption is below it.
        for j in range(i + 1, min(len(blocks), i + 1 + max_lookahead)):
            nb = blocks[j]
            if nb.type == "table" and nb.caption is None:
                cap.method = "binary-top-caption-lookahead-v19"
                cap.position = "after"
                nb.caption = cap
                nb.raw.setdefault("caption_attached_from_top_block_index_v19", i)
                consumed.add(i)
                break
            if nb.type in {"image", "equation"}:
                break
    return [b for idx, b in enumerate(blocks) if idx not in consumed]


_old_binary_postprocess_v19 = BinaryHwpParser._postprocess


def _v19_binary_postprocess(self: "BinaryHwpParser") -> None:
    # Start from the previous postprocess behavior, then repair caption placement
    # inside table-cell mixed content and at top level.
    _old_binary_postprocess_v19(self)
    for b in self.blocks:
        _v19_relocate_captions_in_block(b)
    self.blocks = _v19_relocate_top_level_table_captions(self.blocks)
    for i, b in enumerate(self.blocks, 1):
        b.order = i
    self.metadata["binary_caption_relocation_v19"] = True


BinaryHwpParser._postprocess = _v19_binary_postprocess  # type: ignore[method-assign]

# UI title update.
try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v19", "HWP Parser Verification UI v19")
except Exception:
    pass

# =============================================================================
# v20 table-caption de-duplication and nested-table caption binding fixes
# =============================================================================
# v19 correctly kept Shift+Enter captions, but it could still misplace captions
# in binary fallback documents that use a large outer layout table containing
# ordinary paragraphs and nested data tables. The root causes were:
#   1) the parent table stole the first table caption found anywhere in its cell
#      mixed content via _extract_table_caption_from_mixed();
#   2) lookahead relocation skipped over a table that already had a caption and
#      attached the old caption to a later table, producing duplicated/shifted
#      captions.
# v20 makes table captions one-to-one and never searches past the first plausible
# table target in the same mixed-content stream.


def _v20_caption_key_from_caption_obj(cap: Any) -> str:
    if isinstance(cap, Caption):
        return caption_lookup_key(cap.text)
    if isinstance(cap, dict):
        return caption_lookup_key(str(cap.get("text") or ""))
    return caption_lookup_key(str(cap or ""))


def _v20_item_existing_caption_key(item: Dict[str, Any]) -> str:
    if not isinstance(item, dict):
        return ""
    cap = item.get("caption")
    if cap:
        return _v20_caption_key_from_caption_obj(cap)
    return ""


def _v20_collect_object_caption_keys(content: Sequence[Dict[str, Any]], target_type: Optional[str] = None) -> set:
    keys = set()
    for item in content or []:
        if not isinstance(item, dict):
            continue
        typ = item.get("type")
        if target_type is None or typ == target_type:
            k = _v20_item_existing_caption_key(item)
            if k:
                keys.add(k)
        if typ == "table":
            for row in item.get("rows") or []:
                for cell in row or []:
                    if isinstance(cell, dict):
                        keys.update(_v20_collect_object_caption_keys(cell.get("content") or [], target_type))
    return keys


def _v20_caption_item_key(item: Dict[str, Any]) -> str:
    return caption_lookup_key(_v19_item_caption_text(item))


def _v20_normalize_content_captions(content: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    normalized: List[Dict[str, Any]] = []
    for item in content or []:
        if not isinstance(item, dict):
            continue
        item = dict(item)
        if item.get("type") == "table":
            _v19_relocate_captions_in_table_item(item)
        cap_item = _v19_make_caption_item_from_paragraph(item)
        normalized.append(cap_item if cap_item is not None else item)
    return normalized


def _v19_relocate_captions_in_content(content: List[Dict[str, Any]], max_lookahead: int = 14, max_text_chars: int = 9000) -> List[Dict[str, Any]]:  # type: ignore[no-redef]
    """v20 replacement for v19 caption relocation.

    Differences from v19:
    - If the first plausible following table already has a caption, stop instead
      of skipping it and attaching the caption to a later table.
    - Do not keep standalone duplicate caption items when an object in the same
      content stream already owns the same caption text.
    - Enforce one caption text -> one target table/image/equation within the same
      mixed-content scope.
    """
    if not content:
        return content

    normalized = _v20_normalize_content_captions(content)
    consumed = set()
    used_caption_keys = _v20_collect_object_caption_keys(normalized)

    def attach_to_item(target_idx: int, cap_idx: int, target_type: str, method: str, position: str = "after") -> bool:
        if target_idx < 0 or target_idx >= len(normalized) or cap_idx in consumed:
            return False
        target = normalized[target_idx]
        if not isinstance(target, dict) or target.get("type") != target_type:
            return False
        if target.get("caption"):
            return False
        cap_obj = _v19_caption_obj_from_item(normalized[cap_idx], target_type, method, position)
        if not cap_obj:
            return False
        key = _v20_caption_key_from_caption_obj(cap_obj)
        if key and key in used_caption_keys:
            # The same caption is already attached in this scope. Drop the
            # duplicate standalone caption rather than shifting it to another table.
            consumed.add(cap_idx)
            return False
        target["caption"] = cap_obj
        target.setdefault("raw", {})["caption_attached_from_item_index_v20"] = cap_idx
        if key:
            used_caption_keys.add(key)
        consumed.add(cap_idx)
        return True

    for i, item in enumerate(normalized):
        if not isinstance(item, dict) or item.get("type") != "caption":
            continue
        target_type = _v19_caption_target_from_item(item)
        if target_type not in {"table", "image", "equation"}:
            continue
        this_key = _v20_caption_item_key(item)
        if this_key and this_key in used_caption_keys:
            consumed.add(i)
            continue

        # Immediate previous object is the normal visual case for bottom captions.
        for j in range(i - 1, max(-1, i - 4), -1):
            if j < 0:
                break
            prev = normalized[j]
            if not isinstance(prev, dict):
                continue
            if prev.get("type") == target_type:
                if not prev.get("caption"):
                    attach_to_item(j, i, target_type, "binary-inline-caption-previous-v20", "after")
                # Do not skip over the nearest plausible object, even if it is
                # already captioned; otherwise old captions drift to later tables.
                break
            if prev.get("type") == "paragraph" and normalize_space(str(prev.get("text") or "")):
                break
        if i in consumed:
            continue

        # Look ahead only to the first plausible table/image/equation. If it is
        # already captioned, the current caption is a duplicate or belongs to a
        # structure we cannot safely infer; leave it for duplicate cleanup.
        text_budget = 0
        for j in range(i + 1, min(len(normalized), i + 1 + max_lookahead)):
            nxt = normalized[j]
            if not isinstance(nxt, dict):
                continue
            if nxt.get("type") == "caption" and _v19_caption_target_from_item(nxt) == target_type:
                break
            if nxt.get("type") == "paragraph":
                text_budget += len(str(nxt.get("text") or ""))
                if text_budget > max_text_chars:
                    break
                continue
            if nxt.get("type") == target_type:
                if not nxt.get("caption"):
                    attach_to_item(j, i, target_type, f"binary-{target_type}-caption-lookahead-v20", "after")
                # Always stop at the first compatible object.
                break
            if nxt.get("type") in {"table", "image", "equation"}:
                # A different object type between caption and target means unsafe.
                break

    # Final duplicate cleanup: if a caption item has the same normalized text as
    # any object caption in this scope, remove the detached caption item.
    object_keys = _v20_collect_object_caption_keys(normalized)
    out: List[Dict[str, Any]] = []
    seen_standalone_caption_keys = set()
    for idx, item in enumerate(normalized):
        if idx in consumed:
            continue
        if isinstance(item, dict) and item.get("type") == "caption":
            key = _v20_caption_item_key(item)
            if key and key in object_keys:
                continue
            if key and key in seen_standalone_caption_keys:
                continue
            if key:
                seen_standalone_caption_keys.add(key)
        out.append(item)
    return out


def _v19_relocate_top_level_table_captions(blocks: List[Block], max_lookahead: int = 10) -> List[Block]:  # type: ignore[no-redef]
    """v20 replacement for top-level table-caption relocation.

    Stop at the first table candidate even when it already has a caption, and
    remove duplicate detached caption paragraphs when their caption text is
    already owned by a table block.
    """
    if not blocks:
        return blocks
    consumed = set()
    used_caption_keys = set()
    for b in blocks:
        if b.type == "table" and b.caption:
            k = caption_lookup_key(b.caption.text)
            if k:
                used_caption_keys.add(k)

    for i, b in enumerate(blocks):
        if b.type not in {"caption", "paragraph"}:
            continue
        text = normalize_space(b.text or (b.caption.text if b.caption else ""))
        cap = caption_from_text(text, "table")
        if not cap:
            continue
        key = caption_lookup_key(cap.text)
        if key and key in used_caption_keys:
            consumed.add(i)
            continue
        for j in range(i + 1, min(len(blocks), i + 1 + max_lookahead)):
            nb = blocks[j]
            if nb.type == "table":
                if nb.caption is None:
                    cap.method = "binary-top-caption-lookahead-v20"
                    cap.position = "after"
                    nb.caption = cap
                    nb.raw.setdefault("caption_attached_from_top_block_index_v20", i)
                    if key:
                        used_caption_keys.add(key)
                    consumed.add(i)
                break
            if nb.type in {"image", "equation"}:
                break
    return [b for idx, b in enumerate(blocks) if idx not in consumed]


def _v20_extract_table_caption_from_mixed(self: "BinaryHwpParser", mixed: List[Dict[str, Any]]) -> Optional[Caption]:
    """Do not promote cell-internal table captions to the parent table.

    In structured HWP forms, a large outer layout table often contains many
    paragraphs, images, and nested data tables inside its cells. Captions inside
    those cells describe the nested objects, not the outer layout table. v14-v19
    could steal the first such caption and show it as the parent table caption,
    causing downstream caption shifts. Captions are now attached by local
    adjacency/relocation in cell.content instead.
    """
    return None


BinaryHwpParser._extract_table_caption_from_mixed = _v20_extract_table_caption_from_mixed  # type: ignore[method-assign]

# Ensure the existing monkey-patched binary postprocess uses the rebound v20
# relocation globals and records the patch version.
_old_binary_postprocess_v20 = BinaryHwpParser._postprocess

def _v20_binary_postprocess(self: "BinaryHwpParser") -> None:
    _old_binary_postprocess_v20(self)
    self.metadata["binary_caption_deduplication_v20"] = True

BinaryHwpParser._postprocess = _v20_binary_postprocess  # type: ignore[method-assign]

try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v19", "HWP Parser Verification UI v20")
except Exception:
    pass



# =============================================================================
# v21 binary table boundary and caption pairing fixes
# =============================================================================
# v20 still had a critical failure mode for form-heavy HWP files:
# a nested TABLE group could continue scanning after its declared cells had been
# parsed, thereby swallowing the following caption/body paragraphs and attaching
# the next table's caption to the previous nested table. v21 stops a table group
# after its declared cell count has been consumed. This keeps each inline table
# and its caption in the correct local sequence.


def _v21_expected_cell_count(row_count: Optional[int], col_count: Optional[int]) -> Optional[int]:
    try:
        if row_count and col_count and 1 <= int(row_count) <= 1000 and 1 <= int(col_count) <= 300:
            return int(row_count) * int(col_count)
    except Exception:
        pass
    return None


def _v21_parse_table_group(self: "BinaryHwpParser", records: List[HwpRecord], start: int, section: Optional[int]) -> Tuple[Block, int]:
    table_rec = records[start]
    row_count, col_count = self._table_dimensions(table_rec.payload)
    expected_cells = _v21_expected_cell_count(row_count, col_count)
    rows: Dict[int, Dict[int, TableCell]] = {}
    mixed: List[Dict[str, Any]] = []
    raw_cells: List[Dict[str, Any]] = []

    j = start + 1
    cell_no = 0
    while j < len(records):
        r = records[j]
        if r.level < table_rec.level:
            break
        if j != start + 1 and r.tag_id == HWPTAG_TABLE and r.level <= table_rec.level:
            break
        if expected_cells is not None and cell_no >= expected_cells:
            break

        if r.tag_id == HWPTAG_LIST_HEADER:
            cell_meta = self._parse_list_header_as_cell(r.payload, cell_no)
            cell_no += 1
            k = j + 1
            cell_records: List[HwpRecord] = []
            while k < len(records):
                nr = records[k]
                if nr.level < table_rec.level:
                    break
                if k != j + 1 and nr.tag_id == HWPTAG_LIST_HEADER and nr.level == r.level:
                    break
                if k != j + 1 and nr.tag_id == HWPTAG_TABLE and nr.level <= table_rec.level:
                    break
                cell_records.append(nr)
                k += 1

            cell, cell_mixed = self._parse_cell_records(cell_records, cell_meta, section)
            rr = int(cell.row if cell.row is not None else 0)
            cc = int(cell.col if cell.col is not None else 0)
            rows.setdefault(rr, {})[cc] = cell
            raw_cells.append({"row": rr, "col": cc, "row_span": cell.row_span, "col_span": cell.col_span, "text_preview": (cell.text or "")[:120]})
            for item in cell_mixed:
                item.setdefault("row", rr)
                item.setdefault("col", cc)
                mixed.append(item)
            j = k
            if expected_cells is not None and cell_no >= expected_cells:
                break
            continue
        j += 1

    table_rows = self._materialize_rows(rows, row_count, col_count)
    caption = self._extract_table_caption_from_mixed(mixed)
    block = Block(
        "table",
        unique_id("tbl"),
        self._next_order(),
        section=section,
        caption=caption,
        rows=table_rows,
        raw={
            "record": table_rec.to_meta(),
            "binary_table_parser": "v21",
            "row_count_declared": row_count,
            "col_count_declared": col_count,
            "expected_cell_count_v21": expected_cells,
            "actual_cell_count_v21": cell_no,
            "cell_count": sum(len(r) for r in table_rows),
            "cells": raw_cells[:300],
            "mixed_content_count": len(mixed),
            "payload_prefix_hex": table_rec.payload[:96].hex(),
            "table_group_boundary_fix_v21": True,
        },
    )
    if mixed:
        block.raw["mixed_content"] = mixed[:500]
    return block, j


BinaryHwpParser._parse_table_group = _v21_parse_table_group  # type: ignore[method-assign]


def _v21_relocate_captions_in_table_item(item: Dict[str, Any]) -> None:
    for row in item.get("rows") or []:
        for cell in row or []:
            if not isinstance(cell, dict):
                continue
            content = cell.get("content") or []
            if content:
                cell["content"] = _v21_relocate_captions_in_content(content)


def _v21_relocate_captions_in_content(content: List[Dict[str, Any]], max_lookahead: int = 18, max_text_chars: int = 12000) -> List[Dict[str, Any]]:
    if not content:
        return content

    normalized: List[Dict[str, Any]] = []
    for item in content:
        if not isinstance(item, dict):
            continue
        item = dict(item)
        if item.get("type") == "table":
            _v21_relocate_captions_in_table_item(item)
        cap_item = _v19_make_caption_item_from_paragraph(item)
        normalized.append(cap_item if cap_item is not None else item)

    consumed: set[int] = set()
    used_keys = _v20_collect_object_caption_keys(normalized)

    def attach(target_idx: int, cap_idx: int, target_type: str, method: str, position: str) -> bool:
        if target_idx < 0 or target_idx >= len(normalized) or cap_idx in consumed:
            return False
        target = normalized[target_idx]
        if not isinstance(target, dict) or target.get("type") != target_type:
            return False
        if target.get("caption"):
            return False
        cap_obj = _v19_caption_obj_from_item(normalized[cap_idx], target_type, method, position)
        if not cap_obj:
            return False
        key = _v20_caption_key_from_caption_obj(cap_obj)
        if key and key in used_keys:
            consumed.add(cap_idx)
            return False
        target["caption"] = cap_obj
        target.setdefault("raw", {})["caption_attached_from_item_index_v21"] = cap_idx
        if key:
            used_keys.add(key)
        consumed.add(cap_idx)
        return True

    for i, item in enumerate(normalized):
        if not isinstance(item, dict) or item.get("type") != "caption" or i in consumed:
            continue
        target_type = _v19_caption_target_from_item(item)
        if target_type not in {"table", "image", "equation"}:
            continue
        key = _v20_caption_item_key(item)
        if key and key in used_keys:
            consumed.add(i)
            continue

        # Bottom caption: only if immediately after a compatible object or after
        # short caption-neutral spacing. Stop on a non-empty paragraph.
        for j in range(i - 1, max(-1, i - 5), -1):
            prev = normalized[j]
            if not isinstance(prev, dict):
                continue
            typ = prev.get("type")
            if typ == target_type:
                if not prev.get("caption"):
                    attach(j, i, target_type, "binary-inline-caption-previous-v21", "after")
                break
            if typ == "paragraph" and normalize_space(str(prev.get("text") or "")):
                break
            if typ in {"table", "image", "equation"} and typ != target_type:
                break
        if i in consumed:
            continue

        # Preposed/top caption. Attach to the first compatible object within a
        # bounded local scope. Stop at another caption or unrelated object.
        text_budget = 0
        for j in range(i + 1, min(len(normalized), i + 1 + max_lookahead)):
            nxt = normalized[j]
            if not isinstance(nxt, dict):
                continue
            typ = nxt.get("type")
            if typ == "caption":
                break
            if typ == "paragraph":
                t = normalize_space(str(nxt.get("text") or ""))
                text_budget += len(t)
                if text_budget > max_text_chars:
                    break
                continue
            if typ == target_type:
                if not nxt.get("caption"):
                    attach(j, i, target_type, "binary-table-caption-lookahead-v21" if target_type == "table" else f"binary-{target_type}-caption-lookahead-v21", "after")
                break
            if typ in {"table", "image", "equation"}:
                break

    object_keys = _v20_collect_object_caption_keys(normalized)
    out: List[Dict[str, Any]] = []
    seen_caption_keys: set[str] = set()
    for idx, item in enumerate(normalized):
        if idx in consumed:
            continue
        if isinstance(item, dict) and item.get("type") == "caption":
            key = _v20_caption_item_key(item)
            if key and key in object_keys:
                continue
            if key and key in seen_caption_keys:
                continue
            if key:
                seen_caption_keys.add(key)
        out.append(item)
    return out


# Rebind both names because earlier code calls the v19 symbol directly.
_v19_relocate_captions_in_content = _v21_relocate_captions_in_content  # type: ignore[assignment]
_v19_relocate_captions_in_table_item = _v21_relocate_captions_in_table_item  # type: ignore[assignment]


def _content_plain_text_v21(content: Sequence[Dict[str, Any]]) -> str:
    parts: List[str] = []
    for item in content or []:
        if not isinstance(item, dict):
            continue
        typ = item.get("type")
        if typ in {"paragraph", "caption"} and item.get("text"):
            parts.append(str(item.get("text")))
        elif typ == "table":
            if item.get("caption") and isinstance(item.get("caption"), dict):
                parts.append(str(item["caption"].get("text") or ""))
            for r in item.get("rows") or []:
                row_parts = []
                for c in r or []:
                    if isinstance(c, dict):
                        row_parts.append(str(c.get("text") or ""))
                if row_parts:
                    parts.append(" | ".join([p for p in row_parts if p]))
        elif typ == "image":
            cap = item.get("caption")
            if isinstance(cap, dict) and cap.get("text"):
                parts.append(str(cap.get("text")))
    return normalize_space("\n".join([p for p in parts if p]))


def _v21_postprocess_nested_cell_captions_in_block(block: Block) -> None:
    if block.type != "table" or not block.rows:
        return
    for row in block.rows or []:
        for cell in row or []:
            content = getattr(cell, "content", None) or []
            if content:
                cell.content = _v21_relocate_captions_in_content(content)
                cell.text = _content_plain_text_v21(cell.content)
                cell.paragraphs = [str(x.get("text", "")) for x in cell.content if isinstance(x, dict) and x.get("type") == "paragraph" and x.get("text")]


_old_binary_postprocess_v21 = BinaryHwpParser._postprocess

def _v21_binary_postprocess(self: "BinaryHwpParser") -> None:
    _old_binary_postprocess_v21(self)
    for b in self.blocks:
        _v21_postprocess_nested_cell_captions_in_block(b)
    self.metadata["binary_table_group_boundary_fix_v21"] = True
    self.metadata["binary_caption_local_pairing_v21"] = True
    for i, b in enumerate(self.blocks, 1):
        b.order = i

BinaryHwpParser._postprocess = _v21_binary_postprocess  # type: ignore[method-assign]

try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v20", "HWP Parser Verification UI v21")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v19", "HWP Parser Verification UI v21")
except Exception:
    pass


# =============================================================================
# v22 caption-table pairing correction for nested table-heavy HWP forms
# =============================================================================
# Root cause fixed here:
# - v17/v21 attached a table caption to the immediately preceding nested table
#   while parsing a cell. In HWP form documents, table captions can be serialized
#   before the *next* table even when a previous table already appeared in the
#   same outer cell. This caused captions such as
#   "표. 트랜스포머-오토인코더 기반 ... 성능 비교 실험" to be attached to the
#   previous SMD/PSM dataset table, while the actual performance table lost its
#   caption.
#
# v22 policy:
# 1. During binary cell parsing, table captions are never attached immediately to
#    a previous table. They remain as explicit inline caption items.
# 2. A local post-pass pairs table captions to the best table candidate using
#    local order plus weak semantic cues. This handles both "caption before
#    table" and "caption after table" layouts without stealing captions across
#    nested-table boundaries.


def _v22_inline_table_text(item: Dict[str, Any]) -> str:
    parts: List[str] = []
    if not isinstance(item, dict):
        return ""
    if item.get("text"):
        parts.append(str(item.get("text")))
    cap = item.get("caption")
    if isinstance(cap, dict) and cap.get("text"):
        parts.append(str(cap.get("text")))
    for row in item.get("rows") or []:
        for cell in row or []:
            if not isinstance(cell, dict):
                continue
            if cell.get("text"):
                parts.append(str(cell.get("text")))
            for sub in cell.get("content") or []:
                if isinstance(sub, dict):
                    if sub.get("type") in {"paragraph", "caption"} and sub.get("text"):
                        parts.append(str(sub.get("text")))
                    elif sub.get("type") == "table":
                        parts.append(_v22_inline_table_text(sub))
    return normalize_space("\n".join(p for p in parts if p))


def _v22_semantic_table_score(caption_text: str, table_item: Dict[str, Any]) -> float:
    cap = normalize_space(caption_text).lower()
    txt = normalize_space(_v22_inline_table_text(table_item)).lower()
    if not cap or not txt:
        return 0.0
    score = 0.0

    # Generic token overlap. Keep it weak because Korean captions and English
    # table headers often do not share literal tokens.
    cap_tokens = set(re.findall(r"[가-힣A-Za-z0-9]+", cap))
    txt_tokens = set(re.findall(r"[가-힣A-Za-z0-9]+", txt))
    if cap_tokens and txt_tokens:
        score += min(12.0, len(cap_tokens & txt_tokens) * 2.0)

    # Domain-specific but safe cues for analysis/report tables.
    if any(k in cap for k in ["데이터셋", "벤치마크", "로깅", "공개"]):
        if any(k.lower() in txt for k in ["smd", "psm", "number of features", "number of train", "number of test", "anomaly ratio", "csv", "label"]):
            score += 28.0
    if any(k in cap for k in ["성능", "비교", "실험", "탐지 성능", "f1"]):
        if any(k.lower() in txt for k in ["recall", "precision", "f1", "f1 score", "model", "methods", "metric", "knn", "autoencoder", "ours"]):
            score += 30.0
    if any(k in cap for k in ["마진", "margin"]):
        if any(k.lower() in txt for k in ["margin", "f1", "score", "0.85", "0.9", "0.95", "1.0"]):
            score += 22.0

    # Penalize obvious mismatch: dataset caption on performance table or vice versa.
    if any(k in cap for k in ["데이터셋", "벤치마크"]):
        if any(k.lower() in txt for k in ["recall", "precision", "f1", "model", "knn", "autoencoder"]):
            score -= 10.0
    if any(k in cap for k in ["성능", "비교", "실험"]):
        if any(k.lower() in txt for k in ["number of features", "number of train", "number of test", "anomaly ratio", "format", "csv"]):
            score -= 12.0
    return score


def _v22_caption_target_from_item(item: Dict[str, Any]) -> Optional[str]:
    return _v19_caption_target_from_item(item) if '_v19_caption_target_from_item' in globals() else infer_caption_target_type(str(item.get("text") or ""), {}, "")


def _v22_caption_obj_from_item(item: Dict[str, Any], target_type: str, method: str, position: str) -> Optional[Dict[str, Any]]:
    cap_obj = _v19_caption_obj_from_item(item, target_type, method, position) if '_v19_caption_obj_from_item' in globals() else None
    if cap_obj:
        return cap_obj
    txt = normalize_space(str(item.get("text") or ""))
    cap = caption_from_text(txt, target_type)
    if not cap:
        return None
    cap.method = method
    cap.position = position
    return cap.to_dict()


def _v22_make_caption_item_from_paragraph(item: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if '_v19_make_caption_item_from_paragraph' in globals():
        return _v19_make_caption_item_from_paragraph(item)
    if not isinstance(item, dict) or item.get("type") != "paragraph":
        return None
    cap = caption_from_text(str(item.get("text") or ""))
    if not cap:
        return None
    return {"type": "caption", "text": cap.text, "caption": cap.to_dict(), "target_type": infer_caption_target_type(cap.text, {}, "")}


def _v22_existing_caption_score(table_item: Dict[str, Any]) -> float:
    cap = table_item.get("caption")
    if not isinstance(cap, dict) or not cap.get("text"):
        return -1.0
    return _v22_semantic_table_score(str(cap.get("text")), table_item)


def _v22_should_override_caption(table_item: Dict[str, Any], new_score: float) -> bool:
    cap = table_item.get("caption")
    if not cap:
        return True
    old_score = _v22_existing_caption_score(table_item)
    # Replace only when the existing caption is weak and the new local caption
    # has a clearly stronger association.
    return new_score >= max(18.0, old_score + 8.0)


def _v22_relocate_captions_in_table_item(item: Dict[str, Any]) -> None:
    for row in item.get("rows") or []:
        for cell in row or []:
            if not isinstance(cell, dict):
                continue
            content = cell.get("content") or []
            if content:
                cell["content"] = _v22_relocate_captions_in_content(content)
                # Keep text in sync with nested content for UI search/summary.
                cell["text"] = _content_plain_text_v21(cell["content"]) if '_content_plain_text_v21' in globals() else normalize_space("\n".join(str(x.get("text", "")) for x in cell["content"] if isinstance(x, dict)))


def _v22_relocate_captions_in_content(content: List[Dict[str, Any]], max_lookahead: int = 24, max_text_chars: int = 18000) -> List[Dict[str, Any]]:
    if not content:
        return content

    normalized: List[Dict[str, Any]] = []
    for item in content:
        if not isinstance(item, dict):
            continue
        item = dict(item)
        if item.get("type") == "table":
            _v22_relocate_captions_in_table_item(item)
        cap_item = _v22_make_caption_item_from_paragraph(item)
        normalized.append(cap_item if cap_item is not None else item)

    consumed: set[int] = set()

    def candidate_previous(cap_idx: int, target_type: str) -> Optional[Tuple[int, float, str]]:
        # Previous/bottom captions are accepted only when there is no non-empty
        # paragraph between the table and caption. This prevents captions that
        # start the next subsection from being stolen by the previous table.
        distance = 0
        for j in range(cap_idx - 1, max(-1, cap_idx - 5), -1):
            prev = normalized[j]
            if not isinstance(prev, dict):
                continue
            typ = prev.get("type")
            if typ == target_type:
                base = max(4.0, 16.0 - distance * 2.0)
                if target_type == "table":
                    base += _v22_semantic_table_score(str(normalized[cap_idx].get("text") or ""), prev)
                return (j, base, "after")
            if typ == "paragraph" and normalize_space(str(prev.get("text") or "")):
                return None
            if typ in {"table", "image", "equation"} and typ != target_type:
                return None
            distance += 1
        return None

    def candidate_next(cap_idx: int, target_type: str) -> Optional[Tuple[int, float, str]]:
        text_budget = 0
        distance = 0
        for j in range(cap_idx + 1, min(len(normalized), cap_idx + 1 + max_lookahead)):
            nxt = normalized[j]
            if not isinstance(nxt, dict):
                continue
            typ = nxt.get("type")
            if typ == "caption" and _v22_caption_target_from_item(nxt) == target_type:
                return None
            if typ == "paragraph":
                t = normalize_space(str(nxt.get("text") or ""))
                # Major numbered heading often separates a previous bottom
                # caption from the next object, but in this corpus captions for
                # tables are usually right before the explanatory paragraph and
                # then the table. Therefore do not stop immediately; just bound
                # the budget.
                text_budget += len(t)
                if text_budget > max_text_chars:
                    return None
                distance += 1
                continue
            if typ == target_type:
                base = max(6.0, 28.0 - distance * 0.8)
                if target_type == "table":
                    base += _v22_semantic_table_score(str(normalized[cap_idx].get("text") or ""), nxt)
                return (j, base, "before")
            if typ in {"table", "image", "equation"}:
                return None
            distance += 1
        return None

    # First pass: robust table-caption pairing. This is the critical v22 fix.
    for i, item in enumerate(normalized):
        if not isinstance(item, dict) or item.get("type") != "caption" or i in consumed:
            continue
        target_type = _v22_caption_target_from_item(item)
        if target_type != "table":
            continue
        text = normalize_space(str(item.get("text") or ""))
        if not caption_from_text(text, "table"):
            continue
        prev_c = candidate_previous(i, "table")
        next_c = candidate_next(i, "table")
        candidates = [c for c in [prev_c, next_c] if c is not None]
        if not candidates:
            continue
        # Prefer the highest local+semantic score. This fixes the case where a
        # performance caption appears after a dataset table but before the actual
        # performance table.
        target_idx, score, pos = max(candidates, key=lambda x: x[1])
        target = normalized[target_idx]
        if not _v22_should_override_caption(target, score):
            continue
        cap_obj = _v22_caption_obj_from_item(item, "table", "binary-table-caption-paired-v22", "after" if pos in {"before", "after"} else "after")
        if not cap_obj:
            continue
        target["caption"] = cap_obj
        target.setdefault("raw", {})["caption_attached_from_item_index_v22"] = i
        target.setdefault("raw", {})["caption_pairing_score_v22"] = score
        target.setdefault("raw", {})["caption_pairing_direction_v22"] = pos
        consumed.add(i)

    # Second pass: image/equation captions. Keep v21-like locality but do not
    # interfere with table captions.
    for i, item in enumerate(normalized):
        if not isinstance(item, dict) or item.get("type") != "caption" or i in consumed:
            continue
        target_type = _v22_caption_target_from_item(item)
        if target_type not in {"image", "equation"}:
            continue
        # Prefer immediate previous object, otherwise immediate lookahead.
        attached = False
        for j in range(i - 1, max(-1, i - 4), -1):
            prev = normalized[j]
            if not isinstance(prev, dict):
                continue
            typ = prev.get("type")
            if typ == target_type:
                if not prev.get("caption"):
                    cap_obj = _v22_caption_obj_from_item(item, target_type, f"binary-{target_type}-caption-previous-v22", "after")
                    if cap_obj:
                        prev["caption"] = cap_obj
                        consumed.add(i)
                        attached = True
                break
            if typ == "paragraph" and normalize_space(str(prev.get("text") or "")):
                break
        if attached:
            continue
        for j in range(i + 1, min(len(normalized), i + 8)):
            nxt = normalized[j]
            if not isinstance(nxt, dict):
                continue
            typ = nxt.get("type")
            if typ == target_type:
                if not nxt.get("caption"):
                    cap_obj = _v22_caption_obj_from_item(item, target_type, f"binary-{target_type}-caption-next-v22", "before")
                    if cap_obj:
                        nxt["caption"] = cap_obj
                        consumed.add(i)
                break
            if typ in {"table", "image", "equation"}:
                break

    # Remove consumed captions and exact duplicate standalone captions. Do not
    # delete a unique standalone table caption unless it has actually been paired.
    object_keys = _v20_collect_object_caption_keys(normalized) if '_v20_collect_object_caption_keys' in globals() else set()
    out: List[Dict[str, Any]] = []
    seen_standalone: set[str] = set()
    for idx, item in enumerate(normalized):
        if idx in consumed:
            continue
        if isinstance(item, dict) and item.get("type") == "caption":
            key = _v20_caption_item_key(item) if '_v20_caption_item_key' in globals() else caption_lookup_key(str(item.get("text") or ""))
            if key and key in object_keys:
                continue
            if key and key in seen_standalone:
                continue
            if key:
                seen_standalone.add(key)
        out.append(item)
    return out


def _v22_parse_cell_records(self, cell_records, cell_meta, section):
    """v22 replacement: table captions stay explicit until local pairing.

    This removes the v17 behavior that attached any table caption to the most
    recent previous table during low-level record parsing.
    """
    parts = []
    paras = []
    mixed = []
    i = 0
    while i < len(cell_records):
        rec = cell_records[i]

        if rec.tag_id == HWPTAG_TABLE:
            try:
                nested_block, next_i = self._parse_table_group(cell_records, i, section)
                nested_item = _v17_inline_table_item_from_block(nested_block)
                mixed.append(nested_item)
                nested_text = normalize_space(nested_item.get("text", ""))
                if nested_text:
                    parts.append(nested_text)
                    paras.append(nested_text)
                i = max(next_i, i + 1)
                continue
            except Exception as e:
                mixed.append({"type": "table", "rows": [], "raw": {"record": rec.to_meta(), "nested_table_parse_error_v22": str(e)}})
                i += 1
                continue

        if rec.tag_id == HWPTAG_CTRL_HEADER and _v16_ctrl_id(rec.payload) == "gso":
            block, next_i = self._parse_gso_group(cell_records, i, section)
            if block:
                item = {
                    "type": "image",
                    "image_path": block.image_path,
                    "media_type": block.media_type,
                    "caption": block.caption.to_dict() if block.caption else None,
                    "raw": block.raw,
                }
                if hasattr(block, "image_paths"):
                    item["image_paths"] = getattr(block, "image_paths")
                mixed.append(item)
                cap_txt = (item.get("caption") or {}).get("text") if isinstance(item.get("caption"), dict) else None
                if cap_txt:
                    parts.append(str(cap_txt)); paras.append(str(cap_txt))
                i = max(next_i, i + 1)
                continue

        if rec.tag_id == HWPTAG_PARA_TEXT:
            text, ctrls = self.decode_para_text(rec.payload)
            text = normalize_space(text)
            if not text or self._is_control_only_text(text):
                i += 1
                continue
            cap_img = caption_from_text(text, "image")
            cap_tbl = caption_from_text(text, "table")
            if cap_img:
                attached = False
                for item in reversed(mixed):
                    if item.get("type") == "image" and not item.get("caption"):
                        item["caption"] = cap_img.to_dict(); attached = True; break
                if not attached:
                    mixed.append({"type": "caption", "text": cap_img.text, "caption": cap_img.to_dict(), "target_type": "image"})
                parts.append(cap_img.text); paras.append(cap_img.text); i += 1; continue
            if cap_tbl:
                # v22: do NOT attach to previous table here. Keep explicit and
                # let _v22_relocate_captions_in_content pair it with the correct
                # nested table using scope/order/semantic cues.
                mixed.append({"type": "caption", "text": cap_tbl.text, "caption": cap_tbl.to_dict(), "target_type": "table"})
                parts.append(cap_tbl.text); paras.append(cap_tbl.text); i += 1; continue
            parts.append(text); paras.append(text); mixed.append({"type": "paragraph", "text": text})
            i += 1
            continue

        if rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
            image_block = self._image_block_from_picture(rec, section)
            item = {"type": "image", "image_path": image_block.image_path, "media_type": image_block.media_type, "raw": image_block.raw}
            if hasattr(image_block, "image_paths"):
                item["image_paths"] = getattr(image_block, "image_paths")
            mixed.append(item)
            i += 1
            continue

        i += 1

    # Pair captions inside the cell immediately after its local mixed stream is
    # constructed. Postprocess runs the same routine again, which is idempotent.
    mixed = _v22_relocate_captions_in_content(mixed)
    text_for_cell = _content_plain_text_v21(mixed) if '_content_plain_text_v21' in globals() else normalize_space("\n".join(str(x.get("text", "")) for x in mixed if isinstance(x, dict)))
    cell = TableCell(
        text=text_for_cell or normalize_space("\n".join(parts)),
        row=cell_meta.get("row", 0),
        col=cell_meta.get("col", 0),
        row_span=max(1, int(cell_meta.get("row_span", 1))),
        col_span=max(1, int(cell_meta.get("col_span", 1))),
        attrs={"binary_cell_meta": cell_meta, "nested_table_caption_pairing_v22": True},
        paragraphs=_dedupe_adjacent_strings([str(x.get("text", "")) for x in mixed if isinstance(x, dict) and x.get("type") == "paragraph" and x.get("text")]),
        content=mixed,
    )
    return cell, mixed


BinaryHwpParser._parse_cell_records = _v22_parse_cell_records  # type: ignore[method-assign]
_v21_relocate_captions_in_content = _v22_relocate_captions_in_content  # type: ignore[assignment]
_v21_relocate_captions_in_table_item = _v22_relocate_captions_in_table_item  # type: ignore[assignment]
_v19_relocate_captions_in_content = _v22_relocate_captions_in_content  # type: ignore[assignment]
_v19_relocate_captions_in_table_item = _v22_relocate_captions_in_table_item  # type: ignore[assignment]

_old_binary_postprocess_v22 = BinaryHwpParser._postprocess

def _v22_binary_postprocess(self: "BinaryHwpParser") -> None:
    _old_binary_postprocess_v22(self)
    # Run the v22 local pairing one final time after all previous monkey patches.
    for b in self.blocks:
        if b.type == "table" and b.rows:
            for row in b.rows or []:
                for cell in row or []:
                    content = getattr(cell, "content", None) or []
                    if content:
                        cell.content = _v22_relocate_captions_in_content(content)
                        cell.text = _content_plain_text_v21(cell.content) if '_content_plain_text_v21' in globals() else cell.text
    self.metadata["binary_nested_table_caption_pairing_v22"] = True
    for i, b in enumerate(self.blocks, 1):
        b.order = i

BinaryHwpParser._postprocess = _v22_binary_postprocess  # type: ignore[method-assign]

try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v21", "HWP Parser Verification UI v22")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v20", "HWP Parser Verification UI v22")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v19", "HWP Parser Verification UI v22")
except Exception:
    pass

# =============================================================================
# CLI
# =============================================================================

def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    ap = argparse.ArgumentParser(description="Parse HWP into structured JSON and optionally launch verification web UI.")
    ap.add_argument("input", nargs="?", help="Input .hwp or intermediate .xml. Optional with --web.")
    ap.add_argument("-o", "--output-dir", default="hwp_parsed_output", help="Output directory.")
    ap.add_argument("--mode", choices=["auto", "pyhwp", "binary", "xml"], default="auto")
    ap.add_argument("--json-name", default="result.json")
    ap.add_argument("--compact", action="store_true")
    ap.add_argument("--keep-intermediate", action="store_true")
    ap.add_argument("--hwp5proc", default="hwp5proc")
    ap.add_argument("--print-summary", action="store_true")
    ap.add_argument("--web", action="store_true", help="Start built-in verification web UI.")
    ap.add_argument("--host", default="127.0.0.1")
    ap.add_argument("--port", type=int, default=7860)
    ap.add_argument("--no-browser", action="store_true")
    return ap.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    output_dir = ensure_dir(args.output_dir)
    if args.web:
        return start_web_ui(args.input, output_dir, args.mode, True if args.keep_intermediate else True, args.hwp5proc, args.host, args.port, not args.no_browser)
    if not args.input:
        print("ERROR: input is required unless --web is used", file=sys.stderr)
        return 1
    try:
        doc = FullHwpParser(args.input, output_dir, mode=args.mode, keep_intermediate=args.keep_intermediate, hwp5proc_path=args.hwp5proc).parse()
        json_path = write_json(doc, output_dir / args.json_name, pretty=not args.compact)
        if args.print_summary:
            print(json.dumps(summarize(doc), ensure_ascii=False, indent=2))
        else:
            print(f"Wrote: {json_path}")
            print(json.dumps(summarize(doc), ensure_ascii=False, indent=2))
        return 0 if not doc.errors else 2
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1


# =============================================================================
# v23 generalized caption candidates + complete table retention
# =============================================================================
# v22 still assumed that most captions are textually marked with "표." or
# "그림.". That is not how many HWP documents are authored: users often create a
# built-in caption and type only the title text. v23 therefore separates
# "caption candidacy" from textual label parsing. A short structural/title-like
# paragraph near a table/image can become a caption candidate even when it has no
# visible prefix.
#
# v23 also removes debug/raw slicing for table mixed content and cell metadata so
# exported JSON keeps the complete table evidence instead of truncated previews.

CAPTION_GENERIC_MAX_CHARS_V23 = 240
CAPTION_GENERIC_MAX_LINES_V23 = 3


def _v23_norm_key(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()


def _v23_caption_explicit_target(text: str) -> Optional[str]:
    if caption_from_text(text, "table"):
        return "table"
    if caption_from_text(text, "image"):
        return "image"
    if caption_from_text(text, "equation"):
        return "equation"
    return None


def _v23_plain_title_like(text: str) -> bool:
    """Return True for caption-title-like text without requiring 표./그림.

    This intentionally rejects long prose and section headings, while accepting
    common Korean/English technical figure/table titles.
    """
    s = normalize_space(text)
    if not s:
        return False
    if len(s) > CAPTION_GENERIC_MAX_CHARS_V23:
        return False
    lines = [x.strip() for x in str(text or "").splitlines() if x.strip()]
    if len(lines) > CAPTION_GENERIC_MAX_LINES_V23:
        return False

    # Reject obvious section/body starts.
    if re.match(r"^(?:[0-9]+[.)]|[①-⑳]|[가-힣]\)|□|<\s*서식|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.)", s):
        return False
    if re.match(r"^(?:그리고|또한|따라서|그러나|먼저|다음으로|본\s+연구|제안한\s+방법|이를\s+통해)\b", s):
        return False

    # Long narrative sentence endings are usually body text, not captions.
    if len(s) > 80 and re.search(r"(하였다|하였다\.|한다|한다\.|이다|이다\.|있다|있다\.|되었다|되었다\.)$", s):
        return False

    # Reject table data rows / numeric-only fragments.
    tokens = re.findall(r"[가-힣A-Za-z]+", s)
    if not tokens:
        return False
    if len(tokens) == 1 and len(s) < 12:
        return False

    # Titles often contain nominalized nouns and no final prose verb. This is a
    # permissive title-shape check.
    title_cues = [
        "구조", "시스템", "모델", "결과", "비교", "성능", "분석", "화면", "모습",
        "데이터셋", "벤치마크", "실험", "프레임워크", "알고리즘", "파이프라인",
        "dataset", "benchmark", "performance", "comparison", "result", "architecture",
        "framework", "system", "model", "analysis", "experiment", "pipeline",
    ]
    if any(c.lower() in s.lower() for c in title_cues):
        return True

    # Short noun phrase with little punctuation.
    punct = len(re.findall(r"[.!?。！？]", s))
    if len(s) <= 120 and punct == 0 and len(tokens) >= 2:
        return True
    return False


def _v23_caption_candidate_from_item(item: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    if not isinstance(item, dict):
        return None
    typ = item.get("type")
    text = normalize_space(str(item.get("text") or ""))
    cap_dict = item.get("caption") if isinstance(item.get("caption"), dict) else None
    if cap_dict and cap_dict.get("text"):
        text = normalize_space(str(cap_dict.get("text") or text))

    if not text:
        return None

    explicit_target = _v23_caption_explicit_target(text)
    if typ == "caption":
        target = item.get("target_type") or explicit_target or infer_caption_target_type(text, item.get("attrs") if isinstance(item.get("attrs"), dict) else {}, "")
        return {
            "text": text,
            "explicit_target": target,
            "generic": explicit_target is None,
            "structural": True,
            "source_type": typ,
        }

    if typ == "paragraph":
        if explicit_target:
            return {
                "text": text,
                "explicit_target": explicit_target,
                "generic": False,
                "structural": False,
                "source_type": typ,
            }
        if _v23_plain_title_like(text):
            return {
                "text": text,
                "explicit_target": None,
                "generic": True,
                "structural": False,
                "source_type": typ,
            }
    return None


def _v23_caption_to_obj(text: str, target_type: str, method: str, position: str, generic: bool = False) -> Dict[str, Any]:
    cap = caption_from_text(text, target_type)
    if cap:
        cap.method = method
        cap.position = position
        d = cap.to_dict()
        d.setdefault("raw", {})["v23_explicit_caption"] = True
        return d
    return Caption(
        text=normalize_space(text),
        method=method,
        position=position,
        raw={"v23_generic_caption_without_visible_label": bool(generic)},
    ).to_dict()


def _v23_object_text(item: Dict[str, Any]) -> str:
    if not isinstance(item, dict):
        return ""
    if item.get("type") == "table":
        return _v22_inline_table_text(item) if '_v22_inline_table_text' in globals() else normalize_space(str(item.get("text") or ""))
    return normalize_space(str(item.get("text") or ""))


def _v23_semantic_score(target_type: str, caption_text: str, obj: Dict[str, Any]) -> float:
    if target_type == "table" and '_v22_semantic_table_score' in globals():
        return float(_v22_semantic_table_score(caption_text, obj))
    # For images we usually do not have OCR text. Use only weak title cues; the
    # actual pairing is mostly structural/proximity based.
    cap = normalize_space(caption_text).lower()
    score = 0.0
    if target_type == "image":
        if any(k in cap for k in ["구조", "시스템", "모델", "화면", "모습", "결과", "점수", "변화", "architecture", "system", "model", "screen"]):
            score += 8.0
    return score


def _v23_between_stats(items: List[Dict[str, Any]], a: int, b: int, target_type: str) -> Dict[str, Any]:
    lo, hi = sorted((a, b))
    text_chars = 0
    para_count = 0
    hard_objects = 0
    same_type_objects = 0
    caption_count = 0
    for k in range(lo + 1, hi):
        it = items[k]
        if not isinstance(it, dict):
            continue
        typ = it.get("type")
        if typ == "paragraph":
            t = normalize_space(str(it.get("text") or ""))
            text_chars += len(t)
            if t:
                para_count += 1
        elif typ == "caption":
            caption_count += 1
        elif typ in {"table", "image", "equation"}:
            hard_objects += 1
            if typ == target_type:
                same_type_objects += 1
    return {"text_chars": text_chars, "para_count": para_count, "hard_objects": hard_objects, "same_type_objects": same_type_objects, "caption_count": caption_count}


def _v23_pair_score(items: List[Dict[str, Any]], cap_idx: int, obj_idx: int, target_type: str, cand: Dict[str, Any]) -> Optional[Tuple[float, str]]:
    cap_text = cand["text"]
    explicit_target = cand.get("explicit_target")
    generic = bool(cand.get("generic"))
    if explicit_target and explicit_target != target_type:
        return None

    stats = _v23_between_stats(items, cap_idx, obj_idx, target_type)
    if stats["same_type_objects"]:
        return None
    # Do not cross another caption of the same target. This prevents caption
    # shifts across sequential tables.
    if stats["caption_count"] and abs(cap_idx - obj_idx) > 1:
        return None

    direction = "before" if cap_idx < obj_idx else "after"
    item_distance = abs(cap_idx - obj_idx) - 1
    semantic = _v23_semantic_score(target_type, cap_text, items[obj_idx])

    if generic:
        # Generic captions without visible labels must be structurally close or
        # semantically strong. This is the safety guard that avoids swallowing
        # arbitrary headings as captions.
        close = item_distance <= 2 and stats["hard_objects"] == 0 and stats["text_chars"] <= 260
        semantically_close = semantic >= 18.0 and stats["text_chars"] <= 4000 and stats["hard_objects"] == 0
        if not (close or semantically_close or cand.get("structural")):
            return None
        score = 26.0 + semantic
        if close:
            score += 18.0
        if cand.get("structural"):
            score += 18.0
    else:
        score = 52.0 + semantic

    # Position and distance terms.
    if direction == "before":
        score += 5.0
    else:
        score += 3.0
    score -= item_distance * (1.2 if not generic else 2.5)
    score -= min(30.0, stats["text_chars"] / (700.0 if not generic else 180.0))
    score -= stats["hard_objects"] * 12.0

    # For explicit table captions, allow explanatory paragraphs before the table
    # only when the semantic match is clearly stronger than the previous table.
    if not generic and target_type == "table" and stats["text_chars"] > 3000 and semantic < 18.0:
        score -= 20.0

    threshold = 20.0 if generic else 28.0
    if score < threshold:
        return None
    return score, direction


def _v23_relocate_captions_in_table_item(item: Dict[str, Any]) -> None:
    if not isinstance(item, dict):
        return
    for row in item.get("rows") or []:
        for cell in row or []:
            if not isinstance(cell, dict):
                continue
            content = cell.get("content") or []
            if content:
                cell["content"] = _v23_relocate_captions_in_content(content)
                cell["text"] = _content_plain_text_v21(cell["content"]) if '_content_plain_text_v21' in globals() else normalize_space("\n".join(str(x.get("text", "")) for x in cell["content"] if isinstance(x, dict)))


def _v23_relocate_captions_in_content(content: List[Dict[str, Any]], max_window: int = 40) -> List[Dict[str, Any]]:
    if not content:
        return content

    items: List[Dict[str, Any]] = []
    for original in content:
        if not isinstance(original, dict):
            continue
        item = dict(original)
        if item.get("type") == "table":
            _v23_relocate_captions_in_table_item(item)
        items.append(item)

    object_indices = [i for i, it in enumerate(items) if isinstance(it, dict) and it.get("type") in {"table", "image", "equation"}]
    caption_candidates: List[Tuple[int, Dict[str, Any]]] = []
    for i, it in enumerate(items):
        cand = _v23_caption_candidate_from_item(it)
        if cand:
            caption_candidates.append((i, cand))

    pair_options: List[Tuple[float, int, int, str, str, Dict[str, Any]]] = []
    for cap_idx, cand in caption_candidates:
        for obj_idx in object_indices:
            if abs(cap_idx - obj_idx) > max_window:
                continue
            target_type = str(items[obj_idx].get("type"))
            scored = _v23_pair_score(items, cap_idx, obj_idx, target_type, cand)
            if scored is None:
                continue
            score, direction = scored
            # Prefer a following object when the caption is before it and the
            # score is comparable. This matches common HWP caption-before-table
            # forms and prevents stealing by a previous table.
            if direction == "before":
                score += 2.0
            pair_options.append((score, cap_idx, obj_idx, target_type, direction, cand))

    pair_options.sort(key=lambda x: x[0], reverse=True)
    used_caps: set[int] = set()
    used_objs: set[int] = set()
    for score, cap_idx, obj_idx, target_type, direction, cand in pair_options:
        if cap_idx in used_caps or obj_idx in used_objs:
            continue
        obj = items[obj_idx]
        # Replace existing captions when v23 has structural evidence. This fixes
        # earlier versions that attached a later caption to the previous table.
        method = f"binary-{target_type}-caption-generic-v23" if cand.get("generic") else f"binary-{target_type}-caption-explicit-v23"
        position = "before" if direction == "before" else "after"
        obj["caption"] = _v23_caption_to_obj(cand["text"], target_type, method, position, generic=bool(cand.get("generic")))
        obj.setdefault("raw", {})["caption_pairing_v23"] = {
            "caption_index": cap_idx,
            "object_index": obj_idx,
            "score": score,
            "direction": direction,
            "generic": bool(cand.get("generic")),
            "structural": bool(cand.get("structural")),
        }
        used_caps.add(cap_idx)
        used_objs.add(obj_idx)

    # Remove only caption source items that were actually consumed. Generic
    # paragraph captions are removed only if paired; otherwise they remain normal
    # paragraph text.
    out: List[Dict[str, Any]] = []
    seen_unpaired_caption_keys: set[str] = set()
    object_caption_keys = set()
    for it in items:
        if isinstance(it, dict) and it.get("type") in {"table", "image", "equation"} and isinstance(it.get("caption"), dict):
            object_caption_keys.add(caption_lookup_key(str(it["caption"].get("text") or "")))
    for i, it in enumerate(items):
        if i in used_caps:
            continue
        if isinstance(it, dict) and it.get("type") == "caption":
            key = caption_lookup_key(str(it.get("text") or ""))
            if key and key in object_caption_keys:
                continue
            if key and key in seen_unpaired_caption_keys:
                continue
            if key:
                seen_unpaired_caption_keys.add(key)
        out.append(it)
    return out


def _v23_parse_cell_records(self, cell_records, cell_meta, section):
    parts: List[str] = []
    paras: List[str] = []
    mixed: List[Dict[str, Any]] = []
    i = 0
    while i < len(cell_records):
        rec = cell_records[i]

        if rec.tag_id == HWPTAG_TABLE:
            try:
                nested_block, next_i = self._parse_table_group(cell_records, i, section)
                nested_item = _v17_inline_table_item_from_block(nested_block) if '_v17_inline_table_item_from_block' in globals() else {
                    "type": "table",
                    "id": nested_block.id,
                    "text": nested_block.text,
                    "rows": [[cell.to_dict() for cell in row] for row in (nested_block.rows or [])],
                    "caption": nested_block.caption.to_dict() if nested_block.caption else None,
                    "raw": nested_block.raw,
                }
                mixed.append(nested_item)
                nested_text = normalize_space(nested_item.get("text", ""))
                if nested_text:
                    parts.append(nested_text)
                    paras.append(nested_text)
                i = max(next_i, i + 1)
                continue
            except Exception as e:
                mixed.append({"type": "table", "rows": [], "raw": {"record": rec.to_meta(), "nested_table_parse_error_v23": str(e)}})
                i += 1
                continue

        if rec.tag_id == HWPTAG_CTRL_HEADER and '_v16_ctrl_id' in globals() and _v16_ctrl_id(rec.payload) == "gso":
            block, next_i = self._parse_gso_group(cell_records, i, section)
            if block:
                item = {
                    "type": "image",
                    "image_path": block.image_path,
                    "media_type": block.media_type,
                    "caption": block.caption.to_dict() if block.caption else None,
                    "raw": block.raw,
                }
                if hasattr(block, "image_paths"):
                    item["image_paths"] = getattr(block, "image_paths")
                mixed.append(item)
                cap_txt = (item.get("caption") or {}).get("text") if isinstance(item.get("caption"), dict) else None
                if cap_txt:
                    parts.append(str(cap_txt)); paras.append(str(cap_txt))
                i = max(next_i, i + 1)
                continue

        if rec.tag_id == HWPTAG_PARA_TEXT:
            text, ctrls = self.decode_para_text(rec.payload)
            text = normalize_space(text)
            if not text or self._is_control_only_text(text):
                i += 1
                continue
            explicit = _v23_caption_explicit_target(text)
            if explicit:
                cap = caption_from_text(text, explicit)
                cap_dict = cap.to_dict() if cap else Caption(text=text, method="pattern").to_dict()
                mixed.append({"type": "caption", "text": cap_dict.get("text", text), "caption": cap_dict, "target_type": explicit})
                parts.append(cap_dict.get("text", text)); paras.append(cap_dict.get("text", text))
                i += 1
                continue
            # Do not decide generic captions here. Keep them as paragraphs so
            # v23 can use local object context before consuming them.
            parts.append(text); paras.append(text); mixed.append({"type": "paragraph", "text": text})
            i += 1
            continue

        if rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
            image_block = self._image_block_from_picture(rec, section)
            item = {"type": "image", "image_path": image_block.image_path, "media_type": image_block.media_type, "raw": image_block.raw}
            if hasattr(image_block, "image_paths"):
                item["image_paths"] = getattr(image_block, "image_paths")
            mixed.append(item)
            i += 1
            continue

        i += 1

    mixed = _v23_relocate_captions_in_content(mixed)
    text_for_cell = _content_plain_text_v21(mixed) if '_content_plain_text_v21' in globals() else normalize_space("\n".join(str(x.get("text", "")) for x in mixed if isinstance(x, dict)))
    cell = TableCell(
        text=text_for_cell or normalize_space("\n".join(parts)),
        row=cell_meta.get("row", 0),
        col=cell_meta.get("col", 0),
        row_span=max(1, int(cell_meta.get("row_span", 1))),
        col_span=max(1, int(cell_meta.get("col_span", 1))),
        attrs={"binary_cell_meta": cell_meta, "nested_table_caption_pairing_v23": True},
        paragraphs=_dedupe_adjacent_strings([str(x.get("text", "")) for x in mixed if isinstance(x, dict) and x.get("type") == "paragraph" and x.get("text")]),
        content=mixed,
    )
    return cell, mixed


def _v23_parse_table_group(self: "BinaryHwpParser", records: List[HwpRecord], start: int, section: Optional[int]) -> Tuple[Block, int]:
    table_rec = records[start]
    row_count, col_count = self._table_dimensions(table_rec.payload)
    expected_cells = _v21_expected_cell_count(row_count, col_count) if '_v21_expected_cell_count' in globals() else (row_count * col_count if row_count and col_count else None)
    rows: Dict[int, Dict[int, TableCell]] = {}
    mixed: List[Dict[str, Any]] = []
    raw_cells: List[Dict[str, Any]] = []

    j = start + 1
    cell_no = 0
    while j < len(records):
        r = records[j]
        if r.level < table_rec.level:
            break
        if j != start + 1 and r.tag_id == HWPTAG_TABLE and r.level <= table_rec.level:
            break
        if expected_cells is not None and cell_no >= expected_cells:
            break
        if r.tag_id == HWPTAG_LIST_HEADER:
            cell_meta = self._parse_list_header_as_cell(r.payload, cell_no)
            cell_no += 1
            k = j + 1
            cell_records: List[HwpRecord] = []
            while k < len(records):
                nr = records[k]
                if nr.level < table_rec.level:
                    break
                if k != j + 1 and nr.tag_id == HWPTAG_LIST_HEADER and nr.level == r.level:
                    break
                if k != j + 1 and nr.tag_id == HWPTAG_TABLE and nr.level <= table_rec.level:
                    break
                cell_records.append(nr)
                k += 1
            cell, cell_mixed = self._parse_cell_records(cell_records, cell_meta, section)
            rr = int(cell.row if cell.row is not None else 0)
            cc = int(cell.col if cell.col is not None else 0)
            rows.setdefault(rr, {})[cc] = cell
            raw_cells.append({
                "row": rr,
                "col": cc,
                "row_span": cell.row_span,
                "col_span": cell.col_span,
                "text": cell.text or "",
                "text_length": len(cell.text or ""),
            })
            for item in cell_mixed:
                item.setdefault("row", rr)
                item.setdefault("col", cc)
                mixed.append(item)
            j = k
            if expected_cells is not None and cell_no >= expected_cells:
                break
            continue
        j += 1

    table_rows = self._materialize_rows(rows, row_count, col_count)
    block = Block(
        "table",
        unique_id("tbl"),
        self._next_order(),
        section=section,
        caption=None,
        rows=table_rows,
        raw={
            "record": table_rec.to_meta(),
            "binary_table_parser": "v23",
            "row_count_declared": row_count,
            "col_count_declared": col_count,
            "expected_cell_count_v23": expected_cells,
            "actual_cell_count_v23": cell_no,
            "cell_count": sum(len(r) for r in table_rows),
            "cells": raw_cells,
            "mixed_content_count": len(mixed),
            "mixed_content": mixed,
            "payload_prefix_hex": table_rec.payload[:96].hex(),
            "complete_table_raw_export_v23": True,
        },
    )
    # Do not lift child-cell captions to the parent table. Captions are paired
    # locally inside the cell content stream by v23.
    return block, j


BinaryHwpParser._parse_cell_records = _v23_parse_cell_records  # type: ignore[method-assign]
BinaryHwpParser._parse_table_group = _v23_parse_table_group  # type: ignore[method-assign]
_v22_relocate_captions_in_content = _v23_relocate_captions_in_content  # type: ignore[assignment]
_v22_relocate_captions_in_table_item = _v23_relocate_captions_in_table_item  # type: ignore[assignment]
_v21_relocate_captions_in_content = _v23_relocate_captions_in_content  # type: ignore[assignment]
_v21_relocate_captions_in_table_item = _v23_relocate_captions_in_table_item  # type: ignore[assignment]
_v19_relocate_captions_in_content = _v23_relocate_captions_in_content  # type: ignore[assignment]
_v19_relocate_captions_in_table_item = _v23_relocate_captions_in_table_item  # type: ignore[assignment]

_old_binary_postprocess_v23 = _old_binary_postprocess_v22 if '_old_binary_postprocess_v22' in globals() else BinaryHwpParser._postprocess


def _v23_binary_postprocess(self: "BinaryHwpParser") -> None:
    # Call the pre-v22 postprocess path, but since v21/v19 relocation symbols now
    # point to v23, it will not apply the obsolete immediate/prefix-only policy.
    _old_binary_postprocess_v23(self)
    for b in self.blocks:
        if b.type == "table" and b.rows:
            for row in b.rows or []:
                for cell in row or []:
                    content = getattr(cell, "content", None) or []
                    if content:
                        cell.content = _v23_relocate_captions_in_content(content)
                        cell.text = _content_plain_text_v21(cell.content) if '_content_plain_text_v21' in globals() else cell.text
    self.metadata["binary_generalized_caption_pairing_v23"] = True
    self.metadata["complete_table_raw_export_v23"] = True
    for i, b in enumerate(self.blocks, 1):
        b.order = i


BinaryHwpParser._postprocess = _v23_binary_postprocess  # type: ignore[method-assign]

try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v22", "HWP Parser Verification UI v23")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v21", "HWP Parser Verification UI v23")
    VIEWER_HTML = VIEWER_HTML.replace("max-height:360px", "max-height:none")
except Exception:
    pass


# =============================================================================
# v25 generic HWP built-in caption recovery
# =============================================================================
# v24 could detect ordinary text captions only when their visible text looked like
# "표."/"그림." or when pyhwp exposed caption metadata. Real Hancom users often
# use Insert Caption and then type only the title, e.g. "Core-Spoke 공간배치도".
# In binary fallback such caption text is frequently stored as PARA_TEXT inside
# the GSO/picture control group. Earlier versions consumed that group while
# failing caption_from_text(), so the caption disappeared. v25 recovers those
# title-only caption paragraphs from picture/table control-local scope and also
# adds a conservative top-level adjacent pairing pass.

CAPTION_GENERIC_GSO_MAX_CHARS_V25 = 180
CAPTION_GENERIC_ADJACENT_MAX_CHARS_V25 = 160


def _v25_is_section_heading_like(text: str) -> bool:
    s = normalize_space(text)
    if not s:
        return True
    # Section/list headings should remain paragraphs, not captions.
    if re.match(r"^(?:□|■|◆|◇|○|●|<\s*[^>]+\s*>|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.|[0-9]+[.)]|[가-힣]\)|\([0-9]+\)|\([가-힣]\))", s):
        return True
    if re.match(r"^(?:제\s*\d+\s*[장절항]|[0-9]+\.[0-9]+(?:\.[0-9]+)*\s+)", s):
        return True
    return False


def _v25_is_generic_title_caption(text: str, target_type: Optional[str] = None, structural: bool = False) -> bool:
    """Return True for title-only captions without visible 표/그림 marker.

    The function is intentionally stricter outside structural caption scope.
    Inside a GSO/picture caption group, short noun-phrase text is strong evidence
    because ordinary body paragraphs are rarely stored inside the picture group.
    """
    s = normalize_space(text)
    if not s:
        return False
    if len(s) > CAPTION_GENERIC_GSO_MAX_CHARS_V25:
        return False
    lines = [ln.strip() for ln in str(text or "").splitlines() if ln.strip()]
    if len(lines) > 3:
        return False
    if _v25_is_section_heading_like(s):
        return False
    if CAPTION_PATTERN.match(strip_caption_marker_prefix(s)[0]):
        return True
    # Reject obvious prose/body sentences.
    if len(s) > 90 and re.search(r"(다\.|한다\.|하였다\.|입니다\.|있다\.|된다\.|되었다\.|하였다|한다|있다|된다)$", s):
        return False
    if re.match(r"^(?:본|이|이러한|또한|따라서|그러나|그리고|먼저|다음으로|예를 들어|특히)\b", s):
        return False
    # Reject table scalar values.
    tokens = re.findall(r"[가-힣A-Za-z]+", s)
    if len(tokens) < 2:
        return False
    # Strong visual/title cues common in captions.
    cue_words = [
        "배치도", "개념도", "구조도", "구성도", "체계도", "연계도", "흐름도", "프로세스", "로드맵",
        "구조", "구성", "체계", "모델", "시스템", "프레임워크", "화면", "결과", "분석", "비교", "성능",
        "architecture", "layout", "diagram", "framework", "system", "model", "map", "overview", "pipeline",
    ]
    if any(c.lower() in s.lower() for c in cue_words):
        return True
    # In structural GSO caption scope, a short nominal phrase is enough.
    if structural and len(s) <= 120 and len(tokens) >= 2 and not re.search(r"[.!?。！？]", s):
        return True
    # Adjacent captions without structural metadata need some cue to avoid
    # swallowing headings or normal short paragraphs.
    return False


def _v25_make_generic_caption(text: str, target_type: str, method: str, position: str = "after", raw: Optional[Dict[str, Any]] = None) -> Caption:
    explicit = caption_from_text(text, target_type)
    if explicit:
        explicit.method = method
        explicit.position = position
        explicit.raw = {**(explicit.raw or {}), **(raw or {}), "v25_explicit_caption": True}
        return explicit
    return Caption(
        text=normalize_space(text),
        method=method,
        position=position,
        raw={"v25_generic_caption_without_visible_label": True, **(raw or {})},
    )


def _v25_record_texts(parser: "BinaryHwpParser", records: Sequence[HwpRecord]) -> List[Tuple[HwpRecord, str]]:
    out: List[Tuple[HwpRecord, str]] = []
    for rec in records:
        if rec.tag_id != HWPTAG_PARA_TEXT:
            continue
        try:
            text, _ = parser.decode_para_text(rec.payload)
        except Exception:
            continue
        text = normalize_space(text)
        if text and not parser._is_control_only_text(text):
            out.append((rec, text))
    return out


def _v25_caption_for_image_from_records(parser: "BinaryHwpParser", records: Sequence[HwpRecord]) -> Optional[Caption]:
    """Recover both explicit and title-only captions inside an image/GSO group."""
    para_texts = _v25_record_texts(parser, records)

    # First, preserve explicit figure captions.
    for rec, text in para_texts:
        cap = caption_from_text(text, "image")
        if cap:
            cap.method = "binary-gso-structural-explicit-v25"
            cap.position = cap.position or "after"
            cap.raw = {**(cap.raw or {}), "source_record": rec.to_meta(), "gso_group_caption_recovery_v25": True}
            return cap

    # Then, handle Hancom Insert Caption where the visible text is title-only.
    # Prefer the last title-like paragraph in the group because caption lists are
    # usually serialized after the picture shape records.
    candidates: List[Tuple[HwpRecord, str]] = []
    for rec, text in para_texts:
        if _v25_is_generic_title_caption(text, "image", structural=True):
            candidates.append((rec, text))
    if candidates:
        rec, text = candidates[-1]
        return _v25_make_generic_caption(
            text,
            "image",
            "binary-gso-structural-title-only-v25",
            "after",
            {"source_record": rec.to_meta(), "gso_group_caption_recovery_v25": True},
        )
    return None


# _v16_parse_gso_group resolves the caption function through the global name at
# call time. Rebinding this name updates both top-level and cell-level GSO parsing
# without rewriting the full parser.
_v16_caption_for_image_from_records = _v25_caption_for_image_from_records  # type: ignore[assignment]


def _v25_generic_caption_candidate_from_block(block: "Block", target_type: str) -> Optional[Caption]:
    if block.type != "paragraph" or not block.text:
        return None
    text = normalize_space(block.text)
    explicit = caption_from_text(text, target_type)
    if explicit:
        explicit.method = "adjacent-explicit-paragraph-v25"
        return explicit
    if _v25_is_generic_title_caption(text, target_type, structural=False):
        return _v25_make_generic_caption(
            text,
            target_type,
            "adjacent-title-only-paragraph-v25",
            raw={"source_block_id": block.id, "source_order": block.order},
        )
    return None


def _v25_between_has_hard_boundary(blocks: List["Block"], a: int, b: int, target_type: str) -> bool:
    lo, hi = sorted((a, b))
    text_chars = 0
    for k in range(lo + 1, hi):
        typ = blocks[k].type
        if typ in {"table", "image", "equation"}:
            return True
        if typ == "paragraph":
            t = normalize_space(blocks[k].text or "")
            # Section headings are hard boundaries.
            if _v25_is_section_heading_like(t):
                return True
            text_chars += len(t)
            if text_chars > 260:
                return True
    return False


def _v25_resolve_adjacent_title_only_captions(self: "BinaryHwpParser") -> None:
    """Attach title-only paragraph captions adjacent to table/image blocks.

    This pass is deliberately local: it does not jump across another object or a
    section heading. It handles cases where binary parsing emits the HWP caption
    as a standalone paragraph rather than as a GSO child.
    """
    blocks = list(self.blocks)
    consumed: set[int] = set()

    for obj_idx, b in enumerate(blocks):
        if b.type not in {"image", "table", "equation"} or b.caption is not None:
            continue
        target_type = "image" if b.type == "image" else b.type
        candidates: List[Tuple[float, int, str, Caption]] = []
        # Search both sides, but only very locally.
        for cap_idx in (obj_idx - 1, obj_idx + 1, obj_idx - 2, obj_idx + 2):
            if cap_idx < 0 or cap_idx >= len(blocks) or cap_idx in consumed:
                continue
            cap_block = blocks[cap_idx]
            cap = _v25_generic_caption_candidate_from_block(cap_block, target_type)
            if not cap:
                continue
            if _v25_between_has_hard_boundary(blocks, cap_idx, obj_idx, target_type):
                continue
            dist = abs(cap_idx - obj_idx)
            position = "before" if cap_idx < obj_idx else "after"
            cap.position = position
            score = 100.0 - dist * 12.0
            # Prefer after-caption for images, before-caption for tables, but do
            # not make this absolute.
            if target_type == "image" and position == "after":
                score += 8
            if target_type == "table" and position == "before":
                score += 8
            candidates.append((score, cap_idx, position, cap))
        if not candidates:
            continue
        candidates.sort(key=lambda x: x[0], reverse=True)
        score, cap_idx, position, cap = candidates[0]
        cap.raw = {**(cap.raw or {}), "adjacent_pairing_score_v25": score, "object_id": b.id}
        b.caption = cap
        consumed.add(cap_idx)

    if consumed:
        self.blocks = [b for i, b in enumerate(blocks) if i not in consumed]
        for i, b in enumerate(self.blocks, 1):
            b.order = i


_old_binary_postprocess_v25 = BinaryHwpParser._postprocess


def _v25_binary_postprocess(self: "BinaryHwpParser") -> None:
    _old_binary_postprocess_v25(self)
    # The old postprocess may have re-ordered objects; run v25 local recovery at
    # the end so it sees final block adjacency.
    try:
        self._resolve_adjacent_title_only_captions()
    except Exception as e:
        self.warnings.append(f"v25 adjacent generic caption recovery failed: {e}")
    self.metadata["title_only_caption_recovery_v25"] = True
    self.metadata["gso_group_generic_caption_recovery_v25"] = True


BinaryHwpParser._resolve_adjacent_title_only_captions = _v25_resolve_adjacent_title_only_captions  # type: ignore[attr-defined]
BinaryHwpParser._postprocess = _v25_binary_postprocess  # type: ignore[method-assign]

try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v24", "HWP Parser Verification UI v26")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v23", "HWP Parser Verification UI v26")
    VIEWER_HTML = VIEWER_HTML.replace("max-height:360px", "max-height:none")
except Exception:
    pass



# =============================================================================
# v26: table cell fill/background color extraction from HWP binary BorderFill
# =============================================================================
# v26 fixes visual-information loss in tables where empty cells encode meaning
# using background colors, for example schedule bars. HWP stores the visible
# fill in DocInfo BorderFill records, while each table cell references a
# borderfill_id from its LIST_HEADER. Earlier versions parsed only row/column
# spans and text, so colored cells were rendered as empty gray cells.

# Correct DocInfo tag id for BORDER_FILL. Earlier constants in this file are
# intentionally left untouched for backward compatibility because the section
# record constants are already used in many monkey patches.
HWPTAG_BORDER_FILL_DOCINFO_V26 = 0x010 + 4  # 20
TAG_NAMES[HWPTAG_BORDER_FILL_DOCINFO_V26] = "BORDER_FILL"


def _v26_colorref_to_hex(value: int) -> Optional[str]:
    """Convert HWP COLORREF/BGR-ish integer to CSS #rrggbb.

    HWP uses COLORREF-like values. 0xffffffff is commonly used for white/auto
    in BorderFill color-pattern backgrounds; for a table renderer, white is the
    safest visible interpretation.
    """
    try:
        value = int(value) & 0xFFFFFFFF
    except Exception:
        return None
    if value == 0xFFFFFFFF:
        return "#ffffff"
    r = value & 0xFF
    g = (value >> 8) & 0xFF
    b = (value >> 16) & 0xFF
    return f"#{r:02x}{g:02x}{b:02x}"


def _v26_luminance(hex_color: str) -> float:
    try:
        s = hex_color.lstrip('#')
        r, g, b = int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16)
        return 0.2126 * r + 0.7152 * g + 0.0722 * b
    except Exception:
        return 255.0


def _v26_parse_border_fill_payload(payload: bytes, index0: int) -> Dict[str, Any]:
    """Best-effort parser for DocInfo/BORDER_FILL.

    According to the HWP5 model used by pyhwp, BorderFill is:
      UINT16 borderflags
      Border left/right/top/bottom/diagonal, each UINT8 stroke + UINT8 width + COLORREF
      UINT32 fillflags
      optional FillColorPattern: background_color, pattern_color, pattern_type
    This is enough to recover table cell background colors.
    """
    result: Dict[str, Any] = {"index0": index0, "id": index0 + 1, "size": len(payload)}
    try:
        if len(payload) < 36:
            result["parse_error"] = "payload shorter than minimum BorderFill header"
            result["payload_hex"] = payload.hex()
            return result
        o = 0
        result["border_flags"] = int.from_bytes(payload[o:o+2], "little", signed=False)
        o += 2
        borders = []
        for side in ("left", "right", "top", "bottom", "diagonal"):
            if o + 6 > len(payload):
                break
            stroke = payload[o]
            width = payload[o + 1]
            color_val = int.from_bytes(payload[o+2:o+6], "little", signed=False)
            o += 6
            borders.append({"side": side, "stroke": stroke, "width": width, "color": _v26_colorref_to_hex(color_val), "color_raw": color_val})
        result["borders"] = borders
        if o + 4 > len(payload):
            return result
        fillflags = int.from_bytes(payload[o:o+4], "little", signed=False)
        o += 4
        result["fill_flags"] = fillflags
        result["has_color_pattern"] = bool(fillflags & 0x01)
        result["has_image_fill"] = bool(fillflags & 0x02)
        result["has_gradation"] = bool(fillflags & 0x04)
        if fillflags & 0x01 and o + 12 <= len(payload):
            bg_raw = int.from_bytes(payload[o:o+4], "little", signed=False)
            pattern_raw = int.from_bytes(payload[o+4:o+8], "little", signed=False)
            pattern_type = int.from_bytes(payload[o+8:o+12], "little", signed=False)
            bg = _v26_colorref_to_hex(bg_raw)
            result["background_color"] = bg
            result["background_color_raw"] = bg_raw
            result["pattern_color"] = _v26_colorref_to_hex(pattern_raw)
            result["pattern_color_raw"] = pattern_raw
            result["pattern_type"] = pattern_type
            if bg:
                result["css_style"] = f"background-color:{bg};"
                if _v26_luminance(bg) < 96:
                    result["css_style"] += "color:#ffffff;"
        return result
    except Exception as e:
        result["parse_error"] = str(e)
        result["payload_hex"] = payload[:128].hex()
        return result


def _v26_decompress_hwp_stream_for_binary_parser(self: "BinaryHwpParser", raw: bytes) -> bytes:
    if self.metadata.get("compressed"):
        for wbits in (-15, zlib.MAX_WBITS):
            try:
                return zlib.decompress(raw, wbits)
            except Exception:
                pass
    return raw


def _v26_parse_docinfo_styles(self: "BinaryHwpParser") -> None:
    """Read DocInfo/BorderFill style records for table cell backgrounds.

    This function is deliberately independent of pyhwp so binary mode can still
    work when hwp5proc times out.
    """
    self.border_fills_v26 = {}
    self.border_fills_by_zero_index_v26 = {}
    try:
        if not hasattr(self, "ole") or not self.ole.exists("DocInfo"):
            return
        raw = self.ole.openstream("DocInfo").read()
        data = _v26_decompress_hwp_stream_for_binary_parser(self, raw)
        bf_index = 0
        for rec in self.iter_records(data, section=None):
            if rec.tag_id != HWPTAG_BORDER_FILL_DOCINFO_V26:
                continue
            parsed = _v26_parse_border_fill_payload(rec.payload, bf_index)
            parsed["record"] = rec.to_meta()
            # HWP table cell LIST_HEADER borderfill_id is 1-based in the files
            # verified here. Store both maps defensively.
            self.border_fills_by_zero_index_v26[bf_index] = parsed
            self.border_fills_v26[bf_index + 1] = parsed
            bf_index += 1
        self.metadata["border_fill_count_v26"] = bf_index
        self.metadata["table_cell_background_color_support_v26"] = True
    except Exception as e:
        self.warnings.append(f"v26 DocInfo BorderFill parsing failed: {e}")


def _v26_resolve_border_fill_style(self: "BinaryHwpParser", border_fill_id: Optional[int]) -> Dict[str, Any]:
    if border_fill_id is None:
        return {}
    try:
        bfid = int(border_fill_id)
    except Exception:
        return {}
    style = getattr(self, "border_fills_v26", {}).get(bfid)
    if not style:
        # Defensive fallback for rare zero-based references.
        style = getattr(self, "border_fills_by_zero_index_v26", {}).get(bfid)
    return dict(style or {})


_old_binary_parse_v26 = BinaryHwpParser.parse


def _v26_binary_parse(self: "BinaryHwpParser") -> ParsedDocument:
    # Reimplement parse so DocInfo styles are available before table cells are parsed.
    if olefile is None:
        raise RuntimeError("olefile is required for binary parsing. Install: pip install olefile")
    if not olefile.isOleFile(str(self.input_path)):
        raise RuntimeError("not an OLE/CFB HWP file")
    with olefile.OleFileIO(str(self.input_path)) as ole:
        self.ole = ole
        self._parse_header()
        self._parse_docinfo_styles_v26()
        self._parse_sections()
    self._postprocess()
    return ParsedDocument(
        source_path=str(self.input_path),
        method="binary_structured_fallback_v26",
        blocks=self.blocks,
        media_items=self.media_store.items,
        metadata=self.metadata,
        warnings=self.warnings,
        errors=self.errors,
    )


_old_parse_list_header_as_cell_v26 = BinaryHwpParser._parse_list_header_as_cell


def _v26_parse_list_header_as_cell(self: "BinaryHwpParser", payload: bytes, seq: int) -> Dict[str, Any]:
    meta = _old_parse_list_header_as_cell_v26(self, payload, seq)
    try:
        if len(payload) >= 34:
            border_fill_id = int.from_bytes(payload[32:34], "little", signed=False)
            meta["border_fill_id"] = border_fill_id
            bf = self._resolve_border_fill_style_v26(border_fill_id)
            if bf:
                meta["border_fill"] = {k: v for k, v in bf.items() if k not in {"borders"}}
                if bf.get("background_color"):
                    meta["background_color"] = bf.get("background_color")
                if bf.get("css_style"):
                    meta["css_style"] = bf.get("css_style")
        meta["list_header_payload_len"] = len(payload)
    except Exception as e:
        meta["border_fill_parse_error_v26"] = str(e)
    return meta


def _v26_parse_table_payload_visual_info(self: "BinaryHwpParser", payload: bytes) -> Dict[str, Any]:
    """Parse TABLE payload enough to recover default borderfill and zones.

    Zones are important because some HWP tables apply a fill to a rectangular
    range instead of assigning a distinct borderfill to each LIST_HEADER cell.
    """
    out: Dict[str, Any] = {}
    try:
        if len(payload) < 18:
            return out
        o = 0
        flags = int.from_bytes(payload[o:o+4], "little", signed=False); o += 4
        rows = int.from_bytes(payload[o:o+2], "little", signed=False); o += 2
        cols = int.from_bytes(payload[o:o+2], "little", signed=False); o += 2
        cellspacing = int.from_bytes(payload[o:o+2], "little", signed=False); o += 2
        padding = [int.from_bytes(payload[o+i*2:o+i*2+2], "little", signed=False) for i in range(4)]
        o += 8
        rowcols: List[int] = []
        if 0 < rows <= 2000 and o + rows * 2 <= len(payload):
            rowcols = [int.from_bytes(payload[o+i*2:o+i*2+2], "little", signed=False) for i in range(rows)]
            o += rows * 2
        default_bfid = None
        if o + 2 <= len(payload):
            default_bfid = int.from_bytes(payload[o:o+2], "little", signed=False)
            o += 2
        zones: List[Dict[str, Any]] = []
        if o + 2 <= len(payload):
            zone_count = int.from_bytes(payload[o:o+2], "little", signed=False)
            o += 2
            if 0 <= zone_count <= 10000 and o + zone_count * 10 <= len(payload):
                for _ in range(zone_count):
                    sc = int.from_bytes(payload[o:o+2], "little", signed=False)
                    sr = int.from_bytes(payload[o+2:o+4], "little", signed=False)
                    ec = int.from_bytes(payload[o+4:o+6], "little", signed=False)
                    er = int.from_bytes(payload[o+6:o+8], "little", signed=False)
                    bfid = int.from_bytes(payload[o+8:o+10], "little", signed=False)
                    o += 10
                    bf = self._resolve_border_fill_style_v26(bfid)
                    zone = {"starting_column": sc, "starting_row": sr, "end_column": ec, "end_row": er, "border_fill_id": bfid}
                    if bf.get("background_color"):
                        zone["background_color"] = bf.get("background_color")
                    if bf.get("css_style"):
                        zone["css_style"] = bf.get("css_style")
                    zones.append(zone)
        out.update({
            "flags": flags,
            "rows": rows,
            "cols": cols,
            "cellspacing": cellspacing,
            "padding": padding,
            "rowcols": rowcols,
            "default_border_fill_id": default_bfid,
            "valid_zones": zones,
        })
        if default_bfid is not None:
            bf = self._resolve_border_fill_style_v26(default_bfid)
            if bf.get("background_color"):
                out["default_background_color"] = bf.get("background_color")
            if bf.get("css_style"):
                out["default_css_style"] = bf.get("css_style")
    except Exception as e:
        out["parse_error"] = str(e)
    return out


def _v26_apply_table_zone_styles(table_rows: List[List[TableCell]], visual: Dict[str, Any]) -> None:
    zones = visual.get("valid_zones") or []
    if not zones:
        return
    for zone in zones:
        css = zone.get("css_style")
        bg = zone.get("background_color")
        bfid = zone.get("border_fill_id")
        if not css and not bg:
            continue
        try:
            sr, er = int(zone.get("starting_row", 0)), int(zone.get("end_row", -1))
            sc, ec = int(zone.get("starting_column", 0)), int(zone.get("end_column", -1))
        except Exception:
            continue
        for row in table_rows:
            for cell in row:
                r = int(cell.row or 0)
                c = int(cell.col or 0)
                if sr <= r <= er and sc <= c <= ec:
                    meta = cell.attrs.setdefault("binary_cell_meta", {})
                    meta.setdefault("zone_border_fill_id", bfid)
                    if bg and not meta.get("background_color"):
                        meta["background_color"] = bg
                    if css and not meta.get("css_style"):
                        meta["css_style"] = css
                    meta["style_from_table_zone_v26"] = True


_old_v23_parse_table_group_v26 = BinaryHwpParser._parse_table_group


def _v26_parse_table_group(self: "BinaryHwpParser", records: List[HwpRecord], start: int, section: Optional[int]) -> Tuple[Block, int]:
    block, next_i = _old_v23_parse_table_group_v26(self, records, start, section)
    try:
        table_rec = records[start]
        visual = self._parse_table_payload_visual_info_v26(table_rec.payload)
        if visual:
            block.raw.setdefault("table_visual_info_v26", visual)
            if block.rows:
                _v26_apply_table_zone_styles(block.rows, visual)
        # Count cells with visible background for diagnostics.
        styled = 0
        for row in block.rows or []:
            for cell in row:
                meta = (cell.attrs or {}).get("binary_cell_meta", {})
                if meta.get("background_color") or meta.get("css_style"):
                    styled += 1
        if styled:
            block.raw["styled_cell_count_v26"] = styled
    except Exception as e:
        block.raw.setdefault("table_visual_parse_error_v26", str(e))
    return block, next_i


BinaryHwpParser._parse_docinfo_styles_v26 = _v26_parse_docinfo_styles  # type: ignore[attr-defined]
BinaryHwpParser._resolve_border_fill_style_v26 = _v26_resolve_border_fill_style  # type: ignore[attr-defined]
BinaryHwpParser.parse = _v26_binary_parse  # type: ignore[method-assign]
BinaryHwpParser._parse_list_header_as_cell = _v26_parse_list_header_as_cell  # type: ignore[method-assign]
BinaryHwpParser._parse_table_payload_visual_info_v26 = _v26_parse_table_payload_visual_info  # type: ignore[attr-defined]
BinaryHwpParser._parse_table_group = _v26_parse_table_group  # type: ignore[method-assign]

# UI support: render extracted background colors in table cells.
try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v26", "HWP Parser Verification UI v26")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v24", "HWP Parser Verification UI v26")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v23", "HWP Parser Verification UI v26")
    old_js = """function tableHtml(rows){
  if(!rows||!rows.length) return '<div class="empty">표 블록은 감지됐지만 행/열 구조가 비어 있습니다. raw 진단 정보를 확인하세요.</div>';
  return `<div class="table-wrap"><table class="extracted"><tbody>${rows.map(r=>`<tr>${(r||[]).map(c=>`<td rowspan="${esc(c.row_span||1)}" colspan="${esc(c.col_span||1)}" class="${(c.text||(c.content||[]).length)?'':'empty'}">${cellHtml(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}
"""
    new_js = """function tableCellStyle(c){
  const a=(c&&c.attrs)||{}; const meta=a.binary_cell_meta||{};
  const parts=[];
  const css=meta.css_style||a.css_style||'';
  if(css) parts.push(String(css).replace(/"/g,''));
  const bg=meta.background_color||a.background_color||a['background-color']||'';
  if(bg && !parts.join(';').includes('background')) parts.push('background-color:'+String(bg).replace(/"/g,''));
  const fg=meta.color||a.color||'';
  if(fg && !parts.join(';').includes('color:')) parts.push('color:'+String(fg).replace(/"/g,''));
  return parts.length ? ` style="${esc(parts.join(';'))}"` : '';
}
function tableHtml(rows){
  if(!rows||!rows.length) return '<div class="empty">표 블록은 감지됐지만 행/열 구조가 비어 있습니다. raw 진단 정보를 확인하세요.</div>';
  return `<div class="table-wrap"><table class="extracted"><tbody>${rows.map(r=>`<tr>${(r||[]).map(c=>`<td rowspan="${esc(c.row_span||1)}" colspan="${esc(c.col_span||1)}" class="${(c.text||(c.content||[]).length)?'':'empty'}"${tableCellStyle(c)}>${cellHtml(c)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;
}
"""
    if old_js in VIEWER_HTML:
        VIEWER_HTML = VIEWER_HTML.replace(old_js, new_js)
    else:
        VIEWER_HTML = VIEWER_HTML.replace("function tableHtml(rows){", new_js + "\nfunction tableHtml_old_unused_v26(rows){")
except Exception:
    pass


# =============================================================================
# v27: reject footnotes/annotations as captions and preserve them as paragraphs
# =============================================================================
# v25/v26 allowed title-only adjacent paragraphs to become captions. That was
# necessary for real HWP Insert-Caption cases where the visible text is only a
# title, but it was too permissive for footnotes immediately following a table.

def _v27_is_footnote_or_annotation(text: str) -> bool:
    s = normalize_space(text)
    if not s:
        return False
    if re.match(r"^\s*(?:[\*＊※]\s*)?(?:\d+\s*[\)\]】」]|[①-⑳])\s*", s):
        return True
    if re.match(r"^\s*(?:[\*＊※]|주\s*[:：\)]|註\s*[:：\)]|Note\s*[:：\)]|NOTE\s*[:：\)])", s):
        return True
    note_cues = [
        "의미합니다", "말합니다", "합계는", "이어야", "상대적 중요도", "성능판단기준",
        "최종목표", "기술적 성능", "주석", "각주", "해당 시", "단위:", "단위 :",
    ]
    if any(cue in s for cue in note_cues) and re.match(r"^\s*(?:[\*＊※·•]\s*)?(?:\d+\s*[\)\]】」.)]|[①-⑳])", s):
        return True
    if re.match(r"^\s*(?:[\*＊※·•]\s*)", s) and re.search(r"(다\.|합니다\.|됩니다\.|말하며|의미)", s):
        return True
    return False


def _v27_is_bad_caption_text(text: str, method: str = "") -> bool:
    s = normalize_space(text)
    if not s:
        return True
    if _v27_is_footnote_or_annotation(s):
        return True
    m = str(method or "")
    if ("adjacent-title-only" in m or "generic" in m) and re.search(r"(의미합니다|말하며|합계는|이어야|기준이 되는 것)", s):
        return True
    return False


_old_v23_plain_title_like_v27 = _v23_plain_title_like
def _v27_plain_title_like(text: str) -> bool:
    if _v27_is_footnote_or_annotation(text):
        return False
    return _old_v23_plain_title_like_v27(text)
_v23_plain_title_like = _v27_plain_title_like  # type: ignore[assignment]

_old_v25_is_generic_title_caption_v27 = _v25_is_generic_title_caption
def _v27_is_generic_title_caption(text: str, target_type: Optional[str] = None, structural: bool = False) -> bool:
    if _v27_is_footnote_or_annotation(text):
        return False
    return _old_v25_is_generic_title_caption_v27(text, target_type, structural)
_v25_is_generic_title_caption = _v27_is_generic_title_caption  # type: ignore[assignment]


def _v27_caption_dict_is_bad(cap: Any) -> bool:
    if not isinstance(cap, dict):
        return False
    return _v27_is_bad_caption_text(str(cap.get("text") or ""), str(cap.get("method") or ""))


def _v27_caption_obj_is_bad(cap: Any) -> bool:
    if cap is None:
        return False
    if isinstance(cap, Caption):
        return _v27_is_bad_caption_text(cap.text, cap.method)
    if isinstance(cap, dict):
        return _v27_caption_dict_is_bad(cap)
    return False


def _v27_caption_to_restored_paragraph(cap: Any, source: str) -> Dict[str, Any]:
    if isinstance(cap, Caption):
        text = cap.text
        cap_dict = cap.to_dict()
    elif isinstance(cap, dict):
        text = str(cap.get("text") or "")
        cap_dict = dict(cap)
    else:
        text = str(cap or "")
        cap_dict = {"text": text}
    return {"type": "paragraph", "text": normalize_space(text), "raw": {"restored_from_false_caption_v27": True, "source": source, "caption_raw": cap_dict}}


def _v27_sanitize_content_false_captions(content: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    for item in content or []:
        if not isinstance(item, dict):
            continue
        item = dict(item)
        if item.get("type") == "table":
            for row in item.get("rows") or []:
                for cell in row or []:
                    if not isinstance(cell, dict):
                        continue
                    c = cell.get("content") or []
                    if c:
                        new_c = _v27_sanitize_content_false_captions(c)
                        cell["content"] = new_c
                        cell["text"] = _content_plain_text_v21(new_c) if '_content_plain_text_v21' in globals() else normalize_space("\n".join(str(x.get("text", "")) for x in new_c if isinstance(x, dict)))
        cap = item.get("caption")
        if item.get("type") in {"table", "image", "equation"} and _v27_caption_dict_is_bad(cap):
            para = _v27_caption_to_restored_paragraph(cap, "inline-object-caption")
            pos = str((cap or {}).get("position") or "after") if isinstance(cap, dict) else "after"
            item.pop("caption", None)
            item.setdefault("raw", {})["false_caption_removed_v27"] = para["raw"]
            if pos in {"before", "top", "above"}:
                out.append(para); out.append(item)
            else:
                out.append(item); out.append(para)
        else:
            out.append(item)
    return out


def _v27_sanitize_table_cells(block_or_item: Any) -> None:
    rows = block_or_item.rows if isinstance(block_or_item, Block) else (block_or_item.get("rows") if isinstance(block_or_item, dict) else None)
    if not rows:
        return
    for row in rows:
        for cell in row or []:
            if isinstance(cell, TableCell):
                content = cell.content or []
                if content:
                    new_content = _v27_sanitize_content_false_captions(content)
                    cell.content = new_content
                    cell.text = _content_plain_text_v21(new_content) if '_content_plain_text_v21' in globals() else cell.text
                    cell.paragraphs = _dedupe_adjacent_strings([str(x.get("text", "")) for x in new_content if isinstance(x, dict) and x.get("type") == "paragraph" and x.get("text")])
            elif isinstance(cell, dict):
                content = cell.get("content") or []
                if content:
                    new_content = _v27_sanitize_content_false_captions(content)
                    cell["content"] = new_content
                    cell["text"] = _content_plain_text_v21(new_content) if '_content_plain_text_v21' in globals() else cell.get("text", "")


def _v27_restore_false_caption_block(cap: Caption, source_block: Block) -> Block:
    return Block("paragraph", unique_id("p"), 0, section=source_block.section, text=normalize_space(cap.text), raw={"restored_from_false_caption_v27": True, "source_block_id": source_block.id, "source_block_type": source_block.type, "caption_raw": cap.to_dict()})


def _v27_sanitize_top_level_false_captions(self: "BinaryHwpParser") -> None:
    new_blocks: List[Block] = []
    restored = 0
    for block in list(self.blocks):
        if block.type == "table":
            _v27_sanitize_table_cells(block)
        if block.type in {"table", "image", "equation"} and _v27_caption_obj_is_bad(block.caption):
            cap = block.caption
            if isinstance(cap, Caption):
                restored_block = _v27_restore_false_caption_block(cap, block)
                pos = cap.position or "after"
                block.raw.setdefault("false_caption_removed_v27", cap.to_dict())
                block.caption = None
                restored += 1
                if pos in {"before", "top", "above"}:
                    new_blocks.append(restored_block); new_blocks.append(block)
                else:
                    new_blocks.append(block); new_blocks.append(restored_block)
            else:
                new_blocks.append(block)
        else:
            new_blocks.append(block)
    if restored:
        self.blocks = new_blocks
        self.metadata["false_caption_restored_to_paragraph_count_v27"] = restored
        self.warnings.append(f"v27 restored {restored} note/footnote false caption(s) as normal paragraphs")
    for i, b in enumerate(self.blocks, 1):
        b.order = i


_old_binary_postprocess_v27 = BinaryHwpParser._postprocess
def _v27_binary_postprocess(self: "BinaryHwpParser") -> None:
    _old_binary_postprocess_v27(self)
    try:
        _v27_sanitize_top_level_false_captions(self)
    except Exception as e:
        self.warnings.append(f"v27 false-caption sanitation failed: {e}")
    self.metadata["footnote_annotation_caption_guard_v27"] = True
    self.metadata["caption_false_positive_sanitation_v27"] = True
BinaryHwpParser._postprocess = _v27_binary_postprocess  # type: ignore[method-assign]

try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v26", "HWP Parser Verification UI v28")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v25", "HWP Parser Verification UI v28")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v24", "HWP Parser Verification UI v28")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v23", "HWP Parser Verification UI v28")
except Exception:
    pass


# =============================================================================
# v28: robust angle-bracket figure/table captions and local image-caption pairing
# =============================================================================
# Real HWP documents often use Hancom's caption feature but display captions as
# "<그림 3> 제목" / "〈그림 4〉 제목" / "<표 > 제목" rather than "그림." or
# "표.". Earlier versions expected visible 표./그림. markers and therefore missed
# these captions, especially when the images were embedded inside a table cell.
# v28 adds explicit angle-bracket parsing, caption-fragment merging, and safer
# image caption pairing. It keeps v27's footnote/annotation guard.

ANGLE_CAPTION_PATTERN_V28 = re.compile(
    r"^\s*(?:caption|캡션)?\s*[:：\-–—]?\s*"
    r"[<〈《［\[\(]?\s*"
    r"(?P<label>표|그림|사진|도|수식|Table|TABLE|Fig\.?|FIG\.?|Figure|FIGURE|Equation|EQUATION)"
    r"\s*(?P<num>[0-9０-９]+|[IVXivx]+|[A-Za-z가-힣]+)?\s*"
    r"(?:[>〉》］\]\)]|[.)．:：]|[-–—])\s*"
    r"(?P<body>.*)$"
)

LABEL_ONLY_CAPTION_PATTERN_V28 = re.compile(
    r"^\s*(?:caption|캡션)?\s*[:：\-–—]?\s*"
    r"[<〈《［\[\(]?\s*"
    r"(?P<label>표|그림|사진|도|수식|Table|TABLE|Fig\.?|FIG\.?|Figure|FIGURE|Equation|EQUATION)"
    r"\s*(?P<num>[0-9０-９]+|[IVXivx]+|[A-Za-z가-힣]+)?\s*"
    r"[>〉》］\]\)]?\s*$"
)

_old_caption_from_text_v28 = caption_from_text

def _v28_label_target(label: str) -> Optional[str]:
    l = (label or "").lower()
    if l.startswith("표") or l.startswith("table"):
        return "table"
    if l.startswith("그림") or l.startswith("사진") or l.startswith("도") or l.startswith("fig") or l.startswith("figure"):
        return "image"
    if l.startswith("수식") or l.startswith("equation"):
        return "equation"
    return None


def _v28_caption_type_compatible(label: str, object_type: Optional[str]) -> bool:
    if not object_type:
        return True
    return _v28_label_target(label) == object_type


def _v28_normalize_label_text(label: str, num: str, body: str, original: str) -> str:
    # Preserve angle-bracket style when it exists in the source. This makes the
    # verification UI reflect the actual visible HWP caption form.
    s = normalize_space(original)
    if s:
        return s
    core = normalize_space(" ".join(x for x in [label, num] if x))
    return normalize_space((core + " " + (body or "")).strip())


def _v28_caption_from_text(text: str, object_type: Optional[str] = None) -> Optional[Caption]:
    s = normalize_space(text)
    if not s:
        return None
    # Preserve v27's annotation/footnote rejection before any permissive parse.
    try:
        if '_v27_is_footnote_or_annotation' in globals() and _v27_is_footnote_or_annotation(s):
            return None
    except Exception:
        pass

    # First try the existing parser for normal "그림." / "표." / "Fig." cases.
    cap = _old_caption_from_text_v28(s, object_type)
    if cap:
        return cap

    raw_lines = [ln.strip() for ln in str(text or "").splitlines() if ln.strip()]
    if not raw_lines:
        return None
    first, stripped_marker = strip_caption_marker_prefix(raw_lines[0])
    if len(first) > 240:
        return None

    m = ANGLE_CAPTION_PATTERN_V28.match(first)
    if not m:
        return None
    label = m.group("label") or ""
    num = m.group("num") or ""
    body = m.group("body") or ""
    if not _v28_caption_type_compatible(label, object_type):
        return None

    caption_lines = [_v28_normalize_label_text(label, num, body, first)]
    # Soft-line continuation. This covers Shift+Enter or HWP record variants
    # where the visible caption body is serialized as multiple physical lines.
    for extra in raw_lines[1:]:
        extra_stripped, _ = strip_caption_marker_prefix(extra)
        if not extra_stripped:
            continue
        if ANGLE_CAPTION_PATTERN_V28.match(extra_stripped) or CAPTION_PATTERN.match(extra_stripped):
            break
        # Stop at section/head/body starts, but allow short caption continuations.
        if re.match(r"^(?:[0-9]+[.)]|[①-⑳]|□|■|<\s*서식|[가-힣A-Za-z]+\s*[:：])", extra_stripped) and len(" ".join(caption_lines)) > 20:
            break
        candidate = normalize_space(" ".join(caption_lines + [extra_stripped]))
        if len(candidate) > 500:
            break
        caption_lines.append(extra_stripped)

    caption_text = normalize_space(" ".join(caption_lines))
    return Caption(
        text=caption_text,
        method="pattern-angle-v28",
        raw={
            "label": label,
            "num": num,
            "body": body,
            "stripped_caption_marker": stripped_marker,
            "angle_bracket_caption_v28": True,
            "soft_line_joined_v28": len(caption_lines) > 1,
        },
    )

# Rebind the global function used by all later parser methods.
caption_from_text = _v28_caption_from_text  # type: ignore[assignment]


def _v28_explicit_target(text: str) -> Optional[str]:
    if caption_from_text(text, "table"):
        return "table"
    if caption_from_text(text, "image"):
        return "image"
    if caption_from_text(text, "equation"):
        return "equation"
    return None


def _v28_caption_body_is_empty(text: str, target: Optional[str] = None) -> bool:
    cap = caption_from_text(text, target)
    if not cap:
        return False
    raw = cap.raw or {}
    # old CAPTION_PATTERN and new ANGLE pattern both expose body when available.
    body = normalize_space(str(raw.get("body") or ""))
    if body:
        return False
    # Also handle visible strings like "<그림 3>" or "그림 3".
    s = normalize_space(text)
    return bool(LABEL_ONLY_CAPTION_PATTERN_V28.match(strip_caption_marker_prefix(s)[0]))


def _v28_title_continuation_like(text: str, target_type: Optional[str]) -> bool:
    s = normalize_space(text)
    if not s:
        return False
    if '_v27_is_footnote_or_annotation' in globals() and _v27_is_footnote_or_annotation(s):
        return False
    if caption_from_text(s, target_type):
        return False
    # Reject obvious body sentences and list/section headings.
    if re.match(r"^(?:□|■|◆|◇|○|●|[0-9]+[.)]|[①-⑳]|[가-힣]\)|\([0-9]+\)|\([가-힣]\))", s):
        return False
    if len(s) > 180:
        return False
    if len(s) > 80 and re.search(r"(하였다|합니다|있다|된다|되었다|수행|추진|확보)\.?$", s):
        return False
    # Prefer existing generic title classifier where available.
    try:
        if '_v25_is_generic_title_caption' in globals() and _v25_is_generic_title_caption(s, target_type, structural=True):
            return True
    except Exception:
        pass
    # Additional short caption nouns common in Korean technical reports.
    cues = [
        "분석", "결과", "특성", "구조", "구성", "체계", "배치도", "개념도", "흐름도", "파이프라인",
        "시스템", "모델", "화면", "예시", "성능", "비교", "변화", "추이", "대시보드", "데이터셋",
    ]
    if any(c in s for c in cues) and len(re.findall(r"[가-힣A-Za-z]+", s)) >= 2:
        return True
    return False


def _v28_item_text(item: Dict[str, Any]) -> str:
    if not isinstance(item, dict):
        return ""
    if isinstance(item.get("caption"), dict) and item["caption"].get("text"):
        return normalize_space(str(item["caption"].get("text") or ""))
    return normalize_space(str(item.get("text") or ""))


def _v28_merge_caption_fragments(content: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Merge label-only caption fragments with the following title paragraph.

    Handles HWP record sequences such as:
      image, paragraph "<그림 3>", paragraph "열 지연 특성 분석"
    and converts them to a single caption item:
      caption "<그림 3> 열 지연 특성 분석"
    """
    if not content:
        return content
    out: List[Dict[str, Any]] = []
    i = 0
    while i < len(content):
        item = content[i]
        if not isinstance(item, dict):
            i += 1
            continue
        typ = item.get("type")
        text = _v28_item_text(item)
        target = _v28_explicit_target(text)
        if typ in {"paragraph", "caption"} and target and _v28_caption_body_is_empty(text, target):
            j = i + 1
            title_parts: List[str] = []
            consumed = 0
            # Merge up to two short continuation paragraphs. Captions sometimes
            # split at a manual line break or separate HWP para record.
            while j < len(content) and consumed < 2:
                nxt = content[j]
                if not isinstance(nxt, dict) or nxt.get("type") not in {"paragraph", "caption"}:
                    break
                nt = _v28_item_text(nxt)
                if not _v28_title_continuation_like(nt, target):
                    break
                title_parts.append(nt)
                consumed += 1
                j += 1
            if title_parts:
                merged_text = normalize_space(text + " " + " ".join(title_parts))
                cap = caption_from_text(merged_text, target) or Caption(text=merged_text, method="merged-caption-fragment-v28")
                cap.method = "merged-caption-fragment-v28"
                cap.position = "after" if target == "image" else "before"
                cap.raw = {**(cap.raw or {}), "merged_from_label_only_caption_v28": True, "label_fragment": text, "title_fragments": title_parts}
                out.append({"type": "caption", "text": cap.text, "caption": cap.to_dict(), "target_type": target, "raw": {"merged_caption_fragments_v28": True}})
                i = j
                continue
        out.append(item)
        i += 1
    return out


# Update the explicit caption detector used in v23 cell parsing and relocation.
_v23_caption_explicit_target = _v28_explicit_target  # type: ignore[assignment]

_old_v23_pair_score_v28 = _v23_pair_score

def _v28_pair_score(items: List[Dict[str, Any]], cap_idx: int, obj_idx: int, target_type: str, cand: Dict[str, Any]) -> Optional[Tuple[float, str]]:
    scored = _old_v23_pair_score_v28(items, cap_idx, obj_idx, target_type, cand)
    if scored is None:
        return None
    score, direction = scored
    # v23 preferred captions before the following object. That is reasonable for
    # many table captions, but it is wrong for common figure layouts where the
    # caption sits below the image: image -> caption -> image -> caption.
    # Prefer the immediately preceding image unless there is no preceding image.
    if target_type == "image":
        if direction == "after":
            score += 14.0
        else:
            # A before-image caption remains valid, but should not beat an
            # equally close previous-image candidate.
            score -= 4.0
    return score, direction

_v23_pair_score = _v28_pair_score  # type: ignore[assignment]

_old_v23_relocate_captions_in_content_v28 = _v23_relocate_captions_in_content

def _v28_relocate_captions_in_content(content: List[Dict[str, Any]], max_window: int = 40) -> List[Dict[str, Any]]:
    merged = _v28_merge_caption_fragments(content)
    return _old_v23_relocate_captions_in_content_v28(merged, max_window=max_window)

_v23_relocate_captions_in_content = _v28_relocate_captions_in_content  # type: ignore[assignment]
_v22_relocate_captions_in_content = _v28_relocate_captions_in_content  # type: ignore[assignment]
_v21_relocate_captions_in_content = _v28_relocate_captions_in_content  # type: ignore[assignment]
_v19_relocate_captions_in_content = _v28_relocate_captions_in_content  # type: ignore[assignment]

# Rebind GSO caption recovery explicitly, because it calls caption_from_text via
# the global name and now benefits from angle-bracket captions.
try:
    _v16_caption_for_image_from_records = _v25_caption_for_image_from_records  # type: ignore[assignment]
except Exception:
    pass

_old_v27_binary_postprocess_v28 = BinaryHwpParser._postprocess

def _v28_binary_postprocess(self: "BinaryHwpParser") -> None:
    _old_v27_binary_postprocess_v28(self)
    # Apply one more pass over all table cell content after v27 false-caption
    # sanitation. This catches table cells built before the v28 global rebinding
    # and nested tables in raw content streams.
    try:
        for block in self.blocks:
            if block.type == "table":
                def walk_rows(rows):
                    for row in rows or []:
                        for cell in row or []:
                            if isinstance(cell, TableCell):
                                if cell.content:
                                    cell.content = _v28_relocate_captions_in_content(cell.content)
                                    cell.text = _content_plain_text_v21(cell.content) if '_content_plain_text_v21' in globals() else cell.text
                                    cell.paragraphs = _dedupe_adjacent_strings([str(x.get("text", "")) for x in cell.content if isinstance(x, dict) and x.get("type") == "paragraph" and x.get("text")])
                                for it in cell.content or []:
                                    if isinstance(it, dict) and it.get("type") == "table":
                                        # dict rows in inline nested table
                                        for r in it.get("rows") or []:
                                            for c in r or []:
                                                if isinstance(c, dict) and c.get("content"):
                                                    c["content"] = _v28_relocate_captions_in_content(c.get("content") or [])
                                                    c["text"] = _content_plain_text_v21(c["content"]) if '_content_plain_text_v21' in globals() else c.get("text", "")
                    return rows
                walk_rows(block.rows)
        self.metadata["angle_bracket_caption_support_v28"] = True
        self.metadata["image_caption_after_image_preference_v28"] = True
    except Exception as e:
        self.warnings.append(f"v28 caption postprocess failed: {e}")
    for i, b in enumerate(self.blocks, 1):
        b.order = i

BinaryHwpParser._postprocess = _v28_binary_postprocess  # type: ignore[method-assign]

try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v27", "HWP Parser Verification UI v28")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v26", "HWP Parser Verification UI v28")
except Exception:
    pass




# =============================================================================
# v29: delimiter-agnostic / structural caption detection
# =============================================================================
# v28 recognized angle-bracket captions such as "<그림 3> ...". Real HWP
# documents use many visible renderings for the same Insert Caption feature:
#   <그림 3>, {그림 3}, [그림 3], (그림 3), 〈그림 3〉, 【그림 3】, ...
# and frequently omit the visible marker entirely when the binary/XML structure
# already marks the text as a caption. v29 therefore separates caption parsing
# into two layers:
#   1) delimiter-agnostic explicit label recognition using normalized leading
#      punctuation instead of a single bracket literal;
#   2) structural/title-only caption recognition only inside a trusted object
#      scope or immediate local object adjacency.

_old_caption_from_text_v29 = caption_from_text

_CAPTION_OPENERS_V29 = set("<[{(（［｛〈《「『【〔〖﹤＜｢«‹")
_CAPTION_CLOSERS_V29 = set(">]})）］｝〉》」』】〕〗﹥＞｣»›")
_CAPTION_SEPARATORS_V29 = set(" .．。,:：;；-–—_·ㆍ)）]］}｝>＞〉》」』】〕〗")
_CAPTION_LEADING_SKIP_V29 = _CAPTION_OPENERS_V29 | set(" \t\r\n\ufeff\u200b\u3000")
_TABLE_LABELS_V29 = ["표", "table", "tbl"]
_IMAGE_LABELS_V29 = ["그림", "사진", "도", "figure", "fig", "image", "img", "picture", "pic"]
_EQUATION_LABELS_V29 = ["수식", "equation", "eq"]
_ALL_LABELS_V29 = sorted(_TABLE_LABELS_V29 + _IMAGE_LABELS_V29 + _EQUATION_LABELS_V29, key=len, reverse=True)


def _v29_is_hangul(ch: str) -> bool:
    return bool(ch and ("가" <= ch <= "힣"))


def _v29_is_ascii_alpha(ch: str) -> bool:
    return bool(ch and ("a" <= ch.lower() <= "z"))


def _v29_label_target(label: str) -> Optional[str]:
    l = normalize_space(label).lower().rstrip(".")
    if l in _TABLE_LABELS_V29:
        return "table"
    if l in _IMAGE_LABELS_V29:
        return "image"
    if l in _EQUATION_LABELS_V29:
        return "equation"
    return None


def _v29_read_caption_label(s: str) -> Optional[Dict[str, Any]]:
    """Parse a caption label without assuming a specific enclosure character."""
    if not s:
        return None
    original = s
    s, stripped_marker = strip_caption_marker_prefix(s)
    s = s.strip()
    if not s:
        return None
    i = 0
    opener_count = 0
    while i < len(s) and s[i] in _CAPTION_LEADING_SKIP_V29:
        if s[i] in _CAPTION_OPENERS_V29:
            opener_count += 1
        i += 1
    low = s[i:].lower()
    label = None
    for lab in _ALL_LABELS_V29:
        if low.startswith(lab.lower()):
            end = i + len(lab)
            nxt = s[end:end+1]
            if lab in {"표", "그림", "사진", "도", "수식"}:
                if nxt and _v29_is_hangul(nxt):
                    continue
            else:
                if nxt and (_v29_is_ascii_alpha(nxt) or nxt.isdigit()):
                    if not nxt.isdigit():
                        continue
            label = lab
            i = end
            break
    if not label:
        return None

    while i < len(s) and s[i].isspace():
        i += 1
    num_start = i
    while i < len(s) and re.match(r"[0-9０-９IVXivxA-Za-z가-힣\-–—_.]", s[i]):
        if _v29_is_hangul(s[i]) and i - num_start >= 6:
            break
        i += 1
    num = s[num_start:i].strip()
    if num and not re.search(r"[0-9０-９IVXivx]", num) and any(_v29_is_hangul(ch) for ch in num):
        i = num_start
        num = ""

    had_separator = False
    while i < len(s) and (s[i].isspace() or s[i] in _CAPTION_CLOSERS_V29 or s[i] in _CAPTION_SEPARATORS_V29):
        if s[i] in _CAPTION_CLOSERS_V29 or s[i] in _CAPTION_SEPARATORS_V29:
            had_separator = True
        i += 1
    body = s[i:].strip()
    if label == "도" and not (num or opener_count or had_separator):
        return None
    return {
        "original": original,
        "normalized_source": s,
        "label": label,
        "num": num,
        "body": body,
        "target": _v29_label_target(label),
        "stripped_caption_marker": stripped_marker,
        "opener_count": opener_count,
        "had_separator": had_separator,
        "delimiter_agnostic_v29": True,
    }


def _v29_caption_from_text(text: str, object_type: Optional[str] = None) -> Optional[Caption]:
    text = normalize_space(text)
    if not text:
        return None
    raw_lines = [ln.strip() for ln in str(text).splitlines() if ln.strip()]
    if not raw_lines:
        return None
    first_line = raw_lines[0]
    if len(first_line) <= 260:
        parsed = _v29_read_caption_label(first_line)
        if parsed:
            target = parsed.get("target")
            if object_type and target and target != object_type:
                return None
            if object_type == "table" and target != "table":
                return None
            if object_type == "image" and target != "image":
                return None
            if object_type == "equation" and target != "equation":
                return None
            caption_lines = [normalize_space(str(parsed.get("normalized_source") or first_line))]
            for extra in raw_lines[1:]:
                extra_stripped, _ = strip_caption_marker_prefix(extra)
                if not extra_stripped:
                    continue
                if _v29_read_caption_label(extra_stripped) or CAPTION_PATTERN.match(extra_stripped):
                    break
                if re.match(r"^(?:[0-9]+[.)]|[①-⑳]|□|■|<\s*서식|[가-힣A-Za-z]+\s*[:：])", extra_stripped) and len(" ".join(caption_lines)) > 20:
                    break
                candidate = normalize_space(" ".join(caption_lines + [extra_stripped]))
                if len(candidate) > 500:
                    break
                caption_lines.append(extra_stripped)
            caption_text = normalize_space(" ".join(caption_lines))
            return Caption(
                text=caption_text,
                method="pattern-generalized-v29",
                raw={
                    "label": parsed.get("label"),
                    "num": parsed.get("num"),
                    "body": parsed.get("body"),
                    "target": target,
                    "stripped_caption_marker": parsed.get("stripped_caption_marker"),
                    "delimiter_agnostic_caption_v29": True,
                    "soft_line_joined_v29": len(caption_lines) > 1,
                },
            )
    return _old_caption_from_text_v29(text, object_type)


caption_from_text = _v29_caption_from_text  # type: ignore[assignment]


def _v29_explicit_target(text: str) -> Optional[str]:
    if caption_from_text(text, "table"):
        return "table"
    if caption_from_text(text, "image"):
        return "image"
    if caption_from_text(text, "equation"):
        return "equation"
    return None


def _v29_caption_body_is_empty(text: str, target: Optional[str] = None) -> bool:
    cap = caption_from_text(text, target)
    if not cap:
        return False
    body = normalize_space(str((cap.raw or {}).get("body") or ""))
    if body:
        return False
    parsed = _v29_read_caption_label(strip_caption_marker_prefix(str(text or ""))[0])
    if parsed:
        return not normalize_space(str(parsed.get("body") or ""))
    return False


_v23_caption_explicit_target = _v29_explicit_target  # type: ignore[assignment]
try:
    _v28_explicit_target = _v29_explicit_target  # type: ignore[assignment]
    _v28_caption_body_is_empty = _v29_caption_body_is_empty  # type: ignore[assignment]
except Exception:
    pass

_old_v28_merge_caption_fragments_v29 = _v28_merge_caption_fragments

def _v29_merge_caption_fragments(content: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    return _old_v28_merge_caption_fragments_v29(content)

_v28_merge_caption_fragments = _v29_merge_caption_fragments  # type: ignore[assignment]


def _v29_ctrl_id(payload: bytes) -> str:
    try:
        return _v16_ctrl_id(payload).lower().strip().replace("$", "")
    except Exception:
        return ""


def _v29_is_caption_ctrl_id(ctrl: str) -> bool:
    c = (ctrl or "").lower().strip()
    return c in {"cap", "caption"} or c.startswith("cap")


def _v29_caption_from_record_group(parser: "BinaryHwpParser", records: Sequence[HwpRecord], target_type: Optional[str], section: Optional[int]) -> Optional[Dict[str, Any]]:
    texts: List[str] = []
    metas: List[Dict[str, Any]] = []
    for rec in records:
        if rec.tag_id != HWPTAG_PARA_TEXT:
            continue
        try:
            t, _ = parser.decode_para_text(rec.payload)
        except Exception:
            continue
        t = normalize_space(t)
        if not t or parser._is_control_only_text(t):
            continue
        if '_v27_is_footnote_or_annotation' in globals() and _v27_is_footnote_or_annotation(t):
            continue
        texts.append(t)
        metas.append(rec.to_meta())
    if not texts:
        return None
    merged = normalize_space(" ".join(texts))
    target = target_type or _v29_explicit_target(merged) or "image"
    cap = caption_from_text(merged, target)
    if cap:
        cap.method = "binary-caption-control-explicit-v29"
        cap.position = cap.position or ("after" if target == "image" else "before")
    else:
        cap = _v25_make_generic_caption(
            merged,
            target,
            "binary-caption-control-title-only-v29",
            "after" if target == "image" else "before",
            {"caption_control_records_v29": metas, "section": section},
        )
    return {"type": "caption", "text": cap.text, "caption": cap.to_dict(), "target_type": target, "raw": {"caption_control_records_v29": metas}}


_old_v23_parse_cell_records_v29 = BinaryHwpParser._parse_cell_records

def _v29_parse_cell_records(self, cell_records, cell_meta, section):
    parts: List[str] = []
    paras: List[str] = []
    mixed: List[Dict[str, Any]] = []
    i = 0
    while i < len(cell_records):
        rec = cell_records[i]
        if rec.tag_id == HWPTAG_TABLE:
            try:
                nested_block, next_i = self._parse_table_group(cell_records, i, section)
                nested_item = _v17_inline_table_item_from_block(nested_block) if '_v17_inline_table_item_from_block' in globals() else {
                    "type": "table",
                    "id": nested_block.id,
                    "text": nested_block.text,
                    "rows": [[cell.to_dict() for cell in row] for row in (nested_block.rows or [])],
                    "caption": nested_block.caption.to_dict() if nested_block.caption else None,
                    "raw": nested_block.raw,
                }
                mixed.append(nested_item)
                nested_text = normalize_space(nested_item.get("text", ""))
                if nested_text:
                    parts.append(nested_text); paras.append(nested_text)
                i = max(next_i, i + 1)
                continue
            except Exception as e:
                mixed.append({"type": "table", "rows": [], "raw": {"record": rec.to_meta(), "nested_table_parse_error_v29": str(e)}})
                i += 1
                continue
        if rec.tag_id == HWPTAG_CTRL_HEADER:
            ctrl = _v29_ctrl_id(rec.payload)
            if ctrl == "gso":
                block, next_i = self._parse_gso_group(cell_records, i, section)
                if block:
                    item = {
                        "type": "image",
                        "image_path": block.image_path,
                        "media_type": block.media_type,
                        "caption": block.caption.to_dict() if block.caption else None,
                        "raw": block.raw,
                    }
                    if hasattr(block, "image_paths"):
                        item["image_paths"] = getattr(block, "image_paths")
                    mixed.append(item)
                    cap_txt = (item.get("caption") or {}).get("text") if isinstance(item.get("caption"), dict) else None
                    if cap_txt:
                        parts.append(str(cap_txt)); paras.append(str(cap_txt))
                    i = max(next_i, i + 1)
                    continue
            if _v29_is_caption_ctrl_id(ctrl):
                group = [rec]
                j = i + 1
                while j < len(cell_records):
                    nr = cell_records[j]
                    if nr.level <= rec.level and j > i + 1:
                        break
                    group.append(nr)
                    j += 1
                cap_item = _v29_caption_from_record_group(self, group, None, section)
                if cap_item:
                    cap_item.setdefault("raw", {})["ctrl_id"] = ctrl
                    mixed.append(cap_item)
                    parts.append(str(cap_item.get("text", ""))); paras.append(str(cap_item.get("text", "")))
                    i = max(j, i + 1)
                    continue
        if rec.tag_id == HWPTAG_PARA_TEXT:
            text, ctrls = self.decode_para_text(rec.payload)
            text = normalize_space(text)
            if not text or self._is_control_only_text(text):
                i += 1
                continue
            explicit = _v29_explicit_target(text)
            if explicit:
                cap = caption_from_text(text, explicit)
                cap_dict = cap.to_dict() if cap else Caption(text=text, method="pattern-generalized-v29").to_dict()
                mixed.append({"type": "caption", "text": cap_dict.get("text", text), "caption": cap_dict, "target_type": explicit, "raw": {"explicit_caption_v29": True}})
                parts.append(cap_dict.get("text", text)); paras.append(cap_dict.get("text", text))
                i += 1
                continue
            parts.append(text); paras.append(text); mixed.append({"type": "paragraph", "text": text})
            i += 1
            continue
        if rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
            image_block = self._image_block_from_picture(rec, section)
            item = {"type": "image", "image_path": image_block.image_path, "media_type": image_block.media_type, "raw": image_block.raw}
            if image_block.caption:
                item["caption"] = image_block.caption.to_dict()
            if hasattr(image_block, "image_paths"):
                item["image_paths"] = getattr(image_block, "image_paths")
            mixed.append(item)
            i += 1
            continue
        i += 1
    mixed = _v28_relocate_captions_in_content(mixed) if '_v28_relocate_captions_in_content' in globals() else _v23_relocate_captions_in_content(mixed)
    text_for_cell = _content_plain_text_v21(mixed) if '_content_plain_text_v21' in globals() else normalize_space("\n".join(str(x.get("text", "")) for x in mixed if isinstance(x, dict)))
    cell = TableCell(
        text=text_for_cell or normalize_space("\n".join(parts)),
        row=cell_meta.get("row", 0),
        col=cell_meta.get("col", 0),
        row_span=max(1, int(cell_meta.get("row_span", 1))),
        col_span=max(1, int(cell_meta.get("col_span", 1))),
        attrs={"binary_cell_meta": cell_meta, "delimiter_agnostic_caption_pairing_v29": True},
        paragraphs=_dedupe_adjacent_strings([str(x.get("text", "")) for x in mixed if isinstance(x, dict) and x.get("type") == "paragraph" and x.get("text")]),
        content=mixed,
    )
    return cell, mixed

BinaryHwpParser._parse_cell_records = _v29_parse_cell_records  # type: ignore[method-assign]

_old_v25_caption_for_image_from_records_v29 = _v25_caption_for_image_from_records

def _v29_caption_for_image_from_records(parser: "BinaryHwpParser", records: Sequence[HwpRecord]) -> Optional[Caption]:
    for idx, rec in enumerate(records):
        if rec.tag_id == HWPTAG_CTRL_HEADER and _v29_is_caption_ctrl_id(_v29_ctrl_id(rec.payload)):
            group = [rec]
            j = idx + 1
            while j < len(records):
                nr = records[j]
                if nr.level <= rec.level and j > idx + 1:
                    break
                group.append(nr)
                j += 1
            cap_item = _v29_caption_from_record_group(parser, group, "image", rec.section)
            if cap_item and isinstance(cap_item.get("caption"), dict):
                d = cap_item["caption"]
                return Caption(text=str(d.get("text") or cap_item.get("text") or ""), method=str(d.get("method") or "binary-gso-caption-control-v29"), position=d.get("position") or "after", raw={**(d.get("raw") or {}), "gso_caption_control_v29": True})
    return _old_v25_caption_for_image_from_records_v29(parser, records)

_v16_caption_for_image_from_records = _v29_caption_for_image_from_records  # type: ignore[assignment]
try:
    _v25_caption_for_image_from_records = _v29_caption_for_image_from_records  # type: ignore[assignment]
except Exception:
    pass

_old_v28_binary_postprocess_v29 = BinaryHwpParser._postprocess

def _v29_binary_postprocess(self: "BinaryHwpParser") -> None:
    _old_v28_binary_postprocess_v29(self)
    try:
        for block in self.blocks:
            if block.type == "table":
                for row in block.rows or []:
                    for cell in row or []:
                        if isinstance(cell, TableCell) and cell.content:
                            cell.content = _v28_relocate_captions_in_content(cell.content) if '_v28_relocate_captions_in_content' in globals() else _v23_relocate_captions_in_content(cell.content)
                            cell.text = _content_plain_text_v21(cell.content) if '_content_plain_text_v21' in globals() else cell.text
        self.metadata["delimiter_agnostic_caption_support_v29"] = True
        self.metadata["structural_caption_control_support_v29"] = True
    except Exception as e:
        self.warnings.append(f"v29 caption postprocess failed: {e}")
    for i, b in enumerate(self.blocks, 1):
        b.order = i

BinaryHwpParser._postprocess = _v29_binary_postprocess  # type: ignore[method-assign]

try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v28", "HWP Parser Verification UI v29")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v27", "HWP Parser Verification UI v29")
except Exception:
    pass



# v29.1: tighten fallback label validation.  Older CAPTION_PATTERN could treat
# ordinary Korean words such as "표준" as if they were table captions.  Once the
# delimiter-agnostic parser has rejected a token-boundary case, do not allow the
# legacy permissive pattern to re-admit it.
for _extra_label in ("fig.", "eq."):
    if _extra_label not in _ALL_LABELS_V29:
        _ALL_LABELS_V29.append(_extra_label)
_ALL_LABELS_V29 = sorted(_ALL_LABELS_V29, key=len, reverse=True)

_prev_caption_from_text_v291 = caption_from_text

def _v291_caption_from_text(text: str, object_type: Optional[str] = None) -> Optional[Caption]:
    cap = _prev_caption_from_text_v291(text, object_type)
    if cap is None:
        return None
    raw = cap.raw or {}
    label = normalize_space(str(raw.get("label") or "")).lower().rstrip(".")
    if cap.method == "pattern" and label:
        # Reject legacy false positives where the Korean label was swallowed as
        # a longer word, e.g. "표준" or "그림자".
        allowed = set()
        if object_type == "table":
            allowed = {"표", "table", "tbl"}
        elif object_type == "image":
            allowed = {"그림", "사진", "도", "fig", "figure", "image", "img", "picture", "pic"}
        elif object_type == "equation":
            allowed = {"수식", "equation", "eq"}
        if allowed and label not in allowed:
            return None
        # If no object_type was supplied, still reject obvious Korean word
        # expansions beginning with a caption morpheme.
        if not object_type and (
            (label.startswith("표") and label != "표") or
            (label.startswith("그림") and label != "그림") or
            (label.startswith("수식") and label != "수식")
        ):
            return None
    return cap

caption_from_text = _v291_caption_from_text  # type: ignore[assignment]
_v23_caption_explicit_target = _v29_explicit_target  # type: ignore[assignment]
try:
    _v28_explicit_target = _v29_explicit_target  # type: ignore[assignment]
except Exception:
    pass


# =============================================================================
# v30: caption recognition must not depend on visible "표/그림" labels
# =============================================================================
# Visible label forms such as <그림 3>, {표 1}, and "그림." are only optional
# surface text. Hancom Insert Caption can store a structural caption while the
# visible text is title-only. v30 therefore treats caption control / GSO-local
# structure as the primary signal and uses conservative adjacency only as a
# fallback when the binary stream does not expose the caption control.

CAPTION_UNLABELED_MAX_CHARS_V30 = 260


def _v30_text_lines(text: str) -> List[str]:
    return [ln.strip() for ln in str(text or "").splitlines() if ln.strip()]


def _v30_is_numeric_or_scalar_text(text: str) -> bool:
    s = normalize_space(text)
    if not s:
        return True
    if re.fullmatch(r"[0-9０-９,\.\-/:%％\s]+", s):
        return True
    if re.fullmatch(r"[○◯O×xX△▲●\-–—_\s]+", s):
        return True
    if len(s) <= 3 and not re.search(r"[가-힣A-Za-z]", s):
        return True
    return False


def _v30_is_probably_body_sentence(text: str) -> bool:
    s = normalize_space(text)
    if not s:
        return True
    if re.match(r"^(?:본|이|이러한|이를|이에|또한|따라서|그러나|그리고|먼저|다음으로|예를\s+들어|특히|향후|현재|구체적으로|최종적으로)\b", s):
        return True
    if len(s) > 70 and re.search(r"(?:다|한다|하였다|됩니다|입니다|있다|된다|되었다|확인하였다|수행하였다|구성하였다)[.。]?$", s):
        return True
    if len(re.findall(r"[.!?。！？]", s)) >= 2:
        return True
    return False


def _v30_is_section_or_list_heading(text: str) -> bool:
    s = normalize_space(text)
    if not s:
        return True
    # Explicit caption labels like <그림 3> or {표 1} are not section headings,
    # even though the older heading guard treats arbitrary <...> as headings.
    try:
        if _v29_read_caption_label(strip_caption_marker_prefix(s)[0]):
            return False
    except Exception:
        pass
    if '_v25_is_section_heading_like' in globals() and _v25_is_section_heading_like(s):
        return True
    if re.match(r"^(?:[-•·◦▪▫○●□■◆◇]\s+|[0-9]+\s*[.)]\s+|[①-⑳]\s*)", s):
        return True
    if re.match(r"^(?:목\s*표|달\s*성|달성률|추진실적|향후계획|지연사유|해소방안)\b", s):
        return True
    return False


def _v30_is_unlabeled_caption_candidate(text: str, *, structural: bool = False, adjacent: bool = False) -> bool:
    s = normalize_space(text)
    if not s or len(s) > CAPTION_UNLABELED_MAX_CHARS_V30:
        return False
    if '_v27_is_footnote_or_annotation' in globals() and _v27_is_footnote_or_annotation(s):
        return False
    if _v30_is_numeric_or_scalar_text(s):
        return False
    if _v30_is_section_or_list_heading(s):
        return False
    lines = _v30_text_lines(text)
    if len(lines) > (4 if structural else 2):
        return False
    if _v29_read_caption_label(strip_caption_marker_prefix(s)[0]):
        return True
    tokens = re.findall(r"[가-힣A-Za-z0-9]+", s)
    if len(tokens) < 1:
        return False
    if _v30_is_probably_body_sentence(s):
        return False
    if structural:
        return True
    if adjacent:
        if len(s) <= 140 and len(tokens) >= 2 and not re.search(r"[.!?。！？]", s):
            return True
    return False


def _v30_make_unlabeled_caption(text: str, target_type: Optional[str], method: str, position: str = "after", raw: Optional[Dict[str, Any]] = None) -> Caption:
    if target_type:
        explicit = caption_from_text(text, target_type)
        if explicit:
            explicit.method = method
            explicit.position = explicit.position or position
            explicit.raw = {**(explicit.raw or {}), **(raw or {}), "v30_structural_or_unlabeled_caption": True}
            return explicit
    return Caption(
        text=normalize_space(text),
        method=method,
        position=position,
        raw={"v30_unlabeled_caption_without_table_figure_keyword": True, **(raw or {})},
    )


_prev_v25_is_generic_title_caption_v30 = _v25_is_generic_title_caption

def _v30_is_generic_title_caption(text: str, target_type: Optional[str] = None, structural: bool = False) -> bool:
    if _prev_v25_is_generic_title_caption_v30(text, target_type, structural):
        return True
    return _v30_is_unlabeled_caption_candidate(text, structural=structural, adjacent=not structural)

_v25_is_generic_title_caption = _v30_is_generic_title_caption  # type: ignore[assignment]


_prev_v29_caption_from_record_group_v30 = _v29_caption_from_record_group

def _v30_caption_from_record_group(parser: "BinaryHwpParser", records: Sequence[HwpRecord], target_type: Optional[str], section: Optional[int]) -> Optional[Dict[str, Any]]:
    texts: List[str] = []
    metas: List[Dict[str, Any]] = []
    for rec in records:
        if rec.tag_id != HWPTAG_PARA_TEXT:
            continue
        try:
            t, _ = parser.decode_para_text(rec.payload)
        except Exception:
            continue
        t = normalize_space(t)
        if not t or parser._is_control_only_text(t):
            continue
        if '_v27_is_footnote_or_annotation' in globals() and _v27_is_footnote_or_annotation(t):
            continue
        if _v30_is_section_or_list_heading(t):
            continue
        texts.append(t)
        metas.append(rec.to_meta())
    if not texts:
        return None
    merged = normalize_space(" ".join(texts))
    explicit_target = _v29_explicit_target(merged)
    target = target_type or explicit_target
    if target:
        cap = caption_from_text(merged, target) or _v30_make_unlabeled_caption(
            merged, target, "binary-caption-control-unlabeled-v30", "after" if target == "image" else "before", {"caption_control_records_v30": metas, "section": section}
        )
        cap.position = cap.position or ("after" if target == "image" else "before")
    else:
        if not _v30_is_unlabeled_caption_candidate(merged, structural=True):
            return None
        cap = _v30_make_unlabeled_caption(
            merged, None, "binary-caption-control-untyped-title-only-v30", "after", {"caption_control_records_v30": metas, "section": section}
        )
    item: Dict[str, Any] = {"type": "caption", "text": cap.text, "caption": cap.to_dict(), "raw": {"caption_control_records_v30": metas, "caption_control_unlabeled_v30": True}}
    if target:
        item["target_type"] = target
    return item

_v29_caption_from_record_group = _v30_caption_from_record_group  # type: ignore[assignment]


_prev_v29_caption_for_image_from_records_v30 = _v29_caption_for_image_from_records

def _v30_caption_for_image_from_records(parser: "BinaryHwpParser", records: Sequence[HwpRecord]) -> Optional[Caption]:
    for idx, rec in enumerate(records):
        if rec.tag_id == HWPTAG_CTRL_HEADER and _v29_is_caption_ctrl_id(_v29_ctrl_id(rec.payload)):
            group = [rec]
            j = idx + 1
            while j < len(records):
                nr = records[j]
                if nr.level <= rec.level and j > idx + 1:
                    break
                group.append(nr)
                j += 1
            cap_item = _v30_caption_from_record_group(parser, group, "image", rec.section)
            if cap_item and isinstance(cap_item.get("caption"), dict):
                d = cap_item["caption"]
                return Caption(text=str(d.get("text") or cap_item.get("text") or ""), method=str(d.get("method") or "binary-gso-caption-control-unlabeled-v30"), position=d.get("position") or "after", raw={**(d.get("raw") or {}), "gso_caption_control_v30": True})
    cap = _prev_v29_caption_for_image_from_records_v30(parser, records)
    if cap:
        return cap
    candidates: List[Tuple[HwpRecord, str]] = []
    seen_picture = False
    for rec in records:
        if rec.tag_id == HWPTAG_SHAPE_COMPONENT_PICTURE:
            seen_picture = True
            continue
        if rec.tag_id != HWPTAG_PARA_TEXT:
            continue
        try:
            t, _ = parser.decode_para_text(rec.payload)
        except Exception:
            continue
        t = normalize_space(t)
        if not t or parser._is_control_only_text(t) or not seen_picture:
            continue
        if _v30_is_unlabeled_caption_candidate(t, structural=True):
            candidates.append((rec, t))
    if candidates:
        rec, t = candidates[-1]
        return _v30_make_unlabeled_caption(t, "image", "binary-gso-structural-unlabeled-v30", "after", {"source_record": rec.to_meta(), "gso_group_caption_recovery_v30": True})
    return None

_v29_caption_for_image_from_records = _v30_caption_for_image_from_records  # type: ignore[assignment]
_v16_caption_for_image_from_records = _v30_caption_for_image_from_records  # type: ignore[assignment]
try:
    _v25_caption_for_image_from_records = _v30_caption_for_image_from_records  # type: ignore[assignment]
except Exception:
    pass


def _v30_block_caption_candidate(block: "Block", target_type: str) -> Optional[Caption]:
    if block.type != "paragraph" or not block.text:
        return None
    text = normalize_space(block.text)
    explicit = caption_from_text(text, target_type)
    if explicit:
        explicit.method = "adjacent-explicit-paragraph-v30"
        return explicit
    if _v30_is_unlabeled_caption_candidate(text, adjacent=True):
        return _v30_make_unlabeled_caption(text, target_type, "adjacent-unlabeled-title-paragraph-v30", "after", {"source_block_id": block.id, "source_order": block.order})
    return None


def _v30_between_blocks_safe_for_unlabeled_caption(blocks: List["Block"], a: int, b: int) -> bool:
    lo, hi = sorted((a, b))
    for k in range(lo + 1, hi):
        blk = blocks[k]
        if blk.type in {"table", "image", "equation"}:
            return False
        if blk.type == "paragraph":
            t = normalize_space(blk.text or "")
            if not t:
                continue
            if _v30_is_section_or_list_heading(t) or len(t) > 40:
                return False
    return True


def _v30_resolve_adjacent_unlabeled_captions(self: "BinaryHwpParser") -> None:
    blocks = list(self.blocks)
    consumed: set[int] = set()
    for obj_idx, obj in enumerate(blocks):
        if obj.type not in {"image", "table", "equation"} or obj.caption is not None:
            continue
        target_type = "image" if obj.type == "image" else obj.type
        candidates: List[Tuple[float, int, Caption]] = []
        for cap_idx in (obj_idx + 1, obj_idx - 1, obj_idx + 2, obj_idx - 2):
            if cap_idx < 0 or cap_idx >= len(blocks) or cap_idx in consumed:
                continue
            cap = _v30_block_caption_candidate(blocks[cap_idx], target_type)
            if not cap or not _v30_between_blocks_safe_for_unlabeled_caption(blocks, obj_idx, cap_idx):
                continue
            dist = abs(cap_idx - obj_idx)
            pos = "after" if cap_idx > obj_idx else "before"
            cap.position = pos
            score = 120.0 - dist * 20.0
            if target_type == "image" and pos == "after":
                score += 20.0
            if target_type == "table" and pos == "before":
                score += 8.0
            if target_type == "table" and pos == "after":
                score += 5.0
            candidates.append((score, cap_idx, cap))
        if not candidates:
            continue
        candidates.sort(key=lambda x: x[0], reverse=True)
        score, cap_idx, cap = candidates[0]
        cap.raw = {**(cap.raw or {}), "adjacent_unlabeled_pairing_score_v30": score, "object_id": obj.id}
        obj.caption = cap
        consumed.add(cap_idx)
    if consumed:
        self.blocks = [b for idx, b in enumerate(blocks) if idx not in consumed]
        for idx, b in enumerate(self.blocks, 1):
            b.order = idx


def _v30_relocate_unlabeled_captions_in_content(content: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    if not content:
        return content
    try:
        items = _v30_prev_relocate_content(content)
    except Exception:
        items = list(content)
    consumed: set[int] = set()
    for obj_idx, obj in enumerate(items):
        if not isinstance(obj, dict) or obj.get("type") not in {"image", "table", "equation"} or obj.get("caption"):
            continue
        target_type = "image" if obj.get("type") == "image" else str(obj.get("type"))
        candidates: List[Tuple[float, int, Caption]] = []
        for cap_idx in (obj_idx + 1, obj_idx - 1, obj_idx + 2, obj_idx - 2):
            if cap_idx < 0 or cap_idx >= len(items) or cap_idx in consumed:
                continue
            cand_item = items[cap_idx]
            if not isinstance(cand_item, dict) or cand_item.get("type") != "paragraph":
                continue
            text = normalize_space(str(cand_item.get("text") or ""))
            if not _v30_is_unlabeled_caption_candidate(text, adjacent=True):
                continue
            safe = True
            lo, hi = sorted((obj_idx, cap_idx))
            for k in range(lo + 1, hi):
                mid = items[k]
                if not isinstance(mid, dict):
                    continue
                if mid.get("type") in {"image", "table", "equation"}:
                    safe = False; break
                if mid.get("type") == "paragraph":
                    mt = normalize_space(str(mid.get("text") or ""))
                    if _v30_is_section_or_list_heading(mt) or len(mt) > 40:
                        safe = False; break
            if not safe:
                continue
            pos = "after" if cap_idx > obj_idx else "before"
            cap = _v30_make_unlabeled_caption(text, target_type, "inline-adjacent-unlabeled-title-v30", pos, {"source_item_index": cap_idx, "target_item_index": obj_idx})
            score = 110.0 - abs(cap_idx - obj_idx) * 18.0
            if target_type == "image" and pos == "after":
                score += 16.0
            candidates.append((score, cap_idx, cap))
        if not candidates:
            continue
        candidates.sort(key=lambda x: x[0], reverse=True)
        score, cap_idx, cap = candidates[0]
        cap.raw = {**(cap.raw or {}), "inline_adjacent_unlabeled_pairing_score_v30": score}
        obj["caption"] = cap.to_dict()
        consumed.add(cap_idx)
    if consumed:
        return [it for idx, it in enumerate(items) if idx not in consumed]
    return items


_v30_prev_relocate_content = _v28_relocate_captions_in_content if '_v28_relocate_captions_in_content' in globals() else _v23_relocate_captions_in_content

def _v30_relocate_captions_in_content(content: List[Dict[str, Any]], max_window: int = 40) -> List[Dict[str, Any]]:
    return _v30_relocate_unlabeled_captions_in_content(content)

_v28_relocate_captions_in_content = _v30_relocate_captions_in_content  # type: ignore[assignment]
_v23_relocate_captions_in_content = _v30_relocate_captions_in_content  # type: ignore[assignment]

_prev_binary_postprocess_v30 = BinaryHwpParser._postprocess

def _v30_binary_postprocess(self: "BinaryHwpParser") -> None:
    _prev_binary_postprocess_v30(self)
    try:
        self._resolve_adjacent_unlabeled_captions()
        for block in self.blocks:
            if block.type == "table":
                for row in block.rows or []:
                    for cell in row or []:
                        if isinstance(cell, TableCell) and cell.content:
                            cell.content = _v30_relocate_unlabeled_captions_in_content(cell.content)
                            cell.text = _content_plain_text_v21(cell.content) if '_content_plain_text_v21' in globals() else cell.text
        self.metadata["keyword_free_structural_caption_support_v30"] = True
        self.metadata["caption_control_primary_v30"] = True
    except Exception as e:
        self.warnings.append(f"v30 keyword-free caption postprocess failed: {e}")
    for idx, b in enumerate(self.blocks, 1):
        b.order = idx

BinaryHwpParser._resolve_adjacent_unlabeled_captions = _v30_resolve_adjacent_unlabeled_captions  # type: ignore[attr-defined]
BinaryHwpParser._postprocess = _v30_binary_postprocess  # type: ignore[method-assign]

try:
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v29", "HWP Parser Verification UI v30")
    VIEWER_HTML = VIEWER_HTML.replace("HWP Parser Verification UI v28", "HWP Parser Verification UI v30")
except Exception:
    pass

# =============================================================================
# v30 entrypoint
# =============================================================================
if __name__ == "__main__":
    raise SystemExit(main())
