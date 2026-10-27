import csv
import html
import io
import json
import os
import re
import sqlite3
import tempfile
import threading
import time
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import textwrap
from typing import Iterable, Optional

import requests
from flask import Flask, Response, redirect, render_template, request, url_for
from openpyxl import Workbook
from PIL import Image, ImageDraw, ImageFont


BASE_DIR = Path(__file__).resolve().parent
# Fixed directory for WhatsApp exported text files.
EXPORT_DIR = BASE_DIR / "data" / "exports"
# 固定目录内可识别的聊天导出后缀（文本 / HTML / JSON）
EXPORT_SUFFIXES = frozenset({".txt", ".html", ".htm", ".json"})
OUTPUT_DATE_FORMAT = "%Y-%m-%d %H:%M"
GEMINI_MODEL = "gemini-2.5-flash-lite"
GEMINI_API_URL = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
# 单次分析送入模型的消息条数上限（避免输出 JSON 过大被截断导致解析失败）
ANALYSIS_MAX_MESSAGES = 250

# Gemini 结构化输出 schema（TYPE 需大写），用于保证 JSON 语法合法
GEMINI_RESPONSE_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "summary": {"type": "STRING"},
        "hot_models": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "model": {"type": "STRING"},
                    "mentions": {"type": "INTEGER"},
                    "pain_points": {"type": "ARRAY", "items": {"type": "STRING"}},
                    "itch_points": {"type": "ARRAY", "items": {"type": "STRING"}},
                    "delight_points": {"type": "ARRAY", "items": {"type": "STRING"}},
                    "solutions": {"type": "ARRAY", "items": {"type": "STRING"}},
                },
                "required": ["model", "mentions"],
            },
        },
        "users": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "nickname": {"type": "STRING"},
                    "phone": {"type": "STRING"},
                    "message_count": {"type": "INTEGER"},
                    "user_type": {"type": "STRING"},
                    "role": {"type": "STRING"},
                    "influence_score": {"type": "INTEGER"},
                    "analysis": {"type": "STRING"},
                    "suggestion": {"type": "STRING"},
                },
                "required": [
                    "nickname",
                    "phone",
                    "message_count",
                    "user_type",
                    "role",
                    "influence_score",
                    "analysis",
                    "suggestion",
                ],
            },
        },
        "opportunities": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "name": {"type": "STRING"},
                    "reason": {"type": "STRING"},
                    "execution_plan": {"type": "STRING"},
                },
                "required": ["name", "reason", "execution_plan"],
            },
        },
        "local_price_analysis": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "product": {"type": "STRING"},
                    "location": {"type": "STRING"},
                    "price": {"type": "STRING"},
                    "fx_rate": {"type": "STRING"},
                    "price_usd": {"type": "NUMBER"},
                    "user": {"type": "STRING"},
                    "phone": {"type": "STRING"},
                    "message_time": {"type": "STRING"},
                    "remark": {"type": "STRING"},
                },
                "required": [
                    "product",
                    "location",
                    "price",
                    "fx_rate",
                    "price_usd",
                    "user",
                    "phone",
                    "message_time",
                    "remark",
                ],
            },
        },
    },
    "required": ["summary", "hot_models", "users", "opportunities", "local_price_analysis"],
}
DB_PATH = BASE_DIR / "data" / "analysis.db"

app = Flask(__name__)

_last_heartbeat_ts: Optional[float] = None
_heartbeat_lock = threading.Lock()


@dataclass
class MessageRecord:
    group_name: str
    speaker_name: str
    speaker_phone: str
    sent_at: datetime
    content: str


LINE_PATTERNS = [
    re.compile(
        r"^\[(?P<date>\d{1,4}[/-]\d{1,2}[/-]\d{1,4}),\s*(?P<time>\d{1,2}:\d{2})(?::\d{2})?\]\s+"
        r"(?P<speaker>[^:]+):\s(?P<content>.*)$"
    ),
    re.compile(
        r"^(?P<date>\d{1,4}[/-]\d{1,2}[/-]\d{1,4})[,]?\s+(?P<time>\d{1,2}:\d{2})\s*-\s*"
        r"(?P<speaker>[^:]+):\s(?P<content>.*)$"
    ),
]

DATE_PATTERNS = [
    "%Y/%m/%d %H:%M",
    "%Y-%m-%d %H:%M",
    "%d/%m/%Y %H:%M",
    "%d/%m/%y %H:%M",
    "%m/%d/%Y %H:%M",
    "%m/%d/%y %H:%M",
]

PHONE_PATTERN = re.compile(r"(\+?\d[\d\s\-()]{5,}\d)")
PRICE_NUMBER_PATTERN = re.compile(
    r"\b(?:\d{1,3}(?:[,\s]\d{3})+|\d{2,7})(?:\.\d+)?(?:\s*[kK])?\b"
)

COUNTRY_BY_PREFIX = {
    "+971": "阿联酋",
    "+966": "沙特",
    "+973": "巴林",
    "+974": "卡塔尔",
    "+965": "科威特",
    "+968": "阿曼",
    "+212": "摩洛哥",
    "+1": "北美",
}


def parse_datetime(date_text: str, time_text: str) -> Optional[datetime]:
    raw = f"{date_text} {time_text}"
    for pattern in DATE_PATTERNS:
        try:
            return datetime.strptime(raw, pattern)
        except ValueError:
            continue
    return None


def extract_phone(speaker: str) -> str:
    # WhatsApp exports often use phone numbers as contact display names.
    match = PHONE_PATTERN.search(speaker)
    if not match:
        return ""
    return re.sub(r"\s+", "", match.group(1))


def get_group_name(file_path: Path) -> str:
    suffix = file_path.suffix.lower()
    if suffix in (".html", ".htm"):
        try:
            head = file_path.read_text(encoding="utf-8", errors="ignore")[:200_000]
            m_header = re.search(
                r'(?is)<div class="page_header".*?<div class="text bold">\s*(.*?)\s*</div>',
                head,
            )
            if m_header:
                t = html_to_plaintext(m_header.group(1)).strip()
                if t:
                    return t
            m = re.search(r"(?is)<title[^>]*>([^<]+)</title>", head)
            if m:
                t = html.unescape(m.group(1).strip())
                if t and t.lower() != "exported data":
                    return t
        except OSError:
            pass
        return file_path.stem
    if suffix == ".json":
        try:
            head = file_path.read_text(encoding="utf-8", errors="ignore")[:400_000]
            for key in ("chatName", "title", "subject", "groupName", "name"):
                mm = re.search(rf'"{key}"\s*:\s*"((?:[^"\\]|\\.)*)"', head)
                if not mm:
                    continue
                raw_t = mm.group(1)
                try:
                    return str(json.loads(f'"{raw_t}"'))
                except json.JSONDecodeError:
                    t = raw_t.replace('\\"', '"').replace("\\\\", "\\")
                    if t.strip():
                        return t.strip()
        except OSError:
            pass
        return file_path.stem
    # .txt：首行常为「WhatsApp Chat with …」
    try:
        first_line = file_path.read_text(encoding="utf-8", errors="ignore").splitlines()[0]
        cleaned = first_line.strip().replace("\ufeff", "")
        if cleaned:
            return cleaned
    except (IndexError, OSError):
        pass
    return file_path.stem


def is_export_file_path(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in EXPORT_SUFFIXES


def html_to_plaintext(raw: str) -> str:
    s = re.sub(r"(?is)<head[^>]*>.*?</head>", " ", raw)
    s = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", s)
    s = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", s)
    s = re.sub(r"(?i)<br\s*/?>", "\n", s)
    s = re.sub(r"(?i)</(p|div|tr|li|h[1-6])>", "\n", s)
    s = re.sub(r"<[^>]+>", "", s)
    return html.unescape(s)


def _json_collect_message_dicts(data: object, depth: int = 0) -> list[dict]:
    if depth > 8:
        return []
    if isinstance(data, list):
        if data and all(isinstance(x, dict) for x in data):
            keys_lower = {k.lower() for k in data[0]}
            msg_keys = {"content", "text", "body", "message", "msg"}
            if keys_lower & msg_keys:
                return data
        out: list[dict] = []
        for item in data:
            out.extend(_json_collect_message_dicts(item, depth + 1))
        return out
    if isinstance(data, dict):
        for key in ("messages", "messageList", "chatMessages", "data", "items", "history"):
            if key in data:
                got = _json_collect_message_dicts(data[key], depth + 1)
                if got:
                    return got
        chats = data.get("chats")
        if isinstance(chats, list):
            out = []
            for chat in chats:
                if not isinstance(chat, dict):
                    continue
                title = (
                    chat.get("name")
                    or chat.get("title")
                    or chat.get("displayName")
                    or chat.get("chatName")
                    or ""
                )
                msgs = chat.get("messages") or chat.get("messageList") or []
                if not isinstance(msgs, list):
                    continue
                for m in msgs:
                    if isinstance(m, dict):
                        mm = dict(m)
                        if title:
                            mm["_chat_title"] = str(title)
                        out.append(mm)
            if out:
                return out
        for v in data.values():
            got = _json_collect_message_dicts(v, depth + 1)
            if got:
                return got
    return []


def _parse_json_datetime_value(val: object) -> Optional[datetime]:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        ts = float(val)
        if ts > 1e12:
            ts /= 1000.0
        try:
            return datetime.fromtimestamp(ts)
        except (OSError, ValueError, OverflowError):
            return None
    if isinstance(val, str):
        s = val.strip()
        if not s:
            return None
        s_iso = s.replace("Z", "+00:00")
        try:
            dt = datetime.fromisoformat(s_iso)
            return dt.replace(tzinfo=None) if dt.tzinfo else dt
        except ValueError:
            pass
    return None


def _json_message_to_record(m: dict, default_group: str) -> Optional[MessageRecord]:
    group_name = str(m.get("_chat_title") or "").strip() or default_group
    content = ""
    for key in ("content", "text", "body", "message", "msg"):
        v = m.get(key)
        if isinstance(v, str) and v.strip():
            content = v.strip()
            break
        if v is not None and not isinstance(v, (dict, list)):
            s = str(v).strip()
            if s:
                content = s
                break
    if not content:
        return None
    speaker = ""
    for key in ("sender", "from", "author", "name", "senderName", "participant", "user"):
        v = m.get(key)
        if isinstance(v, str) and v.strip():
            speaker = v.strip()
            break
        if isinstance(v, dict):
            speaker = str(v.get("name") or v.get("pushname") or "").strip()
            if speaker:
                break
    if not speaker:
        speaker = "未知"

    sent_at: Optional[datetime] = None
    for key in ("sent_at", "createdAt", "created_at", "timestampMs", "timestamp_ms", "ts"):
        if key in m:
            sent_at = _parse_json_datetime_value(m.get(key))
            if sent_at:
                break
    if sent_at is None and "timestamp" in m:
        sent_at = _parse_json_datetime_value(m.get("timestamp"))
    if sent_at is None:
        d = m.get("date")
        t = m.get("time")
        if d is not None and t is not None:
            sent_at = parse_datetime(str(d).strip(), str(t).strip())
    if sent_at is None:
        for key in ("dateTime", "datetime", "message_time"):
            if key in m:
                sent_at = _parse_json_datetime_value(m.get(key))
                if sent_at:
                    break
    if sent_at is None:
        return None

    return MessageRecord(
        group_name=group_name,
        speaker_name=speaker,
        speaker_phone=extract_phone(speaker),
        sent_at=sent_at,
        content=content,
    )


def _iter_json_message_records(file_path: Path, group_name: str) -> Iterable[MessageRecord]:
    raw = file_path.read_text(encoding="utf-8", errors="ignore")
    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        return
    msgs = _json_collect_message_dicts(data)
    for m in msgs:
        rec = _json_message_to_record(m, group_name)
        if rec:
            yield rec


def _parse_telegram_html_datetime(date_title: str) -> Optional[datetime]:
    raw = html.unescape((date_title or "").strip())
    raw = re.sub(r"\s+UTC[+-]\d{1,2}(?::\d{2})?\s*$", "", raw, flags=re.IGNORECASE)
    for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%y %H:%M:%S", "%d.%m.%y %H:%M"):
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            continue
    return None


def _iter_telegram_html_message_records(file_path: Path, group_name: str) -> Iterable[MessageRecord]:
    raw = file_path.read_text(encoding="utf-8", errors="ignore")
    starts = list(re.finditer(r'(?is)<div class="message\s+[^"]*"', raw))
    if not starts:
        return

    for idx, mt in enumerate(starts):
        start = mt.start()
        end = starts[idx + 1].start() if idx + 1 < len(starts) else len(raw)
        segment = raw[start:end]
        if 'class="message default clearfix"' not in segment:
            continue

        dm = re.search(r'(?is)class="pull_right date details"[^>]*title="([^"]+)"', segment)
        if not dm:
            continue
        sent_at = _parse_telegram_html_datetime(dm.group(1))
        if sent_at is None:
            continue

        sm = re.search(r'(?is)<div class="from_name">\s*(.*?)\s*</div>', segment)
        speaker = html_to_plaintext(sm.group(1)).strip() if sm else ""
        if not speaker:
            speaker = group_name or "未知"

        cm = re.search(r'(?is)<div class="text">\s*(.*?)\s*</div>', segment)
        content = html_to_plaintext(cm.group(1)).strip() if cm else ""
        if not content:
            media_caption = re.search(r'(?is)<div class="media_wrap.*?<div class="title[^"]*">\s*(.*?)\s*</div>', segment)
            if media_caption:
                content = html_to_plaintext(media_caption.group(1)).strip()
        if not content:
            continue

        yield MessageRecord(
            group_name=group_name,
            speaker_name=speaker,
            speaker_phone=extract_phone(speaker),
            sent_at=sent_at,
            content=content,
        )


def parse_chat_file(
    file_path: Path, requested_group: Optional[str] = None
) -> Iterable[MessageRecord]:
    group_name = get_group_name(file_path)
    if requested_group and requested_group.strip():
        requested_group_lower = requested_group.strip().lower()
        matches_group = (
            requested_group_lower in group_name.lower()
            or requested_group_lower in file_path.stem.lower()
        )
        if not matches_group:
            return

    suffix = file_path.suffix.lower()
    if suffix == ".json":
        yield from _iter_json_message_records(file_path, group_name)
        return
    if suffix in (".html", ".htm"):
        yielded = False
        for item in _iter_telegram_html_message_records(file_path, group_name):
            yielded = True
            yield item
        if yielded:
            return

    raw = file_path.read_text(encoding="utf-8", errors="ignore")
    text = html_to_plaintext(raw) if suffix in (".html", ".htm") else raw
    current_record: Optional[MessageRecord] = None

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        parsed = None
        for pattern in LINE_PATTERNS:
            match = pattern.match(line)
            if match:
                parsed = match.groupdict()
                break

        if parsed:
            sent_at = parse_datetime(parsed["date"], parsed["time"])
            if not sent_at:
                current_record = None
                continue

            speaker = parsed["speaker"].strip()
            current_record = MessageRecord(
                group_name=group_name,
                speaker_name=speaker,
                speaker_phone=extract_phone(speaker),
                sent_at=sent_at,
                content=parsed["content"].strip(),
            )
            yield current_record
            continue

        # Support multiline messages by appending non-header lines to last message.
        if current_record is not None:
            current_record.content = f"{current_record.content}\n{line}".strip()


def find_target_files(group_name: str) -> list[Path]:
    if not EXPORT_DIR.exists():
        return []
    files = sorted(
        (p for p in EXPORT_DIR.iterdir() if is_export_file_path(p)),
        key=lambda p: p.name.lower(),
    )
    if not group_name.strip():
        return files
    target = group_name.strip().lower()
    matched = []
    for f in files:
        group_title = get_group_name(f).lower()
        if target in f.stem.lower() or target in group_title:
            matched.append(f)
    return matched


def list_export_files() -> list[dict[str, str]]:
    """Scan fixed dir: each supported export = one chat; label for UI selection."""
    if not EXPORT_DIR.exists():
        return []
    out: list[dict[str, str]] = []
    for f in sorted(
        (p for p in EXPORT_DIR.iterdir() if is_export_file_path(p)),
        key=lambda p: p.name.lower(),
    ):
        title = get_group_name(f)
        out.append(
            {
                "file": f.name,
                "label": f"{title} — {f.name}",
            }
        )
    return out


def resolve_target_files(export_file: str, group_name: str) -> tuple[list[Path], str]:
    """
    Prefer exact file selection; otherwise match by group name substring.
    Returns (files, error_message).
    """
    if not EXPORT_DIR.exists():
        return [], f"固定目录不存在：{EXPORT_DIR}"

    name = (export_file or "").strip()
    if name:
        path = EXPORT_DIR / name
        if not is_export_file_path(path):
            return [], "选择的导出文件无效或不在固定目录内。"
        return [path], ""

    g = (group_name or "").strip()
    if not g:
        return [], "请选择具体导出文件，或填写群名称进行匹配。"

    matched = find_target_files(g)
    if not matched:
        return [], f"在固定目录 {EXPORT_DIR} 未找到匹配群名称的导出文件。"
    return matched, ""


# Excel 单格最大约 32767 字符；过长会导致写入异常或文件损坏
XLSX_MAX_CELL_CHARS = 32000


def sanitize_xlsx_text(value: object) -> str:
    """Strip characters illegal in SpreadsheetML / XML 1.0 (otherwise Excel reports file corrupt)."""
    if value is None:
        return ""
    text = str(value)
    out: list[str] = []
    for ch in text:
        o = ord(ch)
        if o < 0x20 and ch not in "\t\n\r":
            continue
        if 0xD800 <= o <= 0xDFFF:
            continue
        if o in (0xFFFE, 0xFFFF):
            continue
        out.append(ch)
    cleaned = "".join(out)
    if len(cleaned) > XLSX_MAX_CELL_CHARS:
        cleaned = cleaned[:XLSX_MAX_CELL_CHARS]
    return cleaned


def parse_dt_local(value: str) -> Optional[datetime]:
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%dT%H:%M")
    except ValueError:
        return None


def infer_country_by_phone(phone: str) -> str:
    p = (phone or "").replace(" ", "")
    for prefix, name in sorted(COUNTRY_BY_PREFIX.items(), key=lambda x: -len(x[0])):
        if p.startswith(prefix):
            return name
    return "未知"


def infer_currency_and_usd(price_text: str, location: str) -> tuple[str, float, str]:
    raw = (price_text or "").lower()
    number_match = PRICE_NUMBER_PATTERN.search(raw)
    if not number_match:
        return "USD", 0.0, "1 USD = 1 USD"
    token = number_match.group(0).replace(",", "").replace(" ", "")
    has_k = token.lower().endswith("k")
    if has_k:
        token = token[:-1]
    try:
        num = float(token)
    except ValueError:
        return "USD", 0.0, "1 USD = 1 USD"
    if has_k:
        num *= 1000.0
    if "aed" in raw or "dh" in raw or location == "阿联酋":
        return "AED", round(num / 3.6725, 2), "1 USD = 3.6725 AED"
    if "sar" in raw or location == "沙特":
        return "SAR", round(num / 3.75, 2), "1 USD = 3.75 SAR"
    if "qar" in raw or location == "卡塔尔":
        return "QAR", round(num / 3.64, 2), "1 USD = 3.64 QAR"
    if "kwd" in raw or location == "科威特":
        return "KWD", round(num * 3.25, 2), "1 KWD ≈ 3.25 USD"
    if "omr" in raw or location == "阿曼":
        return "OMR", round(num * 2.60, 2), "1 OMR ≈ 2.60 USD"
    if "bhd" in raw or location == "巴林":
        return "BHD", round(num * 2.65, 2), "1 BHD ≈ 2.65 USD"
    if "usd" in raw or "$" in raw:
        return "USD", round(num, 2), "1 USD = 1 USD"
    # If currency is not explicit, infer by location first.
    if location == "阿联酋":
        return "AED", round(num / 3.6725, 2), "1 USD = 3.6725 AED（按地点推断）"
    if location == "沙特":
        return "SAR", round(num / 3.75, 2), "1 USD = 3.75 SAR（按地点推断）"
    if location == "卡塔尔":
        return "QAR", round(num / 3.64, 2), "1 USD = 3.64 QAR（按地点推断）"
    if location == "科威特":
        return "KWD", round(num * 3.25, 2), "1 KWD ≈ 3.25 USD（按地点推断）"
    if location == "阿曼":
        return "OMR", round(num * 2.60, 2), "1 OMR ≈ 2.60 USD（按地点推断）"
    if location == "巴林":
        return "BHD", round(num * 2.65, 2), "1 BHD ≈ 2.65 USD（按地点推断）"
    return "USD", round(num, 2), "1 USD = 1 USD"


def heuristic_extract_price_analysis(rows: list[MessageRecord]) -> list[dict]:
    out: list[dict] = []
    keywords = [
        "price",
        "aed",
        "usd",
        "sar",
        "qar",
        "kwd",
        "omr",
        "bhd",
        "预算",
        "价格",
        "合同",
        "warranty",
        "contract",
        "bought",
    ]
    product_tokens = ["rox", "prado", "defender", "leopard", "lynk", "jetour", "denza", "byd", "land cruiser", "sierra"]

    def infer_product_text(text: str) -> str:
        low_text = (text or "").lower()
        for tok in product_tokens:
            if tok in low_text:
                if tok == "lynk":
                    return "LYNK&CO"
                return tok.upper() if tok.isascii() else tok
        return ""

    # Group-level dominant product fallback
    dominant_counter: dict[str, int] = {}
    for r in rows:
        p = infer_product_text(r.content)
        if p:
            dominant_counter[p] = dominant_counter.get(p, 0) + 1
    dominant_product = ""
    if dominant_counter:
        dominant_product = max(dominant_counter.items(), key=lambda x: x[1])[0]

    def infer_product_from_context(index: int) -> str:
        # Search nearby messages for nearest explicit product mention.
        for step in range(1, 6):
            prev_i = index - step
            if prev_i >= 0:
                p = infer_product_text(rows[prev_i].content)
                if p:
                    return p
            next_i = index + step
            if next_i < len(rows):
                p = infer_product_text(rows[next_i].content)
                if p:
                    return p
        return dominant_product

    for idx, item in enumerate(rows):
        text = (item.content or "").strip()
        low = text.lower()
        if not any(k in low for k in keywords):
            continue
        nums = PRICE_NUMBER_PATTERN.findall(text)
        if not nums:
            continue
        product = infer_product_text(text) or infer_product_from_context(idx) or "未明确商品"
        location = infer_country_by_phone(item.speaker_phone)
        for n in nums:
            raw_price = n.replace(",", "").replace(" ", "")
            has_k = raw_price.lower().endswith("k")
            if has_k:
                raw_price = raw_price[:-1]
            try:
                numeric_price = float(raw_price)
            except ValueError:
                continue
            if has_k:
                numeric_price *= 1000.0
            # Filter obvious non-price values (e.g. 5.35s, 140kph-ish numbers)
            if numeric_price < 1000 or numeric_price > 3000000:
                continue

            currency, usd, fx_rate = infer_currency_and_usd(f"{n} {text}", location)
            price_value = (
                f"{int(numeric_price)}" if float(numeric_price).is_integer() else f"{numeric_price:.2f}"
            )
            if currency and currency != "USD":
                price_text = f"{price_value} {currency}"
            elif "usd" in low or "$" in low:
                price_text = f"{price_value} USD"
            else:
                price_text = f"{price_value} {currency}"
            out.append(
                {
                    "product": product,
                    "location": location,
                    "price": price_text,
                    "fx_rate": fx_rate,
                    "price_usd": usd,
                    "user": item.speaker_name,
                    "phone": item.speaker_phone or "",
                    "message_time": item.sent_at.strftime("%Y-%m-%d %H:%M"),
                    "remark": "基于群聊文本规则提取，建议结合经销渠道报价进一步校准。",
                }
            )
    # de-dup by key fields
    seen = set()
    dedup = []
    for r in out:
        key = (r["product"], r["price"], r["phone"], r["message_time"])
        if key in seen:
            continue
        seen.add(key)
        dedup.append(r)
    return dedup


def normalize_and_merge_price_analysis(model_items: list[dict], rows: list[MessageRecord]) -> list[dict]:
    heuristic_items = heuristic_extract_price_analysis(rows)
    merged: list[dict] = []
    seen = set()

    def normalize_item(src: dict) -> Optional[dict]:
        product = (
            sanitize_xlsx_text(src.get("product", "")).strip()
            or sanitize_xlsx_text(src.get("car_model", "")).strip()
            or "未明确商品"
        )
        location = sanitize_xlsx_text(src.get("location", "")).strip() or "未知"
        price = sanitize_xlsx_text(src.get("price", "")).strip()
        user = sanitize_xlsx_text(src.get("user", "")).strip()
        phone = sanitize_xlsx_text(src.get("phone", "")).strip()
        message_time = sanitize_xlsx_text(src.get("message_time", "")).strip()
        remark = sanitize_xlsx_text(src.get("remark", "")).strip() or "根据群聊语境推断。"
        if not price:
            return None

        currency, usd_calc, fx_calc = infer_currency_and_usd(price, location)
        fx_rate = sanitize_xlsx_text(src.get("fx_rate", "")).strip() or fx_calc
        try:
            usd_val = float(src.get("price_usd"))
            if usd_val <= 0:
                usd_val = usd_calc
        except (TypeError, ValueError):
            usd_val = usd_calc

        if currency != "USD" and currency not in price:
            # keep price string aligned with inferred currency
            m = PRICE_NUMBER_PATTERN.search(price)
            if m:
                token = m.group(0).replace(",", "").replace(" ", "")
                if token.lower().endswith("k"):
                    token = token[:-1]
                try:
                    num = float(token)
                    if m.group(0).lower().endswith("k"):
                        num *= 1000
                    price_num = f"{int(num)}" if float(num).is_integer() else f"{num:.2f}"
                    price = f"{price_num} {currency}"
                except ValueError:
                    pass

        return {
            "product": product,
            "location": location,
            "price": price,
            "fx_rate": fx_rate,
            "price_usd": round(float(usd_val), 2),
            "user": user,
            "phone": phone,
            "message_time": message_time,
            "remark": remark,
        }

    for item in (model_items or []):
        n = normalize_item(item)
        if not n:
            continue
        key = (n["product"], n["price"], n["phone"], n["message_time"])
        if key in seen:
            continue
        seen.add(key)
        merged.append(n)

    for item in heuristic_items:
        n = normalize_item(item)
        if not n:
            continue
        key = (n["product"], n["price"], n["phone"], n["message_time"])
        if key in seen:
            continue
        seen.add(key)
        merged.append(n)

    return merged


def build_analysis_prompt(rows: list[MessageRecord], start: datetime, end: datetime) -> str:
    sample = rows[-ANALYSIS_MAX_MESSAGES:] if len(rows) > ANALYSIS_MAX_MESSAGES else rows
    lines: list[str] = []
    for row in sample:
        lines.append(
            f"[{row.sent_at.strftime('%Y-%m-%d %H:%M')}] "
            f"{row.speaker_name} ({row.speaker_phone or 'N/A'}): {row.content}"
        )
    transcript = "\n".join(lines)
    return (
        "你是汽车社群商业分析师。请基于聊天记录输出 JSON，不要输出额外解释。\n"
        "语言要求：除车型英文名/专有名词外，所有内容必须使用简体中文输出。\n"
        f"说明：以下仅为最近 {len(sample)} 条消息样本（全量共 {len(rows)} 条），请据此做归纳，不要编造未出现的事实。\n"
        "输出约束：users 最多 35 人；每条 analysis/suggestion 尽量简洁（建议各不超过 200 字）。\n"
        "分析目标：\n"
        "1) 热门车型与需求，按痛点/痒点/爽点分类，并给出解决方案建议。\n"
        "2) 用户画像：识别提问者/解答者/意见领袖/活跃者/参观者。\n"
        "3) 商机分析：给出机会点、为什么、怎么做。\n"
        "4) 车型当地售价分析：若群里提到车型价格，请提取并推断地点、币种与美元折算，"
        "并给出当地合理售价区间与建议销售价及原因。\n"
        "输出 JSON 结构：\n"
        "{\n"
        '  "hot_models": [{"model":"", "mentions":0, "pain_points":[], "itch_points":[], "delight_points":[], "solutions":[]}],\n'
        '  "users": [{"nickname":"", "phone":"", "message_count":0, "user_type":"参观者|活跃者|意见领袖", "role":"提问者|解答者|观察者", "influence_score":1, "analysis":"", "suggestion":""}],\n'
        '  "opportunities": [{"name":"", "reason":"", "execution_plan":""}],\n'
        '  "local_price_analysis": [{"product":"", "location":"", "price":"", "fx_rate":"", "price_usd":0, "user":"", "phone":"", "message_time":"", "remark":""}],\n'
        '  "summary": ""\n'
        "}\n"
        "字段值约束：\n"
        '- "user_type" 仅使用：参观者、活跃者、意见领袖。\n'
        '- "role" 仅使用：提问者、解答者、观察者。\n'
        '- "price_usd" 必须为数字；若需推断汇率，请使用合理近似并在 remark 说明依据。\n'
        '- 若未出现价格讨论，"local_price_analysis" 返回空数组 []。\n'
        f"时间范围：{start.strftime('%Y-%m-%d %H:%M')} 到 {end.strftime('%Y-%m-%d %H:%M')}\n"
        f"聊天记录如下：\n{transcript}"
    )


def call_gemini_analysis(prompt: str) -> tuple[str, str]:
    api_key = (os.getenv("GEMINI_API_KEY") or "").strip()
    if not api_key and os.name == "nt":
        # On Windows, newly-set User env vars may not appear in old terminals.
        # Fallback to HKCU\Environment so service can still read configured key.
        try:
            import winreg

            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as key:
                reg_val, _ = winreg.QueryValueEx(key, "GEMINI_API_KEY")
                api_key = str(reg_val or "").strip()
        except OSError:
            api_key = ""
    if not api_key:
        return "", "未配置 GEMINI_API_KEY，请先在系统环境变量中设置。"

    url = GEMINI_API_URL.format(model=GEMINI_MODEL)
    headers = {"Content-Type": "application/json"}
    base_body = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.2,
            "topP": 0.9,
            "maxOutputTokens": 8192,
            "responseMimeType": "application/json",
            "responseSchema": GEMINI_RESPONSE_SCHEMA,
        },
    }
    try:
        resp = requests.post(
            f"{url}?key={api_key}",
            headers=headers,
            data=json.dumps(base_body),
            timeout=300,
        )
        if resp.status_code >= 400:
            fallback = dict(base_body)
            fallback["generationConfig"] = dict(base_body["generationConfig"])
            fallback["generationConfig"].pop("responseSchema", None)
            resp2 = requests.post(
                f"{url}?key={api_key}",
                headers=headers,
                data=json.dumps(fallback),
                timeout=300,
            )
            if resp2.status_code >= 400:
                return "", (
                    f"Gemini 请求失败：HTTP {resp.status_code}，{resp.text[:400]}；"
                    f"回退请求 HTTP {resp2.status_code}，{resp2.text[:400]}"
                )
            resp = resp2
        data = resp.json()
        candidates = data.get("candidates") or []
        if not candidates:
            return "", "Gemini 返回为空。"
        parts = candidates[0].get("content", {}).get("parts", [])
        text = "".join((p.get("text") or "") for p in parts).strip()
        if not text:
            return "", "Gemini 未返回文本结果。"
        return text, ""
    except requests.RequestException as exc:
        return "", f"Gemini 请求异常：{exc}"


def extract_json_object(text: str) -> dict:
    content = (text or "").strip()
    if content.startswith("```"):
        content = re.sub(r"^```(?:json)?\s*", "", content)
        content = re.sub(r"\s*```$", "", content)
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        start = content.find("{")
        if start < 0:
            raise
        depth = 0
        in_string = False
        escaped = False
        end = -1
        for idx, ch in enumerate(content[start:], start=start):
            if in_string:
                if escaped:
                    escaped = False
                elif ch == "\\":
                    escaped = True
                elif ch == '"':
                    in_string = False
                continue
            if ch == '"':
                in_string = True
            elif ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    end = idx
                    break
        if end > start:
            return json.loads(content[start : end + 1])
        raise


def get_db_conn() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with get_db_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS analysis_runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_name TEXT NOT NULL,
                analyzed_at TEXT NOT NULL,
                result_json TEXT NOT NULL
            )
            """
        )


def list_analysis_runs(file_name: str, start_date: str, end_date: str) -> list[sqlite3.Row]:
    sql = "SELECT id, file_name, analyzed_at FROM analysis_runs WHERE 1=1"
    params: list[str] = []
    if file_name.strip():
        sql += " AND file_name LIKE ?"
        params.append(f"%{file_name.strip()}%")
    if start_date.strip():
        sql += " AND analyzed_at >= ?"
        params.append(f"{start_date.strip()} 00:00")
    if end_date.strip():
        sql += " AND analyzed_at <= ?"
        params.append(f"{end_date.strip()} 23:59")
    sql += " ORDER BY id DESC"
    with get_db_conn() as conn:
        return list(conn.execute(sql, params).fetchall())


def fetch_analysis_run(run_id: int) -> Optional[sqlite3.Row]:
    with get_db_conn() as conn:
        return conn.execute(
            "SELECT id, file_name, analyzed_at, result_json FROM analysis_runs WHERE id = ?",
            (run_id,),
        ).fetchone()


def save_analysis_run(file_name: str, result_json: dict) -> int:
    now_text = datetime.now().strftime("%Y-%m-%d %H:%M")
    with get_db_conn() as conn:
        cur = conn.execute(
            "INSERT INTO analysis_runs (file_name, analyzed_at, result_json) VALUES (?, ?, ?)",
            (file_name, now_text, json.dumps(result_json, ensure_ascii=False)),
        )
        conn.commit()
        return int(cur.lastrowid)


def build_run_rows(file_path: Path) -> list[MessageRecord]:
    rows: list[MessageRecord] = []
    for item in parse_chat_file(file_path):
        rows.append(item)
    rows.sort(key=lambda x: x.sent_at)
    return rows


def build_run_rows_many(paths: list[Path]) -> list[MessageRecord]:
    rows: list[MessageRecord] = []
    for p in paths:
        for item in parse_chat_file(p):
            rows.append(item)
    rows.sort(key=lambda x: x.sent_at)
    return rows


def load_draw_font(size: int) -> ImageFont.ImageFont:
    # Try common Chinese fonts on Windows first.
    for font_path in [
        r"C:\Windows\Fonts\msyh.ttc",
        r"C:\Windows\Fonts\simhei.ttf",
        r"C:\Windows\Fonts\simsun.ttc",
    ]:
        try:
            return ImageFont.truetype(font_path, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def build_analysis_image_bytes(run_id: int, file_name: str, result: dict) -> bytes:
    summary = sanitize_xlsx_text(result.get("summary", "")) or "无摘要"
    users = result.get("users") or []
    opportunities = result.get("opportunities") or []
    prices = result.get("local_price_analysis") or []

    users_sorted = sorted(users, key=lambda x: x.get("message_count", 0), reverse=True)[:5]

    lines: list[str] = [
        f"分析结果 #{run_id}",
        f"文件: {file_name}",
        "",
        "摘要:",
    ]
    lines.extend(textwrap.wrap(summary, width=60) or ["无"])
    lines.extend(["", "Top用户(前5):"])
    if users_sorted:
        for u in users_sorted:
            lines.append(
                f"- {u.get('nickname','')} {u.get('phone','')} "
                f"发言{u.get('message_count',0)}次 影响力{u.get('influence_score','')}"
            )
    else:
        lines.append("- 无")

    lines.extend(["", "商机分析:"])
    if opportunities:
        for o in opportunities[:3]:
            name = sanitize_xlsx_text(o.get("name", "")) or "未命名"
            reason = sanitize_xlsx_text(o.get("reason", "")) or ""
            lines.append(f"- {name}")
            lines.extend([f"  {w}" for w in textwrap.wrap(reason, width=56)[:2]])
    else:
        lines.append("- 无")

    lines.extend(["", "车型当地售价分析(全部):"])
    if prices:
        for p in prices:
            lines.append(
                f"- {p.get('car_model','')} | {p.get('location','')} | "
                f"{p.get('price','')} | USD {p.get('price_usd','')}"
            )
    else:
        lines.append("- 无")

    # Measure needed height
    font = load_draw_font(24)
    line_h = 34
    pad_x = 30
    pad_y = 26
    width = 1400
    height = pad_y * 2 + max(10, len(lines)) * line_h

    img = Image.new("RGB", (width, height), color=(248, 250, 252))
    draw = ImageDraw.Draw(img)
    draw.rectangle((0, 0, width, 64), fill=(28, 68, 180))
    title_font = load_draw_font(30)
    draw.text((pad_x, 14), "OpenClaw 会话分析结果", fill=(255, 255, 255), font=title_font)

    y = 78
    for line in lines:
        draw.text((pad_x, y), sanitize_xlsx_text(line), fill=(18, 26, 38), font=font)
        y += line_h

    output = io.BytesIO()
    img.save(output, format="PNG")
    output.seek(0)
    return output.getvalue()


def sanitize_base_name(name: str) -> str:
    cleaned = re.sub(r"\.(txt|html|htm|json)$", "", name, flags=re.IGNORECASE).strip()
    cleaned = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff._\- ]+", "_", cleaned)
    cleaned = re.sub(r"\s+", "_", cleaned).strip("._")
    return cleaned or "chat"


def touch_heartbeat() -> None:
    global _last_heartbeat_ts
    with _heartbeat_lock:
        _last_heartbeat_ts = time.time()


@app.route("/raw", methods=["GET", "POST"])
def raw_index():
    touch_heartbeat()
    rows: list[MessageRecord] = []
    error = ""
    analysis_result = ""
    analysis_error = ""

    export_file = request.form.get("export_file", "").strip()
    group_name = request.form.get("group_name", "").strip()
    start = parse_dt_local(request.form.get("start_at", ""))
    end = parse_dt_local(request.form.get("end_at", ""))

    if request.method == "POST":
        files, resolve_err = resolve_target_files(export_file, group_name)
        if resolve_err:
            error = resolve_err
        elif start is None or end is None:
            error = "请选择完整的起止时间（到分钟）。"
        elif start > end:
            error = "起始时间不能晚于结束时间。"
        else:
            group_filter: Optional[str] = None if export_file else group_name
            for file_path in files:
                for item in parse_chat_file(file_path, group_filter):
                    if start <= item.sent_at <= end:
                        rows.append(item)
            rows.sort(key=lambda x: x.sent_at)
            if request.form.get("action") == "analyze" and rows:
                prompt = build_analysis_prompt(rows, start, end)
                analysis_result, analysis_error = call_gemini_analysis(prompt)

    return render_template(
        "index.html",
        rows=rows,
        error=error,
        analysis_result=analysis_result,
        analysis_error=analysis_error,
        upload_msg=request.args.get("upload_msg", "").strip(),
        group_name=group_name,
        export_file=export_file,
        available_exports=list_export_files(),
        start_at=request.form.get("start_at", ""),
        end_at=request.form.get("end_at", ""),
        export_dir=str(EXPORT_DIR),
    )


@app.route("/", methods=["GET"])
def home():
    touch_heartbeat()
    return redirect(url_for("analysis_index"))


@app.route("/analysis", methods=["GET"])
def analysis_index():
    touch_heartbeat()
    q_file = request.args.get("file_name", "").strip()
    q_start = request.args.get("start_date", "").strip()
    q_end = request.args.get("end_date", "").strip()
    runs = list_analysis_runs(q_file, q_start, q_end)
    return render_template(
        "analysis_index.html",
        runs=runs,
        file_name=q_file,
        start_date=q_start,
        end_date=q_end,
        available_exports=list_export_files(),
    )


@app.route("/analysis/run", methods=["POST"])
def analysis_run():
    touch_heartbeat()
    selected = [x.strip() for x in request.form.getlist("export_files") if x.strip()]
    if not selected:
        legacy = request.form.get("export_file", "").strip()
        if legacy:
            selected = [legacy]
    if not selected:
        return Response("请选择要分析的文件（可多选）。", status=400, mimetype="text/plain")

    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    paths: list[Path] = []
    for name in selected:
        path = EXPORT_DIR / name
        if not is_export_file_path(path):
            return Response(f"无效或不允许的文件：{name}", status=400, mimetype="text/plain")
        paths.append(path)

    rows = build_run_rows_many(paths)
    if not rows:
        return Response("该文件没有可解析消息。", status=400, mimetype="text/plain")

    prompt = build_analysis_prompt(rows, rows[0].sent_at, rows[-1].sent_at)
    text, err = call_gemini_analysis(prompt)
    if err:
        return Response(err, status=502, mimetype="text/plain")
    try:
        parsed = extract_json_object(text)
    except Exception as exc:  # noqa: BLE001
        return Response(f"分析结果 JSON 解析失败：{exc}", status=502, mimetype="text/plain")

    parsed["local_price_analysis"] = normalize_and_merge_price_analysis(
        parsed.get("local_price_analysis") or [],
        rows,
    )

    label = paths[0].name if len(paths) == 1 else " + ".join(p.name for p in paths)
    run_id = save_analysis_run(label, parsed)
    return redirect(url_for("analysis_detail", run_id=run_id))


@app.route("/analysis/open-exports", methods=["POST"])
def open_exports_folder():
    touch_heartbeat()
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    try:
        os.startfile(str(EXPORT_DIR))
    except OSError as exc:
        return Response(f"打开目录失败：{exc}", status=500, mimetype="text/plain")
    return Response("ok", status=200, mimetype="text/plain")


@app.route("/analysis/<int:run_id>", methods=["GET"])
def analysis_detail(run_id: int):
    touch_heartbeat()
    run = fetch_analysis_run(run_id)
    if not run:
        return Response("记录不存在。", status=404, mimetype="text/plain")
    result = json.loads(run["result_json"])
    users = result.get("users") or []
    users = sorted(users, key=lambda x: x.get("message_count", 0), reverse=True)
    opportunities = result.get("opportunities") or []
    local_price_analysis = result.get("local_price_analysis") or []
    return render_template(
        "analysis_detail.html",
        run=run,
        users=users,
        opportunities=opportunities,
        local_price_analysis=local_price_analysis,
        summary=result.get("summary", ""),
    )


@app.route("/analysis/<int:run_id>/export", methods=["GET"])
def analysis_export(run_id: int):
    touch_heartbeat()
    run = fetch_analysis_run(run_id)
    if not run:
        return Response("记录不存在。", status=404, mimetype="text/plain")
    result = json.loads(run["result_json"])
    users = result.get("users") or []

    wb = Workbook()
    ws = wb.active
    ws.title = "用户数据分析"
    ws.append(["编号", "手机号", "昵称", "发言次数", "用户类型", "影响力", "特点分析", "建议"])
    total = len(users)
    for idx, user in enumerate(users, start=1):
        mc = user.get("message_count", 0)
        try:
            mc = int(mc)
        except (TypeError, ValueError):
            mc = 0
        inf = user.get("influence_score", "")
        try:
            inf = int(inf) if inf != "" and inf is not None else None
        except (TypeError, ValueError):
            inf = sanitize_xlsx_text(inf) or None

        ws.append(
            [
                total - idx + 1,
                sanitize_xlsx_text(user.get("phone", "")),
                sanitize_xlsx_text(user.get("nickname", "")),
                mc,
                sanitize_xlsx_text(user.get("user_type", "")),
                inf,
                sanitize_xlsx_text(user.get("analysis", "")),
                sanitize_xlsx_text(user.get("suggestion", "")),
            ]
        )

    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", prefix="analysis_export_")
    os.close(fd)
    try:
        wb.save(tmp_path)
    except Exception:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass
        raise

    filename = f"analysis_{run_id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    try:
        with open(tmp_path, "rb") as handle:
            payload = handle.read()
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    if not zipfile.is_zipfile(io.BytesIO(payload)):
        return Response("导出文件生成失败（非有效 xlsx）。", status=500, mimetype="text/plain; charset=utf-8")

    headers = {
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": f'attachment; filename="{filename}"',
        "Content-Length": str(len(payload)),
        "Cache-Control": "no-store",
    }
    return Response(payload, headers=headers, direct_passthrough=True)


@app.route("/export", methods=["POST"])
def export_csv():
    touch_heartbeat()
    export_file = request.form.get("export_file", "").strip()
    group_name = request.form.get("group_name", "").strip()
    start = parse_dt_local(request.form.get("start_at", ""))
    end = parse_dt_local(request.form.get("end_at", ""))

    files, resolve_err = resolve_target_files(export_file, group_name)
    if resolve_err or start is None or end is None or start > end:
        return Response("参数无效，请返回页面重新设置。", status=400, mimetype="text/plain")

    if not files:
        return Response("未找到导出文件。", status=404, mimetype="text/plain")

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["发言人名称", "发言人手机号码", "发言时间", "发言内容"])

    collected: list[MessageRecord] = []
    group_filter: Optional[str] = None if export_file else group_name
    for file_path in files:
        for item in parse_chat_file(file_path, group_filter):
            if start <= item.sent_at <= end:
                collected.append(item)

    collected.sort(key=lambda x: x.sent_at)
    for row in collected:
        writer.writerow(
            [
                row.speaker_name,
                row.speaker_phone,
                row.sent_at.strftime(OUTPUT_DATE_FORMAT),
                row.content,
            ]
        )

    csv_data = "\ufeff" + output.getvalue()
    output.close()

    slug = Path(export_file).stem if export_file else group_name or "chat"
    slug_safe = re.sub(r'[^A-Za-z0-9._\-]+', "_", slug)[:80] or "chat"
    filename = f"whatsapp_export_{slug_safe}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response(
        csv_data,
        mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.route("/heartbeat", methods=["POST"])
def heartbeat():
    touch_heartbeat()
    return Response(status=204)


@app.route("/import", methods=["POST"])
def import_exports():
    touch_heartbeat()
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    incoming_files = request.files.getlist("chat_files")
    if not incoming_files:
        return redirect(url_for("raw_index", upload_msg="未选择文件。"))

    imported = 0
    skipped = 0
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    for incoming in incoming_files:
        original_name = (incoming.filename or "").strip()
        if not original_name:
            skipped += 1
            continue
        ext = Path(original_name).suffix.lower()
        if ext not in EXPORT_SUFFIXES:
            skipped += 1
            continue
        safe_name = sanitize_base_name(Path(original_name).stem)
        save_name = f"{timestamp}_{safe_name}{ext}"
        save_path = EXPORT_DIR / save_name
        counter = 1
        while save_path.exists():
            save_path = EXPORT_DIR / f"{timestamp}_{safe_name}_{counter}{ext}"
            counter += 1
        incoming.save(save_path)
        imported += 1

    msg = f"已导入 {imported} 个文件"
    if skipped:
        msg += f"，跳过 {skipped} 个不支持的格式或无效文件"
    return redirect(url_for("raw_index", upload_msg=msg))


@app.route("/api/openclaw/analyze", methods=["POST"])
def openclaw_analyze():
    """
    Channel integration endpoint:
    1) Receive uploaded chat file (.txt / .html / .json)
    2) Save to fixed exports dir
    3) Run Gemini analysis and store run history
    4) Return one PNG image as analysis result
    """
    touch_heartbeat()
    expected = (os.getenv("OPENCLAW_AUTH_TOKEN") or "").strip()
    if not expected:
        return Response(
            "服务端未配置 OPENCLAW_AUTH_TOKEN。",
            status=500,
            mimetype="text/plain; charset=utf-8",
        )

    provided = (
        request.headers.get("X-OpenClaw-Token", "").strip()
        or request.headers.get("Authorization", "").replace("Bearer ", "").strip()
    )
    if provided != expected:
        return Response("未授权：token 无效。", status=401, mimetype="text/plain; charset=utf-8")

    incoming = request.files.get("file")
    if incoming is None:
        return Response("缺少文件字段 file。", status=400, mimetype="text/plain; charset=utf-8")
    original_name = (incoming.filename or "").strip()
    ext = Path(original_name).suffix.lower()
    if ext not in EXPORT_SUFFIXES:
        return Response(
            f"仅支持以下扩展名：{', '.join(sorted(EXPORT_SUFFIXES))}。",
            status=400,
            mimetype="text/plain; charset=utf-8",
        )

    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    safe_name = sanitize_base_name(Path(original_name).stem)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    save_path = EXPORT_DIR / f"{ts}_{safe_name}{ext}"
    n = 1
    while save_path.exists():
        save_path = EXPORT_DIR / f"{ts}_{safe_name}_{n}{ext}"
        n += 1
    incoming.save(save_path)

    rows = build_run_rows(save_path)
    if not rows:
        return Response("该文件没有可解析消息。", status=400, mimetype="text/plain; charset=utf-8")

    prompt = build_analysis_prompt(rows, rows[0].sent_at, rows[-1].sent_at)
    text, err = call_gemini_analysis(prompt)
    if err:
        return Response(err, status=502, mimetype="text/plain; charset=utf-8")
    try:
        parsed = extract_json_object(text)
    except Exception as exc:  # noqa: BLE001
        return Response(f"分析结果 JSON 解析失败：{exc}", status=502, mimetype="text/plain; charset=utf-8")

    parsed["local_price_analysis"] = normalize_and_merge_price_analysis(
        parsed.get("local_price_analysis") or [],
        rows,
    )

    run_id = save_analysis_run(save_path.name, parsed)
    image_bytes = build_analysis_image_bytes(run_id, save_path.name, parsed)
    headers = {
        "Content-Type": "image/png",
        "Content-Disposition": f'inline; filename="analysis_{run_id}.png"',
        "X-Analysis-Run-Id": str(run_id),
        "X-Analysis-Detail-URL": f"/analysis/{run_id}",
    }
    return Response(image_bytes, headers=headers)


if __name__ == "__main__":
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    init_db()
    print(
        "\nServer started: open http://127.0.0.1:5000 in your browser.\n"
        "Keep this terminal running; press Ctrl+C to stop.\n"
    )
    app.run(host="127.0.0.1", port=5000, debug=False)
