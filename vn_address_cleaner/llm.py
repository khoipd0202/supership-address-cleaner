from __future__ import annotations

import json
import os
import re
import ssl
import time
import unicodedata
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

from parse_address import strip_diacritics

try:
    from clean_address import (
        _ensure_level4_prefix,
        _ensure_street_prefix,
        _normalize_abbrev_v2,
        _pretty_piece_v2,
        _pretty_street_v2,
        _trim_poi_value_v2,
    )
except ImportError:
    _ensure_level4_prefix = None
    _ensure_street_prefix = None
    _normalize_abbrev_v2 = None
    _pretty_piece_v2 = None
    _pretty_street_v2 = None
    _trim_poi_value_v2 = None

_DOTENV_LOADED = False
CEREBRAS_API_URL = "https://api.cerebras.ai/v1/chat/completions"
CEREBRAS_DEFAULT_MODEL = "gpt-oss-120b"


class CerebrasRateLimitError(RuntimeError):
    """Raised when Cerebras free-tier rate/token quota is reached."""


def _load_local_dotenv() -> None:
    global _DOTENV_LOADED
    if _DOTENV_LOADED:
        return
    _DOTENV_LOADED = True
    candidates = [
        Path.cwd() / ".env",
        Path(__file__).resolve().parent.parent / ".env",
    ]
    for path in candidates:
        if not path.exists():
            continue
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value

ADDRESS_SCHEMA = {
    "type": "object",
    "properties": {
        "results": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "row_id": {"type": "integer"},
                    "poi": {"type": "string", "nullable": True},
                    "house_number": {"type": "string", "nullable": True},
                    "street": {"type": "string", "nullable": True},
                    "level4": {"type": "string", "nullable": True},
                    "ward": {"type": "string"},
                    "district": {"type": "string"},
                    "province": {"type": "string"},
                    "clean_address": {"type": "string"},
                    "confidence": {"type": "number"},
                    "flags": {
                        "type": "array",
                        "items": {"type": "string"}
                    },
                    "evidence": {
                        "type": "object",
                        "properties": {
                            "poi_span": {"type": "string", "nullable": True},
                            "house_number_span": {"type": "string", "nullable": True},
                            "street_span": {"type": "string", "nullable": True},
                            "level4_span": {"type": "string", "nullable": True}
                        },
                        "required": [
                            "poi_span",
                            "house_number_span",
                            "street_span",
                            "level4_span"
                        ]
                    }
                },
                "required": [
                    "row_id",
                    "poi",
                    "house_number",
                    "street",
                    "level4",
                    "ward",
                    "district",
                    "province",
                    "clean_address",
                    "confidence",
                    "flags",
                    "evidence"
                ]
            }
        }
    },
    "required": ["results"]
}

SYSTEM_RULES = """
Bạn là address parser cho địa chỉ giao vận Việt Nam.

Nhiệm vụ:
- Chỉ bóc POI, số nhà, đường/phố, cấp 4 từ raw_address.
- ward, district, province từ cột riêng là source of truth.
- Không được tự sửa ward, district, province.
- Nếu raw_address có admin khác với cột riêng, chỉ gắn flag, không ghi đè.
- Chỉ trả component nếu có bằng chứng nguyên văn trong raw_address.
- house_number là field riêng; không được gộp house_number vào street.
- street chỉ là tên đường/phố/quốc lộ/tỉnh lộ/hương lộ hoặc tên đường sau số nhà.

Luật bắt buộc:
1. POI phải dừng trước số nhà, đường/phố, ngõ/hẻm/ngách/kiệt, xã/phường, huyện/quận, tỉnh/tp.
2. Cấp 4 phải dừng trước admin alias.
3. Cấp 4 hợp lệ gồm: thôn, xóm, ấp, khu phố, KP, tổ, tổ dân phố, TDP, khối, bản, buôn, làng, sóc, khóm, đội, tiểu khu, cụm dân cư, khu dân cư.
4. Không bắt "Bản" trong "Huyện Vụ Bản" làm cấp 4.
5. Không bắt "Sở" trong "Trụ sở" làm POI loại Sở.
6. Không bắt câu giao tiếp như "em gửi về", "địa chỉ", "ship về" làm POI.
7. Nếu có "cũ", "mới", "nay thuộc", "sáp nhập", "trước khi sáp nhập", "sau khi sáp nhập" thì phải gắn flag review và confidence <= 0.60 nếu có conflict.
8. Số nhà có thể là: 40, 122, 1F, 259/C3, 259/ C3, 63/2/1, 18C/1N, 12A/5B.
9. Nếu sau số nhà là cụm chữ trước admin hoặc dấu phẩy, cụm đó thường là đường/phố.
10. Nếu không chắc, để null và gắn flag. Không được đoán.
11. evidence.*_span phải là đoạn xuất hiện trong raw_address. Nếu không có span thì field tương ứng phải là null.
12. Không suy luận POI/street/level4 từ ward, district, province.
13. Bỏ các landmark chỉ hướng/vị trí tương đối như: đối diện, gần, cạnh, bên cạnh, sau, trước, cổng trường, cổng bệnh viện.
14. Không trả POI nếu chỉ có loại hình chung mà không có tên riêng, ví dụ: Quầy thuốc, Nhà thuốc, Công ty, Shop, Cửa hàng, Trường, Bệnh viện.
15. Output phải sạch và cụ thể, hạn chế viết tắt: Cty -> Công ty, BV -> Bệnh viện, PK -> Phòng khám, MN -> Mầm non, TTYT -> Trung tâm y tế, KCN -> Khu công nghiệp, KCX -> Khu chế xuất, KP -> Khu phố, TDP -> Tổ dân phố.
16. Nếu cụm không tiền tố như "Đan Thầm" có thể là đường hoặc thôn/xóm/ấp nhưng raw không có bằng chứng, để null thay vì đoán.
18. Nếu candidate cấp 4 không có loại đơn vị (vd "Thụy Lôi") mà theo ngữ cảnh chắc chắn là thôn/xóm/ấp, trả về kèm tiền tố đúng (vd "Thôn Thụy Lôi"); không chắc loại thì null + flag.
17. Trả về đúng JSON schema, không giải thích ngoài JSON.
"""

CEREBRAS_COMPACT_RULES = """Bạn là parser địa chỉ Việt Nam. Chỉ trả JSON object:
{"results":[{"row_id":int,"poi":str|null,"house_number":str|null,"street":str|null,"level4":str|null,"ward":str,"district":str,"province":str,"clean_address":str,"confidence":number,"flags":[],"evidence":{"poi_span":str|null,"house_number_span":str|null,"street_span":str|null,"level4_span":str|null}}]}
Luật gắt:
- ward/district/province lấy nguyên từ input, không sửa.
- Chỉ trả poi/street/level4/house_number nếu có evidence_span nguyên văn trong raw.
- Không đoán. Không chắc thì null và thêm flag.
- house_number tách riêng, không gộp vào street.
- street chỉ là đường/phố/QL/TL/HL hoặc tên đường rõ sau số nhà; không trả ngõ/ngách/hẻm/kiệt làm street.
- Không được trả một mảnh bị cắt cụt của candidate (vd "Đê" từ "Đê Long Biên", "Bến" từ "Bến Phú Định", "Chi" từ "Chi Quan"). Nếu không xác định được trọn tên đường thì street=null.
- Cụm thôn/xóm/ấp, chung cư, nội thất, tiệm/cửa hàng sửa xe, tòa/lô/căn hộ không phải tên đường.
- Nếu raw có ngách/ngõ/hẻm/kiệt trước đường thì bỏ phần ngách/ngõ/hẻm/kiệt, chỉ xét tên đường thật.
- Nếu raw có nhiều POI dính nhau bằng dấu chấm/dấu câu thì xác định ranh giới, không gộp bừa.
- Chuẩn hóa dấu tiếng Việt cho địa danh nếu chắc chắn từ ngữ trong raw.
- Cấp 4 chỉ gồm thôn/xóm/ấp/khu phố/KP/tổ/TDP/khối/bản/buôn/làng/sóc/khóm/đội/tiểu khu/cụm dân cư/khu dân cư.
- Nếu candidate cấp 4 KHÔNG có loại đơn vị (vd "Thụy Lôi") mà theo ngữ cảnh chắc chắn là thôn/xóm/ấp, trả về KÈM tiền tố đúng (vd "Thôn Thụy Lôi"). Không chắc loại nào thì null + flag, không đoán.
- Nếu có nhiều candidate cấp 4, chọn đơn vị đúng cấp thôn/xóm/ấp làm level4, loại candidate không phải cấp 4.
- POI phải có tên riêng; bỏ loại chung như Công ty/Nhà thuốc/Quầy thuốc/Shop/Cửa hàng/Trường/Bệnh viện nếu không có tên.
- Bỏ landmark tương đối như đối diện/gần/cạnh/trước/sau/cổng trường/cổng bệnh viện.
- Mở rộng viết tắt trong output: Cty=Công ty, BV=Bệnh viện, PK=Phòng khám, MN=Mầm non, THCS=Trung học cơ sở, KCN=Khu công nghiệp, KP=Khu phố, TDP=Tổ dân phố.
- Nếu có cũ/mới/nay thuộc/sáp nhập thì confidence <=0.60 và thêm flag.
- Khôi phục dấu tiếng Việt và sửa lỗi gõ nhỏ 1-2 ký tự nếu chắc chắn: "hòa bih"->"Hòa Bình", "nguyen van troi"->"Nguyễn Văn Trỗi". evidence_span vẫn là đoạn NGUYÊN VĂN trong raw.
- Mã đường viết tắt (ĐX 7, TCH 18, XTT 59, N5...) là tên đường hợp lệ, giữ dạng mã và thêm "Đường " phía trước.
- Nếu rule candidate bị cắt cụt (vd rule đưa "Thôn Chiên" mà raw là "Thôn Chiên Noi..."), trả về tên đầy đủ đúng theo raw, chuẩn hóa dấu.
- Nếu raw drop hết mà vẫn còn nội dung ngoài admin, cố bóc POI/đường/cấp 4 từ phần đó theo đúng luật evidence.
- Không markdown, không giải thích, không thêm key ngoài schema."""

SEND_LLM_FLAGS = {
    "INFERRED_LEVEL4_NO_KEYWORD",
    "POSSIBLE_POI_MISSED",
    "POSSIBLE_LEVEL4_MISSED",
    "MULTIPLE_POI_FOUND",
    "MULTIPLE_LEVEL4_FOUND",
    "POI_STREET_BOUNDARY_AMBIGUOUS",
    "LEVEL4_ADMIN_BOUNDARY_AMBIGUOUS",
    "LEVEL4_NAME_BOUNDARY_AMBIGUOUS",
    "POI_LEVEL4_OVERLAP",
    "STREET_LEVEL4_OVERLAP",
    "UNPREFIXED_STREET_FROM_HOUSE_NUMBER",
    "UNPREFIXED_STREET_AMBIGUOUS_DROPPED",
    "HOUSE_NUMBER_BIS_AMBIGUOUS",
    "HOUSE_CODE_STREET_AMBIGUOUS",
    "LEVEL4_LONG_NAME_AMBIGUOUS",
    "LEVEL4_MISSING_UNIT_PREFIX",
    "POI_ADMIN_OVERLAP_TRIMMED",
    "POI_ADMIN_TAIL_TRIMMED",
    "POI_MULTI_OBJECT_AMBIGUOUS",
    "STREET_NAME_SUSPICIOUS",
    # rule drop sạch nhưng raw vẫn còn nội dung -> LLM cứu
    "EMPTY_RESULT_RAW_HAS_CONTENT",
    "STRICT_DROPPED_DETAIL",
    "RAW_BAD_SPELLING_DROPPED_DETAIL",
    "ONLY_STREET_LEVEL_FOUND",
}

REVIEW_MANUAL_FLAGS = {
    "ADMIN_CONFLICT_HIGH_RISK",
    "MISSING_ADMIN_COLUMN",
    "MULTIPLE_ADDRESS_BLOCKS",
    "NEED_REVIEW_ADMIN_VERSION",
    "RAW_HAS_OLD_AND_NEW_ADDRESS",
    "UNKNOWN_LONG_SEGMENT_BEFORE_ADMIN",
}

ACCEPT_WITH_FLAGS = {
    "ADMIN_FILLED_FROM_RAW",
    "ADMIN_VERSION_AMBIGUOUS",
    "INVALID_POI_BLACKLIST",
    "OLD_NEW_ADMIN_CONFLICT",
    "RAW_ADMIN_CONFLICT_WITH_COLUMNS",
    "RAW_CONTAINS_NAY_THUOC",
    "TRU_SO_DETECTED",
    "UNKNOWN_BEFORE_ADMIN",
}

ACCEPT_FLAGS = {
    "ADMIN_NEW_FORMAT",
    "COLON_PREFIX_REMOVED",
    "ORDER_CODE_REMOVED",
    "LEADING_NOTE_REMOVED",
    "LEADING_PERSON_REMOVED",
    "LOW_CONFIDENCE",
    "NOTE_AFTER_ADDRESS_REMOVED",
    "NO_LEVEL4_FOUND",
    "NO_POI_FOUND",
    "ONLY_ADMIN_UNITS_FOUND",
    "PRODUCT_NOTE_REMOVED",
    "RAW_CONTAINS_PHONE",
    "RAW_CONTAINS_PERSON_NAME",
    "STREET_FOUND_FROM_HOUSE_NUMBER",
    "TRAILING_AFTER_PROVINCE_REMOVED",
}

# Backward-compatible names for callers/tests that imported these helpers.
HARD_FLAGS = SEND_LLM_FLAGS
AMBIGUOUS_FLAGS = SEND_LLM_FLAGS

OLD_NEW_KEYWORDS = [
    "cũ", "mới", "nay thuộc", "nay là",
    "sáp nhập", "trước khi sáp nhập", "sau khi sáp nhập"
]
OLD_NEW_KEYWORD_RE = re.compile(
    r"\b(?:nay\s+thuoc|nay\s+la|sap\s+nhap|truoc\s+khi\s+sap\s+nhap|"
    r"sau\s+khi\s+sap\s+nhap)\b",
    re.I,
)
ACCENTED_OLD_NEW_RE = re.compile(r"\b(?:cũ|mới)\b", re.I)

POI_SIGNAL_PHRASES = {
    "agribank", "atm", "bank", "benh vien", "bhxh", "bidv", "buu dien",
    "cafe", "caffe", "cao dang", "cay xang", "cho", "chua", "chung cu", "cong an",
    "doc",
    "cong ty", "cty", "cum cong nghiep", "dai hoc", "den", "garage",
    "gara", "khach san", "kho", "khu che xuat", "khu cong nghiep", "khu do thi",
    "mam non", "mbbank", "mieu", "nha may", "nha nghi", "nha sach",
    "nha tro", "noi that", "tiem sua xe", "sua xe",
    "nha thuoc", "nha tho", "nha van hoa", "phong kham", "pk", "plaza",
    "quay thuoc", "sieu thi", "shop", "toa nha", "tower", "tram bom", "tram y te",
    "truong", "trung tam", "ttyt", "ubnd", "uy ban", "vietcombank",
    "vincom", "vinhomes", "vpbank", "xuong",
}

GENERIC_POI_ONLY = {
    "atm", "bank", "benh vien", "buu dien", "cua hang", "cho",
    "chung cu", "cong ty", "cty", "doanh nghiep", "ga", "khach san",
    "kho", "ngan hang", "nha may", "nha nghi", "nha sach", "nha thuoc", "nha tro",
    "phong kham", "pk", "quay thuoc", "shop", "sieu thi", "toa nha",
    "tram bom", "tram y te", "truong", "trung tam", "ttyt", "ubnd", "uy ban",
    "uy ban nhan dan", "xi nghiep",
}

KNOWN_RESIDENTIAL_POI_NAMES = {
    "lang viet kieu chau au",
    "tran anh ashita",
}

RELATIVE_LANDMARK_RE = re.compile(
    r"\b(?:doi\s+dien|gan|canh|ben\s+canh|ke\s+ben|phia\s+sau|dang\s+sau|"
    r"sau|phia\s+truoc|dang\s+truoc|truoc|cach|cong|"
    r"nga\s+(?:ba|tu|\d))\b",
    re.I,
)

LEVEL4_SIGNAL_PHRASES = {
    "ap", "ban", "buon", "cum dan cu", "doi", "khom", "khoi", "khu",
    "khu dan cu", "khu pho", "lang", "soc", "tdp", "thon", "tieu khu",
    "to", "to dan pho", "xom",
}

STREET_SIGNAL_PHRASES = {
    "cao toc", "dt", "duong", "hem", "hl", "huong lo", "kiet", "km",
    "ngach", "ngo", "pho", "ql", "quoc lo", "tinh lo", "tl",
}

COMMUNICATION_BLACKLIST = {
    "dia chi",
    "dia chi nguoi nhan",
    "doi dia chi ve",
    "em gui ve",
    "em goi ve",
    "giao den",
    "giao ve",
    "gui cho chi",
    "gui ve",
    "nguoi nhan",
    "ship den",
    "ship ve",
}

ADMIN_PREFIXES = {
    "phuong", "xa", "quan", "huyen", "tinh", "tp", "tt", "tx",
    "thanh pho", "thi xa", "thi tran",
}

LEVEL4_PREFIXES = {
    "ap", "ban", "buon", "cum dan cu", "doi", "khom", "khoi", "khu",
    "khu dan cu", "khu pho", "lang", "soc", "tdp", "thon", "tieu khu",
    "to", "to dan pho", "xom",
}

# Số nhà: bắt đầu bằng SỐ (25, 25A, 63/2/1); chữ đứng trước chỉ hợp lệ khi
# là dạng ghép có "/" (B3/24S, C1/29, k29/19 - kiểu Bình Chánh/kiệt Đà Nẵng).
# Chữ + >=4 số đứng trơ (A1103) là mã vận đơn, không phải số nhà.
HOUSE_NUMBER_RE = re.compile(
    r"^(?:"
    r"\d+[A-Za-zĐđ0-9]*"
    r"(?:\s*[/\-.]\s*[A-Za-zĐđ]{0,4}\d+[A-Za-zĐđ0-9]*)*"
    r"|"
    r"[A-Za-zĐđ]{1,4}\d+[A-Za-zĐđ0-9]*"
    r"(?:\s*[/\-.]\s*[A-Za-zĐđ]{0,4}\d+[A-Za-zĐđ0-9]*)+"
    r")"
    r"(?:\s+(?:[A-Za-zĐđ]|bis)(?=\s|$))?$",
    re.I,
)


def _norm(value: Any) -> str:
    text = unicodedata.normalize("NFC", str(value or ""))
    text = strip_diacritics(text.lower())
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _compact(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", _norm(value))


def _clean_optional_text(value: Any) -> str | None:
    if value is None:
        return None
    text = unicodedata.normalize("NFC", str(value))
    text = re.sub(r"\s+", " ", text).strip(" ,.-–—/:;")
    return text or None


def _pretty_llm_component(value: Any) -> str | None:
    text = _clean_optional_text(value)
    if not text:
        return None
    if _pretty_piece_v2:
        text = _pretty_piece_v2(text)
        expansions = [
            (r"\bKhách Sạn\b", "Khách sạn"),
            (r"\bCty\b", "Công ty"),
            (r"\bBV\b", "Bệnh viện"),
            (r"\bPK\b", "Phòng khám"),
            (r"\bMN\b", "Mầm non"),
            (r"\bTTYT\b", "Trung tâm y tế"),
            (r"\bKCN\b", "Khu công nghiệp"),
            (r"\bKCX\b", "Khu chế xuất"),
            (r"\bCCN\b", "Cụm công nghiệp"),
            (r"\bTDP\b", "Tổ dân phố"),
            (r"\bKP\b", "Khu phố"),
            (r"\bUBND\b", "Ủy ban nhân dân"),
            (r"\bTHCS\b", "Trung học cơ sở"),
            (r"\bTHPT\b", "Trung học phổ thông"),
            (r"\bCấp\s*2\b", "Cấp 2"),
            (r"\bCấp\s*3\b", "Cấp 3"),
        ]
        for pat, repl in expansions:
            text = re.sub(pat, repl, text)
        return text
    return text


def _pretty_llm_street(value: Any) -> str | None:
    text = _clean_optional_text(value)
    if not text:
        return None
    if _pretty_street_v2:
        return _pretty_street_v2(text)
    return _pretty_llm_component(text)


def _pretty_llm_level4(value: Any, raw_address: Any) -> str | None:
    text = _pretty_llm_component(value)
    if not text:
        return None
    if _ensure_level4_prefix:
        text = _ensure_level4_prefix(text, raw_address)
    return _accent_known_level4(text)


LEVEL4_ACCENT_FIXES = {
    "hong chau": "Hồng Châu",
}


def _accent_known_level4(value: Any) -> str | None:
    text = _clean_optional_text(value)
    if not text:
        return None
    norm_value = _norm(text)
    prefix_map = {
        "thon": "Thôn",
        "xom": "Xóm",
        "ap": "Ấp",
        "khu pho": "Khu phố",
        "to dan pho": "Tổ dân phố",
        "to": "Tổ",
    }
    for prefix_norm, prefix_display in sorted(prefix_map.items(), key=lambda item: len(item[0]), reverse=True):
        if norm_value == prefix_norm:
            return prefix_display
        if norm_value.startswith(prefix_norm + " "):
            core = norm_value[len(prefix_norm):].strip()
            fixed = LEVEL4_ACCENT_FIXES.get(core)
            if fixed:
                return f"{prefix_display} {fixed}"
    return LEVEL4_ACCENT_FIXES.get(norm_value) or text


def _contains_phrase(norm_text: str, phrases: set[str]) -> bool:
    padded = f" {norm_text} "
    return any(re.search(r"(?<!\w)" + re.escape(p) + r"(?!\w)", padded) for p in phrases)


def _is_generic_poi_only(value: Any) -> bool:
    return _norm(value) in GENERIC_POI_ONLY


def _raw_poi_signal_is_relative_or_generic(raw_address: Any) -> bool:
    norm_raw = _norm(raw_address)
    if not norm_raw:
        return False
    if norm_raw in GENERIC_POI_ONLY:
        return True
    marker = RELATIVE_LANDMARK_RE.search(norm_raw)
    if not marker:
        return False
    before_marker = norm_raw[:marker.start()].strip()
    if before_marker in GENERIC_POI_ONLY:
        return True
    return bool(re.search(r"\b(?:doi\s+dien|gan|canh|ben\s+canh|ke\s+ben)\b", norm_raw))


def _has_relative_landmark_context(value: Any, raw_address: Any) -> bool:
    value_norm = _norm(value)
    raw_norm = _norm(raw_address)
    if not value_norm or not raw_norm:
        return False
    idx = raw_norm.find(value_norm)
    if idx < 0:
        return False
    prefix_window = raw_norm[max(0, idx - 60):idx]
    return bool(RELATIVE_LANDMARK_RE.search(prefix_window))


def _to_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _raw_has_old_new_keyword(raw_address: Any) -> bool:
    raw_text = unicodedata.normalize("NFC", str(raw_address or "")).lower()
    if ACCENTED_OLD_NEW_RE.search(raw_text):
        return True
    norm_raw = _norm(raw_address)
    return bool(OLD_NEW_KEYWORD_RE.search(norm_raw))


LEVEL4_SIGNAL_RE = re.compile(
    r"\b(?:thon|xom|ap|khu\s+pho|kp|to\s+dan\s+pho|tdp|to\s+\d|"
    r"khoi|ban\s+(?:\d|[a-z]{2,})|buon|lang|khom|doi|tieu\s+khu|cum\s+dan\s+cu|"
    r"khu\s+dan\s+cu|khu\s+\d)\b",
    re.I,
)


def _raw_semantic_llm_flags(raw_address: Any, rule_result: dict[str, Any]) -> set[str]:
    norm_raw = _norm(raw_address)
    semantic_flags: set[str] = set()
    if not norm_raw:
        return semantic_flags
    if (
        not rule_result.get("poi")
        and _contains_phrase(norm_raw, POI_SIGNAL_PHRASES)
        and not _raw_poi_signal_is_relative_or_generic(raw_address)
    ):
        semantic_flags.add("POSSIBLE_POI_MISSED")
    level4_scan = re.sub(r"\bdoi\s+dien\b", " ", norm_raw)
    # Không coi chữ Làng/Khu nằm trong một POI dự án đã nhận diện là dấu hiệu
    # bỏ sót cấp 4 (vd Khu đô thị Làng Việt Kiều Châu Âu).
    poi_norm = _norm(rule_result.get("poi"))
    if poi_norm:
        poi_variants = {poi_norm}
        poi_core = re.sub(
            r"^(?:khu do thi|chung cu|khu dan cu|toa nha)\s+",
            "",
            poi_norm,
        ).strip()
        if poi_core:
            poi_variants.add(poi_core)
        for phrase in sorted(poi_variants, key=len, reverse=True):
            level4_scan = level4_scan.replace(phrase, " ")
    if not rule_result.get("level4") and LEVEL4_SIGNAL_RE.search(level4_scan):
        semantic_flags.add("POSSIBLE_LEVEL4_MISSED")
    if not any(rule_result.get(k) for k in ("poi", "street", "level4")):
        if _raw_has_content_beyond_admin(norm_raw, rule_result):
            semantic_flags.add("EMPTY_RESULT_RAW_HAS_CONTENT")
    return semantic_flags


def _raw_has_content_beyond_admin(norm_raw: str, rule_result: dict[str, Any]) -> bool:
    """Rule không bóc được gì nhưng raw vẫn còn chữ ngoài phần hành chính
    -> có thông tin bị bỏ sót, phải đưa LLM."""
    leftover = f" {norm_raw} "
    for field in ("ward", "district", "province"):
        value = rule_result.get(field)
        for form in (_norm(value), _strip_admin_prefix_norm(value)):
            if form and len(form) >= 2:
                leftover = leftover.replace(f" {form} ", " ")
    tokens = [
        t for t in leftover.split()
        if len(t) >= 2
        and not t.isdigit()
        and t not in ADMIN_PREFIXES
        and t not in {"viet", "nam", "vn", "so", "sn", "nha", "dia", "chi"}
    ]
    return len(tokens) >= 2


def has_hard_flags(flags: list[str]) -> bool:
    return bool(set(flags or []) & HARD_FLAGS)


def route_after_rules(rule_result: dict[str, Any]) -> str:
    flags = set(rule_result.get("flags", []) or [])
    raw_address = rule_result.get("raw_address", "")
    flags.update(_raw_semantic_llm_flags(raw_address, rule_result))

    if flags & SEND_LLM_FLAGS:
        return "SEND_LLM"
    if flags & REVIEW_MANUAL_FLAGS or _raw_has_old_new_keyword(raw_address):
        return "REVIEW_MANUAL"
    if flags & ACCEPT_WITH_FLAGS:
        return "ACCEPT_RULE_WITH_FLAGS"
    return "ACCEPT_RULE"


def should_send_to_llm(rule_result: dict[str, Any]) -> bool:
    return route_after_rules(rule_result) == "SEND_LLM"


def _llm_text(value: Any, limit: int = 420) -> str:
    text = unicodedata.normalize("NFC", str(value or "")).strip()
    return text[:limit]


def build_prompt(rows: list[dict[str, Any]]) -> str:
    payload = []
    for row in rows:
        rule_candidates = row.get("rule_candidates") or {}
        rule = {}
        for key in ("poi", "street", "level4"):
            value = rule_candidates.get(key)
            if value:
                rule[key] = value
        flags = list(dict.fromkeys((row.get("flags") or []) + (rule_candidates.get("flags") or [])))
        if flags:
            rule["flags"] = flags
        confidence = row.get("confidence")
        if confidence is not None:
            rule["confidence"] = confidence
        payload.append({
            "row_id": int(row["row_id"]),
            "raw": _llm_text(row.get("raw_address")),
            "ward": _llm_text(row.get("ward"), 120),
            "district": _llm_text(row.get("district"), 120),
            "province": _llm_text(row.get("province"), 120),
            "rule": rule,
        })

    return "Rows JSON:\n" + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))


def _dedupe_text_key(value: Any) -> str:
    text = _norm(value)
    return re.sub(r"\s+", " ", text).strip()


def dedupe_llm_rows(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], dict[tuple[str, str, str, str], list[int]]]:
    """Dedupe ambiguous rows before LLM calls while preserving original row ids."""
    dedup_map: dict[tuple[str, str, str, str], list[int]] = {}
    unique_rows: list[dict[str, Any]] = []
    for item in rows:
        key = (
            _dedupe_text_key(item.get("raw_address")),
            _dedupe_text_key(item.get("ward")),
            _dedupe_text_key(item.get("district")),
            _dedupe_text_key(item.get("province")),
        )
        if key not in dedup_map:
            dedup_map[key] = []
            unique_rows.append(item)
        dedup_map[key].append(int(item["row_id"]))
    return unique_rows, dedup_map


def _extract_json_object(text: Any) -> dict[str, Any]:
    raw = str(text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.I)
        raw = re.sub(r"\s*```$", "", raw)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        start = raw.find("{")
        end = raw.rfind("}")
        if start >= 0 and end > start:
            return json.loads(raw[start:end + 1])
        raise


def _normalise_cerebras_response(parsed: dict[str, Any], rows: list[dict[str, Any]]) -> dict[str, Any]:
    results = parsed.get("results")
    if isinstance(results, list):
        return parsed

    if len(rows) == 1 and any(
        key in parsed
        for key in ("poi", "house_number", "street", "level4", "ward", "district", "province", "evidence")
    ):
        item = dict(parsed)
        item.setdefault("row_id", int(rows[0]["row_id"]))
        item.setdefault("flags", [])
        return {"results": [item]}

    raise ValueError("Cerebras response missing required 'results' array.")


def _cerebras_headers(api_key: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {api_key}",
        "Accept": "application/json",
        "Content-Type": "application/json",
        "User-Agent": "vn-address-cleaner/0.1",
    }


def _https_context() -> ssl.SSLContext | None:
    try:
        import certifi
    except ImportError:
        return None
    return ssl.create_default_context(cafile=certifi.where())


def _post_cerebras(payload: dict[str, Any], api_key: str, timeout: int = 60) -> dict[str, Any]:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urllib.request.Request(
        os.environ.get("CEREBRAS_API_URL", CEREBRAS_API_URL),
        data=data,
        headers=_cerebras_headers(api_key),
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout, context=_https_context()) as response:
            body = response.read().decode("utf-8")
            return json.loads(body)
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", "replace")
        message = f"Cerebras API HTTP {exc.code}: {body}"
        if exc.code == 429:
            raise CerebrasRateLimitError(message) from exc
        raise RuntimeError(message) from exc


def parse_rows_with_cerebras(
    rows: list[dict[str, Any]],
    model: str | None = None,
    api_key: str | None = None,
) -> dict[str, Any]:
    if not rows:
        return {"results": []}

    _load_local_dotenv()
    key = api_key or os.environ.get("CEREBRAS_API_KEY") or os.environ.get("CEREBRAS_API_TOKEN")
    if not key:
        raise ValueError("CEREBRAS_API_KEY environment variable is not set and no API key was provided.")

    if not model:
        model = os.environ.get("CEREBRAS_MODEL") or CEREBRAS_DEFAULT_MODEL
    prompt = build_prompt(rows)
    payload = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": CEREBRAS_COMPACT_RULES,
            },
            {"role": "user", "content": prompt},
        ],
        "temperature": 0,
        "response_format": {
            "type": "json_schema",
            "json_schema": {
                "name": "address_results",
                "strict": True,
                "schema": ADDRESS_SCHEMA,
            },
        },
    }

    max_retries = 2
    last_error: Exception | None = None
    for attempt in range(max_retries):
        try:
            response = _post_cerebras(payload, key)
            content = response["choices"][0]["message"]["content"]
            parsed = _extract_json_object(content)
            return _normalise_cerebras_response(parsed, rows)
        except Exception as e:
            last_error = e
            if isinstance(e, CerebrasRateLimitError):
                raise
            is_decode_error = isinstance(e, json.JSONDecodeError)
            if is_decode_error and len(rows) > 1:
                break

            err_str = str(e)
            if "response_format" in err_str and "400" in err_str and "response_format" in payload:
                # Hạ cấp dần: json_schema -> json_object -> bỏ hẳn
                if payload["response_format"].get("type") == "json_schema":
                    payload["response_format"] = {"type": "json_object"}
                else:
                    payload = {k: v for k, v in payload.items() if k != "response_format"}
                continue

            if "HTTP 401" in err_str or "HTTP 403" in err_str:
                raise e

            is_rate_limited = "429" in err_str or "rate" in err_str.lower() or "quota" in err_str.lower()
            if is_rate_limited:
                raise CerebrasRateLimitError(err_str) from e

            if attempt < max_retries - 1:
                sleep_sec = (attempt + 1) * 2
                print(f"Cerebras API call attempt {attempt + 1} failed: {e}. Retrying in {sleep_sec:.2f}s...")
                time.sleep(sleep_sec)
            else:
                if len(rows) > 1:
                    print(f"Cerebras API call exhausted retries. Error: {e}. Splitting batch...")
                    break
                else:
                    raise e

    if len(rows) <= 1:
        raise last_error or RuntimeError("Cerebras API call failed.")

    # Split batch of size > 1 in half and process recursively
    mid = len(rows) // 2
    left_rows = rows[:mid]
    right_rows = rows[mid:]
    
    print(f"Splitting batch of size {len(rows)} into two sub-batches of size {len(left_rows)} and {len(right_rows)}...")
    left_res = parse_rows_with_cerebras(left_rows, model=model, api_key=api_key)
    right_res = parse_rows_with_cerebras(right_rows, model=model, api_key=api_key)
    
    combined_results = left_res.get("results", []) + right_res.get("results", [])
    return {"results": combined_results}

def _source_admin(original_row: dict[str, Any], field: str) -> str:
    rule_candidates = original_row.get("rule_candidates") or {}
    return _clean_optional_text(original_row.get(field) or rule_candidates.get(field)) or ""


def _strip_admin_prefix_norm(value: Any) -> str:
    norm = _norm(value)
    for prefix in sorted(ADMIN_PREFIXES, key=len, reverse=True):
        if norm == prefix:
            return ""
        if norm.startswith(prefix + " "):
            return norm[len(prefix):].strip()
    return norm


def _admin_aliases_from_row(original_row: dict[str, Any]) -> set[str]:
    aliases = set()
    for field in ("ward", "district", "province"):
        value = _source_admin(original_row, field)
        norm_value = _norm(value)
        core = _strip_admin_prefix_norm(value)
        for alias in (norm_value, core):
            if alias:
                aliases.add(alias)
                aliases.add(alias.replace(" ", ""))
    return aliases


def _is_admin_only(value: Any, original_row: dict[str, Any]) -> bool:
    norm_value = _norm(value)
    compact_value = _compact(value)
    if not norm_value:
        return False
    aliases = _admin_aliases_from_row(original_row)
    core = _strip_admin_prefix_norm(value)
    return (
        norm_value in aliases
        or compact_value in aliases
        or core in aliases
        or core.replace(" ", "") in aliases
    )


def _trim_row_admin_tail(value: Any, original_row: dict[str, Any]) -> str | None:
    """Cắt tên ward/district/province bị LLM nối vào cuối tên đường."""
    text = _clean_optional_text(value)
    if not text:
        return None
    changed = True
    while changed:
        changed = False
        value_norm = _norm(text)
        parts = text.split()
        for field in ("ward", "district", "province"):
            source = _source_admin(original_row, field)
            aliases = {_norm(source), _strip_admin_prefix_norm(source)}
            for alias in sorted((x for x in aliases if x), key=len, reverse=True):
                if not value_norm.endswith(" " + alias):
                    continue
                word_count = len(alias.split())
                remaining = parts[:-word_count]
                remaining_norm = _norm(" ".join(remaining))
                remaining_core = re.sub(
                    r"^(?:duong|pho|de|quoc lo|tinh lo|huong lo)\s+",
                    "",
                    remaining_norm,
                ).strip()
                if len(remaining_core.split()) < 2:
                    continue
                text = " ".join(remaining).strip(" ,.-–—/:;")
                changed = True
                break
            if changed:
                break
    return text or None


def _has_admin_tail(value: Any) -> bool:
    norm_value = _norm(value)
    if not norm_value:
        return False
    padded = f" {norm_value} "
    return any(
        re.search(r"(?<!\w)" + re.escape(prefix) + r"(?!\w)", padded)
        for prefix in ADMIN_PREFIXES
    )


def _is_note_like(value: Any) -> bool:
    norm_value = _norm(value)
    if not norm_value:
        return False
    padded = f" {norm_value} "
    if any(re.search(r"(?<!\w)" + re.escape(x) + r"(?!\w)", padded) for x in COMMUNICATION_BLACKLIST):
        return True
    return bool(re.search(
        r"\b(?:ship|giao|goi|gui|hang|cod|thu ho|size|mau|kg|vnd|"
        r"khong dung|ko dung|tra lai|cho xem|dong kiem)\b",
        norm_value,
    ))


def _raw_contains(raw_address: Any, value: Any) -> bool:
    value_norm = _norm(value)
    if not value_norm:
        return False
    raw_norm = _norm(raw_address)
    if value_norm in raw_norm:
        return True
    value_compact = _compact(value)
    raw_compact = _compact(raw_address)
    return bool(value_compact and value_compact in raw_compact)


def _edit_distance_leq(a: str, b: str, limit: int) -> bool:
    """Levenshtein(a, b) <= limit, cắt sớm để rẻ."""
    if abs(len(a) - len(b)) > limit:
        return False
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        best = i
        for j, cb in enumerate(b, 1):
            cost = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb))
            cur.append(cost)
            best = min(best, cost)
        if best > limit:
            return False
        prev = cur
    return prev[-1] <= limit


def _evidence_supports_value(field: str, value: str, result: dict[str, Any], raw_address: Any) -> bool:
    evidence = result.get("evidence") or {}
    evidence_key = {
        "poi": "poi_span",
        "house_number": "house_number_span",
        "street": "street_span",
        "level4": "level4_span",
    }[field]
    span = _clean_optional_text(evidence.get(evidence_key))
    if not span or not _raw_contains(raw_address, span):
        return False
    if _raw_contains(raw_address, value):
        return True
    value_norm = _norm(value)
    span_norm = _norm(span)
    if value_norm and (value_norm in span_norm or span_norm in value_norm):
        return True
    # LLM sửa lỗi gõ nhỏ ("hòa bih" -> "Hòa Bình"): chấp nhận nếu sai khác
    # tối đa 1-2 ký tự so với evidence span (span vẫn phải nguyên văn trong raw).
    value_compact = value_norm.replace(" ", "")
    span_compact = span_norm.replace(" ", "")
    # bỏ tiền tố loại đường/đơn vị LLM thêm vào trước khi so
    for prefix in ("duong", "pho", "thon", "xom", "ap", "khupho", "todanpho", "to", "tieukhu"):
        if value_compact.startswith(prefix) and not span_compact.startswith(prefix):
            value_compact = value_compact[len(prefix):]
            break
    if len(span_compact) >= 5:
        limit = 1 if len(span_compact) <= 10 else 2
        return _edit_distance_leq(value_compact, span_compact, limit)
    return False


def _starts_with_any(norm_value: str, prefixes: set[str]) -> bool:
    return any(norm_value == prefix or norm_value.startswith(prefix + " ") for prefix in prefixes)


def _is_real_ngo_prefix(value: Any) -> bool:
    norm_value = _norm(value)
    raw_tokens = re.findall(r"\w+", unicodedata.normalize("NFC", str(value or "").lower()), flags=re.UNICODE)
    raw_first = raw_tokens[0] if raw_tokens else ""
    return bool(
        norm_value == "ngo"
        or (
            norm_value.startswith("ngo ")
            and (raw_first == "ngõ" or re.match(r"^ngo\s+\d", norm_value))
        )
    )


def _is_alley_street_value(value: Any) -> bool:
    norm_value = _norm(value)
    raw_tokens = re.findall(r"\w+", unicodedata.normalize("NFC", str(value or "").lower()), flags=re.UNICODE)
    raw_first = raw_tokens[0] if raw_tokens else ""
    if raw_first in {"ngõ", "ngách", "hẻm", "kiệt", "ngach", "hem", "kiet"}:
        return True
    return bool(raw_first == "ngo" and re.match(r"^ngo\s+\d", norm_value))


def _has_street_signal(value: Any) -> bool:
    norm_value = _norm(value)
    signals = set(STREET_SIGNAL_PHRASES)
    signals.discard("ngo")
    return _contains_phrase(norm_value, signals) or _is_real_ngo_prefix(value)


def _is_valid_house_number(value: str) -> bool:
    return bool(HOUSE_NUMBER_RE.fullmatch(value))


def _is_valid_level4(value: str) -> bool:
    norm_value = _norm(value)
    if not _starts_with_any(norm_value, LEVEL4_PREFIXES):
        return False
    return len(norm_value.split()) <= 8


def _is_valid_poi(value: str) -> bool:
    norm_value = _norm(value)
    if _is_generic_poi_only(value):
        return False
    if _is_note_like(value):
        return False
    if not _contains_phrase(norm_value, POI_SIGNAL_PHRASES):
        return False
    return 1 <= len(norm_value.split()) <= 18


COMMON_STREET_NAME_FIRST_TOKENS = {
    "ba", "bach", "bui", "cao", "chu", "cu", "dang", "dinh", "do",
    "dong", "duong", "hai", "ho", "hoa", "hoang", "hung", "huynh",
    "kim", "le", "ly", "mai", "ngo", "nguyen", "phan", "pham",
    "quang", "to", "ton", "tran", "trinh", "trung", "truong", "vo",
    "vu",
}


def _looks_like_clear_named_street(value: Any) -> bool:
    norm_value = _norm(value)
    words = norm_value.split()
    if norm_value in {"ng trai", "nguyen trai"}:
        return True
    if not 2 <= len(words) <= 5:
        return False
    if any(char.isdigit() for char in norm_value):
        return False
    if LEVEL4_SIGNAL_RE.search(norm_value):
        return False
    return words[0] in COMMON_STREET_NAME_FIRST_TOKENS


def _looks_like_local_street_code(value: Any) -> bool:
    norm_value = _norm(value)
    return bool(re.fullmatch(
        r"(?:[a-z]{1,4}\s*\d{1,3}[a-z]?|\d{1,3}[a-z][a-z0-9]*)",
        norm_value,
    ))


def _is_valid_street(value: str, has_house_number: bool) -> bool:
    norm_value = _norm(value)
    if not norm_value or _is_note_like(value):
        return False
    if _starts_with_any(norm_value, LEVEL4_PREFIXES):
        return False
    if norm_value.startswith(("day ", "o ", "lo ", "nhanh ")):
        return False
    if any(name in norm_value for name in KNOWN_RESIDENTIAL_POI_NAMES):
        return False
    if _is_alley_street_value(value):
        return False
    # 16B1 có thể giống số nhà, nhưng khi model đã tách được số nhà 20 thì
    # token kế tiếp là mã dãy/đường nội khu hợp lệ.
    if _is_valid_house_number(value) and not has_house_number:
        return False
    words = norm_value.split()
    if len(words) > 8:
        return False
    # Một chữ trần sau số nhà không đủ chứng minh là đường. Đây là nguồn của
    # các kết quả bị cắt cụt kiểu "Ninh", "Âu", "Chi", "Bến", "Tiệm".
    if len(words) == 1 and words[0].isalpha() and not _has_street_signal(value):
        return False
    if re.fullmatch(r"(?:ql|tl|hl|dt)\s*\d+[a-z]?", norm_value):
        return True
    if re.fullmatch(r"(?:ql|tl|hl|dt)\d+[a-z]?", _compact(value)):
        return True
    # Mã đường địa phương: ĐX 7, TCH 18, XTT 59, N5... (chữ + số)
    if _looks_like_local_street_code(value):
        return True
    if _has_street_signal(value):
        return True
    return has_house_number and len(words) >= 2 and any(w.isalpha() for w in words)


def _ensure_output_street_prefix(value: Any) -> str | None:
    street = _pretty_llm_street(value)
    if not street:
        return None
    if _ensure_street_prefix:
        return _ensure_street_prefix(street) or None
    norm_value = _norm(street)
    if norm_value == "ngo" or norm_value.startswith("ngo "):
        if _is_real_ngo_prefix(street):
            return street
    explicit_prefixes = {
        "cao toc", "dai lo", "duong", "hem", "huong lo", "kiet",
        "ngach", "pho", "quoc lo", "tinh lo",
    }
    if any(norm_value == prefix or norm_value.startswith(prefix + " ") for prefix in explicit_prefixes):
        return street
    if re.fullmatch(r"(?:ql|tl|hl|dt|ddt)\s*\d+[a-z]?", norm_value):
        return street
    if re.fullmatch(r"(?:ql|tl|hl|dt|ddt)\d+[a-z]?", _compact(street)):
        return street
    # Không tự thêm "Đường" khi raw không ghi loại đường
    return street


def _reject_component(
    result: dict[str, Any],
    flags: list[str],
    field: str,
    reason: str,
) -> None:
    result[field] = None
    flags.append(f"LLM_{field.upper()}_{reason}")


def validate_llm_result(result: dict[str, Any], original_row: dict[str, Any]) -> dict[str, Any]:
    result = dict(result or {})
    flags = list(result.get("flags") or [])
    raw_address = original_row.get("raw_address", "")
    rule_candidates = original_row.get("rule_candidates") or {}
    rule_flags = set(original_row.get("flags") or []) | set(rule_candidates.get("flags") or [])

    # Không cho LLM sửa 3 cột admin
    source_ward = _source_admin(original_row, "ward")
    source_district = _source_admin(original_row, "district")
    source_province = _source_admin(original_row, "province")

    if _clean_optional_text(result.get("ward")) != source_ward:
        flags.append("LLM_CHANGED_WARD_REVERTED")
    result["ward"] = source_ward

    if _clean_optional_text(result.get("district")) != source_district:
        flags.append("LLM_CHANGED_DISTRICT_REVERTED")
    result["district"] = source_district

    if _clean_optional_text(result.get("province")) != source_province:
        flags.append("LLM_CHANGED_PROVINCE_REVERTED")
    result["province"] = source_province

    for field in ("poi", "house_number", "street", "level4"):
        result[field] = _clean_optional_text(result.get(field))
    if result.get("poi") and _trim_poi_value_v2:
        poi_candidate = result["poi"]
        if _normalize_abbrev_v2:
            poi_candidate = _normalize_abbrev_v2(poi_candidate)
        trimmed_poi = _trim_poi_value_v2(poi_candidate, set(), set())
        if trimmed_poi and _norm(trimmed_poi) != _norm(result["poi"]):
            result["poi"] = trimmed_poi
            flags.append("LLM_POI_ADDRESS_TAIL_TRIMMED")
    if result.get("street"):
        result["street"] = _trim_row_admin_tail(result["street"], original_row)

    # LLM đôi khi bỏ dấu phân cách rồi nuốt trường kế tiếp vào street
    # ("Lê Lai . Hợp Thành" -> "Lê Lai Hợp Thành"). Dùng evidence nguyên văn
    # để khôi phục ranh giới trước khi kiểm định.
    street_evidence = _clean_optional_text(
        (result.get("evidence") or {}).get("street_span")
    )
    if result.get("street") and street_evidence and re.search(r"\s\.\s", street_evidence):
        evidence_head = re.split(r"\s+\.\s+", street_evidence, maxsplit=1)[0]
        evidence_head = re.sub(
            r"^\s*(?:số|so|sn|nhà|nha)?\s*"
            r"(?:\d+[A-Za-zĐđ0-9]*(?:\s*[/\-.]\s*[A-Za-zĐđ0-9]+)*)\s+",
            "",
            evidence_head,
            flags=re.I,
        ).strip(" ,.-–—/:;")
        if evidence_head and _norm(evidence_head) in _norm(result["street"]):
            result["street"] = evidence_head

    if _raw_has_old_new_keyword(raw_address):
        if "ADMIN_VERSION_AMBIGUOUS" not in flags:
            flags.append("ADMIN_VERSION_AMBIGUOUS")
        result["confidence"] = min(_to_float(result.get("confidence"), 0.6), 0.60)

    for field in ("poi", "house_number", "street", "level4"):
        value = result.get(field)
        if not value:
            continue
        if _is_admin_only(value, original_row):
            _reject_component(result, flags, field, "ADMIN_ONLY_REJECTED")
            continue
        if field in {"house_number", "street", "level4"} and _has_admin_tail(value):
            _reject_component(result, flags, field, "ADMIN_TAIL_REJECTED")
            continue
        if _is_note_like(value):
            result[field] = None
            flags.append("INVALID_POI_BLACKLIST" if field == "poi" else f"LLM_{field.upper()}_NOTE_REJECTED")
            continue
        if field == "poi" and _has_relative_landmark_context(value, raw_address):
            _reject_component(result, flags, field, "RELATIVE_LANDMARK_REJECTED")
            continue
        if not _evidence_supports_value(field, value, result, raw_address):
            _reject_component(result, flags, field, "EVIDENCE_REJECTED")
            continue
        if field == "street" and _is_alley_street_value(value):
            _reject_component(result, flags, field, "ALLEY_REJECTED")
            continue
        if field == "house_number" and not _is_valid_house_number(value):
            _reject_component(result, flags, field, "FORMAT_REJECTED")
        elif field == "poi" and not _is_valid_poi(value):
            _reject_component(result, flags, field, "FORMAT_REJECTED")
        elif field == "level4" and not _is_valid_level4(value):
            _reject_component(result, flags, field, "FORMAT_REJECTED")
        elif field == "street" and not _is_valid_street(value, bool(result.get("house_number"))):
            _reject_component(result, flags, field, "FORMAT_REJECTED")
        elif (
            field == "street"
            and "UNPREFIXED_STREET_FROM_HOUSE_NUMBER" in rule_flags
            and not _has_street_signal(value)
            and not _looks_like_clear_named_street(value)
            and not _looks_like_local_street_code(value)
        ):
            _reject_component(result, flags, field, "UNPREFIXED_AMBIGUOUS_REJECTED")

    if result.get("street"):
        result["street"] = _ensure_output_street_prefix(result["street"])
    if result.get("poi"):
        result["poi"] = _pretty_llm_component(result["poi"])
    if result.get("level4"):
        result["level4"] = _pretty_llm_level4(result["level4"], raw_address)

    detail_fields = ("poi", "street", "level4")
    if result.get("house_number") and not result.get("street"):
        flags.append("LLM_HOUSE_NUMBER_ONLY_NOT_OUTPUT")
    if not any(result.get(field) for field in detail_fields):
        flags.append("LLM_NO_VALID_DETAIL")
        result["confidence"] = min(_to_float(result.get("confidence"), 0.5), 0.50)
    elif any(flag.startswith("LLM_") and flag.endswith("_REJECTED") for flag in flags):
        result["confidence"] = min(_to_float(result.get("confidence"), 0.7), 0.70)
    else:
        flags.append("LLM_VALIDATED")
        result["confidence"] = _to_float(result.get("confidence"), 0.0)

    result["confidence"] = max(0.0, min(1.0, round(_to_float(result.get("confidence")), 3)))
    result["flags"] = sorted(set(flags))
    return result
