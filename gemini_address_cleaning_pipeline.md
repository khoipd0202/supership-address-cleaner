# Tích hợp Gemini API vào pipeline làm sạch địa chỉ Excel

Tài liệu này mô tả cách tích hợp Gemini API theo hướng tiết kiệm chi phí: **rule parser xử lý trước, Gemini chỉ xử lý dòng lỗi hoặc mơ hồ**.

---

## 1. Mục tiêu

Không dùng LLM để parse toàn bộ dữ liệu. Mục tiêu đúng là:

```text
Excel
→ Rule Parser
→ Validator
→ Chỉ dòng lỗi/mơ hồ mới gọi Gemini
→ Validator lần 2
→ Xuất Excel sạch + flags + confidence
```

Kỳ vọng thực tế:

```text
80–95% dòng: rule-only, không gọi Gemini
5–20% dòng: gọi Gemini
1–3% dòng: review thủ công
```

---

## 2. Khi nào gọi Gemini?

Chỉ gọi Gemini nếu dòng có một trong các flag nặng:

```text
POSSIBLE_POI_MISSED
POSSIBLE_LEVEL4_MISSED
MULTIPLE_POI_FOUND
MULTIPLE_LEVEL4_FOUND
RAW_ADMIN_CONFLICT_WITH_COLUMNS
OLD_NEW_ADMIN_CONFLICT
ADMIN_VERSION_AMBIGUOUS
RAW_CONTAINS_NAY_THUOC
INVALID_POI_BLACKLIST
REMOVED_OVERLAPS_KEPT
LOW_CONFIDENCE
```

Không gọi Gemini nếu rule parser đã tự tin:

```text
confidence >= 0.90
và không có hard flag
```

---

## 3. Tạo Gemini API key

Vào Google AI Studio và tạo API key cho Gemini API.

Sau đó lưu key vào biến môi trường.

### macOS / Linux

```bash
export GEMINI_API_KEY="your_api_key_here"
```

### Windows PowerShell

```powershell
setx GEMINI_API_KEY "your_api_key_here"
```

---

## 4. Tạo project Python

```bash
mkdir address-cleaner-llm
cd address-cleaner-llm

python -m venv .venv
source .venv/bin/activate
```

Windows:

```powershell
python -m venv .venv
.venv\Scripts\activate
```

Cài thư viện:

```bash
pip install google-genai pandas openpyxl python-dotenv
```

---

## 5. Cấu trúc thư mục

```text
address-cleaner-llm/
├── .env
├── input/
│   └── orders.xlsx
├── output/
│   └── cleaned_orders.xlsx
├── address_rules.py
├── gemini_parser.py
├── validator.py
├── main.py
└── config.py
```

File `.env`:

```env
GEMINI_API_KEY=your_api_key_here
```

---

## 6. Model nên dùng

Bắt đầu bằng:

```python
GEMINI_MODEL = "gemini-2.5-flash"
```

Nếu muốn tiết kiệm hơn, có thể thử:

```python
GEMINI_MODEL = "gemini-2.5-flash-lite"
```

Khuyến nghị:

```text
Dòng thường: rule-only
Dòng nghi ngờ nhẹ: gemini-2.5-flash-lite
Dòng khó/conflict/cũ-mới: gemini-2.5-flash
```

---

## 7. Output JSON chuẩn

Gemini phải trả về JSON theo schema cố định, không trả text tự do.

Mỗi dòng output cần có:

```json
{
  "row_id": 1,
  "poi": null,
  "house_number": "63/2/1",
  "street": "Nguyễn Phúc Chu",
  "level4": null,
  "ward": "Phường Thành Nhất",
  "district": "Thành phố Buôn Ma Thuột",
  "province": "Tỉnh Đắk Lắk",
  "clean_address": "Nguyễn Phúc Chu, Phường Thành Nhất, Thành phố Buôn Ma Thuột, Tỉnh Đắk Lắk",
  "confidence": 0.88,
  "flags": [],
  "evidence": {
    "poi_span": null,
    "house_number_span": "63/2/1",
    "street_span": "Nguyễn Phúc Chu",
    "level4_span": null
  }
}
```

---

## 8. File `gemini_parser.py`

```python
import os
import json
from typing import Any, Dict, List

from google import genai
from google.genai import types


client = genai.Client(api_key=os.environ["GEMINI_API_KEY"])


ADDRESS_SCHEMA = {
    "type": "object",
    "properties": {
        "results": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "row_id": {"type": "integer"},
                    "poi": {"type": ["string", "null"]},
                    "house_number": {"type": ["string", "null"]},
                    "street": {"type": ["string", "null"]},
                    "level4": {"type": ["string", "null"]},
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
                            "poi_span": {"type": ["string", "null"]},
                            "house_number_span": {"type": ["string", "null"]},
                            "street_span": {"type": ["string", "null"]},
                            "level4_span": {"type": ["string", "null"]}
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
11. Trả về đúng JSON schema, không giải thích ngoài JSON.
"""


def build_prompt(rows: List[Dict[str, Any]]) -> str:
    payload = []
    for row in rows:
        payload.append({
            "row_id": int(row["row_id"]),
            "raw_address": row["raw_address"],
            "ward": row["ward"],
            "district": row["district"],
            "province": row["province"],
            "rule_candidates": row.get("rule_candidates", {}),
            "rule_flags": row.get("flags", []),
            "rule_confidence": row.get("confidence")
        })

    return f"""
{SYSTEM_RULES}

Input rows:
{json.dumps(payload, ensure_ascii=False, indent=2)}
"""


def parse_rows_with_gemini(
    rows: List[Dict[str, Any]],
    model: str = "gemini-2.5-flash"
) -> Dict[str, Any]:
    prompt = build_prompt(rows)

    response = client.models.generate_content(
        model=model,
        contents=prompt,
        config=types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=ADDRESS_SCHEMA,
            temperature=0.0,
        ),
    )

    return json.loads(response.text)
```

---

## 9. File `validator.py`

```python
from typing import Dict, Any, List


HARD_FLAGS = {
    "RAW_ADMIN_CONFLICT_WITH_COLUMNS",
    "OLD_NEW_ADMIN_CONFLICT",
    "ADMIN_VERSION_AMBIGUOUS",
    "RAW_CONTAINS_NAY_THUOC",
    "MULTIPLE_POI_FOUND",
    "MULTIPLE_LEVEL4_FOUND",
    "INVALID_POI_BLACKLIST",
    "REMOVED_OVERLAPS_KEPT",
    "LOW_CONFIDENCE",
}


OLD_NEW_KEYWORDS = [
    "cũ", "mới", "nay thuộc", "nay là",
    "sáp nhập", "trước khi sáp nhập", "sau khi sáp nhập"
]


def has_hard_flags(flags: List[str]) -> bool:
    return bool(set(flags or []) & HARD_FLAGS)


def should_send_to_llm(rule_result: Dict[str, Any]) -> bool:
    confidence = rule_result.get("confidence", 0)
    flags = rule_result.get("flags", [])

    if has_hard_flags(flags):
        return True

    if confidence < 0.90:
        return True

    return False


def validate_llm_result(result: Dict[str, Any], original_row: Dict[str, Any]) -> Dict[str, Any]:
    flags = list(result.get("flags", []))

    # Không cho LLM sửa 3 cột admin
    if result["ward"] != original_row["ward"]:
        result["ward"] = original_row["ward"]
        flags.append("LLM_CHANGED_WARD_REVERTED")

    if result["district"] != original_row["district"]:
        result["district"] = original_row["district"]
        flags.append("LLM_CHANGED_DISTRICT_REVERTED")

    if result["province"] != original_row["province"]:
        result["province"] = original_row["province"]
        flags.append("LLM_CHANGED_PROVINCE_REVERTED")

    raw_lower = str(original_row["raw_address"]).lower()

    if any(k in raw_lower for k in OLD_NEW_KEYWORDS):
        if "ADMIN_VERSION_AMBIGUOUS" not in flags:
            flags.append("ADMIN_VERSION_AMBIGUOUS")
        result["confidence"] = min(float(result.get("confidence", 0.6)), 0.60)

    # Nếu POI trùng câu giao tiếp thì loại
    poi = result.get("poi")
    if poi:
        poi_lower = poi.lower()
        blacklist = [
            "em gửi về", "gửi về", "gửi cho chị",
            "địa chỉ", "địa chỉ người nhận",
            "ship về", "đổi địa chỉ về", "người nhận"
        ]
        if any(x in poi_lower for x in blacklist):
            result["poi"] = None
            flags.append("INVALID_POI_BLACKLIST")

    result["flags"] = sorted(set(flags))
    return result
```

---

## 10. File `address_rules.py`

Đây là rule parser local bản khởi đầu. Sau này nên thay bằng span-based parser tốt hơn.

```python
import re
from typing import Dict, Any


POI_KEYWORDS = [
    "trạm y tế", "bệnh viện", "bv", "phòng khám",
    "trường", "mầm non", "tiểu học", "thcs", "thpt",
    "ubnd", "ủy ban", "công an",
    "ngân hàng", "bank", "vpbank", "vietcombank",
    "công ty", "cty", "học viện",
    "cửa hàng", "quán", "nhà thuốc",
    "tòa nhà", "toà nhà", "chung cư"
]

LEVEL4_KEYWORDS = [
    "thôn", "xóm", "ấp", "khu phố", "kp",
    "tổ dân phố", "tdp", "tổ", "khối",
    "bản", "buôn", "làng", "sóc", "khóm",
    "đội", "tiểu khu", "cụm dân cư", "khu dân cư"
]

HOUSE_PATTERN = re.compile(
    r"(?<!\w)(?:[A-ZĐ]?\d+[A-ZĐ]?|\d+[A-ZĐ]+|[A-ZĐ]+\d+)"
    r"(?:\s*[\/\-]\s*(?:[A-ZĐ]?\d+[A-ZĐ]?|\d+[A-ZĐ]+|[A-ZĐ]+\d+))*"
    r"(?!\w)",
    re.IGNORECASE
)

PHONE_PATTERN = re.compile(r"(\+84|0)\d{8,10}")


def normalize_text(text: str) -> str:
    if text is None:
        return ""
    text = str(text)
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    text = text.strip()
    return text


def simple_rule_parse(row: Dict[str, Any]) -> Dict[str, Any]:
    raw = normalize_text(row["raw_address"])
    raw_lower = raw.lower()
    flags = []
    confidence = 0.95

    if PHONE_PATTERN.search(raw):
        flags.append("RAW_CONTAINS_PHONE")
        confidence -= 0.05

    if any(k in raw_lower for k in ["cũ", "mới", "nay thuộc", "sáp nhập"]):
        flags.append("ADMIN_VERSION_AMBIGUOUS")
        confidence = min(confidence, 0.60)

    house_number = None
    street = None

    # Bắt số nhà đầu segment + đường phía sau
    parts = re.split(r"[,;.\-]+", raw)
    for part in parts:
        part = part.strip()
        m = HOUSE_PATTERN.search(part)
        if m and m.start() <= 5:
            house_number = re.sub(r"\s+", "", m.group(0))
            after = part[m.end():].strip()

            # Cắt đường trước admin keyword
            after = re.split(
                r"\b(phường|xã|thị trấn|quận|huyện|tỉnh|tp|thành phố|p\.|q\.|tt)\b",
                after,
                flags=re.IGNORECASE
            )[0].strip()

            if after and len(after.split()) <= 6:
                street = after
            break

    poi = None
    for kw in POI_KEYWORDS:
        if kw in raw_lower:
            poi = extract_phrase_from_keyword(raw, kw)
            break

    level4 = None
    for kw in LEVEL4_KEYWORDS:
        if kw in raw_lower:
            level4 = extract_phrase_from_keyword(raw, kw)
            break

    if poi is None:
        flags.append("NO_POI_FOUND")

    if level4 is None:
        flags.append("NO_LEVEL4_FOUND")

    # Nếu không có POI/Cấp 4 nhưng có số nhà/đường thì vẫn hợp lệ
    if poi is None and level4 is None and (house_number or street):
        flags.append("ONLY_STREET_LEVEL_FOUND")
        confidence = min(confidence, 0.90)

    if confidence < 0.90:
        flags.append("LOW_CONFIDENCE")

    clean_parts = []
    if poi:
        clean_parts.append(poi)
    if level4:
        clean_parts.append(level4)

    clean_parts.extend([
        row["ward"],
        row["district"],
        row["province"],
    ])

    return {
        "poi": poi,
        "house_number": house_number,
        "street": street,
        "level4": level4,
        "ward": row["ward"],
        "district": row["district"],
        "province": row["province"],
        "clean_address": ", ".join([p for p in clean_parts if p]),
        "confidence": round(max(confidence, 0.0), 2),
        "flags": sorted(set(flags)),
    }


def extract_phrase_from_keyword(raw: str, keyword: str) -> str:
    """
    Bản đơn giản: lấy từ keyword đến trước dấu phân tách/admin.
    Sau này nên thay bằng span-based parser.
    """
    pattern = re.compile(re.escape(keyword), re.IGNORECASE)
    m = pattern.search(raw)
    if not m:
        return None

    text = raw[m.start():]
    text = re.split(
        r"[,;.\-]|\b(phường|xã|thị trấn|quận|huyện|tỉnh|tp|thành phố)\b",
        text,
        flags=re.IGNORECASE
    )[0]

    return text.strip()
```

---

## 11. File `main.py`

Giả sử Excel có cột:

```text
Địa chỉ
Phường/Xã
Quận/Huyện
Tỉnh/Thành Phố
```

```python
import os
import pandas as pd
from dotenv import load_dotenv

from address_rules import simple_rule_parse
from validator import should_send_to_llm, validate_llm_result
from gemini_parser import parse_rows_with_gemini


load_dotenv()


INPUT_FILE = "input/orders.xlsx"
OUTPUT_FILE = "output/cleaned_orders.xlsx"
SHEET_NAME = 0
BATCH_SIZE = 20
GEMINI_MODEL = "gemini-2.5-flash"


def make_row_dict(index, row):
    return {
        "row_id": int(index),
        "raw_address": str(row.get("Địa chỉ", "") or ""),
        "ward": str(row.get("Phường/Xã", "") or ""),
        "district": str(row.get("Quận/Huyện", "") or ""),
        "province": str(row.get("Tỉnh/Thành Phố", "") or ""),
    }


def chunks(items, size):
    for i in range(0, len(items), size):
        yield items[i:i + size]


def main():
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

    final_results = {}
    llm_candidates = []

    # Bước 1: rule parser xử lý tất cả dòng
    for idx, row in df.iterrows():
        base_row = make_row_dict(idx, row)
        rule_result = simple_rule_parse(base_row)

        if should_send_to_llm(rule_result):
            llm_candidates.append({
                **base_row,
                "rule_candidates": rule_result,
                "flags": rule_result["flags"],
                "confidence": rule_result["confidence"],
            })
        else:
            final_results[idx] = rule_result

    print(f"Tổng dòng: {len(df)}")
    print(f"Dòng gọi Gemini: {len(llm_candidates)}")
    print(f"Dòng rule-only: {len(df) - len(llm_candidates)}")

    # Bước 2: dedupe trước khi gọi Gemini
    dedup_map = {}
    unique_candidates = []

    for item in llm_candidates:
        key = (
            item["raw_address"].strip().lower(),
            item["ward"].strip().lower(),
            item["district"].strip().lower(),
            item["province"].strip().lower(),
        )

        if key not in dedup_map:
            dedup_map[key] = []
            unique_candidates.append(item)

        dedup_map[key].append(item["row_id"])

    print(f"Dòng unique gọi Gemini: {len(unique_candidates)}")

    # Bước 3: gọi Gemini theo batch nhỏ
    for batch in chunks(unique_candidates, BATCH_SIZE):
        response = parse_rows_with_gemini(batch, model=GEMINI_MODEL)

        for result in response["results"]:
            source_item = next(x for x in batch if x["row_id"] == result["row_id"])
            validated = validate_llm_result(result, source_item)

            key = (
                source_item["raw_address"].strip().lower(),
                source_item["ward"].strip().lower(),
                source_item["district"].strip().lower(),
                source_item["province"].strip().lower(),
            )

            for original_row_id in dedup_map[key]:
                final_results[original_row_id] = validated

    # Bước 4: ghi kết quả ra dataframe
    df["POI"] = None
    df["Số nhà"] = None
    df["Đường/Phố"] = None
    df["Cấp 4"] = None
    df["ĐỊA CHỈ SẠCH"] = None
    df["confidence"] = None
    df["flags"] = None

    for idx, result in final_results.items():
        df.at[idx, "POI"] = result.get("poi")
        df.at[idx, "Số nhà"] = result.get("house_number")
        df.at[idx, "Đường/Phố"] = result.get("street")
        df.at[idx, "Cấp 4"] = result.get("level4")
        df.at[idx, "ĐỊA CHỈ SẠCH"] = result.get("clean_address")
        df.at[idx, "confidence"] = result.get("confidence")
        df.at[idx, "flags"] = "; ".join(result.get("flags", []))

    os.makedirs("output", exist_ok=True)
    df.to_excel(OUTPUT_FILE, index=False)

    print(f"Đã xuất: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
```

Chạy:

```bash
python main.py
```

---

## 12. Test nhanh với các case lỗi

File `test_gemini.py`:

```python
from gemini_parser import parse_rows_with_gemini

rows = [
    {
        "row_id": 1,
        "raw_address": "Quỳnh Nhi .thôn Tân Lập .cumta .mdrak .dak Lak",
        "ward": "Xã Cư M'ta",
        "district": "Huyện M'Đrắk",
        "province": "Tỉnh Đắk Lắk",
        "rule_candidates": {},
        "flags": ["POSSIBLE_LEVEL4_MISSED"],
        "confidence": 0.6,
    },
    {
        "row_id": 2,
        "raw_address": "Trạm Y Tế Tam Thanh, Huyện Vụ Bản, T. Nam Định",
        "ward": "Xã Tam Thanh",
        "district": "Huyện Vụ Bản",
        "province": "Tỉnh Nam Định",
        "rule_candidates": {},
        "flags": ["POSSIBLE_LEVEL4_FALSE_POSITIVE"],
        "confidence": 0.6,
    },
    {
        "row_id": 3,
        "raw_address": "Số 40 Phố Trần Cung, Phường Cổ Nhuế 1 Quận Bắc Từ( Nay Thuộc Phường Nghĩa Đô, Hn ) Liêm, Hà Nội",
        "ward": "Phường Cổ Nhuế 1",
        "district": "Quận Bắc Từ Liêm",
        "province": "Thành phố Hà Nội",
        "rule_candidates": {},
        "flags": ["ADMIN_VERSION_AMBIGUOUS"],
        "confidence": 0.5,
    },
]

print(parse_rows_with_gemini(rows))
```

Kỳ vọng:

```text
Row 1:
POI = null
Cấp 4 = Thôn Tân Lập

Row 2:
POI = Trạm y tế Tam Thanh
Cấp 4 = null

Row 3:
Số nhà = 40
Đường/Phố = Phố Trần Cung
Flags có ADMIN_VERSION_AMBIGUOUS / RAW_CONTAINS_NAY_THUOC
Confidence <= 0.60
```

---

## 13. Giảm chi phí khi dữ liệu lớn

### 13.1. Chỉ gọi Gemini cho dòng có flag

```text
Rule confidence >= 0.90, không hard flag
→ không gọi Gemini

Rule confidence < 0.90 hoặc có conflict
→ gọi Gemini
```

### 13.2. Dedupe trước khi gọi Gemini

Dedupe theo:

```text
normalized_raw_address + ward + district + province
```

### 13.3. Gọi nhiều dòng trong một request

Bắt đầu với:

```python
BATCH_SIZE = 20
```

Sau khi ổn có thể tăng lên:

```python
BATCH_SIZE = 30 hoặc 50
```

### 13.4. Mask dữ liệu cá nhân trước khi gửi LLM

Trước khi gọi Gemini nên xóa hoặc mask:

```text
số điện thoại
tên người nhận nếu tách được
ghi chú giao hàng
COD/giá tiền
```

Ví dụ:

```text
Phạm Minh 0786297988, 41/1 Nguyễn Tri Phương...
```

Gửi lên LLM thành:

```text
<NAME> <PHONE>, 41/1 Nguyễn Tri Phương...
```

Hoặc tốt hơn:

```text
41/1 Nguyễn Tri Phương...
```

---

## 14. Quy trình vận hành thực tế

Lần đầu:

```text
1. Chạy rule parser trên 1 file mẫu.
2. Xem flags.
3. Gửi 100–300 dòng lỗi sang Gemini.
4. So sánh output.
5. Biến lỗi lặp lại thành rule mới.
6. Chạy lại.
```

Sau vài vòng:

```text
1. Rule parser xử lý phần lớn dòng.
2. Gemini chỉ xử lý case mơ hồ.
3. Review thủ công chỉ còn dòng conflict/cũ-mới.
```

---

## 15. Checklist bắt đầu

```text
1. Tạo Gemini API key.
2. Cài google-genai, pandas, openpyxl, python-dotenv.
3. Tạo các file:
   - address_rules.py
   - validator.py
   - gemini_parser.py
   - main.py
4. Chạy test_gemini.py với các case lỗi.
5. Chạy main.py với file Excel mẫu.
6. Kiểm tra cột flags/confidence.
7. Điều chỉnh rule để giảm số dòng gọi Gemini.
8. Khi dữ liệu lớn, dùng dedupe + batch request.
```

---

## 16. Nguyên tắc quan trọng

```text
Không tin LLM 100%.
LLM chỉ là tầng xử lý ngoại lệ.
3 cột Xã/Phường, Huyện/Quận, Tỉnh/TP luôn là source of truth.
Validator phải chạy sau LLM.
Dòng có conflict phải có flag, không được âm thầm cho confidence cao.
```

