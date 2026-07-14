# supership address cleaner

python library and cli for cleaning supership-style vietnamese delivery address excel files.

## what it does

the cleaner reads an input excel file, parses the raw address fields, and exports a smaller workbook with only these columns:

```text
POI | Tên đường | Cấp 4 | Phường/Xã | Quận/Huyện | Tỉnh/TP
```

by default, excel output splits components into separate rows. this means a row with one poi, one street, and one level 4 becomes three output rows:

```text
POI             | Tên đường | Cấp 4 | Phường/Xã | Quận/Huyện | Tỉnh/TP
Công ty ABC     |           |       | ...
                | Đường A   |       | ...
                |           | Thôn 1| ...
```

rows without any detail in `POI`, `Tên đường`, or `Cấp 4` are removed. use `--include-empty-rows` if you want to keep them.

## installation

install locally from this repository:

```bash
cd /Users/dangkhoii/convert_adderss
python3 -m pip install -e .
```

with Cerebras support for ambiguous rows:

```bash
python3 -m pip install -e ".[llm]"
```

or, after cloning the private repository:

```bash
git clone git@github.com:khoipd0202/supership-address-cleaner.git
cd supership-address-cleaner
python3 -m pip install -e .
```

## input format

the input workbook should follow the same structure as the supership files in the `data` folder.

required columns:

```text
Địa chỉ
Phường/Xã
Quận/Huyện
Tỉnh/Thành Phố
```

similar column names such as `Địa chỉ chi tiết`, `Tỉnh/TP`, `Phường`, or `Huyện` are also detected.

## cli usage

clean one excel file:

```bash
vn-address-clean input.xlsx -o output.xlsx
```

or run without installing the console command:

```bash
python3 -m vn_address_cleaner.cli input.xlsx -o output.xlsx
```

select a worksheet:

```bash
vn-address-clean input.xlsx -o output.xlsx --sheet-name "Sheet1"
```

keep rows that have no poi, street, or level 4:

```bash
vn-address-clean input.xlsx -o output.xlsx --include-empty-rows
```

keep poi, street, and level 4 together in one row, matching the older output format:

```bash
vn-address-clean input.xlsx -o output.xlsx --combined-row
```

enable Cerebras for strict ambiguous-row parsing:

```bash
export CEREBRAS_API_KEY="your_api_key_here"
export CEREBRAS_BATCH_SIZE=5
export CEREBRAS_MAX_ROWS_PER_RUN=30
vn-address-clean input.xlsx -o output.xlsx --cerebras
```

Cerebras is only called for semantic ambiguity that rules cannot decide, such as missed POI/level-4 keywords, multiple POI/level-4 candidates, or unclear POI/street/admin boundaries. Low confidence, missing POI, missing level 4, old/new admin wording, and raw admin conflicts are handled by rule flags or manual-review flags without calling Cerebras by default. The LLM result is accepted only when each returned component has an evidence span in the original raw address. `Phường/Xã`, `Quận/Huyện`, and `Tỉnh/TP` always come from the source columns, never from the LLM. Free-tier defaults are intentionally conservative: 5 rows per request and 30 unique LLM rows per run.

## python usage

clean a full excel file:

```python
from vn_address_cleaner import clean_excel

stats = clean_excel(
    input_path="data/orders.xlsx",
    output_path="outputs/cleaned_orders.xlsx",
)

print(stats.as_dict())
```

by default, `clean_excel(...)` writes one component per output row. to keep all extracted components in the same row:

```python
stats = clean_excel(
    input_path="data/orders.xlsx",
    output_path="outputs/cleaned_orders.xlsx",
    split_components=False,
)
```

use Cerebras from Python:

```python
stats = clean_excel(
    input_path="data/orders.xlsx",
    output_path="outputs/cleaned_orders.xlsx",
    use_cerebras=True,
)
```

clean one address:

```python
from vn_address_cleaner import AddressCleaner

cleaner = AddressCleaner()

result = cleaner.clean(
    raw_address="Trường Cao Đẳng Cơ Giới Ninh Bình, Đường Vũ Duy Thanh, Tổ 2, Phường Yên Bình, Thành phố Tam Điệp, Ninh Bình",
    ward="Yên Bình",
    district="Tam Điệp",
    province="Ninh Bình",
)

print(result.as_output_row())
```

expected output:

```python
[
    "Trường Cao Đẳng Cơ Giới Ninh Bình",
    "Đường Vũ Duy Thanh",
    "Tổ 2",
    "Phường Yên Bình",
    "Thành phố Tam Điệp",
    "Tỉnh Ninh Bình",
]
```

for one-address integrations, you can also split that object into database-ready component rows:

```python
print(result.as_component_rows())
```

an address can validly have empty fields. for example, this address has a street but no poi and no level 4:

```python
result = cleaner.clean(
    raw_address="Số 77-79 Đường Hoàng Quốc Việt, Nghĩa Đô, Cầu Giấy, Hà Nội",
    ward="Nghĩa Đô",
    district="Cầu Giấy",
    province="Hà Nội",
)

print(result.as_output_row())
```

expected output:

```python
[
    "",
    "Đường Hoàng Quốc Việt",
    "",
    "Phường Nghĩa Đô",
    "Quận Cầu Giấy",
    "Thành phố Hà Nội",
]
```

## api

main imports:

```python
from vn_address_cleaner import AddressCleaner, clean_excel, clean_workbook_bytes
```

`AddressCleaner.clean(...)` returns a `CleanResult` object with:

```text
poi
street
level4
ward
district
province
confidence
flags
```

`clean_excel(...)` returns a `CleanStats` object with:

```text
input_n
output_n
removed
full_admin
mapped_new
```

## tests

run the parser and library tests:

```bash
python3 -m unittest discover -s scratch -p 'test*.py'
```

## local ui

run the local ui server:

```bash
python3 address_ui.py
```

then open:

```text
http://127.0.0.1:8899
```

when multiple workbooks are uploaded together, the ui downloads one consolidated
`cleaned_address_files.xlsx` workbook. each output sheet includes a `Tên file nguồn`
column so rows from different uploads remain distinguishable.

downloaded address sheets are grouped automatically by `Tỉnh/TP → Quận/Huyện →
Phường/Xã`; rows missing administrative fields are placed at the end for review.

## privacy

do not commit real order files, customer data, generated outputs, cache files, or local sqlite databases.
