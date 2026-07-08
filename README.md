# supership address cleaner

python library and cli for cleaning supership-style vietnamese delivery address excel files.

## what it does

the cleaner reads an input excel file, parses the raw address fields, and exports a smaller workbook with only these columns:

```text
POI | Tên đường | Cấp 4 | Phường/Xã | Quận/Huyện | Tỉnh/TP
```

by default, rows without any detail in `POI`, `Tên đường`, or `Cấp 4` are removed. use `--include-empty-rows` if you want to keep them.

## installation

install locally from this repository:

```bash
cd /Users/dangkhoii/convert_adderss
python3 -m pip install -e .
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

## privacy

do not commit real order files, customer data, generated outputs, cache files, or local sqlite databases.
