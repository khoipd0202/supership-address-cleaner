# VN Address Cleaner

Thư viện làm sạch file Excel địa chỉ giao vận dạng SuperShip.

## Output mặc định

Thư viện xuất workbook chỉ gồm 6 cột nghiệp vụ:

```text
POI | Tên đường | Cấp 4 | Phường/Xã | Quận/Huyện | Tỉnh/TP
```

Các dòng không bóc được `POI`, `Tên đường` hoặc `Cấp 4` sẽ bị loại theo mặc định,
giống luồng lọc dữ liệu sạch hiện tại.

## Cài local

```bash
cd /Users/dangkhoii/convert_adderss
pip install -e .
```

## Dùng trong Python

```python
from vn_address_cleaner import clean_excel

stats = clean_excel(
    "data/SuperShip - DANH SÁCH ĐƠN HÀNG 260416084446.xlsx",
    "outputs/cleaned_minimal.xlsx",
)

print(stats.as_dict())
```

Clean một dòng:

```python
from vn_address_cleaner import AddressCleaner

cleaner = AddressCleaner()
result = cleaner.clean(
    raw_address="Số 77-79 Đường Hoàng Quốc Việt, Nghĩa Đô, Cầu Giấy, Hn",
    ward="Phường Nghĩa Đô",
    district="Quận Cầu Giấy",
    province="Thành phố Hà Nội",
)

print(result.as_output_row())
```

## CLI

```bash
vn-address-clean input.xlsx -o output.xlsx
```

Hoặc không cần cài editable:

```bash
python3 -m vn_address_cleaner.cli input.xlsx -o output.xlsx
```

Giữ cả dòng không có chi tiết:

```bash
vn-address-clean input.xlsx -o output.xlsx --include-empty-rows
```
