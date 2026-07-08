# Tiêu chí dữ liệu sạch cho bài toán chuẩn hóa địa chỉ Việt Nam

Tài liệu này mô tả tiêu chí dữ liệu sạch và rule xử lý cho thư viện làm sạch địa chỉ từ file Excel giao vận.

Mục tiêu chính:

- Dùng các cột `Xã/Phường`, `Huyện/Quận`, `Tỉnh/Thành Phố` trong file làm **source of truth**.
- Từ cột `Địa chỉ chi tiết` chỉ bóc thêm các thành phần chi tiết:
  - `POI`
  - `Số nhà`
  - `Đường/Phố`
  - `Cấp 4`
- Không để raw address ghi đè dữ liệu hành chính đã có trong 3 cột riêng.
- Gắn `flags` và `confidence` để phục vụ review batch lớn.

---

## 1. Định nghĩa dữ liệu sạch

Một dòng địa chỉ được xem là sạch khi parser tách đúng tối thiểu các thành phần sau:

```text
POI | Số nhà | Đường/Phố | Cấp 4 | Xã/Phường | Huyện/Quận | Tỉnh/TP | Flags
```

Output chính cho nghiệp vụ có thể chỉ dùng:

```text
POI | Cấp 4 | Xã/Phường | Huyện/Quận | Tỉnh/TP
```

Tuy nhiên, khi phát triển thư viện và chạy số lượng lớn, bắt buộc nên có thêm các cột debug:

```text
Số nhà | Đường/Phố | removed_parts | confidence | flags
```

Lý do: `Số nhà` là tín hiệu quan trọng để nhận diện `Đường/Phố`. Không nên xóa số nhà ngay từ đầu.

---

## 2. Schema output khuyến nghị

```text
raw_address
POI
Số nhà
Đường/Phố
Cấp 4
Xã/Phường
Huyện/Quận
Tỉnh/TP
ĐỊA CHỈ SẠCH
removed_parts
confidence
flags
```

Nếu chỉ xuất bản cuối cho nghiệp vụ:

```text
POI
Cấp 4
Xã/Phường
Huyện/Quận
Tỉnh/TP
ĐỊA CHỈ SẠCH
confidence
flags
```

Nhưng trong giai đoạn phát triển thư viện, không được bỏ các cột debug:

```text
Số nhà
Đường/Phố
removed_parts
flags
confidence
```

---

## 3. Nguyên tắc source of truth cho hành chính

Ba cột sau trong file Excel được xem là dữ liệu đúng:

```text
Xã/Phường
Huyện/Quận
Tỉnh/Thành Phố
```

Rule:

```text
1. Không cố lấy lại Xã/Huyện/Tỉnh từ địa chỉ chi tiết.
2. Không để raw address ghi đè 3 cột này.
3. Raw address chỉ dùng để:
   - bóc POI
   - bóc số nhà
   - bóc đường/phố
   - bóc cấp 4
   - đối chiếu alias hành chính
   - phát hiện conflict cũ/mới/sáp nhập
```

Ví dụ:

```text
Raw: Đ,c 63/2/1 Nguyễn Phúc Chu, P. Thành Nhất, Đăk Lắk
Cột riêng:
- Xã/Phường: Phường Thành Nhất
- Huyện/Quận: Thành phố Buôn Ma Thuột
- Tỉnh/TP: Tỉnh Đắk Lắk
```

Output hành chính phải lấy từ cột riêng:

```text
Phường Thành Nhất, Thành phố Buôn Ma Thuột, Tỉnh Đắk Lắk
```

Raw chỉ dùng để bóc:

```text
Số nhà: 63/2/1
Đường/Phố: Nguyễn Phúc Chu
```

---

## 4. Tiêu chí sạch cho từng cột

### 4.1. Xã/Phường, Huyện/Quận, Tỉnh/TP

Ba cột này phải được giữ nguyên theo dữ liệu đã chuẩn hóa ở file nguồn.

Parser cần sinh alias từ 3 cột này để nhận diện trong raw address.

Ví dụ alias:

```text
Phường 9:
- phường 9
- p9
- p.9
- p 9

Phường Thành Nhất:
- phường thành nhất
- p. thành nhất
- p thành nhất

Thị trấn Quảng Phú:
- thị trấn quảng phú
- tt quảng phú
- tt quảng

Thành phố Vũng Tàu:
- thành phố vũng tàu
- tp vũng tàu
- tp-vt
- tp vt
- vt

Tỉnh Nam Định:
- tỉnh nam định
- t. nam định
- t nam định
- nam định

Tỉnh Đắk Lắk:
- tỉnh đắk lắk
- đắk lắk
- đăk lăk
- dak lak
- daklak

Xã Cư M'ta:
- xã cư m'ta
- cư m'ta
- cumta
- cu mta
- cư mta

Huyện M'Đrắk:
- huyện m'đrắk
- m'đrắk
- mdrak
- m drak
- m'drak
```

Các alias hành chính chỉ được gắn nhãn `ADMIN`, không được kéo vào `POI` hoặc `Cấp 4`.

---

### 4.2. Số nhà

Một dòng sạch phải nhận diện được số nhà/mã nhà, kể cả dạng phức tạp.

Phải bắt được:

```text
40
122
1F
259/C3
259/ C3
63/2/1
18C/1N
12A/5B
25-27
A12
C3
```

Ví dụ:

```text
Số 40 Phố Trần Cung
```

Tách:

```text
Số nhà: 40
Đường/Phố: Phố Trần Cung
```

Ví dụ:

```text
259/ C3 Phan Bội Châu
```

Tách:

```text
Số nhà: 259/C3
Đường/Phố: Phan Bội Châu
```

Lưu ý quan trọng: số sau keyword cấp 4 không phải số nhà.

Không được đưa các cụm này vào `Số nhà`:

```text
Tổ 1
Khu phố 5
Thôn 3
Xóm 7
Ấp 4
TDP 6
```

Các cụm này là `Cấp 4`.

---

### 4.3. Đường/Phố

Dữ liệu sạch phải bóc được tên đường, kể cả khi raw không có keyword `đường/phố`.

#### Trường hợp có keyword

```text
Số 40 Phố Trần Cung
Đường Nguyễn Văn Cừ
QL 1A
Tỉnh lộ 8
Đại lộ Bình Dương
```

Output:

```text
Đường/Phố: Phố Trần Cung
Đường/Phố: Đường Nguyễn Văn Cừ
Đường/Phố: QL 1A
Đường/Phố: Tỉnh lộ 8
```

#### Trường hợp không có keyword nhưng đứng sau số nhà

```text
63/2/1 Nguyễn Phúc Chu
122 Hùng Vương
1F Lương Thế Vinh
259/C3 Phan Bội Châu
```

Output:

```text
Số nhà: 63/2/1
Đường/Phố: Nguyễn Phúc Chu

Số nhà: 122
Đường/Phố: Hùng Vương

Số nhà: 1F
Đường/Phố: Lương Thế Vinh

Số nhà: 259/C3
Đường/Phố: Phan Bội Châu
```

Tiêu chí quan trọng:

```text
Không bỏ số nhà trước khi dùng nó để nhận diện đường/phố.
```

---

### 4.4. POI

POI là địa điểm, tổ chức, cơ sở, tòa nhà hoặc điểm định vị rõ nghĩa.

Ví dụ POI hợp lệ:

```text
Trạm y tế Tam Thanh
Trường Mầm non Thanh Sơn
UBND Phường Tân Dân
Ngân hàng Bản Việt
Công ty Honda Việt Nam
Học viện Kỹ thuật Quân sự
Tòa nhà VG Building
Nhà thuốc Cường Hải
Cửa hàng Inox Hoàng Vũ
Quán Amity Chợ Đầm
Trung tâm Chăm Sóc Mẹ Và Bé Khánh Trần
```

POI không được nuốt sang:

```text
số nhà
đường/phố
ngõ/hẻm/ngách/kiệt
cấp 4
xã/phường
huyện/quận
tỉnh/thành phố
```

Ví dụ:

```text
Trung Tâm Chăm Sóc Mẹ Và Bé Khánh Trần K 2 Pha Đường Nguyễn Văn Cừ, Phường Lê Hồng Phong...
```

Không nên ra:

```text
POI: Trung tâm Chăm Sóc Mẹ Và Bé Khánh Trần K 2 Pha
```

Nên ra:

```text
POI: Trung tâm Chăm Sóc Mẹ Và Bé Khánh Trần
Đường/Phố: Đường Nguyễn Văn Cừ
Flags: UNKNOWN_TOKEN_BEFORE_STREET
```

Rule:

```text
POI phải dừng trước:
- house_like_token
- street keyword
- admin alias
- level4 keyword
- note keyword
```

---

### 4.5. Cấp 4

Cấp 4 là đơn vị nhỏ hơn xã/phường.

Các loại cấp 4 hợp lệ:

```text
Thôn
Xóm
Ấp
Khu phố
KP
Tổ
Tổ dân phố
TDP
Khối
Bản
Buôn
Làng
Sóc
Khóm
Đội
Tiểu khu
Cụm dân cư
Khu dân cư
```

Ví dụ đúng:

```text
Thôn Tân Lập
Khu phố 5
Xóm Rồng
Ấp Vĩnh Thành
Tổ 1
Khu phố Đồng An 2
Thôn Cự Khánh
Bản Mới
```

Cấp 4 phải dừng trước admin alias.

Ví dụ:

```text
Raw: Quỳnh Nhi .thôn Tân Lập .cumta .mdrak .dak Lak
Cột riêng:
- Xã/Phường: Xã Cư M'ta
- Huyện/Quận: Huyện M'Đrắk
- Tỉnh/TP: Tỉnh Đắk Lắk
```

Output đúng:

```text
Cấp 4: Thôn Tân Lập
```

Không được ra:

```text
Cấp 4: Thôn Tân Lập .cumta .mdrak .dak
```

Vì:

```text
cumta = alias của Xã Cư M'ta
mdrak = alias của Huyện M'Đrắk
dak Lak = alias của Tỉnh Đắk Lắk
```

---

## 5. Rule chống bắt nhầm

### 5.1. Không bắt `Bản` sai

Ví dụ:

```text
Trạm Y Tế Tam Thanh, Huyện Vụ Bản, T. Nam Định
```

Output đúng:

```text
POI: Trạm y tế Tam Thanh
Cấp 4: null
```

Không được bắt:

```text
Cấp 4: Bản T. Nam Định
```

Vì `Bản` nằm trong `Huyện Vụ Bản`, tức là admin alias.

Rule:

```text
Không nhận "Bản" là Cấp 4 nếu:
- nằm trong admin alias
- đứng sau Huyện/Quận/Tỉnh/Thành phố
- nằm trong cụm đã gắn label ADMIN
- nằm trong brand/POI, ví dụ Ngân hàng Bản Việt
```

---

### 5.2. Không bắt `Sở` sai trong `Trụ sở`

Ví dụ:

```text
Trụ Sở Đảng Ủy, UBMTTQ...
```

Không được tách:

```text
POI: Sở Đảng Ủy
```

Vì `sở` ở đây thuộc cụm `trụ sở`, không phải `Sở Tài chính`, `Sở Giáo dục`, `Sở Y tế`.

Rule:

```text
Nếu "sở" đứng sau "trụ" thì không nhận là POI loại Sở.
```

---

### 5.3. Không bắt câu giao tiếp làm POI

Không được lấy các cụm này làm POI:

```text
Em gửi về
Gửi về
Gửi cho chị
Địa chỉ
Địa chỉ người nhận
Ship về
Đổi địa chỉ về
Người nhận
```

Ví dụ sai:

```text
POI: Em Gởi Về
```

Phải sửa thành:

```text
POI: null
Flags: INVALID_POI_BLACKLIST
```

---

## 6. Cấp 4 không có keyword

Một số địa chỉ có cấp 4 nhưng không ghi rõ `thôn/xóm/ấp/bản`.

Ví dụ:

```text
Dc: Quang Yên- Tam Đình - Tương Dương - Nghệ An
```

Với cột riêng:

```text
Xã Tam Đình
Huyện Tương Dương
Tỉnh Nghệ An
```

Output nên là:

```text
Cấp 4: Quang Yên
Flags: INFERRED_LEVEL4_NO_KEYWORD
Confidence: <= 0.75
```

Điều kiện suy luận:

```text
Segment đứng ngay trước xã/phường
Không phải số nhà
Không phải đường/phố
Không phải POI
Không phải tên người rõ ràng
Không phải huyện/tỉnh
Có dạng địa danh 1–4 từ
```

Không cho confidence quá cao vì đây là cấp 4 suy luận, không có keyword.

---

## 7. Địa chỉ cũ/mới/sáp nhập

Nếu raw có các từ:

```text
cũ
mới
nay thuộc
nay là
sáp nhập
trước khi sáp nhập
sau khi sáp nhập
```

Thì dòng đó không được xem là sạch tuyệt đối.

Ví dụ:

```text
Số 40 Phố Trần Cung, Phường Cổ Nhuế 1 Quận Bắc Từ
( Nay Thuộc Phường Nghĩa Đô, HN ) Liêm, Hà Nội
```

Output hành chính vẫn lấy 3 cột riêng:

```text
Phường Cổ Nhuế 1
Quận Bắc Từ Liêm
Thành phố Hà Nội
```

Nhưng bắt buộc gắn cờ:

```text
RAW_CONTAINS_NAY_THUOC
OLD_NEW_ADMIN_CONFLICT
NEED_REVIEW_ADMIN_VERSION
```

Confidence không được cao hơn:

```text
0.60
```

---

## 8. ĐỊA CHỈ SẠCH

Cột `ĐỊA CHỈ SẠCH` chỉ được ghép từ các thành phần đã được xác định đúng.

Nếu muốn giữ đường/phố:

```text
[POI], [Đường/Phố], [Cấp 4], [Xã/Phường], [Huyện/Quận], [Tỉnh/TP]
```

Nếu không muốn giữ đường/phố:

```text
[POI], [Cấp 4], [Xã/Phường], [Huyện/Quận], [Tỉnh/TP]
```

Ví dụ:

```text
Raw: Đ,c 63/2/1 Nguyễn Phúc Chu, P. Thành Nhất, Đăk Lắk
```

Nếu không giữ đường:

```text
Phường Thành Nhất, Thành phố Buôn Ma Thuột, Tỉnh Đắk Lắk
```

Nếu giữ đường:

```text
Nguyễn Phúc Chu, Phường Thành Nhất, Thành phố Buôn Ma Thuột, Tỉnh Đắk Lắk
```

Debug:

```text
Số nhà: 63/2/1
Đường/Phố: Nguyễn Phúc Chu
```

---

## 9. removed_parts

`removed_parts` chỉ được chứa phần thật sự bị loại.

Không được chứa lại phần đã giữ trong:

```text
POI
Cấp 4
Xã/Phường
Huyện/Quận
Tỉnh/TP
Đường/Phố nếu output có giữ đường
```

Ví dụ sai:

```text
Cấp 4: Khu phố 5
removed_parts: Khu phố 5
```

Ví dụ đúng:

```text
Cấp 4: Khu phố 5
removed_parts: số nhà, tên người, số điện thoại, ghi chú giao hàng
```

Khuyến nghị kỹ thuật:

```text
Dùng span-based parser, không dùng string replace đơn giản.
```

Mỗi span nên có dạng:

```json
{
  "text": "Khu phố 5",
  "start": 10,
  "end": 19,
  "label": "LEVEL4",
  "keep": true
}
```

Sau khi parse xong:

```text
removed_parts = tất cả span có keep = false
```

Một span không được vừa là `LEVEL4/POI/STREET/ADMIN` vừa nằm trong `removed_parts`.

---

## 10. Confidence score

Không được để confidence cao nếu có dấu hiệu rủi ro.

Gợi ý mức confidence tối đa:

| Điều kiện | Confidence tối đa |
|---|---:|
| Match rõ POI/Cấp 4/admin, không conflict | 0.95–1.00 |
| Có cấp 4 suy luận không keyword | 0.75 |
| Raw có `cũ/mới/nay thuộc/sáp nhập` | 0.60 |
| Raw admin conflict với 3 cột riêng | 0.60 |
| POI bị nghi sai | 0.70 |
| Multiple POI hoặc multiple Level4 | 0.80 |
| Chỉ còn số nhà + đường + admin, không POI/Cấp 4 | 0.85 |
| POI nằm trong blacklist | 0.40 |
| removed_parts overlap kept_parts | 0.50 |

Công thức penalty gợi ý:

```python
score = 1.0

if RAW_ADMIN_CONFLICT_WITH_COLUMNS:
    score -= 0.35

if ADMIN_VERSION_AMBIGUOUS:
    score -= 0.30

if MULTIPLE_POI_FOUND:
    score -= 0.15

if MULTIPLE_LEVEL4_FOUND:
    score -= 0.15

if POSSIBLE_POI_MISSED:
    score -= 0.20

if POSSIBLE_LEVEL4_MISSED:
    score -= 0.20

if REMOVED_OVERLAPS_KEPT:
    score -= 0.25

if INVALID_POI_BLACKLIST:
    score -= 0.40

if RAW_CONTAINS_OLD_NEW_ADMIN:
    score = min(score, 0.60)
```

---

## 11. Bộ flags bắt buộc

Nếu không có lỗi thì `flags` để rỗng.

Các flag nên có:

```text
NO_POI_FOUND
NO_LEVEL4_FOUND
ONLY_STREET_LEVEL_FOUND
NO_GRANULAR_COMPONENT

POSSIBLE_POI_MISSED
POSSIBLE_LEVEL4_MISSED
MULTIPLE_POI_FOUND
MULTIPLE_LEVEL4_FOUND

INFERRED_LEVEL4_NO_KEYWORD
INVALID_POI_BLACKLIST
REMOVED_OVERLAPS_KEPT

RAW_ADMIN_CONFLICT_WITH_COLUMNS
RAW_CONTAINS_OLD_ADMIN
RAW_CONTAINS_NEW_ADMIN
RAW_CONTAINS_NAY_THUOC
OLD_NEW_ADMIN_CONFLICT
ADMIN_VERSION_AMBIGUOUS
NEED_REVIEW_ADMIN_VERSION

NOTE_AFTER_ADDRESS_REMOVED
RAW_CONTAINS_PHONE
RAW_CONTAINS_PERSON_NAME
LOW_CONFIDENCE
UNKNOWN_TOKEN_BEFORE_STREET
```

---

## 12. Pipeline xử lý khuyến nghị

Không xử lý theo kiểu xóa số nhà/đường trước.

Pipeline đúng:

```text
1. Normalize text
2. Sinh admin aliases từ 3 cột Xã/Huyện/Tỉnh
3. Gắn label ADMIN cho các alias trong raw
4. Cắt note/ghi chú giao hàng nếu nằm sau admin anchor
5. Tách phone/person/noise
6. Tách house_number
7. Tách street_name dựa vào:
   - street keyword
   - house_number + cụm chữ phía sau
8. Tách POI, nhưng phải dừng trước street/house/admin/level4
9. Tách Cấp 4 có keyword, dừng trước admin alias
10. Tách Cấp 4 suy luận không keyword nếu đứng ngay trước ward alias
11. Build ĐỊA CHỈ SẠCH
12. Build removed_parts từ span có keep=false
13. Validate flags
14. Tính confidence
```

---

## 13. Thứ tự ưu tiên label span

Nên dùng span-based parser với các nhãn:

```text
PERSON
PHONE
HOUSE_NUMBER
STREET
POI
LEVEL4
ADMIN_WARD
ADMIN_DISTRICT
ADMIN_PROVINCE
NOTE
UNKNOWN
```

Thứ tự ưu tiên label:

```text
1. ADMIN từ 3 cột có sẵn
2. NOTE / old-new admin note
3. HOUSE_NUMBER
4. STREET
5. POI
6. LEVEL4 có keyword
7. LEVEL4 suy luận không keyword
8. PERSON/NOISE
```

Riêng `ADMIN` nên được detect từ phải sang trái để tránh nhầm.

---

## 14. Test cases bắt buộc

### Case 1: Cấp 4 bị kéo nhầm admin alias

Input:

```text
Quỳnh Nhi .thôn Tân Lập .cumta .mdrak .dak Lak
```

Cột riêng:

```text
Xã/Phường: Xã Cư M'ta
Huyện/Quận: Huyện M'Đrắk
Tỉnh/TP: Tỉnh Đắk Lắk
```

Expected:

```text
POI: null
Số nhà: null
Đường/Phố: null
Cấp 4: Thôn Tân Lập
Flags: RAW_CONTAINS_PERSON_NAME
```

Không được output:

```text
Cấp 4: Thôn Tân Lập .cumta .mdrak .dak
```

---

### Case 2: Không bắt `Bản` trong Huyện Vụ Bản

Input:

```text
Trạm Y Tế Tam Thanh, Huyện Vụ Bản, T. Nam Định
```

Cột riêng:

```text
Xã/Phường: Xã Tam Thanh
Huyện/Quận: Huyện Vụ Bản
Tỉnh/TP: Tỉnh Nam Định
```

Expected:

```text
POI: Trạm y tế Tam Thanh
Cấp 4: null
Flags: NO_LEVEL4_FOUND
```

Không được output:

```text
Cấp 4: Bản T. Nam Định
```

---

### Case 3: Đường/phố có số nhà và địa chỉ cũ/mới

Input:

```text
Số 40 Phố Trần Cung, Phường Cổ Nhuế 1 Quận Bắc Từ( Nay Thuộc Phường Nghĩa Đô, Hn ) Liêm, Hà Nội
```

Cột riêng:

```text
Xã/Phường: Phường Cổ Nhuế 1
Huyện/Quận: Quận Bắc Từ Liêm
Tỉnh/TP: Thành phố Hà Nội
```

Expected:

```text
Số nhà: 40
Đường/Phố: Phố Trần Cung
POI: null
Cấp 4: null
Flags:
- RAW_CONTAINS_NAY_THUOC
- OLD_NEW_ADMIN_CONFLICT
- NEED_REVIEW_ADMIN_VERSION
Confidence: <= 0.60
```

---

### Case 4: Đường không có keyword nhưng đứng sau số nhà

Input:

```text
Đ,c 63/2/1 Nguyễn Phúc Chu, P. Thành Nhất, Đăk Lắk
```

Cột riêng:

```text
Xã/Phường: Phường Thành Nhất
Huyện/Quận: Thành phố Buôn Ma Thuột
Tỉnh/TP: Tỉnh Đắk Lắk
```

Expected:

```text
Số nhà: 63/2/1
Đường/Phố: Nguyễn Phúc Chu
POI: null
Cấp 4: null
```

---

### Case 5: Cấp 4 suy luận không keyword

Input:

```text
Dc: Quang Yên- Tam Đình - Tương Dương - Nghệ An
```

Cột riêng:

```text
Xã/Phường: Xã Tam Đình
Huyện/Quận: Huyện Tương Dương
Tỉnh/TP: Tỉnh Nghệ An
```

Expected:

```text
POI: null
Số nhà: null
Đường/Phố: null
Cấp 4: Quang Yên
Flags: INFERRED_LEVEL4_NO_KEYWORD
Confidence: <= 0.75
```

---

### Case 6: Tên người + số nhà + đường + admin alias

Input:

```text
Quách Thị Huyền, 122 Hùng Vương, Tt Quảng, Cư Mgar, Đăk Lăk
```

Cột riêng:

```text
Xã/Phường: Thị trấn Quảng Phú
Huyện/Quận: Huyện Cư M'gar
Tỉnh/TP: Tỉnh Đắk Lắk
```

Expected:

```text
Số nhà: 122
Đường/Phố: Hùng Vương
POI: null
Cấp 4: null
removed_parts: Quách Thị Huyền
Flags: RAW_CONTAINS_PERSON_NAME
```

---

### Case 7: POI phải dừng trước đường

Input:

```text
Trung Tâm Chăm Sóc Mẹ Và Bé Khánh Trần K 2 Pha Đường Nguyễn Văn Cừ, Phường Lê Hồng Phong, Thành Phố Phủ Lý, Tỉnh Hà Nam
```

Cột riêng:

```text
Xã/Phường: Phường Lê Hồng Phong
Huyện/Quận: Thành phố Phủ Lý
Tỉnh/TP: Tỉnh Hà Nam
```

Expected:

```text
POI: Trung tâm Chăm Sóc Mẹ Và Bé Khánh Trần
Đường/Phố: Đường Nguyễn Văn Cừ
Cấp 4: null
Flags: UNKNOWN_TOKEN_BEFORE_STREET
```

Không được output:

```text
POI: Trung tâm Chăm Sóc Mẹ Và Bé Khánh Trần K 2 Pha
```

---

### Case 8: Số nhà dạng mã nhà

Input:

```text
1F Lương Thế Vinh, P9, Tp- Vt, . Ko Đúng Hàng Như Hình Chị Trả Lại Nhé Em.miễn Ship Nhé
```

Expected:

```text
Số nhà: 1F
Đường/Phố: Lương Thế Vinh
POI: null
Cấp 4: null
Flags:
- ONLY_STREET_LEVEL_FOUND
- NOTE_AFTER_ADDRESS_REMOVED
- NO_POI_FOUND
- NO_LEVEL4_FOUND
```

---

### Case 9: Số nhà dạng slash + chữ

Input:

```text
259/ C3 Phan Bội Châu
```

Expected:

```text
Số nhà: 259/C3
Đường/Phố: Phan Bội Châu
POI: null
Cấp 4: null
```

---

## 15. Tiêu chí cuối cùng để gọi là sạch

Một dòng được xem là sạch khi đạt đủ các điều kiện sau:

```text
1. Xã/Phường, Huyện/Quận, Tỉnh/TP lấy đúng từ 3 cột riêng.

2. Raw address không ghi đè 3 cột admin, chỉ dùng để đối chiếu và bóc phần chi tiết.

3. Nếu có số nhà thì tách được số nhà.

4. Nếu sau số nhà là tên đường thì tách được đường/phố.

5. Nếu có POI thì POI không nuốt số nhà, đường/phố, cấp 4 hoặc admin.

6. Nếu có Cấp 4 thì Cấp 4 dừng trước xã/huyện/tỉnh/admin alias.

7. Không bắt nhầm admin alias thành Cấp 4.
   Ví dụ: Huyện Vụ Bản không được thành Cấp 4 = Bản.

8. Không bắt nhầm câu giao tiếp thành POI.
   Ví dụ: Em gửi về không phải POI.

9. Nếu có địa chỉ cũ/mới/sáp nhập/nay thuộc thì phải gắn flag review.

10. removed_parts không được chứa phần đã được giữ lại.

11. confidence phải phản ánh đúng rủi ro, không được 1.0 cho dòng có conflict.

12. Dòng không có POI/Cấp 4 nhưng chỉ có số nhà + đường + admin vẫn có thể xem là hợp lệ,
    nhưng phải gắn NO_POI_FOUND và NO_LEVEL4_FOUND.
```

---

## 16. Kết luận thiết kế

Tiêu chí sạch mới không phải là:

```text
Bỏ được số nhà/đường là sạch.
```

Mà phải là:

```text
Nhận diện đúng số nhà/đường trước,
dùng chúng làm boundary để không làm bẩn POI và Cấp 4,
rồi mới quyết định giữ hay bỏ đường/phố ở output cuối.
```

Parser cần ba lớp bắt buộc:

```text
1. Admin-anchor layer từ 3 cột có sẵn
2. Street parser layer: số nhà + đường/phố
3. POI/Level4 parser layer có boundary rõ ràng
```

Nếu thiếu `street parser layer`, các lỗi phổ biến sẽ tiếp tục xảy ra:

```text
- POI nuốt nhầm đường
- Cấp 4 kéo nhầm xã/huyện/tỉnh viết tắt
- Không bóc được đường không có keyword
- Không phát hiện được conflict cũ/mới
```
