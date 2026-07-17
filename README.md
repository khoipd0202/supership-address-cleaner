# Hướng dẫn Cài đặt & Chạy ứng dụng SuperShip Address Cleaner (`vn-address-cleaner`)

Chào mừng bạn đến với **SuperShip Address Cleaner**! Đây là tài liệu hướng dẫn chi tiết từng bước (step-by-step) dành cho những thành viên mới clone dự án về máy tính cá nhân để có thể cài đặt môi trường, cấu hình và chạy thử ứng dụng một cách nhanh chóng nhất.

---

## 🚀 Giới thiệu dự án

`vn-address-cleaner` là một thư viện Python và công cụ dòng lệnh (CLI) mạnh mẽ dùng để chuẩn hóa, làm sạch địa chỉ giao hàng tiếng Việt từ các tệp Excel. 

### Điểm nổi bật:
- **Phân tách thông tin chính xác**: Phân chia chuỗi địa chỉ thô thành 6 cột chuẩn: `poi` (điểm định vị/tên riêng), `tên đường`, `cấp 4` (thôn/xóm/ấp/tổ dân phố...), `phường/xã`, `quận/huyện`, `tỉnh/tp`.
- **Cập nhật Địa giới Hành chính mới nhất (2025)**: Tích hợp dữ liệu chuẩn hóa hành chính của Việt Nam, tự động xử lý các trường hợp đổi tên, sáp nhập xã/phường/quận/huyện.
- **Mô hình kết hợp Rule-LLM thông minh**:
  - **Rule-based (80% - 95% số dòng)**: Xử lý cục bộ bằng thuật toán nhanh chóng và chính xác với những địa chỉ rõ ràng, giúp tiết kiệm chi phí API và thời gian.
  - **LLM Fallback (Cerebras API / Llama-3)**: Tự động phát hiện và gửi các dòng địa chỉ mơ hồ, phức tạp hoặc bị lỗi cấu trúc lên LLM để phân tích ngữ nghĩa sâu.
- **Tiết kiệm chi phí**: Hệ thống tự gom nhóm địa chỉ (deduplication) để loại bỏ trùng lặp trước khi gửi lên API của LLM.
- **Local Web UI & CLI**: Cung cấp giao diện Web trực quan (kéo thả và tải về file kết quả gộp) lẫn giao diện dòng lệnh (CLI) mạnh mẽ hỗ trợ xử lý hàng chờ tuần tự (`--queue-all`) cho các file dữ liệu cực lớn (hàng ngàn dòng).

---

## 🛠️ Yêu cầu hệ thống

Trước khi bắt đầu, hãy đảm bảo máy tính của bạn đã cài đặt:
- **Python**: Phiên bản `3.10` trở lên.
- **Git**: Để clone mã nguồn.

---

## 📦 Hướng dẫn cài đặt step-by-step

Thực hiện các bước sau để thiết lập môi trường chạy dự án:

### Bước 1: Clone mã nguồn
Mở Terminal (macOS/Linux) hoặc Command Prompt/PowerShell (Windows) và chạy lệnh:
```bash
git clone git@github.com:khoipd0202/supership-address-cleaner.git
# Hoặc dùng HTTPS nếu bạn chưa thiết lập SSH key:
# git clone https://github.com/khoipd0202/supership-address-cleaner.git

cd supership-address-cleaner
```

### Bước 2: Tạo và kích hoạt môi trường ảo (Virtual Environment)
Việc sử dụng môi trường ảo sẽ giúp cô lập các thư viện của dự án, tránh xung đột với các phiên bản Python khác trên máy của bạn:

- **Trên macOS / Linux:**
  ```bash
  python3 -m venv .venv
  source .venv/bin/activate
  ```
- **Trên Windows (PowerShell):**
  ```powershell
  python -m venv .venv
  .venv\Scripts\activate.ps1
  ```
- **Trên Windows (CMD):**
  ```cmd
  python -m venv .venv
  .venv\Scripts\activate.bat
  ```

Sau khi kích hoạt, bạn sẽ thấy ký hiệu `(.venv)` xuất hiện ở đầu dòng lệnh của Terminal.

### Bước 3: Nâng cấp công cụ quản lý thư viện pip
```bash
python3 -m pip install --upgrade pip
```

### Bước 4: Cài đặt thư viện ở chế độ phát triển (Editable Mode)
Chọn một trong hai chế độ cài đặt dưới đây tùy vào nhu cầu của bạn:

1. **Chế độ Cơ bản (Chỉ dùng Rule-based cục bộ, không gọi LLM):**
   ```bash
   python3 -m pip install -e .
   ```
2. **Chế độ Đầy đủ (Hỗ trợ gọi LLM qua Cerebras API):**
   ```bash
   python3 -m pip install -e ".[llm]"
   ```

---

## 🔑 Cấu hình biến môi trường (Environment Variables)

Nếu bạn muốn sử dụng chức năng nâng cao (gọi LLM xử lý các địa chỉ mơ hồ), bạn cần thiết lập API Key:

1. Sao chép tệp mẫu `.env.example` thành `.env`:
   ```bash
   cp .env.example .env
   ```
2. Mở file `.env` bằng bất kỳ trình soạn thảo văn bản nào và cấu hình các giá trị:
   ```env
   # API Key của Cerebras (lấy miễn phí tại trang chủ Cerebras Cloud)
   CEREBRAS_API_KEY="csk-your_actual_api_key_here"

   # Các cấu hình nâng cao khác (đã cấu hình sẵn mặc định)
   CEREBRAS_MODEL="gpt-oss-120b"
   CEREBRAS_BATCH_SIZE=5
   CEREBRAS_MAX_ROWS_PER_RUN=30
   ```

---

## 🖥️ Hướng dẫn chạy và sử dụng dự án

Dự án cung cấp 3 cách sử dụng chính tùy thuộc vào mục đích của bạn:

### Cách 1: Sử dụng qua Giao diện Web (Local Web UI)
Đây là cách dễ nhất cho người dùng thông thường, hỗ trợ upload nhiều file Excel cùng lúc và tự động gộp kết quả tải xuống.

1. Chạy file giao diện Web cục bộ:
   ```bash
   python3 address_ui.py
   ```
   *(Mặc định server sẽ chạy tại cổng `8899`. Nếu muốn chạy cổng khác, ví dụ cổng `9000`, hãy dùng lệnh: `python3 address_ui.py 9000`)*.
2. Mở trình duyệt web bất kỳ và truy cập địa chỉ:
   ```text
   http://127.0.0.1:8899
   ```
3. Kéo thả hoặc chọn các file Excel cần làm sạch, cấu hình tùy chọn LLM (nếu có API Key), bấm nút chạy và tải về file kết quả gộp cuối cùng.

### Cách 2: Sử dụng qua Dòng lệnh (CLI)
Sau khi cài đặt thành công ở Bước 4, hệ thống đã tự động đăng ký lệnh `vn-address-clean` toàn cục (trong môi trường ảo `.venv`).

- **Xử lý cơ bản (Chỉ dùng Rules):**
  ```bash
  vn-address-clean input.xlsx -o output.xlsx
  ```
- **Xử lý nâng cao (Bật LLM hỗ trợ các dòng khó):**
  ```bash
  vn-address-clean input.xlsx -o output.xlsx --cerebras
  ```
- **Chế độ xử lý tuần tự (Hàng chờ an toàn cho tệp lớn):**
  Khi xử lý file lớn (từ 5,000 đến 10,000 dòng), bạn nên dùng cờ `--queue-all`. Hệ thống sẽ gửi tuần tự các dòng mơ hồ lên LLM, tự động giãn cách nhịp yêu cầu và tự thử lại (backoff) nếu bị chạm giới hạn tần suất gọi API (Rate Limit / HTTP 429).
  ```bash
  vn-address-clean input.xlsx -o output.xlsx --queue-all
  ```
- **Các cờ tùy chọn hữu ích:**
  - `--include-empty-rows`: Giữ nguyên và đưa các dòng không bóc tách được chi tiết (trống POI/Đường/Cấp 4) vào file đầu ra thay vì lọc bỏ.
  - `--combined-row`: Gộp các thành phần địa chỉ chi tiết trong một dòng Excel duy nhất thay vì tách nhỏ ra các dòng con.
  - `--sheet-name "Tên Sheet"`: Chọn tên Sheet cụ thể cần đọc trong file Excel.

*Mẹo: Nếu bạn không muốn cài đặt lệnh toàn cục, bạn vẫn có thể chạy trực tiếp module Python:*
```bash
python3 -m vn_address_cleaner.cli input.xlsx -o output.xlsx
```

### Cách 3: Sử dụng dưới dạng thư viện Python (Library Usage)
Tích hợp quy trình chuẩn hóa địa chỉ trực tiếp vào code Python của bạn:

- **Làm sạch tệp Excel:**
  ```python
  from vn_address_cleaner import clean_excel

  stats = clean_excel(
      input_path="data/orders.xlsx",
      output_path="outputs/cleaned_orders.xlsx",
      use_cerebras=True,          # Bật LLM fallback cho các dòng khó
      split_components=True,      # Tách nhỏ các thành phần POI/Street/Level 4 ra các dòng riêng biệt
  )
  print("Thống kê kết quả:", stats.as_dict())
  ```
- **Làm sạch một chuỗi địa chỉ duy nhất:**
  ```python
  from vn_address_cleaner import AddressCleaner

  cleaner = AddressCleaner()
  result = cleaner.clean(
      raw_address="trường cao đẳng cơ giới ninh bình, đường vũ duy thanh, tổ 2, phường yên bình, thành phố tam điệp, ninh bình",
      ward="yên bình",
      district="tam điệp",
      province="ninh bình",
  )

  print("Mảng xuất Excel:", result.as_output_row())
  # Kỳ vọng xuất ra: ["trường cao đẳng cơ giới ninh bình", "đường vũ duy thanh", "tổ 2", "phường yên bình", "thành phố tam điệp", "tỉnh ninh bình"]
  ```

---

## 🧪 Chạy kiểm thử (Testing)

Để đảm bảo các quy tắc phân tách địa chỉ và việc tích hợp LLM hoạt động đúng như mong đợi, hãy chạy bộ kiểm thử tự động:

- **Chạy kiểm thử các logic cốt lõi & rules:**
  ```bash
  python3 -m unittest discover -s tests -p 'test*.py'
  ```
- **Chạy các kịch bản kiểm thử tích hợp (LLM, thư viện nâng cao):**
  ```bash
  python3 -m unittest discover -s scratch -p 'test*.py'
  ```

---

## 🛡️ Bảo mật và quyền riêng tư (Privacy)

> [!CAUTION]
> Dữ liệu địa chỉ giao hàng và thông tin khách hàng là cực kỳ nhạy cảm.
> - **Không bao giờ** commit các tệp tin Excel thực tế của khách hàng, các tệp kết quả xử lý nằm trong thư mục `outputs/`, hoặc tệp cache SQLite địa phương (`cache_diachi.sqlite`) lên các kho lưu trữ công khai như GitHub.
> - Hãy kiểm tra kỹ tệp `.gitignore` trước khi thực hiện các lệnh push commit lên server.

---
Chúc bạn cài đặt thành công và làm việc hiệu quả với dự án! Nếu gặp bất kỳ khó khăn nào trong quá trình cài đặt, vui lòng liên hệ với quản trị viên dự án hoặc tạo issues trên kho mã nguồn.
