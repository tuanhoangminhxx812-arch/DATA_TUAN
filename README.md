# 📊 Đối Chiếu Trung Gian

Công cụ phân tích & đối chiếu sổ chi tiết tài khoản trung gian — Kiểm tra cân đối Nợ / Có.

## 🎯 Mô tả

Ứng dụng web được xây dựng bằng **Streamlit**, hỗ trợ kế toán viên tải lên file Excel sổ chi tiết tài khoản và tự động phân tích, đối chiếu số liệu phát sinh Nợ và Có. Hệ thống sử dụng thuật toán khớp thông minh nhiều tầng (multi-pass smart matching) để xác định và loại bỏ các cặp bút toán cân đối, từ đó chỉ ra những dòng thực sự gây chênh lệch.

## ✨ Tính năng chính

- **Tải lên & đọc file Excel** — Hỗ trợ `.xlsx` / `.xls`, tự động nhận diện cấu trúc dữ liệu
- **Phân tích thông minh nhiều tầng:**
  - **Pass 1:** Khớp 1-1 Nợ = Có cùng người lập (theo số tiền)
  - **Pass 2:** Khớp Nợ dương + Nợ âm triệt tiêu (cùng người lập)
  - **Pass 3:** Khớp Có dương + Có âm triệt tiêu (cùng người lập)
  - **Pass 4:** Khớp chéo giữa các người lập (Nợ người A = Có người B)
  - **Pass 5:** Khớp nhóm — tổng nhiều dòng Nợ = tổng nhiều dòng Có
- **Báo cáo trực quan** — Bảng tổng hợp theo người lập, thống kê khớp/chưa khớp, highlight dòng chênh lệch
- **Xuất báo cáo Excel** — File gồm 2 sheet: *Dữ liệu gốc* và *Báo cáo chênh lệch*

## 📁 Cấu trúc dự án

```
đối chiếu trung gian/
├── app.py              # Ứng dụng Streamlit chính (giao diện + logic phân tích)
├── analyze_pairs.py    # Script phân tích pattern khớp nhóm (debug/testing)
├── export_report.py    # Script xuất báo cáo đối chiếu (chạy độc lập, CLI)
├── requirements.txt    # Danh sách thư viện Python cần thiết
└── README.md           # Tài liệu hướng dẫn
```

## 🚀 Cài đặt & Chạy

### 1. Cài đặt thư viện

```bash
pip install -r requirements.txt
```

### 2. Chạy ứng dụng web

```bash
streamlit run app.py
```

Ứng dụng sẽ mở tại `http://localhost:8501`.

### 3. Chạy script xuất báo cáo (CLI)

```bash
python export_report.py
```

> **Lưu ý:** Cần chỉnh sửa `INPUT_FILE` và `OUTPUT_DIR` trong file `export_report.py` trước khi chạy.

## 📖 Hướng dẫn sử dụng

1. Mở ứng dụng web (`streamlit run app.py`)
2. **Bước 1:** Tải lên file Excel sổ chi tiết tài khoản (kéo thả hoặc chọn file)
3. **Bước 2:** Nhấn nút **"🔍 Phân tích"** để hệ thống tự động đối chiếu
4. **Bước 3:** Xem kết quả và nhấn **"📥 Xuất file báo cáo Excel"** để tải báo cáo

## 🛠️ Yêu cầu hệ thống

- Python 3.8+
- Trình duyệt web hiện đại (Chrome, Firefox, Edge...)

## 📄 Định dạng file đầu vào

File Excel cần có cấu trúc sổ chi tiết tài khoản với các cột:

| Cột | Nội dung |
|-----|----------|
| A | Nguồn bút toán |
| B | Ngày |
| C | Số CT Phân hệ phụ |
| D | Số CT Phân hệ GL |
| E | Diễn giải |
| F | Số phát sinh Nợ |
| G | Số phát sinh Có |
| H | Người lập |

Hệ thống tự động nhận diện dòng bắt đầu dữ liệu (sau "Số dư đầu kỳ") và dòng kết thúc (trước "Cộng phát sinh").

---

© 2026 — Công cụ phân tích sổ chi tiết tài khoản
