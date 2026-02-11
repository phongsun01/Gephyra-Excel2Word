# Hướng dẫn tạo Template Word với Bảng Động (docxtpl)

## Tổng quan
File này hướng dẫn cách tạo template Word (.docx) với bảng động sử dụng thư viện `docxtpl`.

## Cú pháp cơ bản

### 1. Placeholder đơn giản
```
<<TenBien>>
```
Ví dụ: `<<HoTen>>`, `<<SoHD>>`, `<<DiaChi>>`

### 2. Vòng lặp trong đoạn văn
```
<% for item in DanhSach %>
  - <<item.Ten>>: <<item.GiaTri>>
<% endfor %>
```

### 3. Bảng động (QUAN TRỌNG)

**Cách tạo:**
1. Mở Microsoft Word
2. Tạo bảng với header (hàng tiêu đề)
3. Thêm 1 hàng mẫu bên dưới header
4. Trong hàng mẫu, thêm cú pháp như sau:

**Cột đầu tiên (cell đầu tiên của hàng):**
```
<%tr for item in BangGia%><<item.TenSP>>
```

**Các cột tiếp theo:**
```
<<item.DVT>>
```
```
<<item.Gia>>
```

**Cột cuối cùng (cell cuối cùng của hàng):**
```
<<item.Gia>><%endtr%>
```

> **LƯU Ý:** Tag `<%tr%>` và `<%endtr%>` phải nằm **TRONG** các cell của bảng, không phải ở ngoài.

## Ví dụ Template hoàn chỉnh

### Cấu trúc file Word:

```
HOP DONG MAU
Cong Hoa Xa Hoi Chu Nghia Viet Nam
Doc lap - Tu do - Hanh phuc
---
So Hop Dong: <<SoHD>>
Ben A: Cong Ty ABC
Ben B: <<HoTen>>
Dia Chi: <<DiaChi>>

Danh Sach San Pham
Bang gia chi tiet:

┌─────────────┬─────────┬─────────┐
│   Ten SP    │   DVT   │   Gia   │  ← Header row
├─────────────┼─────────┼─────────┤
│ <%tr for item in BangGia%><<item.TenSP>> │ <<item.DVT>> │ <<item.Gia>><%endtr%> │  ← Template row
└─────────────┴─────────┴─────────┘
```

## Dữ liệu Context (Python)

```python
context = {
    'SoHD': 'HD001',
    'HoTen': 'Nguyen Van A',
    'DiaChi': 'Ha Noi',
    'BangGia': [
        {'TenSP': 'San Pham 1', 'DVT': 'Cai', 'Gia': 100000},
        {'TenSP': 'San Pham 2', 'DVT': 'Hop', 'Gia': 200000},
        {'TenSP': 'San Pham 3', 'DVT': 'Bo', 'Gia': 150000},
    ]
}
```

## Kết quả sau khi render

```
HOP DONG MAU
Cong Hoa Xa Hoi Chu Nghia Viet Nam
Doc lap - Tu do - Hanh phuc
---
So Hop Dong: HD001
Ben A: Cong Ty ABC
Ben B: Nguyen Van A
Dia Chi: Ha Noi

Danh Sach San Pham
Bang gia chi tiet:

┌─────────────┬─────────┬─────────┐
│   Ten SP    │   DVT   │   Gia   │
├─────────────┼─────────┼─────────┤
│ San Pham 1  │   Cai   │ 100000  │
│ San Pham 2  │   Hop   │ 200000  │
│ San Pham 3  │   Bo    │ 150000  │
└─────────────┴─────────┴─────────┘
```

## Lưu ý quan trọng

1. **Delimiter:** Ứng dụng này sử dụng `<<` `>>` thay vì `{{` `}}` mặc định của Jinja2.
2. **Block tags:** Sử dụng `<%` `%>` cho vòng lặp và điều kiện.
3. **Không dùng python-docx:** Không thể tạo template bảng động bằng `python-docx`. Phải tạo thủ công bằng Microsoft Word.
4. **Kiểm tra XML:** Nếu template không hoạt động, mở file .docx bằng 7-Zip và kiểm tra `word/document.xml` để đảm bảo cú pháp đúng.

## Tài liệu tham khảo
- docxtpl documentation: https://docxtpl.readthedocs.io/
- Jinja2 template syntax: https://jinja.palletsprojects.com/
