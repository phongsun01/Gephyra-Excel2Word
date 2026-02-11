# Gephyra - Batch Document Generator

**Gephyra** is a powerful tool to generate multiple Word documents from a standard Excel file using a Word template.

[English](#english) | [Tiếng Việt](#tiếng-việt)

---

# English

## 🚀 Features
- **Hybrid Configuration**: Use `config.yaml` for settings and **Excel** for table mappings.
- **Modern GUI**: Clean, professional interface built with Tkinter Pro.
- **Smart Templates**: Use `<< Variable >>` syntax in your Word templates.
- **Table Support**: Automatically populate tables in Word from Excel ranges.
- **Dynamic Filenames**: Generate filenames based on data (e.g., `Contract-NguyenVanA-HD001.docx`).

## 🛠️ Installation

1.  **Install Python 3.10+**
2.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

## 🏃 Usage

### 1. Generate Sample Data
Run this script to create `examples/Project_Sample` with dummy data and template:
```bash
python generate_samples.py
```

### 2. Run GUI
```bash
python main.py
```
- Click **Load Config** and select `examples/Project_Sample/config.yaml`.
- Click **GENERATE FILES** to start processing.

### 3. Run CLI (Command Line)
```bash
python main.py --config examples/Project_Sample/config.yaml
```

## 📦 Building Executable (.exe)
To create a standalone `.exe` file for Windows:

1.  Ensure `pyinstaller` is installed (`pip install pyinstaller`).
2.  Run the build script:
    ```bash
    build.bat
    ```
3.  The output file `Gephyra.exe` will be in the `dist` folder.

---

# Tiếng Việt

**Gephyra** là công cụ giúp tạo hàng loạt file Word từ file Excel theo template có sẵn, hỗ trợ điền bảng biểu và đặt tên file động.

## 🚀 Tính năng chính
- **Cấu hình lai (Hybrid)**: Dùng `config.yaml` cho cài đặt chung và **Excel** để map dữ liệu bảng.
- **Giao diện hiện đại**: Giao diện GUI đẹp, chuyên nghiệp, dễ sử dụng.
- **Template thông minh**: Sử dụng cú pháp `<< Bien >>` trong file Word để điền dữ liệu.
- **Hỗ trợ Bảng biểu**: Tự động điền dữ liệu vào bảng trong Word từ vùng dữ liệu Excel.
- **Tên file động**: Tự động đặt tên file output theo dữ liệu (VD: `HopDong-NguyenVanA-HD001.docx`).

## 🛠️ Cài đặt

1.  **Cài đặt Python 3.10+**
2.  **Cài đặt thư viện**:
    Mở terminal tại thư mục project và chạy:
    ```bash
    pip install -r requirements.txt
    ```

## 🏃 Hướng dẫn sử dụng sơ lược

### Bước 1: Chuẩn bị dữ liệu mẫu
Chạy script sau để tạo thư mục `examples/Project_Sample` chứa dữ liệu và template mẫu:
```bash
python generate_samples.py
```

### Bước 2: Chạy phần mềm (Giao diện)
```bash
python main.py
```
1.  Bấm **Load Config (YAML)** và chọn file `examples/Project_Sample/config.yaml`.
2.  Xem trước dữ liệu ở bảng bên phải để đảm bảo Excel đã được đọc đúng.
3.  Bấm nút **GENERATE FILES** màu cam để bắt đầu tạo file.
4.  File kết quả sẽ nằm trong thư mục `Output`.

### Bước 3: Chạy bằng dòng lệnh (Advanced)
Dành cho việc tích hợp vào các hệ thống khác:
```bash
python main.py --config examples/Project_Sample/config.yaml
```

## 📦 Đóng gói thành file .exe
Để tạo file chạy `.exe` độc lập trên Windows (không cần cài Python):

1.  Cài đặt PyInstaller: `pip install pyinstaller`
2.  Chạy file script:
    ```bash
    build.bat
    ```
3.  File **Gephyra.exe** sẽ được tạo trong thư mục `dist`.

## 📂 Cấu trúc dự án
```
Gephyra/
├── src/                # Mã nguồn chính
│   ├── gui.py          # Giao diện người dùng (GUI)
│   ├── batch_processor.py # Xử lý logic chính
│   └── ...
├── examples/           # Thư mục chứa ví dụ mẫu
├── main.py            # File chạy chính
├── config_manager.py   # Quản lý cấu hình
└── requirements.txt    # Danh sách thư viện cần thiết
```
