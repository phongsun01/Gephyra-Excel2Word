# Gephyra - Development Roadmap

This document outlines the planned features and improvements for Gephyra in future phases.

## Phase 2: Multi-file & m×n (2 tuần) ⭐⭐⭐⭐⭐

### Chức năng mới
- **m rows × n templates = m×n files:** 1 Excel × nhiều template
- **Batch multi-Excel:** Xử lý nhiều Excel trong `Data/` cùng lúc
- **GUI dynamic:** Dropdown chọn sheet/range runtime, không edit config thủ công
- **Template matrix:** Excel×Template checklist → preview file count

### Config mở rộng
```yaml
batch_configs:
  - excel: "data_hd.xlsx"        # m1=50 rows
    templates:                   # n=2 templates
      - "muaban.docx"
      - "phuluc.docx"
    output_prefix: "{HoTen}_{tpl}"
  - excel: "data_bb.xlsx"        # m2=30
    templates: ["bienban.docx"]  # n=1
# Output: 50×2 + 30×1 = 130 files
```

### GUI mới
```text
[Batch Matrix View]
☑ data_hd.xlsx (50r) × muaban, phuluc = 100 files
☑ data_bb.xlsx (30r) × bienban = 30 files
Total: 130 files | ETA: 5min
[Preview Matrix] [Run Selected]
```

### Use case
Hợp đồng → sinh 3 files: Hợp đồng chính + Phụ lục + Biên bản (m×3).

---

## Phase 3: Data Query & JOIN (2 tuần) ⭐⭐⭐⭐⭐

### Chức năng mới
- **Multi-sheet JOIN:** Kết hợp sheet Text + DanhMuc (LEFT JOIN theo key)
- **Conditional filter:** WHERE Status="Approved" AND Gia>1M
- **Query config:** Từ Excel sheet "Config_Query", "Config_Filter" (zero hard-code)
- **Visual Query Builder:** Drag-drop GUI tạo JOIN/WHERE

### Excel Config sheets
```text
excel_config.xlsx:
├── Config_Query:    Source|Sheet|Columns|JoinKey|Type
├── Config_Filter:   Source|Column|Operator|Value
└── Config_Run:      text_sheet, table_sheet, range...
```

### Workflow JOIN
```text
Text sheet: HoTen, SoHD
DM sheet: MaTB, TenTB, Gia
→ JOIN on SoHD → Context: {HoTen, SoHD, TenTB, Gia}
→ Filter: Status="Approved" AND Gia>1M
→ Render 25/50 rows match
```

### GUI Query Builder
```text
[Visual Builder]
Text --LEFT JOIN--> DanhMuc [on SoHD]
WHERE Status="Approved" AND Gia>1M
[Preview 25 rows] [Generate Config] [Run]
```

---

## Phase 4: Zero Hard-code System (1 tuần) ⭐⭐⭐⭐

### Config 3-cấp (ưu tiên Runtime > Project > Global)
```text
Level 1 Global: config.yaml (paths, workers)
Level 2 Project: project_config.yaml (defaults)
Level 3 Runtime: Excel Config_* sheets (queries, rules)
```

### Excel Config chuẩn (6 sheets)
```text
excel_config.xlsx:
├── Config_Run:        Runtime settings
├── Config_Query:      JOIN logic
├── Config_Filter:     WHERE conditions
├── Config_Tables:     Multi-table mappings (Phase 5)
├── Config_Output:     m×n matrix Excel×Template
└── Config_Expressions: Computed fields (Phase 7)
```

### Ưu điểm
- User chỉnh config trong Excel, không cần edit YAML
- Share config giữa team (1 file excel_config.xlsx)
- Version control dễ (diff Excel sheets)

---

## Phase 5: Advanced m×n Patterns (1 tuần) ⭐⭐⭐

### Chức năng mới
- **Conditional template:** Row "MuaBan" → template A; "BaoHanh" → template B+C
- **Multi-table per template:** 1 template paste nhiều table vào bookmark khác nhau
- **Cross-product matrix:** k Excel × l templates tùy ý

### Config_Output sheet
```text
Excel_File | Template_File | Output_Prefix | Enabled
data_hd    | muaban.docx   | HD_MB        | ✓
data_hd    | phuluc.docx   | HD_PL        | ✓
data_bb    | bienban.docx  | BB           | ✓
```

### Config_Tables (multi-table)
```text
Bookmark | Sheet | Range | Mode
TablePos1| Table | A1:D20| KeepFormat
TablePos2| DM    | A1:F10| KeepFormat
TablePos3| Sum   | A1:C5 | MergeFormat
```

---

## Phase 6: Advanced GUI (3 tuần) ⭐⭐⭐

### 6.1 Template Matrix Manager
```text
        muaban  phuluc  bienban
data_hd  50✓     50✓      -
data_bb   -       -      30✓
TOTAL   100     50      30
[Edit Matrix] [Run All]
```

### 6.2 Data Preview & Validation
```text
Text Data (25/50 after filter):
HoTen | SoHD | Status  | Issues
NgA   | 001  |Approved | ✓
TrB   | 002  |Pending  | ⚠ Skip
[Fix Issues] [Export Fixed Excel]
```

### 6.3 Job Queue & Monitoring
```text
Running: data_hd×muaban 45/50 [█████░]
Queue:   data_hd×phuluc, data_bb×bienban
ETA: 4min | Errors: 2 | Rate: 12files/min
[Pause] [Skip] [Stop]
```

---

## Phase 7: Expression Engine (2 tuần) ⭐⭐

### Computed fields config-driven
```text
Config_Expressions:
Field     | Formula
FullName  | {HoTen} + " - " + {DiaChi}
Total     | {Gia} * {SL}
DateVN    | format_date({Ngay}, "%d/%m/%Y")
IsValid   | {Status} == "OK"
```

### Custom functions (sandbox)
```text
Config_Functions (Python code):
def viet_hoa(ten): return ten.upper()
def tinh_thue(gia, pt=10): return gia*(1+pt/100)
```

### Template render
- **Template:** `{{FullName}} - {{Total}} VNĐ`
- **Output:** `Nguyễn Văn A - Nam Định - 1.500.000.000 VNĐ`

---

## Phase 8: Production Features (2 tuần) ⭐⭐⭐

### Error Handling
- **Modes:** skip | stop | retry | save_partial
- **Checkpoint:** Resume từ interruption (mỗi 10 files)
- **Detailed log:** Per-row error với context

### Performance
- **Parallel workers:** 4 files đồng thời
- **Chunk processing:** 50 rows/batch
- **Memory management:** Stream large Excel

### Template Management
- **Auto-scan placeholders:** `{{HoTen}}` → match Excel columns
- **Template versioning:** v1.0 → v2.0
- **Preview diff:** Compare template versions

---

## Phase 9: Deployment & Share (1 tuần) ⭐⭐

### Config Package Export
```text
Config_Package_v1.2.zip:
├── config.yaml
├── excel_config.xlsx (all Config_*)
├── template_set/ (all .docx)
└── README.md
```

### Enterprise Features
- **API mode:** `excel2word --config project1 --format json`
- **Watch mode:** Monitor `Data/` → auto-process new files
- **Scheduler:** Cron/Task Scheduler integration

---

## Phase 10: Flexible - Scan Theo Giai Đoạn (3 tuần) ⭐⭐⭐

### Workflow Scan Từng Phần
1. **Step 1: Scan Excel First**
   ```text
   [Scan Data/] → 3 Excel found
   ☑ data_hd.xlsx (Text=50rows, Table=20rows)
   ☑ data_bb.xlsx (Info=30rows, KQ=15rows)
   [Select] [Next → Build Workflow]
   ```
2. **Step 2: Build Quy trình từ Excel**
   ```text
   Selected: data_hd.xlsx
   Auto-analyze:
   ✅ Text sheet: HoTen,SoHD → Merge source
   ✅ Table sheet: STT,TB,SL → Table source
   ⚠ DM sheet: Lookup? [Mark JOIN]
   [Confirm] [Next → Templates]
   ```
3. **Step 3: Scan Template (Cuối cùng)**
   ```text
   Scan Template/ for data_hd.xlsx...
   ⭐ template_muaban.docx [100% match]
   ⭐ template_phuluc.docx  [90% match]
   [Select Templates] [Preview Matrix] [Run]
   ```

### Lợi ích
- **Nhanh 10x:** Chỉ scan template cần thiết
- **User-driven:** Excel → workflow → template (step-by-step)
- **Low memory:** Không load toàn bộ cùng lúc
- **Clear feedback:** Progress từng bước

### GUI Flow Hoàn chỉnh
```text
🎯 STEP 1: Data (Excel scan)
🎯 STEP 2: Workflow (analyze sheets)
🎯 STEP 3: Templates (match scan)
🎯 STEP 4: Preview (m×n matrix)
🎯 STEP 5: Run (batch processing)
```

---

## Tổng kết Timeline & Priority

| Phase | Focus | Duration | Priority | Cumulative |
| :--- | :--- | :--- | :--- | :--- |
| 1 | MVP m→m files | 2w | ⭐⭐⭐⭐⭐ | 2w |
| 2 | Multi-file m×n | 2w | ⭐⭐⭐⭐⭐ | 4w |
| 3 | JOIN & Query | 2w | ⭐⭐⭐⭐⭐ | 6w |
| 4 | Zero hard-code | 1w | ⭐⭐⭐⭐ | 7w |
| 5 | Advanced m×n | 1w | ⭐⭐⭐ | 8w |
| 6 | Pro GUI | 3w | ⭐⭐⭐ | 11w |
| 7 | Expression | 2w | ⭐ | 13w |
| 8 | Production | 2w | ⭐⭐⭐ | 15w |
| 9 | Deploy/Share | 1w | ⭐ | 16w |
| 10 | Flexible Scan | 3w | ⭐⭐⭐ | 19w |

> **Khuyến nghị:** Phase 1-3 (6 tuần) đáp ứng 80% nhu cầu procurement/legal thực tế.

### Use Cases Thực tế (Procurement/Legal)
- **Hợp đồng thiết bị y tế (Phase 2):** 50 đơn vị × 3 files (HD+PL+BB) = 150 files
- **Biên bản nghiệm thu (Phase 1):** 30 thiết bị → 30 biên bản riêng
- **Đơn khởi kiện (Phase 3):** Chỉ Approved + filter theo giá trị → selective output
- **Báo cáo thử nghiệm (Phase 7):** Computed Total = Gia×SL, DateVN format

