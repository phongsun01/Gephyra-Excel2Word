import flet as ft
import threading
import queue
import logging
import sys
import traceback
from pathlib import Path

# Debug logging setup
def log_debug(msg):
    try:
        with open("flet_debug.log", "a") as f:
            f.write(msg + "\n")
    except:
        pass

try:
    # Handle relative imports for script vs module
    try:
        from .config_manager import ConfigManager
        from .batch_processor import BatchProcessor
    except ImportError:
        # Fallback for running as script from root
        from src.config_manager import ConfigManager
        from src.batch_processor import BatchProcessor

except Exception as e:
    log_debug(f"Import Error: {e}")
    # Don't exit yet, let FletGUI fail later if needed

# --- Design Constants ---
COLOR_PRIMARY = "#2563EB"
COLOR_SECONDARY = "#3B82F6"
COLOR_CTA = "#F97316"
COLOR_BG = "#F8FAFC"
COLOR_SURFACE = "#FFFFFF"
COLOR_TEXT_MAIN = "#0F172A"
COLOR_TEXT_MUTED = "#64748B"
COLOR_BORDER = "#E2E8F0"
COLOR_SUCCESS = "#10B981"
COLOR_ERROR = "#EF4444"

class FletGUI:
    def __init__(self, page: ft.Page):
        self.page = page
        self.page.title = "Gephyra - Pro (Flet)"
        self.page.theme_mode = ft.ThemeMode.LIGHT
        self.page.padding = 0
        self.page.window_width = 1000
        self.page.window_height = 700
        self.page.bgcolor = COLOR_BG

        # Application State
        self.config_manager = None
        self.is_processing = False
        self.log_queue = queue.Queue()
        
        # UI Components
        self.sidebar = None
        self.content_area = None
        self.log_view = None
        self.data_table = None
        self.btn_run = None
        self.lbl_config_status = None
        self.lbl_project = None

        self._build_layout()
        self._start_log_monitor()

    def _build_layout(self):
        # Sidebar
        self.lbl_config_status = ft.Text("No config loaded", color=COLOR_TEXT_MUTED, size=12)
        
        # Config Path Input
        self.txt_config_path = ft.TextField(
            label="Config Path", 
            text_size=12,
            height=40,
            content_padding=10,
            width=180
        )

        def browse_file(e):
            try:
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes('-topmost', True)
                file_path = filedialog.askopenfilename(
                    title="Select Config File",
                    filetypes=[("YAML files", "*.yaml *.yml")]
                )
                root.destroy()
                if file_path:
                    self.txt_config_path.value = file_path
                    self.txt_config_path.update()
                    self.load_config(file_path)
            except Exception as ex:
                self.log_message(f"Browse Error: {ex}")
                self.page.update()

        btn_browse = ft.ElevatedButton(
            content=ft.Text("...", size=16, weight=ft.FontWeight.BOLD),
            on_click=browse_file,
            tooltip="Browse Config",
            width=50,
            style=ft.ButtonStyle(
                padding=0,
                shape=ft.RoundedRectangleBorder(radius=5),
            )
        )
        
        load_config_btn = ft.OutlinedButton(
            content=ft.Text("Load Config"),
            on_click=lambda e: self.load_config(self.txt_config_path.value),
            style=ft.ButtonStyle(
                color=COLOR_PRIMARY,
                side=ft.BorderSide(1, COLOR_PRIMARY)
            )
        )
        
        # Checkbox for clearing output
        self.chk_clear_output = ft.Checkbox(
            label="Xóa thư mục output trước khi chạy",
            value=False,
            label_style=ft.TextStyle(size=12)
        )
        
        self.btn_run = ft.ElevatedButton(
            content=ft.Text("GENERATE FILES", color="white"),
            bgcolor=COLOR_CTA,
            on_click=self.run_generation,
            disabled=True,
            style=ft.ButtonStyle(shape=ft.RoundedRectangleBorder(radius=8))
        )

        self.sidebar = ft.Container(
            width=280,
            bgcolor=COLOR_SURFACE,
            padding=20,
            content=ft.Column(
                controls=[
                    ft.Text("Gephyra", size=24, weight=ft.FontWeight.BOLD, color=COLOR_TEXT_MAIN),
                    ft.Text("Batch Generator", size=12, color=COLOR_TEXT_MUTED),
                    ft.Divider(height=20, color="transparent"),
                    
                    ft.Text("Configuration", weight=ft.FontWeight.BOLD),
                    ft.Container(
                        content=ft.Column([
                            ft.Container(
                                content=ft.Row([self.txt_config_path, btn_browse], spacing=5),
                                height=50
                            ),
                            load_config_btn,
                            self.lbl_config_status,
                            ft.Divider(height=10, color="transparent"),
                            self.chk_clear_output
                        ]),
                        padding=10,
                        border=ft.border.all(1, COLOR_BORDER),
                        border_radius=8
                    ),
                    
                    ft.Divider(height=20, color="transparent"),
                    
                    ft.Text("Actions", weight=ft.FontWeight.BOLD),
                    ft.Container(
                        content=ft.Column([
                            self.btn_run
                        ]),
                        padding=10,
                        border=ft.border.all(1, COLOR_BORDER),
                        border_radius=8
                    ),
                    
                    ft.Divider(),
                    ft.Text("v1.0.0", size=10, color=COLOR_TEXT_MUTED)
                ]
            )
        )

        # Content Area
        self.lbl_project = ft.Text("Project: --", size=24, weight=ft.FontWeight.BOLD, color=COLOR_TEXT_MAIN)
        
        # Data Table (Preview)
        self.data_table = ft.DataTable(
            columns=[ft.DataColumn(ft.Text("ID")), ft.DataColumn(ft.Text("Info"))],
            rows=[],
            border=ft.border.all(1, COLOR_BORDER),
            vertical_lines=ft.border.BorderSide(1, COLOR_BORDER),
            horizontal_lines=ft.border.BorderSide(1, COLOR_BORDER),
            heading_row_color=ft.Colors.GREY_100,
        )
        
        preview_container = ft.Container(
            content=ft.Column([
                ft.Text("Data Preview (First 50 rows)", weight=ft.FontWeight.BOLD),
                ft.Container(
                    content=ft.ListView(
                        controls=[self.data_table],
                        height=300,
                    ),
                    border=ft.border.all(1, COLOR_BORDER),
                    border_radius=8,
                    padding=0 
                )
            ]),
            expand=True
        )

        # Log Console
        self.log_view = ft.ListView(
            expand=True,
            spacing=2,
            auto_scroll=True,
            height=150
        )
        
        log_container = ft.Container(
            content=ft.Column([
                ft.Text("Execution Log", weight=ft.FontWeight.BOLD),
                ft.Container(
                    content=self.log_view,
                    bgcolor="white",
                    border=ft.border.all(1, COLOR_BORDER),
                    border_radius=8,
                    padding=10,
                    height=150
                )
            ])
        )

        self.content_area = ft.Container(
            padding=20,
            expand=True,
            content=ft.Column([
                self.lbl_project,
                ft.Divider(height=20, color="transparent"),
                preview_container,
                ft.Divider(height=20, color="transparent"),
                log_container
            ])
        )

        # Main Layout
        self.page.add(
            ft.Row(
                controls=[self.sidebar, ft.VerticalDivider(width=1, color=COLOR_BORDER), self.content_area],
                expand=True,
                spacing=0
            )
        )

    def load_config(self, file_path):
        if not file_path:
            return
            
        try:
            self.config_manager = ConfigManager(file_path)
            self.config_manager.load()
            self.config_manager.validate()
            
            # Update UI
            project_name = self.config_manager.config.get('project', {}).get('name', 'Unknown Project')
            self.lbl_project.value = f"Project: {project_name}"
            self.lbl_config_status.value = f"Loaded: {Path(file_path).name}"
            self.lbl_config_status.color = COLOR_SUCCESS
            self.btn_run.disabled = False
            
            self.log_message(f"✅ Configuration loaded successfully: {file_path}")
            self.page.update()
            
            # Preview Data
            threading.Thread(target=self._preview_data, daemon=True).start()
            
        except Exception as ex:
            self.lbl_config_status.value = "Error loading config"
            self.lbl_config_status.color = COLOR_ERROR
            self.log_message(f"❌ Error loading config: {ex}")
            self.page.update()

    def _preview_data(self):
        try:
            from src.excel_loader import ExcelLoader
            
            excel_path = self.config_manager.get_excel_path()
            loader = ExcelLoader(excel_path)
            
            input_config = self.config_manager.get_input_config()
            sheet_name = input_config['text_sheet']['name']
            rows = loader.load_text_sheet(sheet_name)
            
            if not rows:
                self.log_message("⚠️ No data found in Excel.")
                return

            columns = list(rows[0].keys())
            
            new_columns = [ft.DataColumn(ft.Text(col)) for col in columns]
            
            new_rows = []
            for row in rows[:50]:
                cells = [ft.DataCell(ft.Text(str(row.get(col, "")))) for col in columns]
                new_rows.append(ft.DataRow(cells=cells))
            
            self.data_table.columns = new_columns
            self.data_table.rows = new_rows
            
            self.log_message(f"ℹ️ Loaded {len(rows)} rows from Excel.")
            self.page.update()
            
        except Exception as e:
            self.log_message(f"❌ Error previewing data: {e}")
            self.page.update()

    def run_generation(self, e):
        if self.is_processing:
            return
        
        if not self.config_manager:
            self.log_message("❌ Please load a configuration first.")
            return
        
        try:
            # Clear output folder if checkbox is checked
            if self.chk_clear_output.value:
                import shutil
                from pathlib import Path
                output_path = Path(self.config_manager.get_output_path())
                if output_path.exists():
                    shutil.rmtree(output_path)
                    output_path.mkdir(parents=True, exist_ok=True)
                    self.log_message(f"🗑️ Cleared output folder: {output_path}")
            
            self.is_processing = True
            self.btn_run.disabled = True
            self.btn_run.text = "PROCESSING..."
            self.log_message("🚀 Starting batch generation...")
            self.page.update()
        
        except Exception as ex:
            self.log_message(f"❌ Error during pre-generation steps: {ex}")
            self.is_processing = False
            self.btn_run.disabled = False
            self.btn_run.text = "GENERATE FILES"
            self.page.update()
            return
        
        gui_logger = GuiLoggerAdapter(self.log_queue)
        
        def task():
            try:
                processor = BatchProcessor(self.config_manager, gui_logger)
                success, errors = processor.process_all()
                self.log_queue.put(f"COMPLETE: Success={success}, Errors={errors}")
            except Exception as e:
                self.log_queue.put(f"CRITICAL ERROR: {str(e)}")
            finally:
                self.is_processing = False
        
        threading.Thread(target=task, daemon=True).start()

    def log_message(self, msg):
        self.log_view.controls.append(ft.Text(msg, font_family="Consolas", size=12))
        self.page.update()

    def _start_log_monitor(self):
        def monitor():
            while True:
                try:
                    msg = self.log_queue.get()
                    if msg.startswith("COMPLETE"):
                        self.btn_run.disabled = False
                        self.btn_run.text = "GENERATE FILES"
                        self.log_message(msg)
                    else:
                        self.log_message(msg)
                    self.page.update()
                except Exception:
                    pass
        
        threading.Thread(target=monitor, daemon=True).start()

class GuiLoggerAdapter:
    def __init__(self, queue):
        self.queue = queue
        
    def log_success(self, message):
        self.queue.put(f"✅ {message}")
        
    def log_error(self, message):
        self.queue.put(f"❌ {message}")
        
    def log_info(self, message):
        self.queue.put(f"ℹ️ {message}")

if __name__ == "__main__":
    try:
        def main(page: ft.Page):
            app = FletGUI(page)
        ft.app(target=main)
    except Exception as e:
        log_debug(f"Crash: {e}")
        import traceback
        traceback.print_exc()
