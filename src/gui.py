import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
from pathlib import Path

from .config_manager import ConfigManager
from .batch_processor import BatchProcessor
from .logger import Excel2WordLogger

# --- Design System Constants ---
# Colors
COLOR_PRIMARY = "#2563EB"    # Royal Blue
COLOR_SECONDARY = "#3B82F6"  # Blue-500
COLOR_CTA = "#F97316"        # Orange
COLOR_BG = "#F8FAFC"         # Slate-50
COLOR_SURFACE = "#FFFFFF"    # White
COLOR_TEXT_MAIN = "#0F172A"  # Slate-900
COLOR_TEXT_MUTED = "#64748B" # Slate-500
COLOR_BORDER = "#E2E8F0"     # Slate-200
COLOR_SUCCESS = "#10B981"    # Green
COLOR_ERROR = "#EF4444"      # Red

# Fonts
FONT_HEADING = ("Segoe UI", 12, "bold")
FONT_BODY = ("Segoe UI", 10)
FONT_SMALL = ("Segoe UI", 9)

class ModernGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Gephyra - Pro")
        self.root.geometry("1000x700")
        self.root.configure(bg=COLOR_BG)
        
        # Application State
        self.config_manager = None
        self.is_processing = False
        self.log_queue = queue.Queue()
        
        self._setup_styles()
        self._build_layout()
        self._start_log_monitor()

    def _setup_styles(self):
        """Apply custom styles matching ui-design-system.md"""
        style = ttk.Style()
        style.theme_use('clam') # Clean base for customization
        
        # General
        style.configure(".", background=COLOR_BG, foreground=COLOR_TEXT_MAIN, font=FONT_BODY)
        style.configure("TFrame", background=COLOR_BG)
        style.configure("Surface.TFrame", background=COLOR_SURFACE, relief="flat")
        
        # Buttons
        style.configure("Primary.TButton",
                        background=COLOR_PRIMARY,
                        foreground="white",
                        borderwidth=0,
                        focuscolor=COLOR_PRIMARY,
                        font=("Segoe UI", 10, "bold"))
        style.map("Primary.TButton", background=[("active", COLOR_SECONDARY)])
        
        style.configure("Secondary.TButton",
                        background="white",
                        foreground=COLOR_TEXT_MAIN,
                        borderwidth=1,
                        bordercolor=COLOR_BORDER,
                        focuscolor=COLOR_BORDER)
        style.map("Secondary.TButton", background=[("active", "#F1F5F9")])
        
        style.configure("CTA.TButton",
                        background=COLOR_CTA,
                        foreground="white",
                        font=("Segoe UI", 11, "bold"))
        style.map("CTA.TButton", 
            background=[("active", "#EA580C"), ("disabled", "#E2E8F0")], 
            foreground=[("disabled", "#94A3B8")]
        )

        # Labels
        style.configure("Heading.TLabel", font=FONT_HEADING, background=COLOR_BG, foreground=COLOR_TEXT_MAIN)
        style.configure("Muted.TLabel", font=FONT_SMALL, background=COLOR_BG, foreground=COLOR_TEXT_MUTED)
        style.configure("Status.TLabel", background="white", font=("Consolas", 9))

        # Treeview
        style.configure("Treeview", 
                        background="white",
                        fieldbackground="white",
                        foreground=COLOR_TEXT_MAIN,
                        rowheight=28,
                        font=FONT_BODY)
        style.configure("Treeview.Heading", 
                        background="#F1F5F9", 
                        foreground=COLOR_TEXT_MAIN, 
                        font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", "#E0F2FE")], foreground=[("selected", COLOR_PRIMARY)])

    def _build_layout(self):
        """Create the Main Window Layout (Sidebar + Content)."""
        # Main Container
        main_container = ttk.Frame(self.root)
        main_container.pack(fill="both", expand=True)
        
        # --- Sidebar (Left) ---
        sidebar = ttk.Frame(main_container, width=260, style="Surface.TFrame")
        sidebar.pack(side="left", fill="y", padx=(0, 1), pady=0)
        
        # Sidebar Content
        ttk.Label(sidebar, text="Gephyra", style="Heading.TLabel", background=COLOR_SURFACE).pack(pady=(24, 8), padx=20, anchor="w")
        ttk.Label(sidebar, text="Batch Generator", style="Muted.TLabel", background=COLOR_SURFACE).pack(pady=(0, 24), padx=20, anchor="w")
        
        # Config Section in Sidebar
        config_frame = ttk.LabelFrame(sidebar, text="Configuration", padding=16)
        config_frame.pack(fill="x", padx=16, pady=8)
        
        self.btn_load_config = ttk.Button(config_frame, text="Load Config (YAML)", command=self.load_config, style="Secondary.TButton")
        self.btn_load_config.pack(fill="x", pady=4)
        
        self.lbl_config_status = ttk.Label(config_frame, text="No config loaded", style="Muted.TLabel")
        self.lbl_config_status.pack(fill="x", pady=4)

        # Actions in Sidebar
        action_frame = ttk.LabelFrame(sidebar, text="Actions", padding=16)
        action_frame.pack(fill="x", padx=16, pady=16)
        
        self.btn_run = ttk.Button(action_frame, text="GENERATE FILES", command=self.run_generation, style="CTA.TButton", state="disabled")
        self.btn_run.pack(fill="x", pady=8)

        # Footer
        ttk.Label(sidebar, text="v1.0.0", style="Muted.TLabel", background=COLOR_SURFACE).pack(side="bottom", pady=16)

        # --- Content Area (Right) ---
        content = ttk.Frame(main_container, padding=24)
        content.pack(side="right", fill="both", expand=True)
        
        # Top Stats / Info
        info_frame = ttk.Frame(content)
        info_frame.pack(fill="x", pady=(0, 16))
        
        self.lbl_project = ttk.Label(info_frame, text="Project: --", font=("Segoe UI", 16))
        self.lbl_project.pack(anchor="w")
        
        # Data Preview (Treeview)
        preview_frame = ttk.LabelFrame(content, text="Data Preview (First 50 rows)", padding=10)
        preview_frame.pack(fill="both", expand=True, pady=8)
        
        self.tree = ttk.Treeview(preview_frame, columns=("ID", "Info"), show="headings")
        self.tree.heading("ID", text="ID")
        self.tree.heading("Info", text="Info")
        
        # Scrollbar for tree
        scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Log Console
        log_frame = ttk.LabelFrame(content, text="Execution Log", padding=10)
        log_frame.pack(fill="x", pady=(16, 0), ipady=40)
        
        self.log_text = tk.Text(log_frame, height=8, state="disabled", font=("Consolas", 9), bg="white", fg="#334155", relief="flat")
        self.log_text.pack(fill="both", expand=True)
        
    def load_config(self):
        """Open file dialog to load config."""
        file_path = filedialog.askopenfilename(filetypes=[("YAML Config", "*.yaml"), ("All Files", "*.*")])
        if not file_path:
            return
            
        try:
            self.config_manager = ConfigManager(file_path)
            self.config_manager.load()
            self.config_manager.validate()
            
            # Update UI
            project_name = self.config_manager.config.get('project', {}).get('name', 'Unknown Project')
            self.lbl_project.config(text=f"Project: {project_name}")
            self.lbl_config_status.config(text=f"Loaded: {Path(file_path).name}", foreground=COLOR_SUCCESS)
            self.btn_run.config(state="normal")
            
            self.log_message(f"✅ Configuration loaded successfully: {file_path}")
            
            # Preview Data (Async to avoid freeze)
            threading.Thread(target=self._preview_data, daemon=True).start()
            
        except Exception as e:
            self.lbl_config_status.config(text="Error loading config", foreground=COLOR_ERROR)
            messagebox.showerror("Config Error", str(e))
            self.log_message(f"❌ Error loading config: {e}")

    def _preview_data(self):
        """Load and show data in Treeview."""
        try:
            from .excel_loader import ExcelLoader
            
            excel_path = self.config_manager.get_excel_path()
            loader = ExcelLoader(excel_path)
            
            input_config = self.config_manager.get_input_config()
            sheet_name = input_config['text_sheet']['name']
            rows = loader.load_text_sheet(sheet_name)
            
            if not rows:
                self.log_message("⚠️ No data found in Excel.")
                return

            # Setup Columns dynamically
            columns = list(rows[0].keys())
            
            # Update Treeview in main thread
            def update_tree():
                self.tree["columns"] = columns
                for col in columns:
                    self.tree.heading(col, text=col)
                    self.tree.column(col, width=100)
                
                # Clear existing
                for i in self.tree.get_children():
                    self.tree.delete(i)
                
                # Insert rows (limit 50)
                for row in rows[:50]:
                    values = [row.get(col, "") for col in columns]
                    self.tree.insert("", "end", values=values)
                    
                self.log_message(f"ℹ️ Loaded {len(rows)} rows from Excel.")
            
            self.root.after(0, update_tree)
            
        except Exception as e:
            self.log_message(f"❌ Error previewing data: {e}")

    def run_generation(self):
        """Start batch generation in background thread."""
        if self.is_processing:
            return
            
        self.is_processing = True
        self.btn_run.config(state="disabled", text="PROCESSING...")
        self.log_message("🚀 Starting batch generation...")
        
        # Create a GUI Logger adapter
        gui_logger = GuiLoggerAdapter(self.log_queue)
        
        def task():
            try:
                processor = BatchProcessor(self.config_manager, gui_logger)
                success, errors = processor.process_all()
                self.log_queue.put(f"COMPLETE: Success={success}, Errors={errors}")
            except Exception as e:
                self.log_queue.put(f"CRITICAL ERROR: {str(e)}")
            finally:
                self.root.after(0, self._on_generation_complete)

        threading.Thread(target=task, daemon=True).start()

    def _on_generation_complete(self):
        """Reset UI after generation."""
        self.is_processing = False
        self.btn_run.config(state="normal", text="GENERATE FILES")
        messagebox.showinfo("Done", "Processing Complete! Check logs for details.")

    def log_message(self, msg):
        """Add message to log console."""
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"{msg}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _start_log_monitor(self):
        """Check queue for logs from background threads."""
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_message(msg)
        except queue.Empty:
            pass
        self.root.after(100, self._start_log_monitor)


class GuiLoggerAdapter:
    """Adapter to route Logger calls to GUI Queue."""
    def __init__(self, queue):
        self.queue = queue
        
    def log_success(self, message):
        self.queue.put(f"✅ {message}")
        
    def log_error(self, message):
        self.queue.put(f"❌ {message}")
        
    def log_info(self, message):
        self.queue.put(f"ℹ️ {message}")

def main():
    root = tk.Tk()
    # Optional: Set icon if available
    # root.iconbitmap("app.ico") 
    app = ModernGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

