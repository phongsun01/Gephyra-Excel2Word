import sys
print("DEBUG: Starting main.py")
import argparse
from src.cli import main as cli_main
from src.gui import main as gui_main

def main():
    """Launcher for Gephyra."""
    parser = argparse.ArgumentParser(description="Gephyra - Batch Excel to Word Generator")
    parser.add_argument("--gui", choices=['tk', 'flet'], default='tk', help="Choose GUI framework (tk or flet)")
    parser.add_argument("--version", action="store_true", help="Show version and exit")
    parser.add_argument("--check-imports", action="store_true", help="Check required imports and exit")
    parser.add_argument("--generate-sample", action="store_true", help="Generate sample config and exit")
    parser.add_argument('--config', help='Path to config file for CLI mode')
    
    args = parser.parse_args()

    from src import __version__

    # 1. Show Version
    if args.version:
        print(f"Gephyra-Excel2Word v{__version__}")
        return

    # 2. Launch GUI if no args (or explicit flag)
    # Check if any action args are present
    action_args = [args.check_imports, args.generate_sample, args.config] # Added args.config to action_args
    if not any(action_args):
        if args.gui == 'flet':
            try:
                import flet as ft
                from src.flet_gui import FletGUI
                print("Launching Flet GUI...")
                def main_flet(page: ft.Page):
                    app = FletGUI(page)
                ft.app(target=main_flet)
            except ImportError:
                print("Error: flet not installed. Run 'pip install flet'")
            return
        else:
            # Default to Tkinter
            # from src.gui import main as gui_main # Already imported at top
            print("Launching GUI...")
            gui_main()
            return
    
    # 3. Handle other CLI actions
    if args.check_imports:
        # Placeholder for import check logic
        print("Checking imports...")
        # Add actual import check logic here
        return
    
    if args.generate_sample:
        # Placeholder for sample generation logic
        print("Generating sample config...")
        # Add actual sample generation logic here
        return

    # 4. If --config is provided, run CLI mode
    if args.config:
        print(f"Launching CLI mode with config: {args.config}")
        cli_main(config_path=args.config) # Assuming cli_main can take config_path

if __name__ == "__main__":
    main()
