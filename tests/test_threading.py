import sys
import threading
import time
from pathlib import Path

# Add project root to path
sys.path.append(str(Path(__file__).parent.parent))

from src.batch_processor import BatchProcessor
from src.config_manager import ConfigManager

def run_batch_in_thread():
    print(f"[Thread-{threading.get_ident()}] Starting Batch Process...")
    try:
        base_dir = Path("examples/Project_Sample")
        config_path = base_dir / "config.yaml"
        
        config_manager = ConfigManager(config_path)
        config_manager.load()
        config_manager.validate()
        
        processor = BatchProcessor(config_manager)
        success, errors = processor.process_all()
        print(f"[Thread-{threading.get_ident()}] Finished. Success: {success}, Errors: {errors}")
    except Exception as e:
        print(f"[Thread-{threading.get_ident()}] Error: {e}")
        import traceback
        traceback.print_exc()

def test_threading():
    print("🚀 Starting Threading Test...")
    
    # Run in a separate thread like Flet does
    t = threading.Thread(target=run_batch_in_thread)
    t.start()
    t.join()
    
    print("✅ Threading Test Complete")

if __name__ == "__main__":
    test_threading()
