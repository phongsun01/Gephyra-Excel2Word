import sys
import os
from pathlib import Path

# Add project root to path
sys.path.append(str(Path(__file__).parent.parent))

from src.batch_processor import BatchProcessor
from src.config_manager import ConfigManager

def test_integration():
    print("Starting Integration Test (pywin32)...")
    
    # Setup paths
    base_dir = Path("examples/Project_Sample")
    config_path = base_dir / "config.yaml"
    
    if not config_path.exists():
        print(f"❌ Config not found: {config_path}")
        return

    try:
        # Initialize Config
        config_manager = ConfigManager(config_path)
        config_manager.load()
        config_manager.validate()
        print("✅ Config loaded and validated")
        
        # Initialize Processor
        processor = BatchProcessor(config_manager)
        print("✅ BatchProcessor initialized")
        
        # Run Process
        print("🚀 Running process_all()...")
        success, errors = processor.process_all()
        
        print(f"✅ Processing complete. Success: {success}, Errors: {errors}")
        
        if errors > 0:
            print("⚠️ Some files failed processing.")
            
    except Exception as e:
        print(f"❌ Integration Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_integration()
