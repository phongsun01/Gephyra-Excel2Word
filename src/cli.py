import argparse
import sys
import traceback
from .config_manager import ConfigManager
from .batch_processor import BatchProcessor
from .logger import Excel2WordLogger

def parse_arguments():
    parser = argparse.ArgumentParser(description="Gephyra CLI")
    parser.add_argument('--config', required=True, help='Path to config.yaml')
    parser.add_argument('--dry-run', action='store_true', help='Validate without generating files')
    parser.add_argument('--verbose', action='store_true', help='Enable verbose logging')
    return parser.parse_args()

def main():
    args = parse_arguments()
    
    # Init Logger
    logger = Excel2WordLogger(log_file="gephyra.log", verbose=args.verbose)
    logger.log_info(f"Starting Gephyra with config: {args.config}")
    
    try:
        # Load Config
        config_mgr = ConfigManager(args.config)
        config_mgr.load()
        config_mgr.validate()
        
        # Init Processor
        processor = BatchProcessor(config_mgr, logger)
        
        if args.dry_run:
            logger.log_info("Dry-run mode: Configuration valid. Skipping processing.")
            return

        # Run
        success, errors = processor.process_all()
        
        # Summary
        logger.log_info("="*30)
        logger.log_info(f"Processing Complete.")
        logger.log_info(f"✅ Success: {success}")
        logger.log_info(f"❌ Errors: {errors}")
        logger.log_info("="*30)
        
        if errors > 0:
            sys.exit(1)
            
    except Exception as e:
        logger.log_error(f"Critical Error: {str(e)}")
        if args.verbose:
            traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()

