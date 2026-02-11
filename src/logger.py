import logging
import sys
from pathlib import Path

class Excel2WordLogger:
    def __init__(self, log_file=None, verbose=False):
        self.logger = logging.getLogger("Gephyra")
        self.logger.setLevel(logging.DEBUG if verbose else logging.INFO)
        
        # Formatter
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        
        # Console Handler
        ch = logging.StreamHandler(sys.stdout)
        ch.setLevel(logging.DEBUG if verbose else logging.INFO)
        ch.setFormatter(formatter)
        self.logger.addHandler(ch)
        
        # File Handler
        if log_file:
            fh = logging.FileHandler(log_file, encoding='utf-8')
            fh.setLevel(logging.DEBUG) # Always log debug to file
            fh.setFormatter(formatter)
            self.logger.addHandler(fh)
            
        self.stats = {
            'success': 0,
            'error': 0
        }

    def log_success(self, message):
        # Specific success log logic if needed
        self.logger.info(f"✅ {message}")
        self.stats['success'] += 1

    def log_error(self, message):
        self.logger.error(f"❌ {message}")
        self.stats['error'] += 1
        
    def log_info(self, message):
        self.logger.info(message)

    def get_summary(self):
        return self.stats

