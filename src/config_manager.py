import yaml
import os
from pathlib import Path

class ConfigManager:
    def __init__(self, config_path):
        self.config_path = Path(config_path).resolve()
        self.project_root = self.config_path.parent
        self.config = {}

    def load(self):
        """Load YAML config and resolve paths."""
        if not self.config_path.exists():
            raise FileNotFoundError(f"Config file not found: {self.config_path}")

        with open(self.config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)
        
        return self.config

    def validate(self):
        """Validate config structure and paths."""
        required_sections = ['input', 'template', 'output']
        for section in required_sections:
            if section not in self.config:
                raise ValueError(f"Missing required config section: '{section}'")

        # Validate Input
        input_config = self.config['input']
        if 'excel_file' not in input_config:
            raise ValueError("Missing 'excel_file' in input config")
        
        excel_path = self._resolve_path(input_config['excel_file'])
        if not excel_path.exists():
            raise FileNotFoundError(f"Input Excel file not found: {excel_path}")
        
        # Validate Template
        template_config = self.config['template']
        if 'file' not in template_config:
            raise ValueError("Missing 'file' in template config")
            
        template_path = self._resolve_path(template_config['file'])
        if not template_path.exists():
            raise FileNotFoundError(f"Template file not found: {template_path}")

        # Ensure Output Directory exists
        output_config = self.config['output']
        if 'folder' not in output_config:
            raise ValueError("Missing 'folder' in output config")
            
        output_path = self._resolve_path(output_config['folder'])
        output_path.mkdir(parents=True, exist_ok=True)

        return True

    def _resolve_path(self, relative_path):
        """Resolve path relative to config file location."""
        return (self.project_root / relative_path).resolve()

    def get_input_config(self):
        return self.config.get('input', {})

    def get_template_config(self):
        return self.config.get('template', {})

    def get_output_config(self):
        return self.config.get('output', {})
        
    def get_excel_path(self):
        return self._resolve_path(self.config['input']['excel_file'])
        
    def get_template_path(self):
        return self._resolve_path(self.config['template']['file'])
        
    def get_output_path(self):
        return self._resolve_path(self.config['output']['folder'])

