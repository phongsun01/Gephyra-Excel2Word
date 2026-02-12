import re
from pathlib import Path
from tqdm import tqdm
from .excel_loader import ExcelLoader
from .template_renderer_pywin32 import TemplateRenderer

class BatchProcessor:
    def __init__(self, config_manager, logger=None):
        self.config_manager = config_manager
        self.logger = logger
        self.input_config = config_manager.get_input_config()
        self.output_config = config_manager.get_output_config()
        self.template_config = config_manager.get_template_config()

    def process_all(self):
        """Orchestrate the batch processing."""
        # 1. Load Data
        excel_path = self.config_manager.get_excel_path()
        loader = ExcelLoader(excel_path)
        
        # Load main text data
        text_sheet_name = self.input_config['text_sheet']['name']
        required_cols = self.input_config['text_sheet'].get('required_columns', [])
        rows = loader.load_text_sheet(text_sheet_name, required_cols)
        
        # Load table config (optional)
        table_mappings = []
        if 'config_sheet' in self.input_config:
            config_sheet_name = self.input_config['config_sheet']['name']
            try:
                table_mappings = loader.load_config_sheet(config_sheet_name)
            except ValueError:
                if self.logger:
                    self.logger.log_error(f"Config sheet '{config_sheet_name}' not found or invalid. Proceeding without tables.")
        
        # Pre-load tables (optimization: load once if static, or per row? 
        # Usually tables in mail merge are static per run, or dynamic per row?
        # Requirement says: "Copy range Excel... mapping defined in Excel sheet"
        # If range is static A1:D20, it's the SAME table for all files?? 
        # Or does the range change? 
        # "1 Excel -> m Word files... Table paste... mappings defined in Excel sheet"
        # Usually for mail merge, tables are static reference data (like Price List).
        # OR they are dynamic. 
        # Docxtpl context: 'table_data' key.
        # Let's assume static range for now as per "BangTB | Table | A1:D20" example.
        
        template_path = self.config_manager.get_template_path()
        renderer = TemplateRenderer(template_path)
        # Note: New renderer currently supports <<key>> only. 
        # Delimiter config from yaml is ignored for now to ensure reliability.
        
        try:
            # 3. Process Rows
            success_count = 0
            error_count = 0
            
            output_dir = self.config_manager.get_output_path()
            
            # Use tqdm for progress bar if available and generic usage
            iterator = tqdm(rows, desc="Processing files") if rows else []
            
            for i, row in enumerate(iterator):
                try:
                    # Prepare Context (Text only)
                    context = row.copy()
                    
                    # Render with table support
                    renderer.render(
                        text_context=context,
                        excel_path=str(excel_path),
                        table_mappings=table_mappings
                    )
                    
                    # Generate Filename
                    output_filename = self.generate_filename(template_path, row)
                    output_file = output_dir / output_filename
                    
                    # Check overwrite
                    if output_file.exists() and not self.output_config.get('overwrite', False):
                        if self.logger:
                            self.logger.log_error(f"Skipped existing file: {output_filename}")
                        continue
                    
                    # Save
                    renderer.save(output_file)
                    success_count += 1
                    
                    if self.logger:
                        self.logger.log_success(f"Generated: {output_filename}")
                        
                except Exception as e:
                    error_count += 1
                    if self.logger:
                        self.logger.log_error(f"Error row {i+1}: {str(e)}")
                    
                    # Force cleanup on error to be safe? 
                    # Actually render() handles doc close on error.
                    
                    if self.config_manager.config.get('processing', {}).get('fail_fast', False):
                        raise e
        finally:
            # Ensure Word/Excel close after batch
            renderer.cleanup()

        return success_count, error_count

    def generate_filename(self, template_path, row):
        """Generate output filename based on pattern."""
        # 1. Determine Prefix
        prefix = ""
        if self.output_config.get('template_prefix_mode', True):
            prefix = template_path.stem.replace("-template", "") + "-"
        
        # 2. Apply Pattern
        pattern = self.output_config.get('filename_pattern', '{ID}')
        try:
            filename_body = pattern.format(**row)
        except KeyError as e:
            raise ValueError(f"Filename pattern contains unknown key: {e}")
            
        # 3. Sanitize
        if self.output_config.get('sanitize_filename', True):
            filename_body = self._sanitize_filename(filename_body)
            
        # 4. Suffix
        suffix = self.output_config.get('suffix', '.docx')
        
        return f"{prefix}{filename_body}{suffix}"

    def _sanitize_filename(self, filename):
        """Remove invalid characters from filename."""
        # Remove invalid chars: / \ : * ? " < > |
        return re.sub(r'[/\\:*?"<>|]', '', filename)

