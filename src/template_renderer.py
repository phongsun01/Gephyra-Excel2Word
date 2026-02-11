from pathlib import Path
import zipfile
import re
from io import BytesIO

class TemplateRenderer:
    def __init__(self, template_path):
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found: {self.template_path}")
        
        self.delimiter_start = "<<"
        self.delimiter_end = ">>"
        self.rendered_content = None
        
    def configure_delimiters(self, start="<<", end=">>"):
        """Configure custom delimiters for find-replace."""
        self.delimiter_start = start
        self.delimiter_end = end

    def render(self, context):
        """Render template by run-aware XML manipulation to handle fragmented text."""
        try:
            import re
            from lxml import etree
            
            # Read the .docx file (which is a ZIP archive)
            with zipfile.ZipFile(self.template_path, 'r') as zip_ref:
                # Read document.xml
                xml_bytes = zip_ref.read('word/document.xml')
                
                # Parse XML with lxml to handle runs properly
                root = etree.fromstring(xml_bytes)
                
                # Define namespace
                namespaces = {
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                }
                
                # Find all paragraphs
                for paragraph in root.findall('.//w:p', namespaces):
                    # Get all text runs in this paragraph
                    runs = paragraph.findall('.//w:r', namespaces)
                    
                    # Concatenate all text from runs to check for placeholders
                    full_text = ''
                    for run in runs:
                        t_elem = run.find('.//w:t', namespaces)
                        if t_elem is not None and t_elem.text:
                            full_text += t_elem.text
                    
                    # Check if any placeholder exists in this paragraph
                    for key, value in context.items():
                        placeholder = f"{self.delimiter_start}{key}{self.delimiter_end}"
                        
                        if placeholder in full_text:
                            # Replace the placeholder in the concatenated text
                            new_text = full_text.replace(placeholder, str(value))
                            
                            # Clear all existing text runs
                            for run in runs:
                                t_elem = run.find('.//w:t', namespaces)
                                if t_elem is not None:
                                    run.remove(t_elem)
                            
                            # Put the new text in the first run only
                            if runs:
                                first_run = runs[0]
                                new_t = etree.SubElement(first_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                                new_t.text = new_text
                                # Preserve spaces
                                new_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                            
                            # Update full_text for next iteration
                            full_text = new_text
                
                # Convert back to bytes
                xml_content = etree.tostring(root, encoding='utf-8', xml_declaration=True)
                
                # Store modified content
                self.rendered_content = BytesIO()
                with zipfile.ZipFile(self.rendered_content, 'w', zipfile.ZIP_DEFLATED) as new_zip:
                    for item in zip_ref.namelist():
                        if item == 'word/document.xml':
                            new_zip.writestr(item, xml_content)
                        else:
                            new_zip.writestr(item, zip_ref.read(item))
                
        except Exception as e:
            raise RuntimeError(f"Error rendering template: {str(e)}")


    def save(self, output_path):
        """Save rendered document."""
        try:
            if self.rendered_content is None:
                raise RuntimeError("No rendered content to save. Call render() first.")
            
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(output_path, 'wb') as f:
                f.write(self.rendered_content.getvalue())
                
        except Exception as e:
            raise IOError(f"Error saving file to {output_path}: {str(e)}")
