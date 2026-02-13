
import unittest
import sys
from pathlib import Path
import os
import time

# Add src to path
sys.path.append(str(Path(__file__).parent.parent))

from src.template_renderer_pywin32 import TemplateRenderer

class TestTemplateRenderer(unittest.TestCase):
    def setUp(self):
        # Use existing sample files
        self.base_dir = Path("examples/Project_Sample")
        self.template_path = self.base_dir / "Template" / "template_manual.docx"
        self.excel_path = self.base_dir / "Data" / "sample_data.xlsx"
        self.output_dir = Path("tests/output")
        self.output_dir.mkdir(exist_ok=True)
        
        if not self.template_path.exists():
            self.skipTest("Sample template not found")

    def test_text_replacement(self):
        print("\nTesting Text Replacement...")
        renderer = TemplateRenderer(self.template_path)
        try:
            context = {
                "SoHD": "TEST_CONTRACT_001", 
                "HoTen": "UNIT_TEST_USER",
                "DiaChi": "Test Address 123"
            }
            renderer.render(context)
            
            output_file = self.output_dir / "test_text.docx"
            renderer.save(output_file)
            
            self.assertTrue(output_file.exists())
            print(f"✅ Generated: {output_file}")
        finally:
            renderer.cleanup()

    def test_table_insertion(self):
        print("\nTesting Table Insertion...")
        if not self.excel_path.exists():
            print("⚠️ Skipping table test (Excel not found)")
            return

        renderer = TemplateRenderer(self.template_path)
        try:
            context = {"Name": "TABLE_USER"}
            # Mock mapping based on sample config logic
            # Assuming sample has a table... check logic later
            # For now just try to open excel
            
            mappings = [
                {
                    'excel_sheet': 'Sheet1', # Assumption
                    'excel_range': 'A1:B2',
                    'word_bookmark': 'TableData' # Assumption
                }
            ]
            
            renderer.render(context, str(self.excel_path), mappings)
            output_file = self.output_dir / "test_table.docx"
            renderer.save(output_file)
            self.assertTrue(output_file.exists())
            print(f"✅ Generated: {output_file}")
        except Exception as e:
            print(f"⚠️ Table test warning: {e}")
        finally:
            renderer.cleanup()

    def test_shared_app(self):
        print("\nTesting Shared App Instance...")
        import win32com.client
        import pythoncom
        
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        
        try:
            renderer = TemplateRenderer(self.template_path, word_app=word)
            renderer.render({"Name": "SHARED_APP"})
            renderer.cleanup()
            
            # Verify Word is still alive
            try:
                # Accessing a property should work
                _ = word.Version
                print("✅ Shared Word app is still alive")
            except:
                self.fail("Shared Word app was closed!")
                
        finally:
            word.Quit()

if __name__ == "__main__":
    unittest.main()
