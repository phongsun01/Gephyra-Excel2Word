
from docx import Document
from pathlib import Path
import sys

def check_output():
    output_path = Path("tests/output/test_text.docx")
    if not output_path.exists():
        print(f"File not found: {output_path}")
        return

    try:
        doc = Document(output_path)
        full_text = []
        for para in doc.paragraphs:
            full_text.append(para.text)
        
        text = "\n".join(full_text)
        print("--- Document Content ---")
        print(text)
        print("------------------------")
        
        if "TEST_CONTRACT_001" in text:
            print("✅ SUCCESS: Found validation text 'TEST_CONTRACT_001'")
        else:
            print("❌ FAILURE: Validation text NOT found")
            
        if "<<SoHD>>" in text:
            print("❌ FAILURE: Placeholder <<SoHD>> still exists")
        else:
            print("✅ SUCCESS: Placeholder <<SoHD>> is gone")

    except Exception as e:
        print(f"Error reading docx: {e}")

if __name__ == "__main__":
    check_output()
