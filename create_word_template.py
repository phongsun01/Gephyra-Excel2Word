"""
Create a proper Word template using COM automation.
This ensures placeholders are not fragmented in the XML.
"""
try:
    import win32com.client
    from pathlib import Path
    
    print("Creating Word application...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    print("Creating new document...")
    doc = word.Documents.Add()
    
    # Add content
    selection = word.Selection
    
    # Title
    selection.Font.Size = 16
    selection.Font.Bold = True
    selection.TypeText("HOP DONG MAU")
    selection.TypeParagraph()
    
    # Reset formatting
    selection.Font.Size = 11
    selection.Font.Bold = False
    
    selection.TypeText("Cong Hoa Xa Hoi Chu Nghia Viet Nam")
    selection.TypeParagraph()
    selection.TypeText("Doc lap - Tu do - Hanh phuc")
    selection.TypeParagraph()
    selection.TypeText("---")
    selection.TypeParagraph()
    selection.TypeParagraph()
    
    # Placeholders - type as single string to avoid fragmentation
    selection.TypeText("So Hop Dong: <<SoHD>>")
    selection.TypeParagraph()
    selection.TypeText("Ben A: Cong Ty ABC")
    selection.TypeParagraph()
    selection.TypeText("Ben B: <<HoTen>>")
    selection.TypeParagraph()
    selection.TypeText("Dia Chi: <<DiaChi>>")
    selection.TypeParagraph()
    
    # Save
    output_path = Path("examples/Project_Sample/Template/template_manual.docx").absolute()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    print(f"Saving to: {output_path}")
    doc.SaveAs(str(output_path))
    doc.Close()
    word.Quit()
    
    print(f"✅ SUCCESS! Created: {output_path}")
    print("Update config.yaml to use 'Template/template_manual.docx'")
    
except ImportError:
    print("❌ pywin32 not installed.")
    print("Install with: pip install pywin32")
    print("\nAlternatively, create the template manually in Microsoft Word:")
    print("1. Open Word")
    print("2. Type the content with placeholders like <<HoTen>>")
    print("3. Save as template_manual.docx")
except Exception as e:
    print(f"❌ Error: {e}")
    import traceback
    traceback.print_exc()
