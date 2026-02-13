
import win32com.client
from pathlib import Path
import os

def repro_replace():
    base_dir = Path("examples/Project_Sample")
    template_path = base_dir / "Template" / "template_manual.docx"
    output_path = Path("repro_output.docx")
    
    print(f"Template exists: {template_path.exists()}")
    
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True # Show to debug
    
    try:
        doc = word.Documents.Open(str(template_path.absolute()), ReadOnly=True)
        
        # Method 1: Standard Find/Replace
        print("Attempting validation replacement for <<SoHD>>...")
        find = doc.Content.Find
        find.ClearFormatting()
        success = find.Execute(
            FindText="<<SoHD>>",
            ReplaceWith="TEST_SUCCESS",
            Replace=2
        )
        print(f"Method 1 Result: {success}")
        
        # Method 2: Wildcard attempt if standard fails
        if not success:
            print("Attempting Wildcard replacement for <<*>>...")
            find.ClearFormatting()
            # Wildcard for <<...>>
            # \< and \> escape the brackets
            # But Word wildcards use different syntax. < and > are word boundaries.
            # Real brackets are \[ and \]. Wait, no.
            # Word Wildcards: 
            # ? = any char
            # * = any string
            # < = beginning of word
            # > = end of word
            # [ ] = one of chars
            # {n,m} = count
            # @ = 1 or more
            # Escape char is \
            
            # So <<SoHD>> would be \<\<SoHD\>\> ? No, < and > are special.
            # Let's try finding just "SoHD" first to see if it exists contiguously.
            
            find.Execute(FindText="SoHD")
            if find.Found:
               print("Found 'SoHD' text!")
            else:
               print("Could not find 'SoHD' text even without brackets!")
               
            # Check for split runs
            # Iterate through characters/words? Too slow.
            
        doc.SaveAs(str(output_path.absolute()))
        doc.Close()
        
    except Exception as e:
        print(f"Error: {e}")
    finally:
        word.Quit()

if __name__ == "__main__":
    repro_replace()
