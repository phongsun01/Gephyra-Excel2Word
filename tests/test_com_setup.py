import win32com.client
import sys
import os

def check_com():
    print("Checking COM availability...")
    
    # Check Word
    try:
        print("Attempting to dispatch Word.Application...")
        word = win32com.client.Dispatch("Word.Application")
        print(f"✅ Word Version: {word.Version}")
        word.Quit()
    except Exception as e:
        print(f"❌ Word Error: {e}")
        sys.exit(1)

    # Check Excel
    try:
        print("Attempting to dispatch Excel.Application...")
        excel = win32com.client.Dispatch("Excel.Application")
        print(f"✅ Excel Version: {excel.Version}")
        excel.Quit()
    except Exception as e:
        print(f"❌ Excel Error: {e}")
        sys.exit(1)

    print("✅ COM Check Successful!")

if __name__ == "__main__":
    check_com()
