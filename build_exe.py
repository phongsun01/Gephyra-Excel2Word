import PyInstaller.__main__
from pathlib import Path
import shutil

# Clean build dirs
for d in ['build', 'dist']:
    if Path(d).exists():
        shutil.rmtree(d)

PyInstaller.__main__.run([
    'main.py',
    '--name=Gephyra',
    '--onefile',
    '--noconsole',
    '--clean',
    '--add-data=src;src',
    '--hidden-import=docxtpl',
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=yaml',
    '--hidden-import=tkinter',
    '--hidden-import=PIL.Image',
    '--hidden-import=PIL.ImageTk',
    # '--icon=app.ico', # Uncomment if icon exists
])
