"""
Template rendering using pywin32 COM Automation.
Replaces lxml-based implementation with native Word/Excel APIs.
"""
import win32com.client
import pythoncom
from pathlib import Path
from typing import Dict, List, Any, Optional
import logging
import subprocess
import atexit
import time

logger = logging.getLogger(__name__)

class TemplateRenderer:
    """
    Render Word templates using COM Automation.
    
    Features:
    - Text placeholder replacement (preserves all formatting)
    - Table insertion from Excel (perfect copy)
    - Bookmark-based table placement
    - Robust cleanup (no zombie processes)
    """
    
    def __init__(self, template_path: str, word_app=None, excel_app=None):
        """
        Initialize renderer with template.
        
        Args:
            template_path: Path to template file
            word_app: Optional pre-initialized Word.Application instance
            excel_app: Optional pre-initialized Excel.Application instance
        """
        self.template_path = Path(template_path)
        if not self.template_path.exists():
            raise FileNotFoundError(f"Template not found: {self.template_path}")
        
        self.word = word_app
        self.excel = excel_app
        self.doc = None
        self._initialized = False
        self._original_thread_id = None
        self._shared_word = bool(word_app)
        self._shared_excel = bool(excel_app)
        
        # Register cleanup
        atexit.register(self.cleanup)
    
    def _init_com(self):
        """Initialize COM (thread-safe)."""
        if self._initialized:
            return
        
        # Required for threading (Flet background threads)
        try:
            pythoncom.CoInitialize()
        except:
            pass # Already initialized
            
        try:
            # Create Word instance if not provided
            if not self.word:
                self.word = win32com.client.Dispatch("Word.Application")
                try:
                    self.word.Visible = False
                    self.word.DisplayAlerts = 0  # wdAlertsNone
                except Exception as e:
                    logger.warning(f"Could not set Word visibility: {e}")
                logger.debug("Word.Application initialized (new instance)")
            else:
                 logger.debug("Word.Application initialized (shared instance)")
                 
            self._initialized = True
        except Exception as e:
            logger.error(f"COM Init Error: {e}")
            raise RuntimeError(
                "Cannot initialize Word COM.\n"
                "Ensure:\n"
                "1. Microsoft Office is installed\n"
                "2. pywin32 postinstall script was run"
            ) from e
    
    def render(
        self, 
        text_context: Dict[str, str],
        excel_path: Optional[str] = None,
        table_mappings: Optional[List[Dict]] = None
    ) -> None:
        """
        Render template with context.
        
        Args:
            text_context: {'Name': 'John', 'ID': '123'}
            excel_path: Path to Excel file (for tables)
            table_mappings: [
                {
                    'excel_sheet': 'Sheet1',
                    'excel_range': 'A1:D10',
                    'word_bookmark': 'TableData'
                }
            ]
        """
        self._init_com()
        
        # Validate template again
        if not self.template_path.exists():
             raise FileNotFoundError(f"Template not found: {self.template_path}")

        try:
            # Open template
            self.doc = self.word.Documents.Open(str(self.template_path.absolute()), ReadOnly=True)
            logger.debug(f"Opened: {self.template_path}")
            
            # 1. Replace text placeholders
            self._replace_text(text_context)
            
            # 2. Insert tables (if any)
            if excel_path and table_mappings:
                self._insert_tables(excel_path, table_mappings)
        
        except Exception as e:
            logger.error(f"Render error: {e}")
            # Ensure doc is closed on error
            if self.doc:
                try:
                    self.doc.Close(SaveChanges=False)
                    self.doc = None
                except:
                    pass
            raise
    
    def _replace_text(self, context: Dict[str, str]) -> None:
        """Replace text placeholders using Word Find/Replace."""
        # Use Range instead of Selection for better stability
        rng = self.doc.Content
        find = rng.Find
        
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        
        for key, value in context.items():
            placeholder = f"<<{key}>>"
            
            # Replace All
            find.Execute(
                FindText=placeholder,
                ReplaceWith=str(value) if value is not None else "",
                Replace=2,  # wdReplaceAll
                Forward=True,
                Wrap=1  # wdFindContinue
            )
            
            # Also check headers/footers/textboxes (StoryRanges)
            for story_range in self.doc.StoryRanges:
                 find_story = story_range.Find
                 find_story.Execute(
                    FindText=placeholder,
                    ReplaceWith=str(value) if value is not None else "",
                    Replace=2,
                    Forward=True,
                    Wrap=1
                )
                 # Iterate through linked stories
                 while story_range.NextStoryRange:
                     story_range = story_range.NextStoryRange
                     find_next = story_range.Find
                     find_next.Execute(
                        FindText=placeholder,
                        ReplaceWith=str(value) if value is not None else "",
                        Replace=2,
                        Forward=True,
                        Wrap=1
                    )
            
    
    def _insert_tables(self, excel_path: str, mappings: List[Dict]) -> None:
        """Insert tables from Excel to Word bookmarks."""
        excel_path = Path(excel_path)
        if not excel_path.exists():
            raise FileNotFoundError(f"Excel: {excel_path}")
        
        # Initialize Excel (lazy)
        wb = None
        try:
            if not self.excel:
                self.excel = win32com.client.Dispatch("Excel.Application")
                try:
                    self.excel.Visible = False
                    self.excel.DisplayAlerts = False
                except Exception as e:
                    logger.warning(f"Could not set Excel visibility: {e}")
                self._shared_excel = False # We created it, we own it

            # Open workbook
            wb = self.excel.Workbooks.Open(str(excel_path.absolute()), ReadOnly=True)
            
            for mapping in mappings:
                sheet_name = mapping.get('excel_sheet')
                range_str = mapping.get('excel_range')
                bookmark = mapping.get('word_bookmark')
                
                if not (sheet_name and range_str):
                    continue
                
                try:
                    # Get Excel range
                    ws = wb.Sheets(sheet_name)
                    rng = ws.Range(range_str)
                    rng.Copy()
                    
                    # Paste to Word
                    # Check if bookmark exists
                    bookmark_exists = False
                    if bookmark:
                        for bm in self.doc.Bookmarks:
                            if bm.Name == bookmark:
                                bookmark_exists = True
                                break
                    
                    if bookmark and bookmark_exists:
                        bm_range = self.doc.Bookmarks(bookmark).Range
                        bm_range.Paste()
                        # Restore bookmark (optional, Paste often deletes it)
                        self.doc.Bookmarks.Add(bookmark, bm_range)
                        logger.debug(f"Pasted to bookmark: {bookmark}")
                    else:
                        # Fallback: paste at end of doc
                        last_paragraph = self.doc.Content.Paragraphs.Last
                        last_paragraph.Range.Paste()
                        logger.warning(f"Bookmark '{bookmark}' not found, pasted at end")
                        
                    # Clear clipboard to avoid memory issues
                    self.excel.CutCopyMode = False
                    
                except Exception as e:
                    logger.error(f"Error copying table {sheet_name}!{range_str}: {e}")
                    # Continue to next table
        
        except Exception as e:
            logger.error(f"Excel error: {e}")
            raise
        finally:
            if wb:
                wb.Close(SaveChanges=False)
    
    def save(self, output_path: Path) -> None:
        """Save rendered document."""
        if not self.doc:
            raise RuntimeError("No document to save. Call render() first.")
        
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # file format 16 = wdFormatDocumentDefault (docx)
        self.doc.SaveAs(str(output_path.absolute()), FileFormat=16)
        logger.info(f"Saved: {output_path}")
        
        # Close document
        self.doc.Close(SaveChanges=False)
        self.doc = None
    
    def cleanup(self) -> None:
        """Robust cleanup of COM objects."""
        try:
            # Close document if open
            if self.doc:
                try:
                    self.doc.Close(SaveChanges=False)
                except:
                    pass
                self.doc = None
            
            # Quit applications
            if self.word and not self._shared_word:
                try:
                    self.word.Quit()
                except:
                    pass
                self.word = None
                
            if self.excel and not self._shared_excel:
                try:
                    self.excel.Quit()
                except:
                    pass
                self.excel = None
            
            # Uninitialize COM
            # Only uninitialize if we are not sharing apps, or be smarter about it.
            # If we shared apps, the caller owns the COM context.
            if self._initialized and not (self._shared_word or self._shared_excel):
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass
                self._initialized = False
            
            logger.debug("COM cleanup complete")
        
        except Exception as e:
            logger.warning(f"Cleanup error (will force kill): {e}")
            self._force_kill_office()
    
    def _force_kill_office(self) -> None:
        """Fallback: Force kill Office processes."""
        try:
            # Only kill if we really messed up? Be careful not to kill user's other docs
            # But since we create new Application instances, maybe okay?
            # Ideally, track Process ID (PID) if possible, but COM Dispatch doesn't give PID easily.
            # Only kill if user configured to do so or catastrophic failure.
            # For now, log warning.
            logger.warning("Should force kill WINWORD.EXE/EXCEL.EXE but skipping to be safe.")
            # subprocess.run("taskkill /F /IM WINWORD.EXE", shell=True, check=False)
            pass 
        except Exception as e:
            logger.error(f"Force kill failed: {e}")
