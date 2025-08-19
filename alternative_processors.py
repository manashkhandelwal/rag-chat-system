"""
Alternative document processors that don't depend on textract
"""
import os
import subprocess
import tempfile
import zipfile
from pathlib import Path
import streamlit as st

# Try to import optional dependencies
try:
    import docx2txt
    DOCX2TXT_AVAILABLE = True
except ImportError:
    DOCX2TXT_AVAILABLE = False

try:
    from pdfminer.high_level import extract_text as pdf_extract_text
    from pdfminer.layout import LAParams
    PDFMINER_AVAILABLE = True
except ImportError:
    PDFMINER_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_AVAILABLE = True
except ImportError:
    PYMUPDF_AVAILABLE = False

class AlternativeDocumentProcessor:
    """Alternative document processor without textract dependency"""
    
    def __init__(self):
        self.available_processors = self._check_available_processors()
    
    def _check_available_processors(self):
        """Check which processors are available"""
        processors = {
            "docx2txt": DOCX2TXT_AVAILABLE,
            "pdfminer": PDFMINER_AVAILABLE,
            "pymupdf": PYMUPDF_AVAILABLE,
            "antiword": self._check_antiword(),
            "libreoffice": self._check_libreoffice()
        }
        return processors
    
    def _check_antiword(self):
        """Check if antiword is available (for DOC files)"""
        try:
            result = subprocess.run(['antiword', '-v'], 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=5)
            return result.returncode == 0
        except (subprocess.TimeoutExpired, FileNotFoundError, subprocess.SubprocessError):
            return False
    
    def _check_libreoffice(self):
        """Check if LibreOffice is available (for DOC/PPT files)"""
        try:
            result = subprocess.run(['libreoffice', '--version'], 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=5)
            return result.returncode == 0
        except (subprocess.TimeoutExpired, FileNotFoundError, subprocess.SubprocessError):
            return False
    
    def extract_doc_text(self, file_path: str) -> str:
        """Extract text from DOC files using multiple fallback methods"""
        
        # Method 1: Try docx2txt (if available)
        if self.available_processors["docx2txt"]:
            try:
                text = docx2txt.process(file_path)
                if text and text.strip():
                    return text
            except Exception as e:
                st.warning(f"docx2txt failed: {e}")
        
        # Method 2: Try antiword (Linux/Mac)
        if self.available_processors["antiword"]:
            try:
                result = subprocess.run(['antiword', file_path], 
                                      capture_output=True, 
                                      text=True, 
                                      timeout=30)
                if result.returncode == 0 and result.stdout.strip():
                    return result.stdout
            except Exception as e:
                st.warning(f"antiword failed: {e}")
        
        # Method 3: Try LibreOffice conversion
        if self.available_processors["libreoffice"]:
            try:
                return self._convert_with_libreoffice(file_path, "doc")
            except Exception as e:
                st.warning(f"LibreOffice conversion failed: {e}")
        
        # Method 4: Try to read as ZIP and extract XML (some DOC files)
        try:
            return self._try_zip_extraction(file_path)
        except Exception:
            pass
        
        st.error("❌ Cannot process DOC files. Please install one of:")
        st.error("   - antiword: `sudo apt-get install antiword` (Linux)")
        st.error("   - LibreOffice: Download from https://libreoffice.org")
        st.error("   - Or convert to DOCX format")
        
        return ""
    
    def extract_ppt_text(self, file_path: str) -> str:
        """Extract text from PPT files"""
        
        # Method 1: Try LibreOffice conversion
        if self.available_processors["libreoffice"]:
            try:
                return self._convert_with_libreoffice(file_path, "ppt")
            except Exception as e:
                st.warning(f"LibreOffice PPT conversion failed: {e}")
        
        # Method 2: Basic structure reading (limited)
        try:
            return self._basic_ppt_reading(file_path)
        except Exception:
            pass
        
        st.error("❌ Cannot process PPT files. Please install LibreOffice:")
        st.error("   - Download from https://libreoffice.org")
        st.error("   - Or convert to PPTX format")
        
        return ""
    
    def extract_pdf_text_alternative(self, file_path: str) -> str:
        """Alternative PDF text extraction methods"""
        
        # Method 1: PyMuPDF (fitz)
        if self.available_processors["pymupdf"]:
            try:
                doc = fitz.open(file_path)
                text = ""
                for page in doc:
                    text += page.get_text() + "\n"
                doc.close()
                if text.strip():
                    return text
            except Exception as e:
                st.warning(f"PyMuPDF extraction failed: {e}")
        
        # Method 2: pdfminer
        if self.available_processors["pdfminer"]:
            try:
                laparams = LAParams(
                    boxes_flow=0.5,
                    word_margin=0.1,
                    char_margin=2.0,
                    line_margin=0.5
                )
                text = pdf_extract_text(file_path, laparams=laparams)
                if text and text.strip():
                    return text
            except Exception as e:
                st.warning(f"pdfminer extraction failed: {e}")
        
        return ""
    
    def _convert_with_libreoffice(self, file_path: str, file_type: str) -> str:
        """Convert document to text using LibreOffice"""
        with tempfile.TemporaryDirectory() as temp_dir:
            # Convert to text format
            cmd = [
                'libreoffice',
                '--headless',
                '--convert-to', 'txt:Text',
                '--outdir', temp_dir,
                file_path
            ]
            
            result = subprocess.run(cmd, 
                                  capture_output=True, 
                                  text=True, 
                                  timeout=60)
            
            if result.returncode == 0:
                # Find the output file
                output_file = Path(temp_dir) / f"{Path(file_path).stem}.txt"
                if output_file.exists():
                    with open(output_file, 'r', encoding='utf-8') as f:
                        return f.read()
            
            raise Exception(f"LibreOffice conversion failed: {result.stderr}")
    
    def _try_zip_extraction(self, file_path: str) -> str:
        """Try to extract text from ZIP-based formats"""
        try:
            with zipfile.ZipFile(file_path, 'r') as zip_file:
                # Look for text content in XML files
                text_content = []
                for file_name in zip_file.namelist():
                    if file_name.endswith('.xml') and ('document' in file_name or 'content' in file_name):
                        try:
                            content = zip_file.read(file_name).decode('utf-8')
                            # Basic XML tag removal (very simple)
                            import re
                            clean_content = re.sub(r'<[^>]+>', ' ', content)
                            clean_content = re.sub(r'\s+', ' ', clean_content).strip()
                            if clean_content:
                                text_content.append(clean_content)
                        except Exception:
                            continue
                
                return '\n'.join(text_content) if text_content else ""
        except Exception:
            return ""
    
    def _basic_ppt_reading(self, file_path: str) -> str:
        """Basic PPT structure reading (very limited)"""
        try:
            with open(file_path, 'rb') as f:
                content = f.read()
                
            # Very basic text extraction - look for readable text
            import re
            # Find potential text strings (this is very basic)
            text_matches = re.findall(r'[a-zA-Z\s]{10,}', content.decode('latin-1', errors='ignore'))
            
            if text_matches:
                return '\n'.join(text_matches[:50])  # Limit to first 50 matches
            
            return ""
        except Exception:
            return ""
    
    def get_installation_suggestions(self, file_type: str) -> str:
        """Get installation suggestions for missing processors"""
        suggestions = {
            "doc": [
                "Install antiword: `sudo apt-get install antiword` (Linux)",
                "Install LibreOffice: https://libreoffice.org",
                "Convert files to DOCX format"
            ],
            "ppt": [
                "Install LibreOffice: https://libreoffice.org",
                "Convert files to PPTX format"
            ],
            "pdf": [
                "Install PyMuPDF: `pip install PyMuPDF`",
                "Install pdfminer: `pip install pdfminer.six`"
            ]
        }
        
        return suggestions.get(file_type, ["No specific suggestions available"])

# Utility function to display processor status
def display_processor_status():
    """Display available processors in Streamlit sidebar"""
    processor = AlternativeDocumentProcessor()
    
    with st.sidebar.expander("🔧 Document Processors Status"):
        st.write("**Available processors:**")
        
        processors_info = {
            "PDF (PyMuPDF)": processor.available_processors["pymupdf"],
            "PDF (pdfminer)": processor.available_processors["pdfminer"], 
            "DOC (docx2txt)": processor.available_processors["docx2txt"],
            "DOC (antiword)": processor.available_processors["antiword"],
            "DOC/PPT (LibreOffice)": processor.available_processors["libreoffice"]
        }
        
        for name, available in processors_info.items():
            status = "✅" if available else "❌"
            st.write(f"{status} {name}")
        
        missing_count = sum(1 for available in processors_info.values() if not available)
        if missing_count > 0:
            st.warning(f"{missing_count} optional processors not available")
            with st.expander("Installation Help"):
                st.code("""
# Install optional processors:

# For better PDF processing:
pip install PyMuPDF pdfminer.six

# For DOC files (Linux/Mac):
sudo apt-get install antiword

# For DOC/PPT files:
# Download LibreOffice from https://libreoffice.org
                """)

# Function to get the best available processor
def get_best_processor_for_format(file_format: str) -> str:
    """Get the best available processor for a file format"""
    processor = AlternativeDocumentProcessor()
    
    if file_format == "pdf":
        if processor.available_processors["pymupdf"]:
            return "PyMuPDF (Recommended)"
        elif processor.available_processors["pdfminer"]:
            return "pdfminer"
        else:
            return "PyPDF2 (Basic)"
    
    elif file_format == "doc":
        if processor.available_processors["libreoffice"]:
            return "LibreOffice (Best)"
        elif processor.available_processors["antiword"]:
            return "antiword"
        elif processor.available_processors["docx2txt"]:
            return "docx2txt (Limited)"
        else:
            return "Not supported - please convert to DOCX"
    
    elif file_format == "ppt":
        if processor.available_processors["libreoffice"]:
            return "LibreOffice (Best)"
        else:
            return "Not supported - please convert to PPTX"
    
    return "Standard processor"