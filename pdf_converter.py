"""
PDF Converter Module
--------------------
Handles DOCX→PDF using docx2pdf with COM init fix + ReportLab fallback.

Author: Automation Team
"""

import os
try:
    from docx2pdf import convert
except ImportError:
    convert = None

try:
    import comtypes.client
    import pythoncom
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False
    pythoncom = None

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import inch
    from docx import Document
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


class PDFConverter:
    def __init__(self, output_folder='offer_letters'):
        self.output_folder = output_folder
    
    def convert_single(self, word_path, pdf_path=None):
        try:
            if not os.path.exists(word_path):
                print(f"✗ File not found: {word_path}")
                return None
            
            if pdf_path is None:
                pdf_path = os.path.splitext(word_path)[0] + '.pdf'
            
            # Method 1: docx2pdf with COM init
            if convert and COM_AVAILABLE and pythoncom:
                try:
                    pythoncom.CoInitialize()
                    convert(word_path, pdf_path)
                    pythoncom.CoUninitialize()
                    print(f"  ✓ Converted (docx2pdf): {os.path.basename(pdf_path)}")
                    return pdf_path
                except Exception as com_e:
                    print(f"  ⚠️ docx2pdf failed ({com_e}), trying fallback")
            
            # Method 2: ReportLab text PDF fallback
            if REPORTLAB_AVAILABLE:
                return self._reportlab_text_pdf(word_path, pdf_path)
            
            print(f"✗ No converter: {os.path.basename(word_path)}")
            return None
            
        except Exception as e:
            print(f"✗ Convert error {os.path.basename(word_path)}: {str(e)}")
            return None
    
    def _reportlab_text_pdf(self, word_path, pdf_path):
        try:
            doc = Document(word_path)
            text_lines = [para.text for para in doc.paragraphs if para.text.strip()]
            text = '\n'.join(text_lines)
            
            c = canvas.Canvas(pdf_path, pagesize=letter)
            width, height = letter
            y = height - inch
            
            for line in text.split('\n'):
                if y < inch:
                    c.showPage()
                    y = height - inch
                c.drawString(inch, y, line[:85])
                y -= 14
            
            c.save()
            print(f"  ✓ Text PDF: {os.path.basename(pdf_path)}")
            return pdf_path
        except Exception as e:
            print(f"✗ Fallback fail: {str(e)}")
            return None
    
    def convert_all_in_folder(self, folder_path=None):
        folder = folder_path or self.output_folder
        word_files = [f for f in os.listdir(folder) if f.endswith('.docx') and not f.startswith('~')]
        
        results = {'converted': 0, 'total': len(word_files), 'errors': 0}
        for f in word_files:
            if self.convert_single(os.path.join(folder, f)):
                results['converted'] += 1
            else:
                results['errors'] += 1
        print(f"Converted {results['converted']}/{results['total']}")
        return results


if __name__ == '__main__':
    print("PDF Converter: docx2pdf (COM fixed) + ReportLab fallback ready.")

