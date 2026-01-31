import re
import os

try:
    import fitz
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False


class SDSCASReader:
    def __init__(self):
        self.cas_pattern = r'\b[1-9]\d{1,6}-\d{2}-\d\b'

    def extract_text_from_pdf(self, pdf_path):
        text = ""
        if HAS_PYMUPDF:
            try:
                doc = fitz.open(pdf_path)
                for page in doc:
                    text += page.get_text("text", sort=True) + "\n"
                doc.close()
                return text
            except:
                pass
        if HAS_PDFPLUMBER:
            try:
                with pdfplumber.open(pdf_path) as pdf:
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + "\n"
                return text
            except Exception as e:
                raise Exception(f"Error reading PDF: {str(e)}")
        raise Exception("No PDF library available")

    def find_cas_numbers(self, text):
        matches = re.findall(self.cas_pattern, text)
        valid_cas = []
        seen = set()
        for match in matches:
            if self.validate_cas_number(match) and match not in seen:
                valid_cas.append(match)
                seen.add(match)
        return {'valid': valid_cas, 'invalid': [], 'all_ordered': []}

    def validate_cas_number(self, cas):
        try:
            digits = cas.replace('-', '')
            checksum = int(digits[-1])
            digits = digits[:-1]
            total = sum(int(d) * (i + 1) for i, d in enumerate(reversed(digits)))
            return total % 10 == checksum
        except:
            return False
