"""
Generate Letters Module
-----------------------
Generate PDF offer letters from Excel using Word template.

Author: Automation Team
"""

import pandas as pd
from docxtpl import DocxTemplate
import os
import re
from datetime import datetime
from pdf_converter import PDFConverter


class OfferLetterGenerator:
    REQUIRED_COLUMNS = ['Name', 'Email', 'Domain', 'Duration', 'Start Date', 'College Name', 'TechieHelp Student Id']
    
    def __init__(self, excel_file, template_file, output_folder='offer_letters'):
        self.excel_file = excel_file
        self.template_file = template_file
        self.output_folder = output_folder
        
        os.makedirs(output_folder, exist_ok=True)
    
    def validate_excel_file(self, df):
        if df.empty:
            return False, "Excel file is empty"
        
        missing = [col for col in self.REQUIRED_COLUMNS if col not in df.columns]
        if missing:
            return False, f"Missing columns: {', '.join(missing)}"
        
        return True, ""
    
    def clean_data(self, df):
        df = df.dropna(how='all').dropna(subset=['Name', 'Email'])
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].str.strip()
        return df
    
    def sanitize_filename(self, name):
        name = re.sub(r'[^\w\s-]', '', name).strip().replace(' ', '_')
        return name
    
    def generate_single_letter(self, student_data, doc):
        context = {
            'name': student_data['Name'],
            'domain': student_data['Domain'],
            'duration': student_data['Duration'],
            'start_date': student_data['Start Date'],
            'college_name': student_data.get('College Name', ''),
            'student_id': student_data.get('TechieHelp Student Id', ''),
            'end_date': student_data.get('End Date', ''),
            'current_date': datetime.now().strftime("%d %B %Y")
        }
        
        doc.render(context)
        
        safe_name = self.sanitize_filename(student_data['Name'])
        docx_path = os.path.join(self.output_folder, f"{safe_name}_Offer_Letter.docx")
        doc.save(docx_path)
        
        # Convert to PDF (handles fallback)
        converter = PDFConverter(self.output_folder)
        pdf_path = os.path.join(self.output_folder, f"offer_letter_{safe_name}.pdf")
        converter.convert_single(docx_path, pdf_path)
        os.remove(docx_path)
        
        return pdf_path
    
    def generate_all_letters(self):
        results = {'success': [], 'errors': [], 'total': 0, 'generated': 0}
        
        try:
            df = pd.read_excel(self.excel_file, engine='openpyxl')
            results['total'] = len(df)
            
            is_valid, err = self.validate_excel_file(df)
            if not is_valid:
                results['errors'].append(err)
                return results
            
            df = self.clean_data(df)
            results['total'] = len(df)
            
            if results['total'] == 0:
                results['errors'].append("No valid data")
                return results
            
            doc = DocxTemplate(self.template_file)
            
            for _, row in df.iterrows():
                try:
                    student_data = row.to_dict()
                    if isinstance(student_data.get('Start Date'), datetime):
                        student_data['Start Date'] = student_data['Start Date'].strftime('%B %d, %Y')
                    
                    pdf_path = self.generate_single_letter(student_data, doc)
                    
                    if pdf_path and os.path.exists(pdf_path):
                        results['success'].append({
                            'name': student_data['Name'],
                            'email': student_data['Email'],
                            'file': pdf_path
                        })
                        results['generated'] += 1
                        
                except Exception as e:
                    results['errors'].append(str(e))
            
            return results
            
        except Exception as e:
            results['errors'].append(str(e))
            return results


def generate_letters_from_excel(excel_file, template_file, output_folder='offer_letters'):
    generator = OfferLetterGenerator(excel_file, template_file, output_folder)
    return generator.generate_all_letters()


if __name__ == '__main__':
    results = generate_letters_from_excel('students.xlsx', 'offer_template.docx')
    print(f"Generated {results['generated']}/{results['total']} PDFs")

