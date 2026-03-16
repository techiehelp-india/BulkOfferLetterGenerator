# TODO: PDF Output Instead of DOCX (Approved)

**Status**: Ready to implement

## Breakdown

1. **Install dep**: `pip install docx2pdf` (requires Word)
2. **generate_letters.py**: In `generate_single_letter()`:
   - After `doc.save(docx_path)`
   - pdf*filename = 'offer_letter*' + safe_name + '.pdf'
   - converter = PDFConverter(output_folder)
   - converter.convert_single(docx_path, pdf_path)
   - os.remove(docx_path)
   - return pdf_path
3. **app.py**: Same modification in `generate_single_letter()`
4. **app.py**: Update `create_zip_file()`: `if file.endswith('.pdf')`
5. **gui_app.py**: Update filename expectations/results to .pdf
6. **Test**: Generate, check `offer_letters/` has PDFs only, ZIP PDFs.

**Next**: Step 1 & 2.
