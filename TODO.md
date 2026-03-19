# TODO: Fix Certificate Generation in app.py

## Plan Summary

Fix column normalization, lenient filtering (require only name/email/domain), optional student_id merging, add debug logs for columns/rows/candidates/skipped, ensure PDF/ZIP.

## Steps (check off as completed):

- [x] 1. Update CERT_REQUIRED_COLUMNS to minimal ['name', 'email', 'domain']
- [x] 2. Add 'help_stud' → 'techiehelp_student_id' mapping in data normalization
- [x] 3. Add debug st.info in generate_certificate: columns, first 2 rows, total, candidates
- [x] 4. Replace strict all() filter with check only required non-empty, collect/log skipped rows with missing fields
- [x] 5. Update template context student_id to merged logic or "N/A"
- [ ] 6. Test generation with sample_certificates.xlsx → expect >0 candidates, PDFs in certificates/, ZIP works
- [ ] 7. Update TODO.md with completion

**All code fixes complete!** Test with `streamlit run app.py`, upload sample_certificates.xlsx → Certificates tab → Generate Certs → expect "Found 2 cert candidates", PDFs in certificates/, download ZIP.
