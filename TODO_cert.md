# Fix 0 Certificates Generated

## Analysis:

CERT_REQUIRED_COLUMNS = ['name', 'email', 'student_id', 'college_name', 'domain', 'start_date', 'end_date']

**sample_certificates.xlsx columns:** Name,Student_ID,College_Name,Domain,Start_Date,End_Date,Email → normalized 'student_id', 'college_name' ✓

**Issue:** Likely uploaded offer sample (missing 'student_id') or normalization mismatch.

## Steps:

- [ ] Add debug st.dataframe(shared_data[0].keys()) in gen_certificate
- [ ] Test with sample_certificates.xlsx
- [ ] Relax filter or map 'techiehelp_student_id' → 'student_id'

**Status:** Adding debug...
