# Internship Offer Letter Generator

A modern **Streamlit web application** that generates internship offer letters in bulk from an Excel file using a Microsoft Word template.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Streamlit](https://img.shields.io/badge/Streamlit-1.30+-red.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

## 📋 Features

- **Modern Web UI** - Clean and responsive interface built with Streamlit
- **Excel Integration** - Read student data from Excel spreadsheets
- **Word Template** - Use customizable Word templates with placeholders
- **Batch Processing** - Generate multiple offer letters at once
- **Progress Tracking** - Real-time progress bar during generation
- **ZIP Download** - Download all generated letters as a ZIP file
- **Error Handling** - Comprehensive validation with clear error messages

## 📁 Project Structure

```
bulkofferlettergenerator/
│
├── app.py                   # Main Streamlit application
├── requirements.txt         # Python dependencies
├── README.md                # Project documentation
├── students.xlsx            # Sample Excel file (5 students)
├── offer_template.docx      # Word template with placeholders
├── offer_letters/          # Generated letters folder
│   ├── John_Smith_Offer_Letter.docx
│   ├── Emily_Johnson_Offer_Letter.docx
│   └── ...
└── __pycache__/            # Python cache
```

## 🛠️ Installation

### Prerequisites

- Python 3.8 or higher
- Microsoft Word (optional, for PDF conversion)

### Step 1: Install Dependencies

```bash
# Navigate to project directory
cd bulkofferlettergenerator

# Install all dependencies
pip install -r requirements.txt
```

### Step 2: Run the Application

```bash
# Run Streamlit app
streamlit run app.py
```

The application will open in your default browser at `http://localhost:8501`

## 📖 Excel File Format

The Excel file should contain these columns:

| Column Name | Description             | Example         |
| ----------- | ----------------------- | --------------- |
| Name        | Student full name       | John Smith      |
| Email       | Student email address   | john@email.com  |
| Domain      | Internship domain       | Web Development |
| Duration    | Internship duration     | 6 months        |
| Start Date  | Start date (YYYY-MM-DD) | 2024-01-15      |

### Example Data

| Name             | Email                | Domain              | Duration | Start Date |
| ---------------- | -------------------- | ------------------- | -------- | ---------- |
| John Smith       | john.smith@email.com | Web Development     | 6 months | 2024-01-15 |
| Emily Johnson    | emily.j@email.com    | Data Science        | 3 months | 2024-02-01 |
| Michael Williams | m.williams@email.com | Machine Learning    | 6 months | 2024-01-20 |
| Sarah Brown      | s.brown@email.com    | Android Development | 4 months | 2024-03-01 |
| David Jones      | d.jones@email.com    | Cloud Computing     | 5 months | 2024-02-15 |

## 📝 Word Template Format

Create a Word document with these placeholders:

- `{{name}}` - Student name
- `{{domain}}` - Internship domain
- `{{duration}}` - Internship duration
- `{{start_date}}` - Start date

### Example Template

```
INTERNSHIP OFFER LETTER

Dear {{name}},

We are pleased to offer you an internship position in {{domain}}.

Duration: {{duration}}
Start Date: {{start_date}}

We look forward to having you on our team!

Sincerely,
HR Manager
Company Name
```

## 🚀 How to Use

1. **Launch the app**: Run `streamlit run app.py`
2. **Upload Excel**: Click "Browse" to upload your Excel file with student data
3. **Generate**: Click "🚀 Generate Offer Letters" button
4. **Download**: Click "📥 Download All Letters (ZIP)" to download all generated letters

## 📥 Output

Generated files are saved in the `offer_letters/` folder with naming format:

```
StudentName_Offer_Letter.docx
```

Example:

- `John_Smith_Offer_Letter.docx`
- `Emily_Johnson_Offer_Letter.docx`
- `Michael_Williams_Offer_Letter.docx`

## 🔧 Configuration

### Changing the Template

1. Place your Word template in the project folder
2. Name it `offer_template.docx` or update the code to use your filename
3. Ensure placeholders match: `{{name}}`, `{{domain}}`, `{{duration}}`, `{{start_date}}`

### Custom Output Folder

Edit the `OUTPUT_FOLDER` variable in `app.py`:

```python
OUTPUT_FOLDER = 'my_custom_folder'
```

## ⚠️ Error Handling

The application handles these error scenarios:

- **Missing columns**: Shows which required columns are missing
- **Empty file**: Alerts when Excel file has no data
- **Template not found**: Prompts to ensure template exists
- **Invalid data**: Skips empty rows automatically

## 📦 Dependencies

```
streamlit>=1.30.0
pandas>=1.5.0
docxtpl>=0.16.0
openpyxl>=3.0.0
python-docx>=0.8.0
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📝 License

This project is licensed under the MIT License.

## 👨‍💻 Author

Automation Team

---

**Note**: This tool is for legitimate business use only.
