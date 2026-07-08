# Excel2Doc2PDF 📄⚡

Automate the creation of personalized documents from spreadsheet data.

Excel2Doc2PDF reads information from an Excel spreadsheet, fills a Word template with the corresponding data, and converts the generated documents into PDF files.

Designed for certificates, letters, reports, invoices, forms, and other repetitive document generation tasks.

---

## ✨ Features

- 📊 Read data directly from Excel spreadsheets
- 📝 Automatically fill Word templates
- 🎨 Preserve template formatting, fonts, and layout
- 📄 Generate personalized DOCX files
- 🔄 Convert DOCX files into PDF
- 🖥 Supports Microsoft Word conversion
- 🆓 Supports LibreOffice conversion (no Microsoft Office required)
- 🎯 Process only selected spreadsheet rows or conditions
- 📁 Batch generate hundreds of documents

---

## 🛠 Requirements

### Python

Python 3.10+

### Python Libraries

Install dependencies:

```bash
pip install openpyxl python-docx
````

---

## PDF Conversion

The project supports two conversion methods:

### Option 1: Microsoft Word (Windows)

Requires:

* Microsoft Word installed
* Word desktop application

The script uses Word's built-in PDF export.

---

### Option 2: LibreOffice (Recommended)

Works without Microsoft Office.

Install LibreOffice:

[https://www.libreoffice.org/download/](https://www.libreoffice.org/download/)

The converter uses LibreOffice's headless mode:

```bash
libreoffice --headless --convert-to pdf file.docx
```

---

## 📂 Example Structure

```
Excel2Doc2PDF/

│
├── main.py
├── template.docx
├── students.xlsx
│
├── output/
│   ├── John_Doe.docx
│   ├── John_Doe.pdf
│
└── requirements.txt
```

---

## ⚙️ How It Works

1. Load Excel spreadsheet

```
Name        Course
John Doe    Python
Mary Smith  JavaScript
```

2. Load Word template

```
Certificate awarded to:

{{Name}}

For completing:

{{Course}}
```

3. Replace placeholders

```
Certificate awarded to:

John Doe

For completing:

Python
```

4. Export PDF

---

## 🎯 Use Cases

* Certificates
* Student reports
* Employee documents
* Invitations
* Invoices
* Personalized letters
