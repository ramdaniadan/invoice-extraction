# 📄 **Invoice Data Extraction**

This repository provides a Python script to extract structured data from invoice PDFs, process it, and save the output into an Excel file. The script uses **regex patterns**, **PDFPlumber**, and **pandas** for efficient and accurate data processing.

---

## 🚀 **Features**

- **Automated PDF Processing**: Extracts data fields such as invoice number, dates, customer information, and more.
- **Data Validation**: Converts and validates data types like dates, numbers, and text.
- **Backup Management**: Moves processed files to a backup folder with timestamped filenames.
- **Excel Export**: Consolidates all extracted data into a timestamped Excel file for easy reporting.
- **Error Handling**: Skips files with errors and continues processing others.

---

## 🛠️ **Technologies Used**

| Technology      | Purpose                           |
|------------------|-----------------------------------|
| **Python**      | Core scripting language          |
| **pdfplumber**  | Extracts text from PDF files      |
| **pandas**      | Data manipulation and export     |
| **os & shutil** | File and directory management    |
| **re (regex)**  | Extracts specific patterns        |

---

## 📂 **Directory Structure**

```plaintext
.
|-- Data/
|   |-- Sample.pdf       # Example input files
|   |-- Backup/          # Backup folder for processed PDFs
|-- Result/
|   |-- result_*.xlsx    # Output Excel files
|-- extract_invoice_data.py # Main script
|-- run_invoice_extraction.bat # Running script
|-- README.md            # Project documentation
```

---

## 📋 **How It Works**

1. **Folder Setup**:
   - Place all invoice PDFs in the `Data` folder.
   - Ensure subfolders `Backup` and `Result` exist, or the script will create them automatically.

2. **Data Extraction**:
   - Extracts specific fields such as:
     - Invoice Number
     - Invoice Date
     - Due Date
     - Customer Number
     - Net Value
     - Amount Due
     - Total Packages, Weight, Volume, and more.

3. **File Backup**:
   - Moves processed PDF files to the `Backup` folder with a timestamped filename.

4. **Excel Export**:
   - Compiles extracted data from all PDFs into a timestamped Excel file in the `Result` folder.

---

## ⚙️ **Getting Started**

### Prerequisites

Ensure the following libraries are installed:

```bash
pip install pdfplumber pandas
```

### Running the Script

1. Clone this repository:

   ```bash
   git clone https://github.com/ramdaniadan/invoice-extraction.git
   cd invoice-extraction
   ```

2. Place your PDF files in the `Data` folder.

3. Run the script:

   ```bash
   python extract_invoice_data.py
   ```

4. Check the `Result` folder for the generated Excel file.

---

## 🛡️ **Error Handling**

- Skips files with missing or malformed fields.
- Logs errors to the console for debugging purposes.

---

## 🛠️ **Future Improvements**

- Add support for multi-page PDFs.
- Integrate a logging system for better error tracking.
- Expand field extractions based on additional invoice formats.

---

## 📬 **Contact**

Have questions or feedback? Feel free to reach out:
- **Email**: ramdaniadan1212@gmail.com
- **GitHub**: [https://github.com/ramdaniadan](https://github.com/ramdaniadan/)

---

Thank you for visiting this repository! 🌟 If you find it useful, please give it a star and share it with others. 🙌
