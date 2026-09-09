# AI Invoice Accounting Assistant 🧾

An AI-powered invoice processing application built with **Python** and **Streamlit**.

The application analyzes invoice images using the **Google Gemini API**, extracts structured accounting data, validates the results, and generates an Excel report for further accounting workflows.

## ✨ Features

- 🤖 AI-powered invoice data extraction
- 📄 Batch processing of multiple invoices
- ✅ VAT and amount consistency checks
- 🔢 Polish NIP number validation
- ⚠️ Automatic warning detection for inconsistent or suspicious data
- 🌍 Multilingual interface: Turkish, English, and Polish
- 📊 Excel report generation
- ⚡ Parallel processing of multiple invoices

## 🛠️ Tech Stack

- **Python**
- **Streamlit**
- **Google Gemini API**
- **Pandas**
- **Pillow**
- **OpenPyXL**

## 🔄 How It Works

```text
Invoice Image
      ↓
Google Gemini API
      ↓
Structured JSON Data
      ↓
Validation & Error Checks
      ↓
Data Preview
      ↓
Excel Report
```

## ✅ Validation

The application performs several validation checks:

- Netto + VAT = Brutto consistency
- Polish NIP number length validation
- Detection of unusually high invoice amounts
- Preservation of AI-generated warnings for manual review

## 🌍 Supported Languages

The application interface is available in:

- Turkish 🇹🇷
- English 🇬🇧
- Polish 🇵🇱

## 🚀 Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/julwenis/ai-invoice-accounting.git
cd ai-invoice-accounting
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure the Gemini API

Create the following file:

```text
.streamlit/secrets.toml
```

Add your API key:

```toml
GEMINI_API_KEY = "your-api-key"
```

### 4. Run the application

```bash
streamlit run web_ai.py
```

## 📊 Output

Processed invoice data is displayed inside the application and can be exported as an Excel workbook for further accounting workflows.

## ⚠️ Disclaimer

This project is intended for software demonstration and automation purposes.

AI-generated accounting data should be manually reviewed before being used for official accounting or tax reporting.
