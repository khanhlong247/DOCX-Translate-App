# DOCX-Translate-App

A desktop application (PyQt5) for viewing – selecting – translating individual paragraphs from .docx files into multiple languages using Google Cloud Translate. The app displays two panes side by side (original & translated), normalizes text into a single column, and allows re-translating the same paragraph into different languages without reopening the file.

## 🔥 Key Features

- Side-by-side panes: left (original), right (translated).

- Normalize DOCX → HTML into a single column.

- Translate selected text in the left pane, replacing the corresponding paragraph in the translated pane.

- Re-translate the same paragraph into different languages.

- Save translated file as .docx.

## 🧱 Architecture & Technologies

- Python 3.9+

- PyQt5 + Qt WebEngine (Chromium) for HTML rendering.

- python-docx for reading/writing DOCX.

- LibreOffice (headless) to convert DOCX → HTML (preferred).

- Mammoth as a fallback for DOCX → HTML conversion.

- BeautifulSoup4 for post-processing HTML/CSS.

- Google Cloud Translate v2 for text translation.

## Main Files

├─ main.py # Entry point

├─ ui_mainwindow.py # UI & translation workflow

├─ translator_base.py # Google Translate integration

├─ translator_columns.py # DOCX → HTML conversion

├─ utils.py # Utility functions

└─ translate-tool.json # Service Account JSON

## 🛠️ Installation

**1) Install LibreOffice**

- Download and install LibreOffice (Standard).

- Ensure the soffice command is available in PATH.

- Windows: add folder ...\LibreOffice\program\ to PATH.

**2) Create Translate API**

- Open Google Cloud Console → create a Project.

- Enable Cloud Translation API.

- Create a Service Account → generate JSON key.

- Place the JSON file next to the source code, default name: translate-tool.json (or set the environment variable GOOGLE_APPLICATION_CREDENTIALS pointing to this file).

**3) Create virtual environment & install libraries**

``
python -m venv venv
``

# Windows
``
venv\Scripts\activate
``

# macOS/Linux
``
source venv/bin/activate
``

``
pip install -r requirements.txt
``

## ▶️ Run the Application

``
python main.py
``

## 🚀 How to Use

- Upload DOCX → select a .docx file.

- The app will display the text in two panes.

- Choose Target language (e.g., Vietnamese).

- Highlight the paragraph to translate in the left pane → click Translate selection.

- The corresponding paragraph in the translated pane will be replaced.

**Re-translate the same paragraph into another language:**

- Simply change the target language → re-select the same paragraph on the left → click Translate selection.

- The app will locate the previous translation in the correct paragraph and replace it with the new one.

- Download the translated file to save as .docx.

**Note:** Highlight & translate only applies to paragraph text (does not yet support text inside images/shapes/table cells complex structures).

## 🧩 Customization

- Language list: modify the langs array in **ui_mainwindow.py**.

- Default language: **self.lang_combo.setCurrentIndex(1)** → change the index as needed.

## 🤝 Contributing

**PRs/Issues are welcome. Please provide:**

- Operating system, Python version.

- Error logs/screenshots.

- Sample file (if any) to reproduce the issue.

## 🛡️ License

**MIT license**

## 📣 Acknowledgements

Thanks to LibreOffice, Mammoth, Google Cloud, and the PyQt5/Qt community for providing excellent tools that made this application possible.
