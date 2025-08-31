import sys
from io import BytesIO
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QSplitter, QPushButton,
    QComboBox, QLabel, QHBoxLayout, QFileDialog, QMessageBox
)
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtCore import QUrl
from docx import Document
from translator_columns import TranslatorColumns
from utils import replace_text_in_paragraph


def _norm_key(s: str) -> str:
    """Create normalized key for snippet (remove extra spaces, limit length)."""
    return " ".join((s or "").split())[:400]


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("DOCX Translator App (1-column layout)")
        self.setGeometry(100, 100, 1200, 800)

        # --- UI ---
        self.original_view = QWebEngineView()
        self.original_view.selectionChanged.connect(self.update_selection)
        self.translated_view = QWebEngineView()

        splitter = QSplitter()
        splitter.addWidget(self.original_view)
        splitter.addWidget(self.translated_view)
        splitter.setSizes([600, 600])

        central = QWidget()
        layout = QVBoxLayout()

        # --- Language selector ---
        lang_layout = QHBoxLayout()
        lang_label = QLabel("Target language:")
        self.lang_combo = QComboBox()
        langs = [
            ("English", "en"),
            ("Vietnamese", "vi"),
            ("French", "fr"),
            ("Spanish", "es"),
            ("German", "de"),
            ("Chinese (Simplified)", "zh-CN"),
        ]
        for name, code in langs:
            self.lang_combo.addItem(name, code)
        self.lang_combo.setCurrentIndex(1)
        lang_layout.addWidget(lang_label)
        lang_layout.addWidget(self.lang_combo)
        layout.addLayout(lang_layout)

        # --- Buttons ---
        btn_layout = QHBoxLayout()
        upload_btn = QPushButton("Upload DOCX")
        upload_btn.clicked.connect(self.upload_file)
        translate_btn = QPushButton("Translate selection")
        translate_btn.clicked.connect(self.translate_selected)
        download_btn = QPushButton("Download translated file")
        download_btn.clicked.connect(self.download_file)
        btn_layout.addWidget(upload_btn)
        btn_layout.addWidget(translate_btn)
        btn_layout.addWidget(download_btn)
        layout.addLayout(btn_layout)

        layout.addWidget(splitter)
        central.setLayout(layout)
        self.setCentralWidget(central)

        # --- State ---
        self.original_doc: Document | None = None
        self.translated_doc: Document | None = None
        self.selection_start = -1
        self.selection_end = -1

        self.segment_map: dict[str, dict] = {}

        try:
            self.translator = TranslatorColumns("translate-tool.json")
        except Exception as e:
            QMessageBox.critical(self, "Translator error", str(e))
            self.translator = None

    # ---------- Upload DOCX ----------
    def upload_file(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Select DOCX file", "", "DOCX (*.docx)")
        if not fname:
            return

        self.original_doc = Document(fname)

        buffer = BytesIO()
        self.original_doc.save(buffer)
        buffer.seek(0)
        self.translated_doc = Document(buffer)

        self.segment_map.clear()

        html1, base1 = self.translator.docx_to_html(self.original_doc)
        self.original_view.setHtml(html1, QUrl.fromLocalFile(base1 + "/"))

        html2, base2 = self.translator.docx_to_html(self.translated_doc)
        self.translated_view.setHtml(html2, QUrl.fromLocalFile(base2 + "/"))

    # ---------- Selection ----------
    def update_selection(self):
        js = """
        (function() {
            var sel = window.getSelection();
            if (sel.rangeCount > 0) {
                var r = sel.getRangeAt(0);
                var pre = r.cloneRange();
                pre.selectNodeContents(document.body);
                pre.setEnd(r.startContainer, r.startOffset);
                var start = pre.toString().length;
                var end = start + r.toString().length;
                return [start,end];
            }
            return [-1,-1];
        })();
        """
        self.original_view.page().runJavaScript(js, self.handle_selection_result)

    def handle_selection_result(self, result):
        if result and len(result) == 2:
            self.selection_start, self.selection_end = result
        else:
            self.selection_start = self.selection_end = -1

    # ---------- Translate selected ----------
    def translate_selected(self):
        if not self.translator:
            return
        if self.selection_start == -1 or self.selection_end <= self.selection_start:
            QMessageBox.warning(self, "Warning", "Please highlight some text on the left pane.")
            return

        self.original_view.page().runJavaScript(
            "window.getSelection().toString();",
            self._translate_callback
        )

    def _translate_callback(self, selected_text: str):
        if not selected_text or not selected_text.strip():
            return

        key = _norm_key(selected_text)
        try:
            target_lang = self.lang_combo.currentData()
            new_text = self.translator.translate_text(selected_text, target_lang)

            replaced = False
            para_idx_used = None

            for idx, p in enumerate(self.translated_doc.paragraphs):
                pos = p.text.find(selected_text.strip())
                if pos != -1:
                    replace_text_in_paragraph(p, pos, pos + len(selected_text.strip()), new_text)
                    replaced = True
                    para_idx_used = idx
                    break

            if not replaced and key in self.segment_map:
                info = self.segment_map[key]
                pi = info.get("para_idx")
                last = info.get("last_text", "")
                if isinstance(pi, int) and 0 <= pi < len(self.translated_doc.paragraphs) and last:
                    p = self.translated_doc.paragraphs[pi]
                    pos = p.text.find(last)
                    if pos != -1:
                        replace_text_in_paragraph(p, pos, pos + len(last), new_text)
                        replaced = True
                        para_idx_used = pi

            if not replaced:
                QMessageBox.warning(self, "Not found",
                                    "Could not locate the segment to replace in the translated document.")
                return

            self.segment_map[key] = {"para_idx": para_idx_used, "last_text": new_text}

            html, base = self.translator.docx_to_html(self.translated_doc)
            self.translated_view.setHtml(html, QUrl.fromLocalFile(base + "/"))

        except Exception as e:
            QMessageBox.critical(self, "Translation error", str(e))

    # ---------- Download DOCX ----------
    def download_file(self):
        if not self.translated_doc:
            QMessageBox.warning(self, "Warning", "No translated content to save.")
            return
        fname, _ = QFileDialog.getSaveFileName(self, "Save translated file", "", "DOCX (*.docx)")
        if fname:
            self.translated_doc.save(fname)
            QMessageBox.information(self, "Success", "Translated file has been saved.")

    def closeEvent(self, event):
        try:
            if self.translator and hasattr(self.translator, "cleanup_all_tmp"):
                self.translator.cleanup_all_tmp()
        finally:
            super().closeEvent(event)
