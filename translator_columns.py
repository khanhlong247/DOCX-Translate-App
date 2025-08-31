import os
import re
import tempfile
import shutil
import subprocess
from typing import Tuple
from bs4 import BeautifulSoup
import mammoth
from docx.document import Document as DocxDocument
from translator_base import TranslatorBase


class TranslatorColumns(TranslatorBase):
    def __init__(self, credential_json="translate-tool.json"):
        super().__init__(credential_json)
        self._tmp_html_dir = None
        self._last_html_plain = ""
        self._all_tmp_dirs = []

    def _cleanup_tmpdir(self):
        if self._tmp_html_dir:
            shutil.rmtree(self._tmp_html_dir, ignore_errors=True)
            self._tmp_html_dir = None

    def cleanup_all_tmp(self):
        for d in self._all_tmp_dirs:
            shutil.rmtree(d, ignore_errors=True)
        self._all_tmp_dirs.clear()

    def _save_doc_to_tmp(self, doc):
        tmpdir = tempfile.mkdtemp(prefix="docx_html_")
        self._all_tmp_dirs.append(tmpdir)
        self._tmp_html_dir = tmpdir
        tmp_path = os.path.join(tmpdir, "temp.docx")

        if isinstance(doc, DocxDocument):
            doc.save(tmp_path)
        elif hasattr(doc, "read"):
            with open(tmp_path, "wb") as f:
                f.write(doc.read())
        elif isinstance(doc, (bytes, bytearray)):
            with open(tmp_path, "wb") as f:
                f.write(doc)
        elif isinstance(doc, str) and os.path.isfile(doc):
            shutil.copy(doc, tmp_path)
        else:
            raise TypeError(f"Unsupported doc type: {type(doc)}")
        return tmp_path

    def _export_with_libreoffice(self, docx_path):
        tmpdir = self._tmp_html_dir
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            raise FileNotFoundError("LibreOffice not found")

        subprocess.run(
            [soffice, "--headless", "--convert-to", "html", docx_path, "--outdir", tmpdir],
            check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE
        )
        html_files = [f for f in os.listdir(tmpdir) if f.lower().endswith(".html")]
        if not html_files:
            raise RuntimeError("LibreOffice failed to convert file HTML")
        html_path = os.path.join(tmpdir, html_files[0])
        with open(html_path, "r", encoding="utf-8", errors="ignore") as f:
            html = f.read()
        return html

    def _fallback_with_mammoth(self, docx_path):
        with open(docx_path, "rb") as f:
            content = mammoth.convert_to_html(f).value
        html = f"<html><head><meta charset='utf-8'></head><body>{content}</body></html>"
        return html

    def _strip_multicolumn_styles(self, soup: BeautifulSoup):
        for tag in soup.find_all(style=True):
            style = tag.get("style") or ""
            cleaned = re.sub(
                r"(?:-webkit-)?column-(?:count|width|gap|fill)\s*:\s*[^;]+;?\s*",
                "",
                style,
                flags=re.IGNORECASE
            )
            if cleaned.strip():
                tag["style"] = cleaned
            else:
                del tag["style"]

        for style_tag in soup.find_all("style"):
            css = style_tag.string or ""
            cleaned_css = re.sub(
                r"(?:-webkit-)?column-(?:count|width|gap|fill)\s*:\s*[^;]+;?\s*",
                "",
                css,
                flags=re.IGNORECASE
            )
            style_tag.string = cleaned_css

    def _normalize_floats_and_absolute(self, soup: BeautifulSoup):
        bad_props = re.compile(
            r"(?:^|;)\s*(?:"
            r"(?:-webkit-)?position|top|left|right|bottom|z-index|float|clear|"
            r"text-wrap|wrap-(?:flow|through|margin|distance)|"
            r"mso-position-[^:;]+|mso-wrap-[^:;]+"
            r")\s*:\s*[^;]+;?",
            re.IGNORECASE,
        )

        def _clean_style(style: str) -> str:
            if not style:
                return style
            s = re.sub(bad_props, ";", style)
            s = re.sub(r";{2,}", ";", s).strip(" ;")
            return s

        candidate_tags = soup.find_all(
            lambda t: t.has_attr("style") and t.name in ("img", "figure", "svg", "object", "span", "div")
        )
        for tag in candidate_tags:
            style = tag.get("style", "")
            cleaned = _clean_style(style)
            if cleaned:
                tag["style"] = cleaned
            else:
                del tag["style"]

        for tag in soup.find_all(["img", "figure", "svg", "object"]):
            existing = tag.get("style", "")
            additions = [
                "display:block",
                "position:static",
                "float:none",
                "z-index:auto",
                "top:auto", "left:auto", "right:auto", "bottom:auto",
                "max-width:100%",
                "height:auto",
            ]
            merged = ";".join([s for s in (existing, ";".join(additions)) if s]).strip(";")
            tag["style"] = merged

        extra_css = soup.new_tag("style")
        extra_css.string = """
            *[style*="position"], *[style*="float"], *[style*="z-index"] {
                position: static !important;
                float: none !important;
                z-index: auto !important;
                top: auto !important; left: auto !important; right: auto !important; bottom: auto !important;
            }
            img, figure, svg, object {
                position: static !important;
                float: none !important;
                z-index: auto !important;
                max-width: 100%;
                height: auto;
                display: block;
            }
        """
        if soup.head:
            soup.head.append(extra_css)
        else:
            head_tag = soup.new_tag("head")
            head_tag.append(extra_css)
            soup.insert(0, head_tag)

    def docx_to_html(self, doc) -> Tuple[str, str]:
        docx_path = self._save_doc_to_tmp(doc)
        try:
            html = self._export_with_libreoffice(docx_path)
        except Exception:
            html = self._fallback_with_mammoth(docx_path)
        finally:
            try:
                os.unlink(docx_path)
            except Exception:
                pass

        soup = BeautifulSoup(html, "html.parser")

        self._strip_multicolumn_styles(soup)

        self._normalize_floats_and_absolute(soup)

        base_css = soup.new_tag("style")
        base_css.string = """
            html, body {
                width: 100% !important;
                margin: 0;
                padding: 0;
                column-count: 1 !important;
                -webkit-column-count: 1 !important;
                column-width: auto !important;
                -webkit-column-width: auto !important;
                column-gap: normal !important;
                -webkit-column-gap: normal !important;
                column-fill: auto !important;
                -webkit-column-fill: auto !important;
            }
            div, p { break-inside: avoid; }
            img, table, figure { max-width: 100%; height: auto; }
        """
        if soup.head:
            soup.head.append(base_css)
        else:
            head_tag = soup.new_tag("head")
            head_tag.append(base_css)
            soup.insert(0, head_tag)

        try:
            self._last_html_plain = soup.get_text().replace("\r\n", "\n").replace("\r", "\n")
        except Exception:
            self._last_html_plain = ""

        return str(soup), self._tmp_html_dir
