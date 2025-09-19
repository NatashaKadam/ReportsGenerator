import os
import sys
import atexit
import hashlib
import json
from decimal import Decimal
from PyQt6.QtWidgets import QApplication, QMessageBox
from PyQt6.QtCore import QTranslator
from .constants import SCRIPT_DIR, TRANSLATIONS

TEMP_FILES = []

def cleanup_temp_files():
    for f in TEMP_FILES:
        try:
            if os.path.exists(f):
                os.remove(f)
        except OSError:
            pass
atexit.register(cleanup_temp_files)

def num_to_words_indian(num_str):
    def convert_to_words(n):
        ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"]
        tens = ["", "Ten", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
        teens = ["Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
        words = ""
        if n >= 10000000:
            words += convert_to_words(n // 10000000) + " Crore "
            n %= 10000000
        if n >= 100000:
            words += convert_to_words(n // 100000) + " Lakh "
            n %= 100000
        if n >= 1000:
            words += convert_to_words(n // 1000) + " Thousand "
            n %= 1000
        if n >= 100:
            words += ones[n // 100] + " Hundred "
            n %= 100
        if n >= 20:
            words += tens[n // 10] + " "
            n %= 10
        elif n >= 10:
            words += teens[n - 10] + " "
            n = 0
        if n > 0:
            words += ones[n] + " "
        return words
    try:
        num = Decimal(str(num_str).replace("₹", "").replace(",", ""))
        rupees = int(num)
        paise = int(round((num - rupees) * 100))
        rupees_words = "Rupees " + convert_to_words(rupees).strip() if rupees > 0 else ""
        paise_words = "Paise " + convert_to_words(paise).strip() if paise > 0 else ""
        result = (rupees_words + " " + paise_words).strip()
        if not result:
            return "Zero"
        return result + " Only"
    except Exception:
        return ""

class CustomTranslator(QTranslator):
    def __init__(self, parent=None, language_code='en'):
        super().__init__(parent)
        self.language_code = language_code
        self._translations = TRANSLATIONS.get(language_code, {})
    def translate(self, context, sourceText, disambiguation=None, n=-1):
        if self.language_code == 'en':
            return sourceText
        if '%' not in sourceText and '{' not in sourceText:
            return self._translations.get(sourceText, sourceText)
        translated_text = self._translations.get(sourceText, sourceText)
        if n != -1:
            try:
                return translated_text % n
            except (TypeError, ValueError):
                pass
        return translated_text

class OperationCanceledError(Exception):
    pass

def setup_assets():
    assets_dir = os.path.join(SCRIPT_DIR, "assets")
    if not os.path.exists(assets_dir): os.makedirs(assets_dir)
    required_assets = [
        os.path.join(SCRIPT_DIR, "assets", "template_merged.docx"), 
        os.path.join(SCRIPT_DIR, "assets", "ssr_data.xlsx"),
        os.path.join(assets_dir, "app_icon.png"),
        os.path.join(assets_dir, "settings_icon.png"),
        os.path.join(assets_dir, "left_arrow_icon.png"),
        os.path.join(assets_dir, "right_arrow_icon.png")
    ]
    missing_assets = [path for path in required_assets if not os.path.exists(path)]
    if missing_assets:
        temp_app = QApplication.instance() or QApplication(sys.argv)
        msg_box = QMessageBox()
        msg_box.setIcon(QMessageBox.Icon.Critical)
        msg_box.setText("Asset Missing")
        msg_box.setInformativeText(f"Required asset files not found:\n" + "\n".join(missing_assets))
        msg_box.setWindowTitle("Error")
        msg_box.exec()
        return False
    return True