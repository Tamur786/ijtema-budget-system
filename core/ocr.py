import os
import re

import pytesseract
import numpy as np
import cv2
from pdf2image import convert_from_path
from pypdf import PdfReader


# ================= CONFIG =================

TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
POPPLER_PATH = r"C:\poppler\Library\bin"

pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


# ================= TEXTEXTRAKTION =================

def _ocr_image(arr, psm=6):
    gray = cv2.cvtColor(arr, cv2.COLOR_BGR2GRAY)

    thresh = cv2.adaptiveThreshold(
        gray,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        8
    )

    return pytesseract.image_to_string(
        thresh,
        lang="deu+eng",
        config=f"--psm {psm}"
    )


def extract_text_ocr(pdf_path):
    images = convert_from_path(
        pdf_path,
        dpi=250,
        poppler_path=POPPLER_PATH
    )

    full_text = ""

    for img in images:
        arr = np.array(img)
        h = arr.shape[0]

        top = arr[:int(h * 0.08), :]

        top_text = pytesseract.image_to_string(
            cv2.threshold(
                cv2.cvtColor(top, cv2.COLOR_BGR2GRAY),
                200,
                255,
                cv2.THRESH_BINARY
            )[1],
            lang="deu+eng",
            config="--psm 6"
        )

        full_text += top_text + "\n"
        full_text += _ocr_image(arr) + "\n"

    return full_text


def extract_text_pypdf(pdf_path):
    try:
        reader = PdfReader(pdf_path)
        text = "".join(page.extract_text() or "" for page in reader.pages)

        if len(text.strip()) > 50:
            return text

    except Exception:
        pass

    return ""


def get_full_text(pdf_path):
    return extract_text_ocr(pdf_path) + "\n" + extract_text_pypdf(pdf_path)


# ================= SMART EXTRACTION =================

_SUPPLIER_SKIP = {
    "khuddam", "ahmadiyya", "jamaat", "mkad",
    "rechnung", "seite", "datum", "telefon", "fax",
    "www", "e-mail", "email",
    "bankverbindung", "bankaccount", "kontonummer",
    "zahlungsbedingung", "lieferbedingung", "versandart",
    "eigentumsvorbehalt", "amtsgericht", "geschäftsführer",
    "steuernummer", "umsatzsteuer", "umsatz", "vat",
}

_COMPANY_SUFFIX = re.compile(
    r'\b(?:GmbH|AG|KG|e\.V\.|eV|OHG|UG|Ltd\.?|Inc\.?|Corp\.?|'
    r'Trading|Group|Handels?|Services?|Solutions?|Vertriebs?|GMBH)\b',
    re.IGNORECASE,
)

_IBAN_RAW = re.compile(
    r'[A-Z]{2}[\s\-]?\d{2}(?:[\s\-]?[A-Z0-9]{4}){3,7}',
    re.IGNORECASE,
)


def extract_supplier(text: str) -> str:
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    lines = [re.split(r'\s*(?:IBAN|BIC|\|)', line)[0].strip() for line in lines]

    candidates_with_suffix = []
    candidates_allcaps = []

    for line in lines:
        if len(line) < 4 or len(line) > 100:
            continue

        if "@" in line or "http" in line.lower():
            continue

        if re.match(r'^[\d\s\+\-\(\)\/\.\,]+$', line):
            continue

        words_lower = {w.lower() for w in re.split(r'\W+', line) if w}

        if words_lower and words_lower.issubset(_SUPPLIER_SKIP):
            continue

        if _COMPANY_SUFFIX.search(line):
            clean = re.split(
                r'\s{3,}|\t|IBAN|BIC|Tel\.|Fax|info@|\|',
                line
            )[0].strip()
            candidates_with_suffix.append(clean)

        letters = [c for c in line if c.isalpha()]
        if letters:
            upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
            if upper_ratio >= 0.6 and len(line) >= 5:
                candidates_allcaps.append(line)

    if candidates_with_suffix:
        return candidates_with_suffix[0]

    if candidates_allcaps:
        return candidates_allcaps[0]

    for line in lines[:20]:
        if len(line) < 4 or len(line) > 80:
            continue

        if "@" in line or "http" in line.lower():
            continue

        if re.match(r'^[\d\s\+\-\(\)\/\.\,]+$', line):
            continue

        words_lower = {w.lower() for w in re.split(r'\W+', line) if w}

        if words_lower and words_lower.issubset(_SUPPLIER_SKIP):
            continue

        return line

    return ""


def extract_amount(text: str) -> str:
    patterns = [
        r'Endbetrag\s*:?\s*(?:EUR\s*)?([\d]{1,3}(?:[.\s]?\d{3})*,\d{2})',
        r'Gesamtbetrag\s*:?\s*(?:EUR\s*)?([\d]{1,3}(?:[.\s]?\d{3})*,\d{2})',
        r'Rechnungsbetrag\s*:?\s*(?:EUR\s*)?([\d]{1,3}(?:[.\s]?\d{3})*,\d{2})',
        r'Bruttobetrag\s*:?\s*(?:EUR\s*)?([\d]{1,3}(?:[.\s]?\d{3})*,\d{2})',
        r'Total\s*:?\s*(?:EUR\s*)?([\d]{1,3}(?:[.\s]?\d{3})*,\d{2})',
        r'Summe\s*:?\s*(?:EUR\s*)?([\d]{1,3}(?:[.\s]?\d{3})*,\d{2})',
        r'(?:EUR\s+)([\d]{1,3}(?:[.\s]?\d{3})*,\d{2})',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).replace(" ", "")

    nums = re.findall(r'\d{1,3}(?:\.\d{3})*,\d{2}', text)

    if nums:
        return max(
            nums,
            key=lambda x: float(x.replace(".", "").replace(",", "."))
        )

    return ""


def extract_ibans(text: str) -> list:
    fixed = (
        text
        .replace("DES", "DE5")
        .replace("DE8?", "DE87")
    )

    seen = set()
    result = []

    for match in _IBAN_RAW.finditer(fixed):
        raw = match.group(0)
        clean = re.sub(r'[^A-Za-z0-9]', '', raw).upper()

        if len(clean) >= 15 and clean not in seen:
            seen.add(clean)
            result.append(clean)

    return result


def extract_rechnungsnr(text: str) -> str:
    patterns = [
        r'Rechnungs(?:nummer|nr\.?)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)',
        r'Rechn?\.\s*Nr\.?\s*[:\-]?\s*([A-Za-z0-9\-\/]+)',
        r'Rg\.?\s*[Nn]r\.?\s*[:\-]?\s*([A-Za-z0-9\-\/]+)',
        r'Invoice\s*(?:No\.?|Number|#)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)',
        r'Document\s*(?:No\.?|Number|#)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)',
        r'Beleg(?:nummer|nr\.?)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)',
        r'Auftrags(?:nummer|nr\.?)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)',
        r'Bestell(?:nummer|nr\.?)\s*[:\-]?\s*([A-Za-z0-9\-\/]+)',
        r'\b((?:RE|INV|RG|RN)\d{4,})\b',
    ]

    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)

        if match:
            value = match.group(1).strip().rstrip(".,;:")

            if re.search(r'\d', value):
                return value

    return ""


def extract_date(text: str) -> str:
    match = re.search(
        r'(?:Rechnungs|Auftrags|Liefer|Beleg)?[Dd]atum\s*[:\-]?\s*'
        r'(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})',
        text,
        re.IGNORECASE,
    )

    if match:
        return match.group(1)

    match = re.search(r'\b(\d{1,2}\.\d{1,2}\.\d{4})\b', text)

    return match.group(1) if match else ""


def analyze_invoice(pdf_path: str) -> dict:
    text = get_full_text(pdf_path)

    ibans = extract_ibans(text)

    return {
        "supplier": extract_supplier(text),
        "amount": extract_amount(text),
        "ibans": ibans,
        "iban": ibans[0] if ibans else "",
        "invoice_number": extract_rechnungsnr(text),
        "date": extract_date(text),
        "raw_text": text,
    }