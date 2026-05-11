import os
import re
import io
import base64
import tempfile

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(BASE_DIR, "ExpenseVoucher2024.pdf")

W, H = 612.283, 858.898

CHECKBOX_NATIONAL = (43.5, 775.5)
CHECKBOX_REGIONAL = (117.3, 775.5)
CHECKBOX_MAJLIS   = (188.5, 775.5)

POS_NAME_AMILA  = (51.0, 710.0)
POS_ID_AMILA    = (261.1, 710.0)
POS_NAME_ANTRAG = (51.0, 674.0)
POS_ID_ANTRAG   = (261.1, 674.0)
POS_TEL_ANTRAG  = (372.6, 674.0)
POS_ABTEILUNG   = (51.0, 638.0)
POS_ZWECK       = (261.1, 638.0)
POS_TEILNEHMER  = (540.0, 640.0)

AUSGABEN_LEFT_X  = 280.0
AUSGABEN_RIGHT_X = 535.0
AUSGABEN_Y = {
    1: 528.0, 2: 505.5, 3: 483.0, 4: 460.5, 5: 438.0,
    6: 528.0, 7: 505.5, 8: 483.0, 9: 460.5, 10: 438.0,
}

POS_TOTAL      = (535.0, 400.0)
POS_NAME_KONTO = (51.0, 335.0)
POS_VERWENDUNG = (348.7, 335.0)

IBAN_X_START   = 51.3
IBAN_BOX_WIDTH = 23.2
IBAN_Y         = 291.5
IBAN_OFFSET    = 5.5

POS_DATUM = (60.0, 90.0)

# Neue Positionen
CHECKBOX_ADVANCE_JA   = (443.0, 708.0)
CHECKBOX_ADVANCE_NEIN = (466.0, 708.0)

POS_ADVANCE_BETRAG    = (505.0, 708.0)

SIG_APPLICANT_POS = (145.0, 82.0, 110.0, 22.0)

SIG_AMILA_POS     = (405.0, 82.0, 110.0, 22.0)


def draw_signature(can, signature_data, pos):
    if not signature_data:
        return

    if "base64," not in signature_data:
        return

    try:
        raw = signature_data.split("base64,", 1)[1]
        image_bytes = base64.b64decode(raw)

        img = ImageReader(io.BytesIO(image_bytes))

        x, y, w, h = pos
        can.drawImage(
            img,
            x,
            y,
            width=w,
            height=h,
            mask="auto",
            preserveAspectRatio=True
        )

    except Exception:
        pass


def fill_voucher(output_path: str, data: dict, template_path: str = TEMPLATE_PATH):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=(W, H))
    can.setFont("Helvetica", 9)

    scope = data.get("scope", "")
    for val, pos in [
        ("National", CHECKBOX_NATIONAL),
        ("Regional", CHECKBOX_REGIONAL),
        ("Majlis", CHECKBOX_MAJLIS),
    ]:
        if scope == val:
            can.drawString(pos[0] + 1, pos[1], "X")

    advance = data.get("advance_erhalten", "")
    if advance == "Ja":
        can.drawString(CHECKBOX_ADVANCE_JA[0], CHECKBOX_ADVANCE_JA[1], "X")
    elif advance == "Nein":
        can.drawString(CHECKBOX_ADVANCE_NEIN[0], CHECKBOX_ADVANCE_NEIN[1], "X")

    advance_betrag = str(data.get("advance_betrag", "") or "").strip()
    if advance_betrag:
        can.drawString(POS_ADVANCE_BETRAG[0], POS_ADVANCE_BETRAG[1], advance_betrag)

    def ds(pos, key):
        value = str(data.get(key, "") or "").strip()
        if value:
            can.drawString(pos[0], pos[1], value)

    ds(POS_NAME_AMILA, "name_amila")
    ds(POS_ID_AMILA, "id_amila")
    ds(POS_NAME_ANTRAG, "name_antrag")
    ds(POS_ID_ANTRAG, "id_antrag")
    ds(POS_TEL_ANTRAG, "tel_antrag")
    ds(POS_ABTEILUNG, "abteilung")
    ds(POS_ZWECK, "zweck")

    teilnehmer = str(data.get("teilnehmer", "") or "").strip()
    if teilnehmer:
        can.drawRightString(POS_TEILNEHMER[0], POS_TEILNEHMER[1], teilnehmer)

    ds(POS_DATUM, "datum")

    for i in range(1, 11):
        value = str(data.get(f"pos{i}_betrag", "") or "").strip()
        if value:
            x = AUSGABEN_LEFT_X if i <= 5 else AUSGABEN_RIGHT_X
            can.drawRightString(x, AUSGABEN_Y[i], value)

    total = str(data.get("total", "") or "").strip()
    if total:
        can.drawRightString(POS_TOTAL[0], POS_TOTAL[1], total)

    ds(POS_NAME_KONTO, "supplier")
    ds(POS_VERWENDUNG, "verwendungszweck")

    iban = re.sub(r"[^A-Za-z0-9]", "", str(data.get("iban", "") or "")).upper()
    chars = iban if len(iban) >= 15 else ""

    for i, ch in enumerate(chars[:34]):
        can.drawString(IBAN_X_START + i * IBAN_BOX_WIDTH + IBAN_OFFSET, IBAN_Y, ch)

    draw_signature(can, data.get("signature_applicant"), SIG_APPLICANT_POS)
    draw_signature(can, data.get("signature_amila"), SIG_AMILA_POS)

    can.save()
    packet.seek(0)

    overlay = PdfReader(packet)

    with open(template_path, "rb") as f:
        template = PdfReader(f)
        writer = PdfWriter()

        page = template.pages[0]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)

        with open(output_path, "wb") as out:
            writer.write(out)


def merge_pdfs(voucher_path: str, invoice_path: str, output_path: str):
    writer = PdfWriter()

    with open(voucher_path, "rb") as vf:
        voucher_reader = PdfReader(vf)
        for page in voucher_reader.pages:
            writer.add_page(page)

    if invoice_path and os.path.exists(invoice_path):
        with open(invoice_path, "rb") as inf:
            invoice_reader = PdfReader(inf)
            for page in invoice_reader.pages:
                writer.add_page(page)

    with open(output_path, "wb") as out:
        writer.write(out)

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import HexColor


def create_advance_pdf(output_path, data):

    c = canvas.Canvas(output_path, pagesize=A4)

    width, height = A4

    # Header
    c.setFillColor(HexColor("#1b2d42"))
    c.rect(0, height - 90, width, 90, fill=1)

    c.setFillColor(HexColor("#ffffff"))
    c.setFont("Helvetica-Bold", 24)
    c.drawString(50, height - 55, "ADVANCE ANTRAG")

    c.setFont("Helvetica", 12)
    c.drawString(
        50,
        height - 78,
        "Vorschuss-Antrag vor finaler Rechnungsabrechnung"
    )

    # Box
    c.setFillColor(HexColor("#ffffff"))
    c.roundRect(40, 180, width - 80, 520, 10, fill=0)

    c.setFillColor(HexColor("#000000"))

    y = 650

    def row(label, value):
        nonlocal y

        c.setFont("Helvetica-Bold", 12)
        c.drawString(60, y, label)

        c.setFont("Helvetica", 12)
        c.drawString(260, y, str(value))

        y -= 40

    row("Advance Code:", data.get("code"))
    row("Datum:", data.get("date"))
    row("Antragsteller:", data.get("name"))
    row("ID:", data.get("member_id"))
    row("Telefon:", data.get("phone"))
    row("Abteilung / Region:", data.get("department"))

    c.setFont("Helvetica-Bold", 12)
    c.drawString(60, y, "Beantragter Betrag:")

    c.setFont("Helvetica-Bold", 16)
    c.drawString(
        260,
        y,
        f"{data.get('advance_amount')} €"
    )

    y -= 60

    c.setFont("Helvetica-Bold", 12)
    c.drawString(60, y, "Zweck des Advances:")

    y -= 25

    text = c.beginText(60, y)
    text.setFont("Helvetica", 12)

    purpose = data.get("purpose", "")

    for line in purpose.split("\n"):
        text.textLine(line)

    c.drawText(text)

    # Hinweis
    c.setFillColor(HexColor("#b03a2e"))
    c.setFont("Helvetica-Bold", 11)

    c.drawString(
        60,
        220,
        "Dies ist KEIN finaler Expense Voucher."
    )

    c.drawString(
        60,
        200,
        "Die Rechnungen müssen später nachgereicht werden."
    )

    # Linien
    c.setFillColor(HexColor("#000000"))

    c.line(60, 120, 240, 120)
    c.line(330, 120, 510, 120)

    c.setFont("Helvetica", 10)

    c.drawString(60, 105, "Unterschrift Antragsteller")
    c.drawString(330, 105, "Genehmigung Nazim / Maal")

    c.save()