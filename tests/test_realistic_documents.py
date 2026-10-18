"""Integration-style scenarios using entirely fictional financial documents."""
from io import BytesIO
from pathlib import Path
import zipfile

from openpyxl import Workbook
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas

from docflow.core import (
    extract_fields, match_by_spreadsheet, merge_pair, merge_receipt_to_many,
    read_pdf, rename_suggestions, search_pdfs,
)


def document(path: Path, lines: list[str], pages: int = 1):
    """Create text-selectable, synthetic invoice/receipt PDFs."""
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=(595, 842))
    for page in range(pages):
        pdf.setFont("Helvetica", 12)
        pdf.drawString(48, 790, "DOCUMENTO FICTICIO - APENAS TESTES")
        for index, line in enumerate(lines):
            pdf.drawString(48, 745 - index * 25, line)
        pdf.drawString(48, 70, f"Pagina {page + 1} de {pages}")
        pdf.showPage()
    pdf.save()
    path.write_bytes(buffer.getvalue())
    return path


def workbook(path, numbers):
    book = Workbook()
    sheet = book.active
    sheet.append(["NF", "Fornecedor", "Valor"])
    for number in numbers:
        sheet.append([number, "Empresa Exemplo LTDA", 1234.56])
    book.save(path)
    return path


def test_invoice_fields_with_text_pdf(tmp_path):
    invoice = document(tmp_path / "nota_000123.pdf", [
        "NOTA FISCAL N 123", "CNPJ 12.345.678/0001-90",
        "Valor total: R$ 1.234,56",
    ])
    assert "NOTA FISCAL" in read_pdf(invoice)
    fields = extract_fields(invoice)
    assert fields["cnpj"] == "12.345.678/0001-90"
    assert str(fields["valor"]) == "1234.56"
    assert fields["numero"] == "123"


def test_receipt_precedes_invoice_even_for_multiple_pages(tmp_path):
    receipt = document(tmp_path / "comprovante.pdf", ["PAGAMENTO APROVADO", "Valor pago: R$ 1.234,56"], pages=2)
    invoice = document(tmp_path / "nota.pdf", ["NOTA FISCAL N 123"], pages=3)
    out = tmp_path / "combinado.pdf"
    merge_pair(receipt, invoice, out)
    reader = PdfReader(out)
    assert len(reader.pages) == 5
    assert "PAGAMENTO APROVADO" in reader.pages[0].extract_text()
    assert "PAGAMENTO APROVADO" in reader.pages[1].extract_text()
    assert "NOTA FISCAL" in reader.pages[2].extract_text()


def test_spreadsheet_missing_and_ambiguous_matches(tmp_path):
    invoice = document(tmp_path / "NF_123.pdf", ["NOTA FISCAL N 123"])
    receipt = document(tmp_path / "COMPROVANTE_123.pdf", ["PAGAMENTO APROVADO"])
    duplicate = document(tmp_path / "OUTRO_COMPROVANTE_123.pdf", ["PAGAMENTO APROVADO"])
    sheet = workbook(tmp_path / "controle.xlsx", [123, 999])
    rows = match_by_spreadsheet(sheet, [invoice], [receipt, duplicate], tmp_path / "saida")
    assert [row.status for row in rows] == ["REVISAR", "REVISAR"]
    assert "2 comprovantes" in rows[0].detail
    assert "0 notas" in rows[1].detail


def test_numbers_do_not_match_part_of_larger_invoice_id(tmp_path):
    invoice = document(tmp_path / "NF_11001.pdf", ["NOTA FISCAL N 11001"])
    receipt = document(tmp_path / "COMPROVANTE_1001.pdf", ["PAGAMENTO APROVADO"])
    sheet = workbook(tmp_path / "controle.xlsx", [1001])
    rows = match_by_spreadsheet(sheet, [invoice], [receipt], tmp_path / "saida")
    assert rows[0].status == "REVISAR"
    assert "0 notas" in rows[0].detail


def test_search_pdf_text_and_filename(tmp_path):
    invoice = document(tmp_path / "documento_2025.pdf", ["NOTA FISCAL N 31415", "Empresa Exemplo LTDA"])
    assert search_pdfs(tmp_path, "31415", content=False) == []
    found = search_pdfs(tmp_path, "31415", content=True)
    assert len(found) == 1 and found[0].source == invoice.name


def test_batch_merge_outputs_are_distinct(tmp_path):
    receipt = document(tmp_path / "comprovante.pdf", ["PAGAMENTO APROVADO"])
    invoices = [document(tmp_path / f"NF_{n}.pdf", [f"NOTA FISCAL N {n}"]) for n in (101, 102, 103)]
    rows = merge_receipt_to_many(receipt, invoices, tmp_path / "saida")
    assert all(row.status == "OK" for row in rows)
    assert len({row.output for row in rows}) == 3
    assert all(len(PdfReader(row.output).pages) == 2 for row in rows)


def test_rename_suggestion_needs_manual_review_when_number_absent(tmp_path):
    pdf = document(tmp_path / "arquivo.pdf", ["COMPROVANTE DE TRANSFERENCIA", "Valor pago: R$ 80,00"])
    row = rename_suggestions([pdf])[0]
    assert row.status == "REVISAR"
    assert "SEM_NUMERO" in row.detail


def test_corrupt_pdf_is_reported_instead_of_crashing_batch(tmp_path):
    receipt = document(tmp_path / "comprovante.pdf", ["PAGAMENTO APROVADO"])
    broken = tmp_path / "NF_999.pdf"
    broken.write_bytes(b"isto nao e um pdf")
    rows = merge_receipt_to_many(receipt, [broken], tmp_path / "saida")
    assert rows[0].status == "ERRO"
