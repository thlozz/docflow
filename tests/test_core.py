from pathlib import Path
from pypdf import PdfReader, PdfWriter
from openpyxl import Workbook
from docflow.core import merge_pair, merge_receipt_to_many, search_pdfs, match_by_spreadsheet, money, normalize

def make_pdf(path, pages=1):
    writer=PdfWriter()
    for _ in range(pages): writer.add_blank_page(width=200,height=200)
    with path.open('wb') as out: writer.write(out)

def test_merge_order(tmp_path):
    a=tmp_path/'comprovante.pdf'; b=tmp_path/'nota.pdf'; out=tmp_path/'resultado.pdf'
    make_pdf(a,2); make_pdf(b,1)
    merge_pair(a,b,out)
    assert len(PdfReader(out).pages)==3

def test_no_overwrite(tmp_path):
    a=tmp_path/'c.pdf'; b=tmp_path/'n.pdf'; make_pdf(a);make_pdf(b)
    try: merge_pair(a,b,a)
    except ValueError: pass
    else: assert False, 'expected ValueError'

def test_search_name_year(tmp_path):
    make_pdf(tmp_path/'Nota_123_2025.pdf'); make_pdf(tmp_path/'Nota_123_2024.pdf')
    result=search_pdfs(tmp_path,'nota_123',2025)
    assert len(result)==1 and result[0].source=='Nota_123_2025.pdf'

def test_money_and_accents():
    assert str(money('R$ 1.234,56'))=='1234.56'
    assert normalize('PRESTAÇÃO')=='prestacao'

def test_spreadsheet_match(tmp_path):
    n=tmp_path/'nota_NF_123.pdf'; p=tmp_path/'comprovante_123.pdf'
    make_pdf(n);make_pdf(p)
    wb=Workbook();wb.active.append(['NF']);wb.active.append([123]);excel=tmp_path/'lista.xlsx';wb.save(excel)
    results=match_by_spreadsheet(excel,[n],[p],tmp_path/'saida')
    assert results[0].status=='OK'
    assert len(PdfReader(results[0].output).pages)==2

def test_ambiguous_matches(tmp_path):
    n=tmp_path/'nota_123.pdf'; p=tmp_path/'comprovante_123.pdf'; p2=tmp_path/'comprovante_outro_123.pdf'
    for f in (n,p,p2): make_pdf(f)
    wb=Workbook();wb.active.append(['NF']);wb.active.append([123]);excel=tmp_path/'lista.xlsx';wb.save(excel)
    result=match_by_spreadsheet(excel,[n],[p,p2],tmp_path/'saida')
    assert result[0].status=='REVISAR'
