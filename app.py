from pathlib import Path
import tempfile
import zipfile
import io
import streamlit as st
from docflow.core import (merge_receipt_to_many, search_pdfs, rename_suggestions,
                          export_report, match_by_spreadsheet, Result)

st.set_page_config(page_title='DocFlow', page_icon='📄', layout='wide')
st.title('DocFlow')
st.caption('Automação de documentos • PDFs • Prestação de contas')
st.info('Projeto demonstrativo. Use documentos fictícios ou autorizados. Arquivos são processados temporariamente nesta sessão.')

def store(upload, folder):
    path = folder / Path(upload.name).name
    path.write_bytes(upload.getvalue())
    return path

def report_download(results):
    with tempfile.TemporaryDirectory() as temp:
        path = export_report(results, Path(temp) / 'relatorio.xlsx')
        st.download_button('Baixar relatório Excel', path.read_bytes(), 'relatorio_docflow.xlsx',
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def show_results(results):
    st.dataframe([vars(r) for r in results], use_container_width=True)
    report_download(results)

def download_zip(folder):
    data = io.BytesIO()
    with zipfile.ZipFile(data, 'w', zipfile.ZIP_DEFLATED) as z:
        for file in sorted(folder.glob('*.pdf')): z.write(file, arcname=file.name)
    st.download_button('Baixar PDFs processados (ZIP)', data.getvalue(), 'docflow_resultados.zip', 'application/zip')

tab_merge, tab_search, tab_rename, tab_excel = st.tabs(['Juntar PDFs', 'Buscar PDFs', 'Analisar / Renomear', 'Mesclar por planilha'])
with tab_merge:
    st.subheader('Adicionar um comprovante antes de várias notas fiscais')
    receipt = st.file_uploader('Comprovante', type='pdf', key='receipt')
    invoices = st.file_uploader('Notas fiscais', type='pdf', accept_multiple_files=True, key='invoices')
    if st.button('Processar PDFs', type='primary', key='merge'):
        if not receipt or not invoices: st.warning('Envie um comprovante e ao menos uma nota.')
        else:
            with tempfile.TemporaryDirectory() as temp:
                root = Path(temp)
                (root/'entrada').mkdir(); (root/'saida').mkdir()
                r = store(receipt, root/'entrada')
                files = []
                (root/'notas').mkdir()
                for item in invoices: files.append(store(item, root/'notas'))
                results = merge_receipt_to_many(r, files, root/'saida')
                show_results(results); download_zip(root/'saida')
with tab_search:
    st.subheader('Localizar PDFs por nome ou conteúdo')
    uploads = st.file_uploader('PDFs para pesquisar', type='pdf', accept_multiple_files=True, key='search_files')
    term = st.text_input('Termo de busca')
    year_str = st.text_input('Ano (opcional, ex.: 2025)')
    content = st.checkbox('Pesquisar também dentro do PDF (texto selecionável)')
    if st.button('Buscar', key='search'):
        if not uploads or not term.strip(): st.warning('Envie PDFs e informe um termo.')
        elif year_str and (not year_str.isdigit() or len(year_str)!=4): st.warning('Informe um ano de quatro dígitos.')
        else:
            with tempfile.TemporaryDirectory() as temp:
                folder = Path(temp)
                for item in uploads: store(item, folder)
                show_results(search_pdfs(folder, term, int(year_str) if year_str else None, content))
with tab_rename:
    st.subheader('Extrair campos e sugerir nomes (sem modificar originais)')
    uploads = st.file_uploader('PDFs para analisar', type='pdf', accept_multiple_files=True, key='rename_files')
    if st.button('Analisar documentos', key='rename'):
        if not uploads: st.warning('Envie ao menos um PDF.')
        else:
            with tempfile.TemporaryDirectory() as temp:
                folder = Path(temp); files = [store(item, folder) for item in uploads]
                show_results(rename_suggestions(files))
            st.caption('Os nomes são sugestões: confira os dados antes de renomear arquivos.')
with tab_excel:
    st.subheader('Relacionar notas e comprovantes pelo número da NF')
    st.caption('A planilha precisa ter uma coluna NF. Os nomes dos PDFs devem conter o número correspondente.')
    excel = st.file_uploader('Planilha Excel', type='xlsx')
    notes = st.file_uploader('Notas fiscais', type='pdf', accept_multiple_files=True, key='excel_notes')
    proofs = st.file_uploader('Comprovantes', type='pdf', accept_multiple_files=True, key='excel_proofs')
    if st.button('Mesclar pela planilha', key='excel_merge'):
        if not excel or not notes or not proofs: st.warning('Envie planilha, notas e comprovantes.')
        else:
            with tempfile.TemporaryDirectory() as temp:
                root=Path(temp)
                for name in ('notas','comprovantes','saida'): (root/name).mkdir()
                x=store(excel,root)
                n=[store(f,root/'notas') for f in notes]
                p=[store(f,root/'comprovantes') for f in proofs]
                try:
                    results=match_by_spreadsheet(x,n,p,root/'saida')
                    show_results(results); download_zip(root/'saida')
                except Exception as exc: st.error(str(exc))
