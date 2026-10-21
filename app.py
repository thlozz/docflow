"""Interface operacional do DocFlow."""
from pathlib import Path
from importlib.metadata import PackageNotFoundError, version
import io
import re
import tempfile
import zipfile

import streamlit as st
from openpyxl import load_workbook
from docflow.core import (
    Result, clean_name, export_report, match_by_spreadsheet,
    merge_receipt_to_many, normalize, pdfs, read_pdf, rename_suggestions,
)

st.set_page_config(page_title="DocFlow | Documentos", layout="wide")
# Os seletores data-testid podem mudar entre versões do Streamlit.
st.markdown("""<style>
[data-testid="stAppViewBlockContainer"] {max-width: 1200px; padding-top: 1.5rem}
[data-testid="stSidebar"] {border-right: 1px solid #777777}
[data-testid="stDataFrame"] {border-radius: 4px}
[data-testid="stFileUploader"] {border-radius: 4px}
[data-testid="stMetric"] {padding-bottom: 8px}
</style>""", unsafe_allow_html=True)

try:
    APP_VERSION = version("docflow-document-automation")
except PackageNotFoundError:
    APP_VERSION = "não instalada"

MAX_FILES = 1200
MAX_UNCOMPRESSED = 250 * 1024 * 1024

def persist_upload(upload, folder):
    name = Path(upload.name).name
    if not name or name in (".", ".."):
        raise ValueError("Nome de arquivo inválido.")
    path = folder / name
    if path.exists():
        raise ValueError(f"Arquivo duplicado: {name}")
    path.write_bytes(upload.getvalue())
    return path

def receive_pdfs(uploaded, target):
    target.mkdir(parents=True, exist_ok=True)
    if uploaded is None:
        return 0
    if uploaded.name.lower().endswith(".zip"):
        with zipfile.ZipFile(io.BytesIO(uploaded.getvalue())) as archive:
            entries = [entry for entry in archive.infolist() if not entry.is_dir() and entry.filename.lower().endswith(".pdf")]
            if len(entries) > MAX_FILES or sum(e.file_size for e in entries) > MAX_UNCOMPRESSED:
                raise ValueError("ZIP muito grande. Limite: 1.200 PDFs e 250 MB descompactados.")
            seen = set()
            for entry in entries:
                parts = Path(entry.filename.replace("\\", "/")).parts
                if ".." in parts or entry.filename.startswith("/"):
                    raise ValueError("O ZIP contém um caminho inválido.")
                # Preserve a safe hierarchy, avoiding path traversal.
                parts = [clean_name(part) for part in parts]
                relative = Path(*parts)
                if str(relative).casefold() in seen:
                    raise ValueError("O ZIP contém nomes duplicados.")
                seen.add(str(relative).casefold())
                destination = target / relative
                destination.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(entry) as source, destination.open("wb") as output:
                    while chunk := source.read(1024 * 1024):
                        output.write(chunk)
            return len(entries)
    if uploaded.name.lower().endswith(".pdf"):
        persist_upload(uploaded, target)
        return 1
    raise ValueError("Envie um arquivo ZIP ou PDF.")

def collect_inputs(single, multiple, folder):
    count = receive_pdfs(single, folder) if single else 0
    for upload in multiple or []:
        count += receive_pdfs(upload, folder)
    return count

def get_terms(mode, text_value, spreadsheet, column):
    if mode == "Nome / lista":
        terms = [line.strip() for line in re.split(r"[,;\n]+", text_value) if line.strip()]
    else:
        if spreadsheet is None:
            raise ValueError("Envie uma planilha Excel.")
        wb = load_workbook(io.BytesIO(spreadsheet.getvalue()), read_only=True, data_only=True)
        try:
            rows = wb.active.iter_rows(values_only=True)
            headers = [normalize(str(v or "")) for v in next(rows, ())]
            key = normalize(column)
            if key not in headers:
                raise ValueError(f"Coluna '{column}' não encontrada. Colunas disponíveis: {', '.join(headers)}")
            idx = headers.index(key)
            terms = [str(row[idx]).strip().removesuffix(".0") for row in rows if len(row) > idx and row[idx] is not None]
        finally:
            wb.close()
    terms = list(dict.fromkeys(t for t in terms if t))
    if not terms:
        raise ValueError("Informe pelo menos um nome ou número para localizar.")
    if len(terms) > 2000:
        raise ValueError("Máximo de 2.000 termos por busca.")
    return terms

def matches(term, file, content, text_cache):
    filename = normalize(file.stem)
    wanted = normalize(term)
    if not wanted:
        return False
    # Numeric IDs must not match part of another number (NF 1001 != 11001).
    if wanted.isdigit():
        pattern = r"(?<!\d)" + re.escape(wanted) + r"(?!\d)"
        if re.search(pattern, filename):
            return True
    elif wanted in filename:
        return True
    if content:
        if file not in text_cache:
            text_cache[file] = normalize(read_pdf(file))
        body = text_cache[file]
        if wanted.isdigit():
            return bool(re.search(r"(?<!\d)" + re.escape(wanted) + r"(?!\d)", body))
        return wanted in body
    return False

def package_results(rows, found_files, source_root, extra_files=None):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for file in sorted(found_files):
            archive.write(file, arcname=str(Path("encontrados") / file.relative_to(source_root)))
        for name, payload in (extra_files or {}).items():
            archive.writestr(name, payload)
        with tempfile.TemporaryDirectory() as tmp:
            report = export_report(rows, Path(tmp) / "relatorio.xlsx")
            archive.write(report, arcname="relatorio.xlsx")
    return output.getvalue()

def report_bytes(rows):
    with tempfile.TemporaryDirectory() as tmp:
        return export_report(rows, Path(tmp) / "relatorio.xlsx").read_bytes()

def result_table(rows):
    ok = sum(r.status in ("OK", "ENCONTRADO", "SUGESTAO") for r in rows)
    revisar = sum(r.status in ("REVISAR", "NÃO ENCONTRADO") for r in rows)
    erro = sum(r.status == "ERRO" for r in rows)
    a, b, c = st.columns(3)
    a.metric("OK", ok)
    b.metric("REVISAR", revisar)
    c.metric("ERRO", erro)
    st.caption("OK: operação concluída. REVISAR: confira documentos ausentes ou ambíguos. ERRO: verifique a entrada e tente novamente.")
    st.dataframe([{"Arquivo / termo":r.source,"Status":r.status,"Detalhes":r.detail,"Saída":r.output} for r in rows],
                 use_container_width=True, hide_index=True)

def output_actions(rows, output_dir=None, key="output"):
    result_table(rows)
    st.download_button("Baixar relatório Excel", report_bytes(rows), file_name="relatorio_docflow.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"{key}_excel")
    if output_dir is not None:
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
            for file in sorted(pdfs(output_dir)):
                archive.write(file, arcname=str(file.relative_to(output_dir)))
        if pdfs(output_dir):
            st.download_button("Baixar PDFs processados", stream.getvalue(), file_name="docflow_processados.zip",
                               mime="application/zip", key=f"{key}_zip")

with st.sidebar:
    st.title("DocFlow")
    st.caption(f"Versão {APP_VERSION}")
    page = st.radio("Ferramentas", [
        "Início", "Localizar e separar PDFs", "Juntar PDFs",
        "Analisar e renomear", "Mesclar por planilha",
    ])
    st.caption("Use apenas documentos fictícios ou autorizados.")

if page == "Início":
    st.title("DocFlow")
    st.write("Operações de conferência e processamento documental para prestação de contas.")
    st.subheader("Ferramentas")
    st.dataframe([
        {"Ferramenta": "Localizar e separar PDFs", "Recebe": "PDFs/ZIP e lista ou Excel", "Gera": "ZIP e relatório", "Quando usar": "Separar documentos de um lote"},
        {"Ferramenta": "Juntar PDFs", "Recebe": "Comprovante e notas", "Gera": "PDFs combinados", "Quando usar": "Anexar um comprovante a várias notas"},
        {"Ferramenta": "Analisar e renomear", "Recebe": "PDFs textuais", "Gera": "Sugestões em Excel", "Quando usar": "Conferir dados e padronizar nomes"},
        {"Ferramenta": "Mesclar por planilha", "Recebe": "Excel, notas e comprovantes", "Gera": "PDFs e relatório", "Quando usar": "Conciliar notas e comprovantes por NF"},
    ], hide_index=True, use_container_width=True)
    st.subheader("Exemplo de demonstração")
    st.write("Os documentos fictícios da pasta examples/ permitem testar as operações sem enviar arquivos próprios.")
    if st.button("Carregar exemplo"):
        st.session_state["use_example"] = True
        st.success("Exemplo selecionado. Abra Mesclar por planilha para executar.")
    st.subheader("Limites")
    st.write("1.200 PDFs por lote · 250 MB descompactados · 2.000 termos por busca · Sem OCR")

elif page == "Localizar e separar PDFs":
    st.title("Localizar e separar PDFs")
    st.write("Envie uma pasta compactada (.zip) ou vários PDFs. Escolha os nomes ou números de uma planilha e baixe apenas os documentos correspondentes.")
    st.info("Exemplo: uma pasta com 200 PDFs + planilha com coluna NF contendo 1001, 1002 e 1003 → ZIP com os documentos encontrados e relatório dos faltantes.")
    source_zip = st.file_uploader("Pasta de origem (.zip) ou PDF", type=["zip","pdf"], key="find_zip")
    uploads = st.file_uploader("Ou selecione vários PDFs", type=["pdf"], accept_multiple_files=True, key="find_multiple")
    mode = st.radio("Como deseja localizar?", ["Nome / lista", "Planilha Excel"], horizontal=True)
    if mode == "Nome / lista":
        names = st.text_area("Nomes ou números (um por linha, ou separados por vírgula)", placeholder="1001\n1002\nFornecedor ABC")
        excel = None
        column = "NF"
    else:
        excel = st.file_uploader("Planilha de referência (.xlsx)", type="xlsx", key="find_excel")
        column = st.text_input("Nome da coluna a procurar", value="NF")
        names = ""
    include_text = st.checkbox("Pesquisar também no texto interno dos PDFs (mais lento)", value=False)
    if st.button("Localizar e separar", type="primary"):
        try:
            terms = get_terms(mode, names, excel, column)
            with tempfile.TemporaryDirectory() as tmp:
                folder = Path(tmp) / "origem"
                count = collect_inputs(source_zip, uploads, folder)
                if not count:
                    st.warning("Envie um ZIP com PDFs ou selecione os arquivos.")
                else:
                    files = pdfs(folder)
                    cache = {}
                    found = set()
                    rows = []
                    for term in terms:
                        hits = []
                        errors = []
                        for file in files:
                            try:
                                if matches(term, file, include_text, cache):
                                    hits.append(file)
                            except Exception as exc:
                                errors.append(f"{file.name}: {exc}")
                        found.update(hits)
                        if hits:
                            rows.append(Result(term, "ENCONTRADO", f"{len(hits)} arquivo(s)", "; ".join(str(p.relative_to(folder)) for p in hits[:8])))
                        elif errors:
                            rows.append(Result(term, "REVISAR", f"Não localizado; {len(errors)} PDF(s) não puderam ser lidos"))
                        else:
                            rows.append(Result(term, "NÃO ENCONTRADO", "Nenhum PDF correspondente"))
                    package = package_results(rows, found, folder)
                    st.session_state["find_result"] = (rows, package, count, len(found))
        except Exception as exc:
            st.error(f"Não foi possível processar: {exc}")
    if "find_result" in st.session_state:
        rows, package, count, nfound = st.session_state["find_result"]
        a,b,c = st.columns(3)
        a.metric("PDFs analisados", count)
        b.metric("PDFs separados", nfound)
        c.metric("Termos não localizados", sum(r.status != "ENCONTRADO" for r in rows))
        result_table(rows)
        st.download_button("Baixar pasta de PDFs encontrados + relatório", package,
                           file_name="docflow_documentos_separados.zip", mime="application/zip", type="primary")

elif page == "Juntar PDFs":
    st.title("Juntar comprovante e notas fiscais")
    st.write("Cria um PDF para cada nota, colocando o mesmo comprovante antes da nota. Os originais não são alterados.")
    receipt = st.file_uploader("Comprovante de pagamento", type="pdf", key="receipt")
    invoices = st.file_uploader("Notas fiscais", type="pdf", accept_multiple_files=True, key="invoices")
    if st.button("Combinar documentos", type="primary"):
        if not receipt or not invoices:
            st.warning("Envie um comprovante e pelo menos uma nota fiscal.")
        else:
            with tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                for name in ("comprovante","notas","saida"): (root/name).mkdir()
                try:
                    proof = persist_upload(receipt, root/"comprovante")
                    notes = [persist_upload(file, root/"notas") for file in invoices]
                    rows = merge_receipt_to_many(proof, notes, root/"saida")
                    package = package_results(rows, set(pdfs(root/"saida")), root/"saida")
                    st.session_state["merge_result"] = (rows, package)
                except Exception as exc:
                    st.error(f"Falha ao combinar: {exc}")
    if "merge_result" in st.session_state:
        rows, package = st.session_state["merge_result"]
        result_table(rows)
        st.download_button("Baixar PDFs combinados e relatório", package, file_name="docflow_combinados.zip", mime="application/zip", type="primary")

elif page == "Analisar e renomear":
    st.title("Analisar e sugerir nomes")
    st.write("Extrai campos de PDFs com texto selecionável e sugere nomes. Revise os resultados antes de utilizar.")
    uploads = st.file_uploader("Documentos PDF", type="pdf", accept_multiple_files=True, key="rename_files")
    if st.button("Analisar documentos", type="primary"):
        if not uploads:
            st.warning("Envie pelo menos um PDF.")
        else:
            with tempfile.TemporaryDirectory() as tmp:
                try:
                    folder = Path(tmp)
                    paths = [persist_upload(file, folder) for file in uploads]
                    st.session_state["rename_result"] = rename_suggestions(paths)
                except Exception as exc:
                    st.error(f"Erro na análise: {exc}")
    if "rename_result" in st.session_state:
        rows = st.session_state["rename_result"]
        result_table(rows)
        st.download_button("Baixar sugestões em Excel", report_bytes(rows), file_name="sugestoes_nomes.xlsx")

elif page == "Mesclar por planilha":
    st.title("Mesclar documentos por planilha")
    st.write("Associa notas e comprovantes pelo número da NF presente nos nomes dos PDFs. Casos ambíguos ficam para revisão.")
    excel = st.file_uploader("Planilha Excel (coluna NF)", type="xlsx", key="merge_excel")
    notes = st.file_uploader("Notas fiscais", type="pdf", accept_multiple_files=True, key="merge_notes")
    proofs = st.file_uploader("Comprovantes", type="pdf", accept_multiple_files=True, key="merge_proofs")
    if st.button("Mesclar pela planilha", type="primary"):
        if not excel or not notes or not proofs:
            st.warning("Envie a planilha, as notas e os comprovantes.")
        else:
            with tempfile.TemporaryDirectory() as tmp:
                root = Path(tmp)
                for name in ("notas","comprovantes","saida"): (root/name).mkdir()
                try:
                    xlsx = persist_upload(excel, root)
                    n = [persist_upload(file, root/"notas") for file in notes]
                    p = [persist_upload(file, root/"comprovantes") for file in proofs]
                    rows = match_by_spreadsheet(xlsx, n, p, root/"saida")
                    package = package_results(rows, set(pdfs(root/"saida")), root/"saida")
                    st.session_state["spreadsheet_result"] = (rows, package)
                except Exception as exc:
                    st.error(f"Falha na mesclagem: {exc}")
    if "spreadsheet_result" in st.session_state:
        rows, package = st.session_state["spreadsheet_result"]
        result_table(rows)
        st.download_button("Baixar PDFs e relatório", package, file_name="docflow_planilha.zip", mime="application/zip", type="primary")

