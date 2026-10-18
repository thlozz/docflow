"""DocFlow web interface — v0.4."""
from pathlib import Path
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

st.set_page_config(page_title="DocFlow | Automação de documentos", page_icon="📁", layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700;800&family=Manrope:wght@500;600;700;800&display=swap');
:root {--accent:#dc2626;--ink:#111827;--muted:#64748b}
html,body,[class*="css"],[data-testid="stApp"] {font-family:'DM Sans',sans-serif}
[data-testid="stAppViewContainer"] {background:radial-gradient(ellipse at 90% 0%,#fff1f2 0%,transparent 34%),#f7f8fb}
.block-container {padding-top:2.2rem;padding-bottom:4rem;max-width:1250px}
h1,h2,h3 {font-family:'Manrope',sans-serif;letter-spacing:-.045em;color:#121826}
h1 {font-size:clamp(2rem,3.3vw,3.2rem)!important;font-weight:800!important}
h2,h3 {font-weight:800!important}
p {line-height:1.65}
[data-testid="stSidebar"] {background:linear-gradient(165deg,#111827 0%,#191b2b 65%,#27151d 100%);border-right:1px solid #2d3443}
[data-testid="stSidebar"] * {color:#e8edf6}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {color:#aeb8c9}
[data-testid="stSidebar"] [role="radiogroup"] label {padding:11px 12px;border-radius:11px;transition:background .15s}
[data-testid="stSidebar"] [role="radiogroup"] label:hover {background:#303548}
[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] {gap:6px}
[data-testid="stMetric"] {border:1px solid #e8eaf0;border-radius:18px;padding:20px;background:white;box-shadow:0 7px 24px rgba(15,23,42,.035)}
[data-testid="stMetricLabel"] {color:#64748b}
[data-testid="stMetricValue"] {font-family:'Manrope',sans-serif;font-weight:800}
[data-testid="stFileUploader"],[data-testid="stDataFrame"],[data-testid="stTextArea"] {border-radius:14px}
.stButton>button,.stDownloadButton>button {border-radius:11px;min-height:44px;font-weight:700;transition:all .15s}
.stButton>button[kind="primary"],[data-testid="stDownloadButton"] button[kind="primary"] {background:#dc2626;border-color:#dc2626;color:white;box-shadow:0 6px 16px rgba(220,38,38,.16)}
.stButton>button[kind="primary"]:hover {background:#b91c1c;border-color:#b91c1c}
div[data-testid="stAlert"] {border-radius:13px}
.df-eyebrow {font-size:.73rem;letter-spacing:.17em;text-transform:uppercase;font-weight:800;color:#ef4444;margin-bottom:12px}
.df-hero {background:linear-gradient(115deg,#111827 0%,#1f2937 62%,#471c28 100%);padding:38px 42px;border-radius:24px;position:relative;overflow:hidden;box-shadow:0 16px 42px rgba(17,24,39,.12);margin-bottom:25px}
.df-hero:after {content:"";position:absolute;width:280px;height:280px;border:55px solid rgba(248,113,113,.09);border-radius:50%;right:-85px;top:-140px}
.df-hero .df-eyebrow {color:#fca5a5}
.df-hero h1 {color:#fff!important;font-size:clamp(2rem,3.4vw,3rem)!important;max-width:680px;margin:0 0 13px}
.df-hero p {color:#cbd5e1;max-width:620px;font-size:1.03rem;margin:0}
.df-grid {display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px;margin:12px 0 28px}
.df-card {background:white;border:1px solid #e8eaf0;border-radius:18px;padding:24px;min-height:158px;box-shadow:0 8px 25px rgba(15,23,42,.035)}
.df-card .df-icon {display:inline-flex;align-items:center;justify-content:center;background:linear-gradient(145deg,#fff1f2,#ffe4e6);color:#be123c;border-radius:14px;padding:12px;margin-bottom:14px;box-shadow:inset 0 1px 0 #fff}.df-card:hover {transform:translateY(-3px);box-shadow:0 16px 36px rgba(15,23,42,.08);border-color:#fecdd3}.df-card {transition:transform .18s,box-shadow .18s,border-color .18s}
.df-card h3 {font-size:1.12rem;margin:0 0 7px;font-weight:800}
.df-card p {color:#64748b;margin:0;font-size:.92rem}
.df-section {font-size:.75rem;color:#64748b;font-weight:800;letter-spacing:.12em;text-transform:uppercase;margin:25px 0 12px}
.df-note {background:#fff;border:1px solid #e8eaf0;border-left:4px solid #ef4444;border-radius:13px;padding:16px 20px;color:#475569;margin:16px 0}
.df-brand {display:flex;align-items:center;gap:11px;font-size:1.55rem;font-weight:800;letter-spacing:-.05em;color:white;margin:10px 0 5px}
.df-logo {background:#dc2626;border-radius:12px;padding:8px 10px}
.df-sidebar-note {font-size:.78rem;color:#94a3b8;border:1px solid #384152;border-radius:12px;padding:14px;margin-top:18px}
@media(max-width:700px) {.df-grid {grid-template-columns:1fr}.df-hero {padding:26px 24px}.block-container {padding-top:1.3rem}}
</style>""", unsafe_allow_html=True)

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
    st.dataframe([{"Arquivo / termo":r.source,"Status":r.status,"Detalhes":r.detail,"Saída":r.output} for r in rows],
                 use_container_width=True, hide_index=True)

def output_actions(rows, output_dir=None, key="output"):
    result_table(rows)
    st.download_button("📊 Baixar relatório Excel", report_bytes(rows), file_name="relatorio_docflow.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"{key}_excel")
    if output_dir is not None:
        stream = io.BytesIO()
        with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
            for file in sorted(pdfs(output_dir)):
                archive.write(file, arcname=str(file.relative_to(output_dir)))
        if pdfs(output_dir):
            st.download_button("📦 Baixar PDFs processados", stream.getvalue(), file_name="docflow_processados.zip",
                               mime="application/zip", key=f"{key}_zip")

with st.sidebar:
    st.markdown('<div class="df-brand"><span class="df-logo">◈</span> DocFlow</div>', unsafe_allow_html=True)
    st.caption("WORKSPACE / V0.4")
    st.markdown('<div class="df-section" style="color:#94a3b8">NAVEGAÇÃO</div>', unsafe_allow_html=True)
    page = st.radio("Ferramentas", ["🏠 Início", "🔎 Localizar e separar PDFs", "🔗 Juntar PDFs",
                                    "🏷️ Analisar e renomear", "📋 Mesclar por planilha"], label_visibility="collapsed")
    st.markdown('<div class="df-sidebar-note">🔒 <b>Privacidade primeiro</b><div style="margin-top:7px">Use documentos fictícios ou autorizados. Os arquivos são processados temporariamente.</div></div>', unsafe_allow_html=True)

if page == "🏠 Início":
    st.markdown("""<section class="df-hero">
      <div class="df-eyebrow">DOCUMENT WORKSPACE · V0.4</div>
      <h1>Menos tarefas repetitivas.<br>Mais tempo para o que importa.</h1>
      <p>Seu centro de automação documental. Encontre, organize e combine PDFs em lote com relatórios para cada operação.</p>
    </section>""", unsafe_allow_html=True)
    a,b,c = st.columns(3)
    a.metric("Ferramentas disponíveis", "04")
    b.metric("Formatos de saída", "PDF · ZIP · XLSX")
    c.metric("Modo de trabalho", "Na nuvem")
    st.markdown('<div class="df-section">EXPLORAR FERRAMENTAS</div>', unsafe_allow_html=True)
    st.markdown("""<div class="df-grid"><div class="df-card"><span class="df-icon"><svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/><path d="M8 11h6M11 8v6"/></svg></span><h3>Localizar e separar</h3><p>Busque centenas de PDFs por nome ou planilha. Exporte os encontrados e um relatório de pendências.</p></div><div class="df-card"><span class="df-icon"><svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 3v7a4 4 0 0 0 4 4h6"/><path d="m14 11 3 3-3 3"/><path d="M7 21v-3a4 4 0 0 1 4-4"/></svg></span><h3>Juntar documentos</h3><p>Combine um comprovante com várias notas, mantendo os arquivos originais intactos.</p></div><div class="df-card"><span class="df-icon"><svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 4h10l8 8-9 9-9-9z"/><circle cx="8" cy="9" r="1.5"/></svg></span><h3>Análise e nomes</h3><p>Extraia informações de PDFs textuais e gere sugestões de nomes para conferência.</p></div><div class="df-card"><span class="df-icon"><svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M4 9h16M10 9v12M4 15h16"/></svg></span><h3>Mesclar por planilha</h3><p>Relacione notas e comprovantes pela coluna NF e revise os casos ambíguos.</p></div></div>""", unsafe_allow_html=True)
    st.markdown('<div class="df-note"><b>Como começar:</b> escolha uma ferramenta no menu lateral. Para enviar uma pasta inteira de PDFs, compacte-a em ZIP e use <b>Localizar e separar PDFs</b>.</div>', unsafe_allow_html=True)
    st.caption("Projeto de portfólio • PDFs digitalizados sem texto selecionável ainda exigem OCR, não disponível nesta versão.")

elif page == "🔎 Localizar e separar PDFs":
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
    if st.button("🔎 Localizar e separar", type="primary"):
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
        st.download_button("📦 Baixar pasta de PDFs encontrados + relatório", package,
                           file_name="docflow_documentos_separados.zip", mime="application/zip", type="primary")

elif page == "🔗 Juntar PDFs":
    st.title("Juntar comprovante e notas fiscais")
    st.write("Cria um PDF para cada nota, colocando o mesmo comprovante antes da nota. Os originais não são alterados.")
    receipt = st.file_uploader("Comprovante de pagamento", type="pdf", key="receipt")
    invoices = st.file_uploader("Notas fiscais", type="pdf", accept_multiple_files=True, key="invoices")
    if st.button("🔗 Combinar documentos", type="primary"):
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
        st.download_button("📦 Baixar PDFs combinados e relatório", package, file_name="docflow_combinados.zip", mime="application/zip", type="primary")

elif page == "🏷️ Analisar e renomear":
    st.title("Analisar e sugerir nomes")
    st.write("Extrai campos de PDFs com texto selecionável e sugere nomes. Revise os resultados antes de utilizar.")
    uploads = st.file_uploader("Documentos PDF", type="pdf", accept_multiple_files=True, key="rename_files")
    if st.button("🏷️ Analisar documentos", type="primary"):
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
        st.download_button("📊 Baixar sugestões em Excel", report_bytes(rows), file_name="sugestoes_nomes.xlsx")

elif page == "📋 Mesclar por planilha":
    st.title("Mesclar documentos por planilha")
    st.write("Associa notas e comprovantes pelo número da NF presente nos nomes dos PDFs. Casos ambíguos ficam para revisão.")
    excel = st.file_uploader("Planilha Excel (coluna NF)", type="xlsx", key="merge_excel")
    notes = st.file_uploader("Notas fiscais", type="pdf", accept_multiple_files=True, key="merge_notes")
    proofs = st.file_uploader("Comprovantes", type="pdf", accept_multiple_files=True, key="merge_proofs")
    if st.button("📋 Mesclar pela planilha", type="primary"):
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
        st.download_button("📦 Baixar PDFs e relatório", package, file_name="docflow_planilha.zip", mime="application/zip", type="primary")
