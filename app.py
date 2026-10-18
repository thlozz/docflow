"""DocFlow web interface — v0.8."""
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

/* Decorative canvas and scroll storytelling */
.df-hero {min-height:340px;display:flex;flex-direction:column;justify-content:center;padding-right:34%;background:radial-gradient(circle at 83% 48%,rgba(244,63,94,.24),transparent 34%),linear-gradient(118deg,#0c1220,#20283b 65%,#481d2a)}
.df-hero::after {display:none}
.df-hero-art {position:absolute;right:0;top:0;height:100%;width:40%;display:flex;align-items:center;justify-content:center;pointer-events:none;opacity:.97}
.df-hero-art svg {width:100%;max-width:370px;height:auto;filter:drop-shadow(0 20px 35px rgba(0,0,0,.2))}
.df-hero .df-eyebrow,.df-hero h1,.df-hero p {position:relative;z-index:1}
.df-scroll-hint {color:#fda4af;font-size:.75rem;letter-spacing:.13em;font-weight:700;margin-top:26px;z-index:1}
.df-story {padding:48px 10px 12px}
.df-story-title {font-size:clamp(1.5rem,2.8vw,2.5rem);font-weight:800;letter-spacing:-.05em;color:#111827;margin:6px 0 14px}
.df-story-subtitle {color:#64748b;max-width:620px;margin:0 0 26px}
.df-timeline {display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:18px;position:relative}
.df-step {position:relative;border-radius:20px;background:white;border:1px solid #e5e7eb;padding:26px 22px;min-height:205px;box-shadow:0 9px 26px rgba(15,23,42,.035)}
.df-step-number {display:inline-block;color:#dc2626;font-size:.83rem;font-weight:800;letter-spacing:.1em;margin-bottom:24px}
.df-step h3 {margin:0 0 10px;font-size:1.15rem}
.df-step p {font-size:.9rem;color:#64748b;margin:0}
.df-showcase {display:grid;grid-template-columns:1fr 1fr;gap:30px;align-items:center;background:#111827;border-radius:24px;padding:34px 38px;margin:38px 0;color:white;overflow:hidden}
.df-showcase h2 {color:#fff!important;font-size:clamp(1.6rem,2.7vw,2.4rem);line-height:1.2}
.df-showcase p {color:#cbd5e1}
.df-preview {border:1px solid #394154;background:#1b2435;border-radius:16px;padding:18px;transform:rotate(2deg);box-shadow:0 16px 34px #080b1260}
.df-preview-bar {display:flex;gap:6px;margin-bottom:17px}
.df-preview-bar i {display:block;width:8px;height:8px;border-radius:50%;background:#fb7185}
.df-preview-bar i:nth-child(2){background:#fbbf24}.df-preview-bar i:nth-child(3){background:#4ade80}
.df-preview-row {display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #334155;padding:11px 4px;font-size:.8rem;color:#dbeafe}
.df-preview-row:last-child {border:0}
.df-preview-ok {color:#86efac;font-weight:700}
.df-preview-review {color:#fcd34d;font-weight:700}
.df-end {text-align:center;padding:35px 20px 55px}
.df-end h2 {font-size:clamp(1.5rem,2.8vw,2.5rem)}
.df-end p {color:#64748b}
@supports (animation-timeline: view()) {
  .df-card,.df-step,.df-showcase,.df-end {animation:df-appear linear both;animation-timeline:view();animation-range:entry 0% entry 90%}
  @keyframes df-appear {from {opacity:.25;translate:0 34px}to {opacity:1;translate:0 0}}
}
@media(prefers-reduced-motion:reduce) {.df-card,.df-step,.df-showcase,.df-end {animation:none!important;transition:none!important}}
@media(max-width:850px) {.df-hero {padding-right:42%}.df-timeline {grid-template-columns:1fr}.df-showcase {grid-template-columns:1fr}}
@media(max-width:650px) {.df-hero {min-height:410px;padding:30px 24px;justify-content:flex-start}.df-hero-art {top:auto;bottom:-65px;right:-55px;width:65%;height:230px;opacity:.55}.df-hero h1 {font-size:2rem!important}.df-showcase {padding:26px}.df-story {padding:30px 0 5px}}

/* V0.8 — layered 3D document scene and progressive scroll narrative */
.df-hero {isolation:isolate;min-height:430px}
.df-scene {perspective:950px;transform-style:preserve-3d;overflow:visible}
.df-glow {position:absolute;width:290px;height:290px;border-radius:50%;background:radial-gradient(circle,#fb718550,transparent 70%);filter:blur(22px);top:50%;left:50%;transform:translate(-50%,-50%)}
.df-orbit {position:absolute;border:1px solid #fda4af45;border-radius:50%;width:320px;height:320px;top:50%;left:50%;transform:translate(-50%,-50%) rotate(-25deg)}
.df-orbit-two {width:245px;height:245px;border-style:dashed;border-color:#fb718575;animation:df-orbit 35s linear infinite}
.df-document {position:absolute;box-sizing:border-box;width:172px;height:226px;top:50%;left:50%;border-radius:15px;padding:23px 20px;background:linear-gradient(145deg,#fff 0%,#f1f5f9 70%,#dbe4f0 100%);box-shadow:12px 25px 0 #11182740,22px 45px 48px #03071280,inset -5px -5px 15px #94a3b825;display:flex;flex-direction:column;gap:13px;transform-style:preserve-3d;backface-visibility:hidden}
.df-document-back {transform:translate(-90%,-61%) rotate(-16deg) rotateY(-22deg) rotateX(8deg);animation:df-float-back 6s ease-in-out infinite}
.df-document-front {transform:translate(-8%,-39%) rotate(11deg) rotateY(-17deg) rotateX(6deg);animation:df-float-front 7s ease-in-out infinite;background:linear-gradient(145deg,#fff 0%,#fff0f2 74%,#fecdd3 100%)}
.df-doc-top {display:block;width:54%;height:11px;background:#e11d48;border-radius:5px;margin-bottom:5px}
.df-doc-line {display:block;width:90%;height:6px;background:#cbd5e1;border-radius:5px}
.df-doc-line.short {width:64%}
.df-doc-chart {display:block;width:100%;height:62px;border-radius:10px;background:linear-gradient(135deg,#ffe4e6,#fda4af55);position:relative}
.df-doc-chart:after {content:"";position:absolute;inset:17px 15px;border-bottom:4px solid #e11d48;border-right:4px solid #e11d48;transform:skew(-20deg) rotate(-8deg);border-radius:3px}
.df-doc-check {position:absolute;bottom:23px;right:17px;width:39px;height:39px;border-radius:50%;display:grid;place-items:center;background:#dcfce7;color:#16a34a;font-size:24px;font-weight:800;box-shadow:0 6px 15px #14532d30}
.df-float {position:absolute;display:grid;place-items:center;border:1px solid #ffffff40;backdrop-filter:blur(14px);box-shadow:0 12px 28px #03071270;color:white;font-weight:800}
.df-float-one {top:13%;right:16%;border-radius:12px;width:65px;height:48px;background:linear-gradient(145deg,#fb7185,#be123c);transform:rotate(13deg);animation:df-bob 4s ease-in-out infinite}
.df-float-two {bottom:10%;left:13%;border-radius:50%;width:53px;height:53px;background:linear-gradient(145deg,#34d399,#047857);font-size:26px;animation:df-bob 5s ease-in-out -2s infinite}
@keyframes df-float-back {0%,100%{transform:translate(-90%,-61%) rotate(-16deg) rotateY(-22deg) rotateX(8deg)}50%{transform:translate(-90%,-68%) rotate(-12deg) rotateY(-14deg) rotateX(13deg)}}
@keyframes df-float-front {0%,100%{transform:translate(-8%,-39%) rotate(11deg) rotateY(-17deg) rotateX(6deg)}50%{transform:translate(-8%,-46%) rotate(8deg) rotateY(-8deg) rotateX(10deg)}}
@keyframes df-bob {0%,100%{translate:0 0}50%{translate:0 -15px}}
@keyframes df-orbit {to{transform:translate(-50%,-50%) rotate(335deg)}}
.df-scroll-progress {height:4px;background:linear-gradient(90deg,#e11d48,#fb7185);width:0%;position:sticky;top:0;z-index:10;border-radius:5px;box-shadow:0 1px 10px #fb718580}
.df-story {position:relative}
.df-timeline:before {content:"";position:absolute;height:3px;left:12%;right:12%;top:42px;background:linear-gradient(90deg,#fecdd3,#e11d48,#fecdd3);opacity:.6}
.df-step {z-index:1;transition:transform .3s ease,box-shadow .3s ease}
.df-step:hover {transform:translateY(-7px);box-shadow:0 18px 40px #11182715}
.df-step-number {background:#fff1f2;border:1px solid #fecdd3;border-radius:20px;padding:8px 13px}
.df-showcase {position:relative;background:radial-gradient(circle at 80% 25%,#be123c40,transparent 35%),linear-gradient(115deg,#0f172a,#20243b);min-height:350px}
.df-preview {transform:perspective(850px) rotateY(-11deg) rotateX(4deg) rotateZ(2deg);transition:transform .5s ease,box-shadow .5s ease}
.df-showcase:hover .df-preview {transform:perspective(850px) rotateY(0deg) rotateX(0deg) rotateZ(0deg);box-shadow:0 24px 48px #0007}
.df-end {background:linear-gradient(180deg,#fff,#fff1f2);border:1px solid #ffe4e6;border-radius:22px;margin-bottom:24px}
@supports (animation-timeline: scroll()) {
 .df-scroll-progress {animation:df-progress linear both;animation-timeline:scroll(root block)}
 @keyframes df-progress {to{width:100%}}
}
@supports (animation-timeline: view()) {
 .df-story-title,.df-story-subtitle,.df-step,.df-showcase,.df-end {animation:df-rise linear both;animation-timeline:view(block);animation-range:entry 0% entry 75%}
 .df-step:nth-child(2) {animation-range:entry 10% entry 85%}
 .df-step:nth-child(3) {animation-range:entry 20% entry 95%}
 .df-preview {animation:df-preview-in linear both;animation-timeline:view(block);animation-range:entry 0% cover 50%}
 @keyframes df-rise {from{opacity:.15;translate:0 65px;filter:blur(5px)}to{opacity:1;translate:0 0;filter:blur(0)}}
 @keyframes df-preview-in {from{opacity:.3;translate:50px 25px}to{opacity:1;translate:0 0}}
}
@media(max-width:850px) {.df-hero {padding-right:40%}.df-scene {width:45%}.df-document {scale:.82}.df-orbit {scale:.8}}
@media(max-width:650px) {.df-hero {min-height:450px;padding-right:24px}.df-scene {width:78%;height:245px;right:-18%;bottom:-68px;top:auto;opacity:.56}.df-document {scale:.7}.df-timeline:before {display:none}.df-showcase {min-height:0}}
@media(prefers-reduced-motion:reduce) {.df-document,.df-float,.df-orbit,.df-scroll-progress,.df-story-title,.df-story-subtitle,.df-step,.df-showcase,.df-end,.df-preview {animation:none!important;transition:none!important;filter:none!important}.df-scroll-progress {width:100%}}

/* V0.8 — consistent gradient design across every workspace */
.df-workspace-hero {position:relative;overflow:hidden;border-radius:22px;background:radial-gradient(circle at 87% 15%,#f43f5e45,transparent 30%),linear-gradient(118deg,#111827 0%,#20293b 62%,#561e32 100%);padding:34px 38px;margin:0 0 28px;box-shadow:0 14px 38px #11182719;isolation:isolate}
.df-workspace-hero:before {content:"";position:absolute;width:235px;height:235px;border:1px solid #fb718555;border-radius:45px;transform:rotate(28deg);right:-38px;top:-70px;box-shadow:0 0 0 35px #fb71850a,0 0 0 75px #fb718508}
.df-workspace-hero:after {content:"";position:absolute;width:105px;height:135px;border-radius:12px;background:linear-gradient(145deg,#ffffff30,#ffffff08);border:1px solid #ffffff25;right:105px;top:30px;transform:rotate(-16deg);box-shadow:18px 20px 0 #fb71851a;pointer-events:none}
.df-workspace-hero .df-eyebrow {color:#fda4af;position:relative;z-index:1}
.df-workspace-hero h1 {color:white!important;font-size:clamp(1.7rem,3vw,2.55rem)!important;margin:0 0 10px;max-width:760px;position:relative;z-index:1}
.df-workspace-hero p {color:#cbd5e1;max-width:660px;margin:0;position:relative;z-index:1}
.df-workspace-tip {background:linear-gradient(105deg,#fff,#fff1f2);border:1px solid #fecdd3;border-radius:15px;padding:15px 18px;margin:10px 0 24px;color:#475569}
.df-workspace-tip strong {color:#be123c}
.df-section-divider {height:1px;background:linear-gradient(90deg,transparent,#fecdd3 20%,#e11d48 50%,#fecdd3 80%,transparent);margin:30px 0}
.df-benefits {display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin:0 0 26px}
.df-benefit {padding:20px;background:linear-gradient(145deg,#fff,#fff7f8);border:1px solid #ffe4e6;border-radius:17px}
.df-benefit strong {display:block;font-size:1rem;color:#9f1239;margin-bottom:8px}
.df-benefit span {font-size:.86rem;color:#64748b}
@media(max-width:700px){.df-workspace-hero {padding:26px 22px}.df-workspace-hero:after{right:-40px;opacity:.3}.df-benefits{grid-template-columns:1fr}}

/* v0.8 — editorial product UI, restrained gradients and visual rhythm */
:root {--df-crimson:#e11d48;--df-night:#101624}
.df-hero,.df-workspace-hero,.df-showcase {border:1px solid #ffffff15}
.df-hero {background:radial-gradient(ellipse at 85% 38%,#f43f5e36,transparent 40%),linear-gradient(122deg,#0b1020 3%,#1e293b 58%,#511b33 100%)}
.df-workspace-hero {background:radial-gradient(ellipse at 88% 28%,#fb718545,transparent 35%),linear-gradient(115deg,#0d1323,#20283b 63%,#571e35)}
.df-section {display:flex;align-items:center;gap:13px}
.df-section:after {content:"";height:1px;flex:1;background:linear-gradient(90deg,#fecdd3,transparent)}
.df-card {border-radius:20px;padding:27px;background:linear-gradient(150deg,#fff 55%,#fff8f9)}
.df-card h3 {font-size:1.14rem}
.df-card .df-icon {border:1px solid #fecdd3}
.df-card:hover .df-icon {background:linear-gradient(145deg,#be123c,#fb7185);color:white;border-color:#fb7185}
.df-benefit {border:1px solid #f1e3e8;background:linear-gradient(135deg,#fff,#fff9fa)}
.df-note {background:linear-gradient(105deg,#fff,#fff7f8)}
.df-workspace-hero:after {background:linear-gradient(145deg,#ffffff2b,#ffffff05);box-shadow:18px 20px 0 #fb71851a,38px 42px 0 #ffffff0a}
.df-utility-head {display:flex;align-items:center;justify-content:space-between;gap:14px;margin:8px 0 15px;color:#64748b;font-size:.8rem;letter-spacing:.08em;text-transform:uppercase;font-weight:800}
.df-utility-head:after {content:"";height:1px;flex:1;background:linear-gradient(90deg,#fecdd3,transparent)}
.df-status {display:inline-block;padding:6px 11px;border-radius:99px;background:#fff1f2;color:#be123c;font-size:.75rem;font-weight:800;border:1px solid #fecdd3}
[data-testid="stFileUploader"] {background:linear-gradient(135deg,#fff,#fffafb);border:1px solid #f1e3e8;border-radius:16px;padding:8px}
[data-testid="stFileUploader"] section {border-radius:12px;border-color:#fda4af}
[data-testid="stDataFrame"] {border:1px solid #e5e7eb;overflow:hidden}
.stButton>button[kind="primary"],[data-testid="stDownloadButton"] button[kind="primary"] {background:linear-gradient(105deg,#be123c,#e11d48 65%,#f43f5e);border-color:#be123c;box-shadow:0 9px 22px #e11d4830}
.stButton>button[kind="primary"]:hover,[data-testid="stDownloadButton"] button[kind="primary"]:hover {background:linear-gradient(105deg,#9f1239,#be123c 65%,#e11d48);transform:translateY(-1px)}
.df-footer {margin:55px 0 12px;padding:22px 0;border-top:1px solid #e5e7eb;display:flex;justify-content:space-between;gap:16px;color:#64748b;font-size:.8rem}
.df-footer b {color:#be123c}
@media(max-width:700px){.df-footer {flex-direction:column}}
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
    st.caption("WORKSPACE / V0.8")
    st.markdown('<div class="df-section" style="color:#94a3b8">NAVEGAÇÃO</div>', unsafe_allow_html=True)
    page = st.radio("Ferramentas", ["🏠 Início", "🔎 Localizar e separar PDFs", "🔗 Juntar PDFs",
                                    "🏷️ Analisar e renomear", "📋 Mesclar por planilha"], label_visibility="collapsed")
    st.markdown('<div class="df-sidebar-note">🔒 <b>Privacidade primeiro</b><div style="margin-top:7px">Use documentos fictícios ou autorizados. Os arquivos são processados temporariamente.</div></div>', unsafe_allow_html=True)

if page == "🏠 Início":
    st.markdown("""<div class="df-scroll-progress" aria-hidden="true"></div><section class="df-hero">
      <div class="df-eyebrow">DOCUMENT WORKSPACE · V0.5</div>
      <h1>Menos tarefas repetitivas.<br>Mais tempo para o que importa.</h1>
      <p>Seu centro de automação documental. Encontre, organize e combine PDFs em lote com relatórios para cada operação.</p>\n      <div class="df-scroll-hint">↓ EXPLORE A EXPERIÊNCIA</div><div class="df-hero-art df-scene" aria-hidden="true">
<div class="df-orbit df-orbit-one"></div><div class="df-orbit df-orbit-two"></div>
<div class="df-glow"></div>
<div class="df-document df-document-back"><span class="df-doc-top"></span><span class="df-doc-line"></span><span class="df-doc-line short"></span><span class="df-doc-chart"></span><span class="df-doc-line"></span></div>
<div class="df-document df-document-front"><span class="df-doc-top"></span><span class="df-doc-line"></span><span class="df-doc-line short"></span><span class="df-doc-chart"></span><span class="df-doc-line"></span><span class="df-doc-check">✓</span></div>
<div class="df-float df-float-one">PDF</div><div class="df-float df-float-two">✓</div>
</div>
    </section>""", unsafe_allow_html=True)
    a,b,c = st.columns(3)
    a.metric("Ferramentas disponíveis", "04")
    b.metric("Formatos de saída", "PDF · ZIP · XLSX")
    c.metric("Modo de trabalho", "Na nuvem")
    st.markdown("""<div class="df-benefits">
      <div class="df-benefit"><strong>◈ Organize em lote</strong><span>Reduza o trabalho repetitivo com operações em vários documentos.</span></div>
      <div class="df-benefit"><strong>◈ Confira resultados</strong><span>Relatórios em Excel ajudam a identificar arquivos não localizados.</span></div>
      <div class="df-benefit"><strong>◈ Preserve originais</strong><span>As operações criam novos arquivos para download, sem alterar os enviados.</span></div>
    </div>""", unsafe_allow_html=True)
    st.markdown('<div class="df-section">EXPLORAR FERRAMENTAS</div>', unsafe_allow_html=True)
    st.markdown("""<div class="df-grid"><div class="df-card"><span class="df-icon"><svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-4-4"/><path d="M8 11h6M11 8v6"/></svg></span><h3>Localizar e separar</h3><p>Busque centenas de PDFs por nome ou planilha. Exporte os encontrados e um relatório de pendências.</p></div><div class="df-card"><span class="df-icon"><svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7 3v7a4 4 0 0 0 4 4h6"/><path d="m14 11 3 3-3 3"/><path d="M7 21v-3a4 4 0 0 1 4-4"/></svg></span><h3>Juntar documentos</h3><p>Combine um comprovante com várias notas, mantendo os arquivos originais intactos.</p></div><div class="df-card"><span class="df-icon"><svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 4h10l8 8-9 9-9-9z"/><circle cx="8" cy="9" r="1.5"/></svg></span><h3>Análise e nomes</h3><p>Extraia informações de PDFs textuais e gere sugestões de nomes para conferência.</p></div><div class="df-card"><span class="df-icon"><svg viewBox="0 0 24 24" width="28" height="28" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M4 9h16M10 9v12M4 15h16"/></svg></span><h3>Mesclar por planilha</h3><p>Relacione notas e comprovantes pela coluna NF e revise os casos ambíguos.</p></div></div>""", unsafe_allow_html=True)
    st.markdown('<div class="df-note"><b>Como começar:</b> escolha uma ferramenta no menu lateral. Para enviar uma pasta inteira de PDFs, compacte-a em ZIP e use <b>Localizar e separar PDFs</b>.</div>', unsafe_allow_html=True)
    st.markdown("""<section class="df-story">
      <div class="df-eyebrow">FLUXO SIMPLES</div>
      <div class="df-story-title">De uma pasta cheia de PDFs a tudo organizado.</div>
      <p class="df-story-subtitle">Um processo em três etapas. Sem precisar renomear ou localizar cada documento manualmente.</p>
      <div class="df-timeline">
        <article class="df-step"><span class="df-step-number">01 / ENVIAR</span><h3>Reúna os documentos</h3><p>Selecione vários PDFs ou envie uma pasta compactada em ZIP.</p></article>
        <article class="df-step"><span class="df-step-number">02 / LOCALIZAR</span><h3>Defina o critério</h3><p>Informe nomes, números ou uma coluna da sua planilha Excel.</p></article>
        <article class="df-step"><span class="df-step-number">03 / EXPORTAR</span><h3>Baixe os resultados</h3><p>Receba os arquivos encontrados em ZIP e um relatório de conferência.</p></article>
      </div>
    </section>
    <section class="df-showcase">
      <div><div class="df-eyebrow" style="color:#fda4af">PRÉVIA ILUSTRATIVA</div>
      <h2>Resultados claros, sem perder o controle.</h2>
      <p>Veja o que foi encontrado, identifique pendências e faça o download dos documentos organizados.</p>
      <p style="font-size:.78rem;color:#94a3b8">Os dados ao lado são fictícios e servem apenas para demonstrar a interface.</p></div>
      <div class="df-preview"><div class="df-preview-bar"><i></i><i></i><i></i></div>
      <div class="df-preview-row"><span>NF_1001.pdf</span><span class="df-preview-ok">✓ Encontrado</span></div>
      <div class="df-preview-row"><span>NF_1002.pdf</span><span class="df-preview-ok">✓ Encontrado</span></div>
      <div class="df-preview-row"><span>NF_1003.pdf</span><span class="df-preview-review">! Revisar</span></div>
      <div class="df-preview-row"><span>relatorio.xlsx</span><span>↓ Exportar</span></div></div>
    </section>
    <section class="df-end"><div class="df-eyebrow">PRONTO PARA COMEÇAR?</div>
    <h2>Seu próximo lote, mais organizado.</h2>
    <p>Selecione uma ferramenta no menu lateral e experimente com documentos fictícios.</p></section>""", unsafe_allow_html=True)
    st.caption("Projeto de portfólio • PDFs digitalizados sem texto selecionável ainda exigem OCR, não disponível nesta versão.")

elif page == "🔎 Localizar e separar PDFs":
    st.markdown("""<section class="df-workspace-hero"><div class="df-eyebrow">PESQUISA INTELIGENTE</div><h1>Localizar e separar PDFs</h1><p>Encontre documentos por nome, número ou planilha e exporte tudo em um pacote organizado.</p></section>""", unsafe_allow_html=True)
    st.write("Envie uma pasta compactada (.zip) ou vários PDFs. Escolha os nomes ou números de uma planilha e baixe apenas os documentos correspondentes.")
    st.info("Exemplo: uma pasta com 200 PDFs + planilha com coluna NF contendo 1001, 1002 e 1003 → ZIP com os documentos encontrados e relatório dos faltantes.")
    st.markdown('<div class="df-utility-head">ÁREA DE TRABALHO <span class="df-status">PROCESSAMENTO LOCAL NA SESSÃO</span></div>', unsafe_allow_html=True)
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
    st.markdown("""<section class="df-workspace-hero"><div class="df-eyebrow">COMBINAÇÃO EM LOTE</div><h1>Juntar comprovante e notas fiscais</h1><p>Anexe um comprovante a cada nota fiscal, preservando os arquivos originais.</p></section>""", unsafe_allow_html=True)
    st.write("Cria um PDF para cada nota, colocando o mesmo comprovante antes da nota. Os originais não são alterados.")
    st.markdown('<div class="df-utility-head">ÁREA DE TRABALHO <span class="df-status">PROCESSAMENTO LOCAL NA SESSÃO</span></div>', unsafe_allow_html=True)
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
    st.markdown("""<section class="df-workspace-hero"><div class="df-eyebrow">ANÁLISE DOCUMENTAL</div><h1>Analisar e sugerir nomes</h1><p>Identifique informações em PDFs textuais e gere sugestões de nomes para conferência.</p></section>""", unsafe_allow_html=True)
    st.write("Extrai campos de PDFs com texto selecionável e sugere nomes. Revise os resultados antes de utilizar.")
    st.markdown('<div class="df-utility-head">ÁREA DE TRABALHO <span class="df-status">PROCESSAMENTO LOCAL NA SESSÃO</span></div>', unsafe_allow_html=True)
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
    st.markdown("""<section class="df-workspace-hero"><div class="df-eyebrow">AUTOMAÇÃO COM EXCEL</div><h1>Mesclar documentos por planilha</h1><p>Relacione notas e comprovantes pelo número da NF e identifique casos que exigem revisão.</p></section>""", unsafe_allow_html=True)
    st.write("Associa notas e comprovantes pelo número da NF presente nos nomes dos PDFs. Casos ambíguos ficam para revisão.")
    st.markdown('<div class="df-utility-head">ÁREA DE TRABALHO <span class="df-status">PROCESSAMENTO LOCAL NA SESSÃO</span></div>', unsafe_allow_html=True)
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

# Shared footer across all workspaces.
st.markdown('<footer class="df-footer"><span><b>DocFlow</b> / Automação documental em Python</span><span>PDF · Excel · ZIP &nbsp; | &nbsp; Projeto de portfólio</span></footer>', unsafe_allow_html=True)
