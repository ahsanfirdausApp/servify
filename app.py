import streamlit as st
import pandas as pd
import numpy as np
import io, os, re, json, copy
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Servify", page_icon="🧩", layout="wide", initial_sidebar_state="collapsed")

# ----------------------------- UI -----------------------------
st.markdown("""
<style>
.block-container {padding-top:1.2rem;padding-bottom:2.5rem;max-width:1280px}
.hero {padding:28px;border-radius:18px;background:linear-gradient(120deg,#102c26,#176b52);color:white}
.hero h1 {font-size:clamp(1.55rem,4vw,2.25rem);margin:0;line-height:1.2}
.hero p {color:#d8eee6;margin:8px 0 0;line-height:1.5}
.service-card {border:1px solid rgba(128,128,128,.25);border-radius:15px;padding:17px;min-height:142px;height:100%;overflow-wrap:anywhere}
.service-card h3 {font-size:1.05rem;line-height:1.3}
.service-card p {line-height:1.45}
.muted {color:#888;font-size:.9rem}
div[data-testid="stMetric"] {border:1px solid rgba(128,128,128,.2);padding:12px;border-radius:12px}
[data-testid="stFileUploader"] {width:100%}
[data-testid="stDataFrame"],[data-testid="stTable"] {max-width:100%;overflow-x:auto}
.stDownloadButton,.stButton {width:100%}
.stDownloadButton button,.stButton button {min-height:2.8rem;border-radius:10px}
textarea,input,[data-baseweb="select"] {font-size:16px !important}
button[role="tab"] {white-space:normal;min-height:2.8rem}
@media (max-width:700px) {
 .block-container {padding:.75rem .75rem 2rem}
 .hero {padding:20px 17px;border-radius:14px}
 .hero p {font-size:.93rem}
 .service-card {padding:14px;min-height:0;margin-bottom:8px}
 .service-card h3 {font-size:1rem}
 div[data-testid="stHorizontalBlock"] {gap:.55rem !important;flex-wrap:wrap !important}
 div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {min-width:100% !important;flex:1 1 100% !important}
 div[data-testid="stMetric"] {padding:10px}
 div[data-testid="stMetricValue"] {font-size:1.35rem}
 [data-testid="stTabs"] [role="tablist"] {gap:.25rem;overflow-x:auto;flex-wrap:nowrap;scrollbar-width:thin}
 [data-testid="stTabs"] button[role="tab"] {flex:0 0 auto;min-width:max-content;padding:.55rem .75rem;font-size:.84rem}
 [data-testid="stSidebar"] {min-width:min(85vw,20rem)}
 .stButton button,.stDownloadButton button {width:100%;min-height:2.9rem}
 [data-testid="stDataFrame"] {font-size:.82rem}
 h1 {font-size:1.55rem !important}
 h2 {font-size:1.3rem !important}
 h3 {font-size:1.1rem !important}
}
@media (max-width:390px) {
 .block-container {padding-left:.55rem;padding-right:.55rem}
 .hero {padding:17px 14px}
 [data-testid="stTabs"] button[role="tab"] {font-size:.78rem;padding:.5rem .6rem}
}
</style>
""", unsafe_allow_html=True)

if "last_result" not in st.session_state: st.session_state.last_result = None
if "last_log" not in st.session_state: st.session_state.last_log = []
if "last_filename" not in st.session_state: st.session_state.last_filename = "servify_result.xlsx"

def read_uploaded(upload):
    upload.seek(0)
    name = upload.name.lower()
    if name.endswith(".csv"):
        return {"Data": pd.read_csv(upload)}
    wb = pd.ExcelFile(upload)
    return {sheet: pd.read_excel(upload, sheet_name=sheet) for sheet in wb.sheet_names}

def clean_text(df):
    out = df.copy()
    for col in out.select_dtypes(include=["object"]).columns:
        out[col] = out[col].map(lambda x: re.sub(r"\s+", " ", x).strip() if isinstance(x, str) else x)
    return out

def normalize_headers(df):
    out = df.copy()
    out.columns = [re.sub(r"\s+", " ", str(c)).strip() for c in out.columns]
    return out

def auto_format_workbook(xlsx_bytes):
    xlsx_bytes.seek(0)
    wb = load_workbook(xlsx_bytes)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        if ws.max_row and ws.max_column:
            ws.auto_filter.ref = ws.dimensions
        header_fill = PatternFill("solid", fgColor="176B52")
        header_font = Font(bold=True, color="FFFFFF")
        line = Side(style="thin", color="D9E2DF")
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in ws.iter_rows():
            for cell in row:
                cell.border = Border(bottom=line)
                if cell.row > 1:
                    cell.alignment = Alignment(vertical="center")
        for col in ws.columns:
            letter = get_column_letter(col[0].column)
            max_len = max([len(str(c.value)) if c.value is not None else 0 for c in col] + [10])
            ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 42)
        ws.row_dimensions[1].height = 25
    out = io.BytesIO()
    wb.save(out); out.seek(0)
    return out.getvalue()

def export_sheets(sheets, formatted=True):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, df in sheets.items():
            safe = str(name)[:31] or "Data"
            df.to_excel(writer, sheet_name=safe, index=False)
    bio.seek(0)
    if formatted:
        return auto_format_workbook(bio)
    return bio.getvalue()

def find_col(df, phrase):
    phrase = phrase.strip().strip('"').strip("'")
    exact = [c for c in df.columns if str(c).lower() == phrase.lower()]
    if exact: return exact[0]
    partial = [c for c in df.columns if phrase.lower() in str(c).lower()]
    return partial[0] if len(partial) == 1 else None

def apply_instruction(sheets, instruction):
    """Rule-based Indonesian instruction planner. Returns transformed sheets + audit log."""
    result = {k: v.copy() for k, v in sheets.items()}
    log = []
    raw = instruction.lower()
    target_name = next(iter(result))
    # User can identify a sheet with "di sheet Nama"
    m_sheet = re.search(r"(?:di|pada)\s+sheet\s+['\"]?([^,'\".]+)", instruction, re.I)
    if m_sheet:
        candidate = m_sheet.group(1).strip()
        matched = [s for s in result if s.lower() == candidate.lower()]
        if matched: target_name = matched[0]
    df = normalize_headers(result[target_name])

    # Header whitespace normalization
    if not df.columns.equals(result[target_name].columns):
        log.append("Menormalkan spasi pada nama kolom.")
    result[target_name] = df

    # trim / normalize text
    if any(x in raw for x in ["rapikan spasi", "hapus spasi berlebih", "bersihkan spasi", "trim"]):
        df = clean_text(df); log.append("Membersihkan spasi berlebih pada seluruh kolom teks.")

    # duplicate removal
    if any(x in raw for x in ["hapus duplikat", "hapus data ganda", "hilangkan duplikat", "remove duplicates"]):
        before = len(df)
        subset = None
        m = re.search(r"duplikat(?:\s+berdasarkan|\s+berdasar)\s+['\"]?([^,'\".]+)", instruction, re.I)
        if m:
            c = find_col(df, m.group(1))
            if c: subset = [c]
        df = df.drop_duplicates(subset=subset, keep="first")
        log.append(f"Menghapus {before-len(df)} baris duplikat" + (f" berdasarkan {subset[0]}." if subset else "."))

    # remove empty rows / columns
    if any(x in raw for x in ["hapus baris kosong", "hapus baris yang kosong", "hilangkan baris kosong"]):
        before = len(df); df = df.dropna(how="all")
        log.append(f"Menghapus {before-len(df)} baris yang seluruh kolomnya kosong.")
    if any(x in raw for x in ["hapus kolom kosong", "hilangkan kolom kosong"]):
        before = len(df.columns); df = df.dropna(axis=1, how="all")
        log.append(f"Menghapus {before-len(df.columns)} kolom kosong.")

    # case normalization
    if "huruf kapital" in raw or "uppercase" in raw or "ubah ke huruf besar" in raw:
        cols = df.select_dtypes(include=["object"]).columns
        for c in cols: df[c] = df[c].map(lambda x: x.upper() if isinstance(x,str) else x)
        log.append(f"Mengubah teks menjadi huruf kapital pada {len(cols)} kolom.")
    elif "huruf kecil" in raw or "lowercase" in raw:
        cols = df.select_dtypes(include=["object"]).columns
        for c in cols: df[c] = df[c].map(lambda x: x.lower() if isinstance(x,str) else x)
        log.append(f"Mengubah teks menjadi huruf kecil pada {len(cols)} kolom.")
    elif "title case" in raw or "kapital di awal" in raw:
        cols = df.select_dtypes(include=["object"]).columns
        for c in cols: df[c] = df[c].map(lambda x: x.title() if isinstance(x,str) else x)
        log.append(f"Mengubah teks menjadi Title Case pada {len(cols)} kolom.")

    # drop a named column
    m_drop = re.search(r"(?:hapus|hilangkan)\s+kolom\s+['\"]?([^,'\".]+)", instruction, re.I)
    if m_drop:
        c = find_col(df, m_drop.group(1))
        if c:
            df = df.drop(columns=[c]); log.append(f"Menghapus kolom '{c}'.")

    # rename a column
    m_rename = re.search(r"(?:ubah nama|ganti nama|rename)\s+kolom\s+['\"]?([^'\"]+?)['\"]?\s+(?:menjadi|ke|jadi)\s+['\"]?([^,'\".]+)", instruction, re.I)
    if m_rename:
        old = find_col(df, m_rename.group(1))
        if old:
            new = m_rename.group(2).strip()
            df = df.rename(columns={old:new}); log.append(f"Mengubah nama kolom '{old}' menjadi '{new}'.")

    # fill blanks in a named column
    m_fill = re.search(r"(?:isi|ganti)\s+(?:sel\s+)?kosong\s+(?:di|pada)\s+kolom\s+['\"]?([^'\"]+?)['\"]?\s+(?:dengan|menjadi)\s+['\"]?([^,'\".]+)", instruction, re.I)
    if m_fill:
        c = find_col(df, m_fill.group(1))
        if c:
            value = m_fill.group(2).strip()
            df[c] = df[c].fillna(value); log.append(f"Mengisi nilai kosong pada '{c}' dengan '{value}'.")

    # sort
    m_sort = re.search(r"(?:urutkan|sortir|sort)\s+(?:data\s+)?(?:berdasarkan|menurut|by)\s+['\"]?([^,'\".]+)", instruction, re.I)
    if m_sort:
        c = find_col(df, m_sort.group(1))
        if c:
            asc = not any(x in raw for x in ["terbesar ke terkecil","tertinggi ke terendah","descending","menurun","z ke a"])
            df = df.sort_values(c, ascending=asc, na_position="last")
            log.append(f"Mengurutkan data berdasarkan '{c}' ({'naik' if asc else 'turun'}).")

    # lookup between sheets: "cocokkan ... dari sheet X berdasarkan kolom Y"
    if any(x in raw for x in ["vlookup", "xlookup", "cocokkan", "cocokan", "ambil data dari sheet", "lookup"]):
        source_match = re.search(r"(?:dari|pada)\s+sheet\s+['\"]?([^,'\".]+)", instruction, re.I)
        source_sheet = None
        if source_match:
            cand = source_match.group(1).strip()
            source_sheet = next((s for s in result if s.lower() == cand.lower() and s != target_name), None)
        if source_sheet:
            source_df = normalize_headers(result[source_sheet])
            key_match = re.search(r"(?:berdasarkan|dengan\s+kunci|key)\s+kolom\s+['\"]?([^,'\".]+)", instruction, re.I)
            key_phrase = key_match.group(1).strip() if key_match else None
            left_key = find_col(df, key_phrase) if key_phrase else None
            right_key = find_col(source_df, key_phrase) if key_phrase else None
            if left_key and right_key:
                added = [c for c in source_df.columns if c != right_key and c not in df.columns]
                df = df.merge(source_df[[right_key]+added], left_on=left_key, right_on=right_key, how="left")
                log.append(f"Mencocokkan data sheet '{source_sheet}' ke '{target_name}' menggunakan kolom '{left_key}'. Kolom ditambahkan: {', '.join(map(str,added)) or 'tidak ada kolom baru'}.")
            else:
                log.append("Lookup belum dijalankan: kolom kunci tidak teridentifikasi. Periksa nama kolom dan tuliskan 'berdasarkan kolom Nama' atau pilih fitur Gabungkan Data.")
        else:
            log.append("Lookup belum dijalankan: sebutkan nama sheet sumber, misalnya 'ambil data dari sheet Referensi berdasarkan kolom ID'.")

    # numeric summary request
    if any(x in raw for x in ["hitung total", "jumlahkan", "total dari", "rata-rata", "rata rata", "average", "sum kolom"]):
        m_col = re.search(r"(?:kolom|dari)\s+['\"]?([^,'\".]+)", instruction, re.I)
        c = find_col(df, m_col.group(1)) if m_col else None
        if c and pd.api.types.is_numeric_dtype(df[c]):
            if "rata-rata" in raw or "rata rata" in raw or "average" in raw:
                val = df[c].mean()
                log.append(f"Rata-rata kolom '{c}': {val:,.2f}.")
            else:
                val = df[c].sum()
                log.append(f"Total kolom '{c}': {val:,.2f}.")
        elif c:
            log.append(f"Kolom '{c}' bukan numerik. Ubah tipe data terlebih dahulu untuk menghitung.")
        else:
            log.append("Permintaan perhitungan terdeteksi, tetapi nama kolom belum jelas.")

    result[target_name] = df.reset_index(drop=True)
    if not log:
        log.append("Belum ada tindakan otomatis yang cocok dengan instruksi. Gunakan kata kerja spesifik atau fitur manual.")
    return result, log, target_name

def file_download_button(sheets, filename):
    data = export_sheets(sheets)
    st.download_button("⬇️ Download hasil Excel", data=data, file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

# ----------------------------- Navigation -----------------------------
with st.sidebar:
    st.title("SERVIFY")
    st.caption("One place. Multiple solutions.")
    page = st.radio("Navigasi", ["Beranda", "Excel & Spreadsheet", "Layanan lainnya", "Tentang"])
    st.divider()
    st.caption("MVP v2 · Spreadsheet workspace")

if page == "Beranda":
    st.markdown('<div class="hero"><h1>Solusi pekerjaan digital, dalam satu tempat.</h1><p>Mulai dari spreadsheet hingga berbagai kebutuhan digital bisnis dan personal.</p></div>', unsafe_allow_html=True)
    st.write("")
    st.subheader("Pilih layanan")
    cols = st.columns(3)
    cards = [
        ("📊","Excel & Spreadsheet","Data cleaning, formula, lookup, formatting, dan laporan.","Tersedia"),
        ("📝","Dokumen & Writing","Word, PDF, proofreading, terjemahan, dan CV.","Segera hadir"),
        ("🎨","Design & Creative","Desain visual, presentasi, dan kebutuhan konten.","Segera hadir"),
        ("🎬","Video & Content","Video editing, subtitle, dan script konten.","Segera hadir"),
        ("⚙️","Technology & Automation","Website, coding, workflow, dan otomasi.","Segera hadir"),
        ("💼","Business Support","Administrasi, riset, dan laporan bisnis.","Segera hadir")
    ]
    for i,(icon,title,desc,status) in enumerate(cards):
        with cols[i%3]:
            st.markdown(f'<div class="service-card"><h3>{icon} {title}</h3><p>{desc}</p><span class="muted">{status}</span></div>', unsafe_allow_html=True)
            if i%3==2: st.write("")
    st.write("")
    st.info("Kategori layanan dirancang modular. Penambahan jasa baru tidak perlu mengubah mesin inti platform.")
    if st.button("Mulai dari Excel & Spreadsheet", type="primary"):
        st.session_state["nav_excel"] = True
        st.rerun()

elif page == "Excel & Spreadsheet":
    st.title("📊 Excel & Spreadsheet")
    st.caption("Unggah workbook, jelaskan pekerjaan dalam bahasa Indonesia, periksa hasil, lalu unduh.")
    st.caption("📱 Tampilan menyesuaikan layar HP. Geser tab ke samping jika semua menu belum terlihat.")
    tabs = st.tabs(["Instruksi otomatis", "Data cleaning", "Gabungkan / lookup", "Formatting", "Katalog kemampuan"])

    with tabs[0]:
        st.subheader("Jelaskan pekerjaanmu")
        st.write("Kamu bisa menggabungkan beberapa instruksi dalam satu permintaan.")
        upload = st.file_uploader("Upload file Excel atau CSV", type=["xlsx","xlsm","csv"], key="nl_upload")
        if upload:
            try:
                sheets = read_uploaded(upload)
                names = list(sheets.keys())
                chosen = st.selectbox("Sheet yang akan diproses", names)
                st.write(f"**Preview: {chosen}** · {len(sheets[chosen]):,} baris × {len(sheets[chosen].columns)} kolom")
                st.dataframe(sheets[chosen].head(20), use_container_width=True)
                prompt = st.text_area("Instruksi pengerjaan", height=135, placeholder="Contoh: rapikan spasi, hapus duplikat berdasarkan kolom ID, hapus baris kosong, lalu urutkan berdasarkan kolom Tanggal dari terbaru ke terlama.")
                st.caption("Contoh lain: ubah nama kolom Harga Lama menjadi Harga; isi kosong pada kolom Kota dengan Tidak diketahui.")
                if st.button("Analisis & proses instruksi", type="primary", disabled=not prompt.strip()):
                    # Ensure selected sheet is first so instruction applies to it.
                    ordered = {chosen:sheets[chosen]}
                    ordered.update({k:v for k,v in sheets.items() if k != chosen})
                    processed, logs, target = apply_instruction(ordered, prompt)
                    st.session_state.last_result = processed
                    st.session_state.last_log = logs
                    st.session_state.last_filename = os.path.splitext(upload.name)[0] + "_servify.xlsx"
                    st.session_state.last_source = upload.name
                    st.success("Instruksi diproses. Periksa ringkasan dan preview sebelum mengunduh.")
                if st.session_state.last_result is not None:
                    st.divider()
                    st.subheader("Ringkasan pengerjaan")
                    for item in st.session_state.last_log:
                        st.write("• " + item)
                    target = list(st.session_state.last_result.keys())[0]
                    before_rows = len(sheets[chosen]); after_rows = len(st.session_state.last_result[target])
                    m1,m2,m3 = st.columns(3)
                    m1.metric("Baris sebelum", f"{before_rows:,}")
                    m2.metric("Baris sesudah", f"{after_rows:,}", delta=f"{after_rows-before_rows:,}")
                    m3.metric("Jumlah sheet", len(st.session_state.last_result))
                    st.dataframe(st.session_state.last_result[target].head(30), use_container_width=True)
                    file_download_button(st.session_state.last_result, st.session_state.last_filename)
            except Exception as e:
                st.error(f"File tidak dapat diproses: {e}")

    with tabs[1]:
        st.subheader("Data cleaning")
        up = st.file_uploader("Upload CSV atau Excel", type=["xlsx","xlsm","csv"], key="clean_v2")
        if up:
            try:
                data = read_uploaded(up)
                sheet = st.selectbox("Pilih sheet", list(data.keys()), key="clean_sheet")
                df = data[sheet]
                st.dataframe(df.head(25), use_container_width=True)
                c1,c2,c3 = st.columns(3)
                with c1: do_trim=st.checkbox("Rapikan spasi",True,key="trim2")
                with c2: do_dup=st.checkbox("Hapus duplikat",False,key="dup2")
                with c3: do_blank=st.checkbox("Hapus baris kosong",True,key="blank2")
                if st.button("Jalankan cleaning",type="primary"):
                    out=df.copy(); old=len(out)
                    if do_trim: out=clean_text(out)
                    if do_dup: out=out.drop_duplicates()
                    if do_blank: out=out.dropna(how="all")
                    data[sheet]=out.reset_index(drop=True)
                    st.success(f"Cleaning selesai. {old-len(out):,} baris dihapus.")
                    st.dataframe(out.head(30),use_container_width=True)
                    file_download_button(data,os.path.splitext(up.name)[0]+"_cleaned.xlsx")
            except Exception as e: st.error(str(e))

    with tabs[2]:
        st.subheader("Gabungkan dua tabel / kebutuhan lookup")
        left=st.file_uploader("File tabel utama",type=["xlsx","csv"],key="left_v2")
        right=st.file_uploader("File tabel referensi",type=["xlsx","csv"],key="right_v2")
        if left and right:
            try:
                l=read_uploaded(left); r=read_uploaded(right)
                ldf=next(iter(l.values())); rdf=next(iter(r.values()))
                st.write("Tabel utama"); st.dataframe(ldf.head(10),use_container_width=True)
                st.write("Tabel referensi"); st.dataframe(rdf.head(10),use_container_width=True)
                a,b=st.columns(2)
                with a: lk=st.selectbox("Kolom kunci tabel utama",list(ldf.columns))
                with b: rk=st.selectbox("Kolom kunci tabel referensi",list(rdf.columns))
                how=st.selectbox("Metode pencocokan",["left","inner","outer","right"],format_func=lambda x:{"left":"Semua data utama + data yang cocok","inner":"Hanya data yang cocok","outer":"Semua data dari kedua tabel","right":"Semua data referensi + data yang cocok"}[x])
                if st.button("Cocokkan dan gabungkan",type="primary"):
                    out=ldf.merge(rdf,left_on=lk,right_on=rk,how=how,suffixes=("_utama","_referensi"))
                    st.success(f"Selesai: {len(out):,} baris.")
                    st.dataframe(out.head(35),use_container_width=True)
                    file_download_button({"Hasil":out},"hasil_lookup_servify.xlsx")
            except Exception as e: st.error(f"Gagal menggabungkan: {e}")

    with tabs[3]:
        st.subheader("Formatting workbook")
        upf=st.file_uploader("Upload workbook .xlsx",type=["xlsx"],key="fmt_v2")
        if upf:
            try:
                wb=load_workbook(upf)
                sheet=st.selectbox("Pilih sheet untuk melihat struktur",wb.sheetnames,key="fmt_sheet")
                ws=wb[sheet]
                st.write(f"{ws.max_row:,} baris × {ws.max_column:,} kolom")
                st.write("Terapkan header profesional, filter, freeze panes, border ringan, alignment, dan lebar kolom otomatis pada seluruh sheet.")
                if st.button("Rapikan seluruh workbook",type="primary"):
                    bio=io.BytesIO(); wb.save(bio); bio.seek(0)
                    formatted=auto_format_workbook(bio)
                    st.success("Formatting diterapkan pada seluruh sheet.")
                    st.download_button("⬇️ Download workbook",formatted,os.path.splitext(upf.name)[0]+"_formatted.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
            except Exception as e: st.error(str(e))

    with tabs[4]:
        st.subheader("Peta kemampuan Excel & Spreadsheet")
        catalog = {
            "Matematika & agregasi":"SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ROUND, ROUNDUP, ROUNDDOWN, ABS, SUBTOTAL, AGGREGATE",
            "Lookup & reference":"VLOOKUP, HLOOKUP, XLOOKUP, LOOKUP, INDEX, MATCH, XMATCH, CHOOSE, INDIRECT",
            "Logika":"IF, IFS, IFERROR, IFNA, AND, OR, NOT, SWITCH, XOR",
            "Teks":"LEFT, RIGHT, MID, LEN, TRIM, CLEAN, UPPER, LOWER, PROPER, CONCAT, TEXTJOIN, FIND, SEARCH, SUBSTITUTE, REPLACE, TEXT, VALUE, TEXTBEFORE, TEXTAFTER, TEXTSPLIT",
            "Tanggal & waktu":"TODAY, NOW, DATE, DAY, MONTH, YEAR, DATEDIF, EDATE, EOMONTH, WORKDAY, NETWORKDAYS, WEEKDAY, WEEKNUM, YEARFRAC",
            "Perhitungan bersyarat":"SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, MAXIFS, MINIFS, SUMPRODUCT",
            "Dynamic arrays":"FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, TAKE, DROP, CHOOSECOLS, CHOOSEROWS, VSTACK, HSTACK, WRAPROWS, LET, LAMBDA, MAP, REDUCE, SCAN",
            "Statistik & keuangan":"MEDIAN, MODE, STDEV, VAR, RANK, PERCENTILE, QUARTILE, PMT, FV, PV, NPV, IRR, RATE",
            "Data operations":"Remove duplicates, missing values, split/merge columns, append, join, pivot/unpivot, sort, filter, grouping, type conversion",
            "Workbook & automation":"Formatting, conditional formatting, data validation, Pivot Table, Power Query, VBA/Macro, charts, dashboard, protection"
        }
        for k,v in catalog.items():
            with st.expander(k): st.write(v)
        st.warning("Katalog ini adalah cakupan kemampuan yang direncanakan. Tidak semua formula sudah dieksekusi oleh MVP. Eksekusi formula Excel penuh membutuhkan calculation engine yang kompatibel dengan workbook.")

elif page == "Layanan lainnya":
    st.title("Layanan lainnya")
    st.caption("Kategori siap dikembangkan sebagai modul terpisah.")
    for title,desc in [
        ("Dokumen & Writing","Word, PDF, proofreading, terjemahan, CV"),
        ("Design & Creative","Desain grafis, presentasi, editing foto"),
        ("Video & Content","Video editing, subtitle, script konten"),
        ("Technology & Automation","Website, coding, workflow automation"),
        ("Business Support","Administrasi, riset, laporan bisnis")
    ]:
        with st.expander(title):
            st.write(desc); st.caption("Segera hadir — belum menerima pesanan.")

else:
    st.title("Tentang Servify")
    st.markdown("""
**Servify** adalah platform multi-service yang menyediakan berbagai jasa digital dalam satu tempat.

### Prinsip produk
- Satu akun dan satu sistem pesanan untuk banyak kategori jasa.
- Setiap kategori memiliki mesin pengerjaan sendiri.
- Pengguna dapat memberi instruksi dengan bahasa sehari-hari.
- Hasil pekerjaan ditampilkan untuk diperiksa sebelum diunduh.
- Kategori dan kemampuan baru dapat ditambahkan secara modular.

### Status MVP v2
Aplikasi saat ini menjalankan beberapa proses spreadsheet berbasis aturan. Ini belum merupakan layanan AI penuh, belum memiliki sistem akun, pembayaran, database pesanan, atau hosting publik.
""")
    st.info("Untuk keamanan, jangan gunakan file berisi data pribadi atau rahasia pada demo lokal ini.")
