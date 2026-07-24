# kamus_app.py
# Streamlit app: Kamus Arab-Indonesia PDF → Excel via Google Gemini AI
# Cara jalankan: streamlit run kamus_app.py

import streamlit as st
import google.generativeai as genai
import fitz          # PyMuPDF
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import json, io, time, re
from PIL import Image
from datetime import datetime

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ═══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="📖 Kamus Arab-Indonesia Parser",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════════════════════════
# CSS — tema buku Arab klasik: krem hangat + tinta gelap + aksen emas
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+Arabic:wght@400;700&family=Merriweather:ital,wght@0,400;0,700;1,400&family=Inter:wght@400;500;600&display=swap');

/* ─── Global ─────────────────────────────────────── */
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.stApp {
    background: #F9F5EE;
}

/* ─── Header ─────────────────────────────────────── */
.app-header {
    background: linear-gradient(135deg, #1C2B3A 0%, #2E4057 100%);
    color: white;
    padding: 28px 36px 22px;
    border-radius: 12px;
    margin-bottom: 24px;
    border-bottom: 4px solid #C9A84C;
    box-shadow: 0 4px 20px rgba(0,0,0,0.15);
}
.app-title {
    font-family: 'Merriweather', serif;
    font-size: 1.9em;
    font-weight: 700;
    margin: 0 0 4px 0;
    letter-spacing: -0.3px;
}
.app-sub {
    color: #C9A84C;
    font-size: 0.9em;
    letter-spacing: 0.8px;
    text-transform: uppercase;
    font-weight: 500;
}

/* ─── Live Feed Cards ─────────────────────────────── */
.feed-wrap {
    max-height: 520px;
    overflow-y: auto;
    padding: 6px 2px;
    scrollbar-width: thin;
    scrollbar-color: #C9A84C #F0EBE1;
}
.entry-card {
    background: #FFFFFF;
    border: 1px solid #E8E0D0;
    border-left: 5px solid #2E4057;
    padding: 10px 14px 8px;
    margin: 6px 0;
    border-radius: 6px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    transition: all 0.2s;
}
.entry-card:hover { box-shadow: 0 2px 8px rgba(0,0,0,0.12); }
.entry-card.sublema {
    border-left-color: #C9A84C;
    margin-left: 18px;
    background: #FDFAF4;
}
.badge {
    display: inline-block;
    padding: 2px 9px;
    border-radius: 3px;
    font-size: 0.68em;
    font-weight: 700;
    letter-spacing: 0.8px;
    margin-right: 8px;
    text-transform: uppercase;
    vertical-align: middle;
}
.badge-lema    { background: #2E4057; color: #FFF; }
.badge-sublema { background: #C9A84C; color: #1C2B3A; }

.arab-text {
    font-family: 'Noto Sans Arabic', 'Traditional Arabic', serif;
    font-size: 1.35em;
    direction: rtl;
    color: #1C2B3A;
    font-weight: 700;
    vertical-align: middle;
    line-height: 1.8;
}
.card-detail {
    margin-top: 3px;
    font-size: 0.82em;
    color: #555;
    display: flex;
    flex-wrap: wrap;
    gap: 8px;
    align-items: center;
}
.pill {
    background: #EEF2F7;
    color: #2E4057;
    padding: 1px 8px;
    border-radius: 10px;
    font-size: 0.92em;
}
.pill.trans { background: #FFF8E7; color: #7A5C00; font-style: italic; }
.pill.eku   { background: #E8F5E9; color: #1B5E20; font-weight: 600; }
.pill.arti  { background: #EDE7F6; color: #4527A0; }
.pill.alt   { background: #FCE4EC; color: #880E4F; }

/* ─── Stats strip ─────────────────────────────────── */
.stat-strip {
    display: flex;
    gap: 12px;
    margin: 16px 0;
}
.stat-box {
    flex: 1;
    background: #2E4057;
    color: white;
    padding: 12px 10px;
    border-radius: 8px;
    text-align: center;
    border-bottom: 3px solid #C9A84C;
}
.stat-num   { font-size: 1.8em; font-weight: 700; line-height: 1.1; }
.stat-label { font-size: 0.72em; opacity: 0.75; text-transform: uppercase; letter-spacing: 0.5px; }

/* ─── Status badges ───────────────────────────────── */
.status-pill {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 5px 14px;
    border-radius: 20px;
    font-size: 0.85em;
    font-weight: 600;
}
.status-idle       { background: #E8ECF0; color: #2E4057; }
.status-processing { background: #FFF8DC; color: #7A5C00; }
.status-done       { background: #D6F0D8; color: #1B5E20; }
.status-error      { background: #FDECEA; color: #B71C1C; }

/* ─── Section titles ──────────────────────────────── */
.section-title {
    font-family: 'Merriweather', serif;
    font-size: 1.05em;
    color: #2E4057;
    font-weight: 700;
    border-bottom: 2px solid #C9A84C;
    padding-bottom: 4px;
    margin: 16px 0 10px;
}

/* ─── Download strip ──────────────────────────────── */
.dl-strip {
    background: linear-gradient(135deg, #1C2B3A, #2E4057);
    color: white;
    padding: 20px 24px;
    border-radius: 10px;
    text-align: center;
    margin-top: 20px;
    border-top: 3px solid #C9A84C;
}

/* ─── Streamlit overrides ─────────────────────────── */
div[data-testid="stMetricValue"] { font-size: 1.6em; color: #1C2B3A; }
.stProgress > div > div { background: #C9A84C !important; }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════
for key, default in [
    ('entries', []),
    ('feed',    []),
    ('done',    False),
    ('pdf_bytes', None),
    ('pdf_name',  ''),
    ('n_pages',   0),
]:
    if key not in st.session_state:
        st.session_state[key] = default

# ═══════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

ARAB_RE   = re.compile(r'[\u0600-\u06FF\u0750-\u077F\uFB50-\uFDFF\uFE70-\uFEFF]')
LATIN_RE  = re.compile(r'[A-Za-z]')


def build_prompt(n_pages: int) -> str:
    return f"""Kamu adalah parser ahli kamus Indonesia-Arab yang sangat teliti dan akurat.

Berikut adalah {n_pages} halaman dari kamus Indonesia-Arab.
Baca SEMUA kolom dari kiri ke kanan, atas ke bawah. Jangan lewati satu entri pun.

═══ FORMAT ENTRI DI KAMUS ═══
  kata_indonesia (keterangan) [aksara_arab] transliterasi_latin

═══ CARA MEMBEDAKAN LEMA vs SUBLEMA ═══
  • LEMA    = entri utama — posisi paling kiri di kolom, biasanya tebal/bold
  • SUBLEMA = entri turunan — menjorok ke kanan, atau diawali ~, atau diawali ::

═══ ATURAN KOLOM (WAJIB DIIKUTI KETAT) ═══
  1. LEMA          → HANYA aksara Arab. Kosong jika baris ini adalah SUBLEMA.
  2. SUBLEMA       → HANYA aksara Arab. Kosong jika baris ini adalah LEMA.
  3. TRANSLITERASI → HANYA huruf Latin (romanisasi Arab). TANPA aksara Arab.
  4. EKUIVALEN     → kata Indonesia MAKSIMAL 2 KATA sebelum [aksara_arab].
                     TIDAK BOLEH berisi aksara Arab. TIDAK BOLEH berisi transliterasi Latin.
  5. ALTERNATIF    → arti Indonesia lain jika ada pilihan ("atau" / "/" / sinonim Indonesia).
                     Kosong jika tidak ada.
  6. ARTI          → keterangan lebih dari 2 kata, atau penjelasan dalam tanda kurung (...).
                     TIDAK BOLEH berisi aksara Arab.
  7. SINONIM       → sinonim bahasa Arab atau Indonesia jika ada. Kosong jika tidak ada.

═══ CONTOH PARSING ═══
  Input : "abad [قَرْنٌ] qarnun ج"
  Output: {{"LEMA":"قَرْنٌ","SUBLEMA":"","TRANSLITERASI":"qarnun","EKUIVALEN":"abad","ALTERNATIF":"","ARTI":"","SINONIM":""}}

  Input : "  berabad-abad [لِعِدَّةِ قُرُونٍ] li 'iddati quruunin"
  Output: {{"LEMA":"","SUBLEMA":"لِعِدَّةِ قُرُونٍ","TRANSLITERASI":"li 'iddati quruunin","EKUIVALEN":"berabad-abad","ALTERNATIF":"","ARTI":"","SINONIM":""}}

  Input : "abai (me-kan) [أَهْمَلَ] ahmala- [تَرَكَ] taraka ~"
  Output: 2 baris →
    {{"LEMA":"أَهْمَلَ","SUBLEMA":"","TRANSLITERASI":"ahmala-","EKUIVALEN":"abai","ALTERNATIF":"","ARTI":"(me-kan)","SINONIM":""}}
    {{"LEMA":"","SUBLEMA":"تَرَكَ","TRANSLITERASI":"taraka","EKUIVALEN":"abai","ALTERNATIF":"","ARTI":"","SINONIM":""}}

  Input : "abdi [عَبْدٌ] 'abdun- ~ (pe-an) [خِدْمَةٌ] khidmatun"
  Output: 2 baris →
    {{"LEMA":"عَبْدٌ","SUBLEMA":"","TRANSLITERASI":"'abdun-","EKUIVALEN":"abdi","ALTERNATIF":"","ARTI":"","SINONIM":""}}
    {{"LEMA":"","SUBLEMA":"خِدْمَةٌ","TRANSLITERASI":"khidmatun","EKUIVALEN":"","ALTERNATIF":"","ARTI":"(pe-an)","SINONIM":""}}

═══ ABAIKAN ═══
  Nomor halaman, header huruf (A B C ...), garis dekoratif, daftar isi, watermark.

RETURN: HANYA array JSON, tanpa penjelasan, tanpa markdown, tanpa komentar apapun.
[{{"LEMA":"","SUBLEMA":"","TRANSLITERASI":"","EKUIVALEN":"","ALTERNATIF":"","ARTI":"","SINONIM":""}}, ...]"""


def clean_arab(v: str) -> str:
    """Pastikan hanya aksara Arab."""
    v = str(v or '').strip()
    # Hapus huruf Latin, digit, tanda baca Latin
    v = re.sub(r'[A-Za-z0-9\.\,\;\:\!\?\-\/\\]+', '', v).strip()
    return v


def clean_latin(v: str) -> str:
    """Pastikan hanya Latin (transliterasi)."""
    v = str(v or '').strip()
    # Hapus aksara Arab
    v = ARAB_RE.sub('', v).strip()
    return v


def clean_indo(v: str) -> str:
    """Pastikan hanya teks Indonesia (bukan Arab)."""
    v = str(v or '').strip()
    v = ARAB_RE.sub('', v).strip()
    return v


def validate_entry(e: dict) -> dict:
    """Bersihkan dan validasi satu entri."""
    row = {
        'LEMA':          clean_arab(e.get('LEMA',  '')),
        'SUBLEMA':       clean_arab(e.get('SUBLEMA', '')),
        'TRANSLITERASI': clean_latin(e.get('TRANSLITERASI', '')),
        'EKUIVALEN':     clean_indo(e.get('EKUIVALEN', '')),
        'ALTERNATIF':    clean_indo(e.get('ALTERNATIF', '')),
        'ARTI':          clean_indo(e.get('ARTI', '')),
        'SINONIM':       str(e.get('SINONIM', '') or '').strip(),
    }
    # Jika keduanya terisi, jadikan SUBLEMA kosong (prioritas LEMA)
    if row['LEMA'] and row['SUBLEMA']:
        row['SUBLEMA'] = ''
    return row


def is_valid_entry(row: dict) -> bool:
    """True jika entri punya konten bermakna."""
    arab = row['LEMA'] or row['SUBLEMA']
    if not arab or len(arab) < 2:
        return False
    # Filter header satu huruf
    eku = row['EKUIVALEN']
    if len(eku) == 1 and eku.isalpha():
        return False
    return True


def pdf_pages_to_images(pdf_bytes: bytes, start: int, end: int, dpi: int) -> list:
    """Konversi halaman PDF ke list PIL Image."""
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')
    imgs = []
    for i in range(start, min(end, len(doc))):
        pix = doc[i].get_pixmap(dpi=dpi)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        imgs.append(img)
    doc.close()
    return imgs


def call_gemini(model, images: list, retries: int = 2) -> list:
    """Panggil Gemini dengan retry; return list entry dict."""
    prompt = build_prompt(len(images))
    for attempt in range(retries):
        try:
            resp = model.generate_content(
                [prompt] + images,
                generation_config=genai.GenerationConfig(
                    temperature=0.05,
                    max_output_tokens=8192,
                )
            )
            raw = resp.text.strip()

            # Bersihkan markdown fence jika ada
            raw = re.sub(r'^```(?:json)?\s*', '', raw)
            raw = re.sub(r'\s*```$', '', raw).strip()

            # Ambil array JSON
            if not raw.startswith('['):
                m = re.search(r'\[.*\]', raw, re.DOTALL)
                if m:
                    raw = m.group(0)
                else:
                    raise ValueError("Respons bukan JSON array")

            data = json.loads(raw)
            results = []
            for e in data:
                if not isinstance(e, dict):
                    continue
                row = validate_entry(e)
                if is_valid_entry(row):
                    results.append(row)
            return results

        except json.JSONDecodeError as ex:
            if attempt < retries - 1:
                time.sleep(3)
                continue
            st.warning(f"⚠️ JSON error: {str(ex)[:80]}")
            return []
        except Exception as ex:
            if attempt < retries - 1:
                time.sleep(4)
                continue
            st.error(f"❌ Gemini error: {str(ex)[:150]}")
            return []
    return []


def entry_card_html(e: dict) -> str:
    """Render satu entri sebagai HTML card."""
    is_lema = bool(e.get('LEMA'))
    arab    = e.get('LEMA') or e.get('SUBLEMA', '')
    trans   = e.get('TRANSLITERASI', '')
    eku     = e.get('EKUIVALEN', '')
    alt     = e.get('ALTERNATIF', '')
    arti    = e.get('ARTI', '')

    cls   = "entry-card" if is_lema else "entry-card sublema"
    badge = f'<span class="badge badge-{"lema" if is_lema else "sublema"}">{"LEMA" if is_lema else "SUBLEMA"}</span>'

    pills = ""
    if eku:   pills += f'<span class="pill eku">📝 {eku}</span>'
    if trans: pills += f'<span class="pill trans">🔤 {trans}</span>'
    if alt:   pills += f'<span class="pill alt">≈ {alt}</span>'
    if arti:  pills += f'<span class="pill arti">ℹ {arti}</span>'

    return f"""<div class="{cls}">
  {badge}<span class="arab-text">{arab}</span>
  <div class="card-detail">{pills}</div>
</div>"""


def make_excel(entries: list) -> bytes:
    """Buat file Excel terformat dari list entri."""
    COLS = ['LEMA','SUBLEMA','TRANSLITERASI','EKUIVALEN','ALTERNATIF','ARTI','SINONIM']
    df = pd.DataFrame(entries, columns=COLS) if entries else pd.DataFrame(columns=COLS)

    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine='openpyxl')
    buf.seek(0)

    wb = openpyxl.load_workbook(buf)
    ws = wb.active

    # Header
    hfill = PatternFill('solid', fgColor='1C2B3A')
    hfont = Font(name='Calibri', bold=True, color='C9A84C', size=11)
    for cell in ws[1]:
        cell.fill = hfill
        cell.font = hfont
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Lebar kolom
    for col, w in zip('ABCDEFG', [22, 22, 25, 20, 20, 42, 18]):
        ws.column_dimensions[col].width = w

    # Font
    f_arab  = Font(name='Noto Sans Arabic', size=13)
    f_ital  = Font(name='Calibri', size=11, italic=True)
    f_body  = Font(name='Calibri', size=11)
    z_fill  = PatternFill('solid', fgColor='F9F5EE')
    z_fill2 = PatternFill('solid', fgColor='FFFFFF')

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        ridx = row[0].row
        for cell in row:
            cl = cell.column_letter
            # Zebra stripes
            cell.fill = z_fill if ridx % 2 == 0 else z_fill2
            if cl in ('A', 'B'):          # LEMA / SUBLEMA — Arab
                if cell.value:
                    cell.font = f_arab
                cell.alignment = Alignment(
                    horizontal='right', vertical='center', reading_order=2
                )
            elif cl == 'C':               # TRANSLITERASI — italic
                cell.font = f_ital
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:                         # Indonesia
                cell.font = f_body
                cell.alignment = Alignment(
                    horizontal='left', vertical='center',
                    wrap_text=(cl == 'F')
                )

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
# UI — HEADER
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="app-header">
  <div class="app-title">📖 Kamus Arab-Indonesia Parser</div>
  <div class="app-sub">Ekstraksi otomatis PDF → Excel · Powered by Google Gemini AI</div>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚙️ Konfigurasi")

    api_key = st.text_input(
        "🔑 Gemini API Key",
        type="password",
        placeholder="AIza...",
        help="Buat di https://aistudio.google.com/app/apikey",
    )

    model_name = st.selectbox(
        "🤖 Model Gemini",
        ["gemini-3.5-flash", "gemini-3.1-pro", "gemini-3.1-flash-lite"],
        help="3.5-flash: tercepat | 3.1-pro: paling akurat",
    )

    st.markdown("---")

    uploaded = st.file_uploader(
        "📄 Upload PDF Kamus",
        type=["pdf"],
        help="Upload file PDF kamus Indonesia-Arab",
    )

    if uploaded:
        # Simpan ke session state agar tidak terbaca ulang saat rerun
        if st.session_state.pdf_name != uploaded.name:
            st.session_state.pdf_bytes = uploaded.read()
            st.session_state.pdf_name  = uploaded.name
            raw_doc = fitz.open(stream=st.session_state.pdf_bytes, filetype='pdf')
            st.session_state.n_pages = len(raw_doc)
            raw_doc.close()
            # Reset hasil lama jika PDF baru
            st.session_state.entries = []
            st.session_state.feed    = []
            st.session_state.done    = False

        n_pages = st.session_state.n_pages
        st.success(f"✅ **{n_pages} halaman** dimuat")
        st.caption(f"Ukuran: {round(len(st.session_state.pdf_bytes)/1024/1024,1)} MB")

    st.markdown("---")

    batch_size = st.select_slider(
        "📦 Halaman per batch",
        options=[3, 5, 7, 10],
        value=5,
        help="5 = standar | 10 = lebih cepat tapi mungkin kurang akurat",
    )

    dpi = st.select_slider(
        "🔍 Resolusi gambar (DPI)",
        options=[100, 120, 150, 200],
        value=150,
        help="150 = standar yang bagus",
    )

    # Range halaman
    p_start, p_end = 1, st.session_state.n_pages or 1
    if uploaded and st.session_state.n_pages > 0:
        use_range = st.checkbox("📌 Proses range halaman tertentu")
        if use_range:
            p_start, p_end = st.slider(
                "Range halaman",
                min_value=1,
                max_value=st.session_state.n_pages,
                value=(1, min(10, st.session_state.n_pages)),
            )

    delay_s = st.slider(
        "⏱️ Jeda antar batch (detik)",
        min_value=1, max_value=10, value=3,
        help="Mencegah rate limit Gemini",
    )

    st.markdown("---")

    btn_proses = st.button(
        "▶️ Mulai Proses",
        type="primary",
        use_container_width=True,
        disabled=not (api_key and uploaded),
    )

    btn_reset = st.button("🗑️ Reset Semua", use_container_width=True)
    if btn_reset:
        st.session_state.entries  = []
        st.session_state.feed     = []
        st.session_state.done     = False
        st.rerun()

    st.markdown("---")
    st.markdown(
        "<small style='color:#888'>Font aksara Arab di Excel:<br>"
        "<b>Noto Sans Arabic</b><br>"
        "<a href='https://fonts.google.com/noto/specimen/Noto+Sans+Arabic' target='_blank'>Download di sini</a> "
        "jika belum terinstall</small>",
        unsafe_allow_html=True,
    )

# ═══════════════════════════════════════════════════════════════════════════════
# MAIN AREA
# ═══════════════════════════════════════════════════════════════════════════════

# ── Stats strip ──────────────────────────────────────────────────────────────
entries = st.session_state.entries
n_lema  = sum(1 for e in entries if e.get('LEMA'))
n_sub   = sum(1 for e in entries if e.get('SUBLEMA'))
n_hal   = st.session_state.n_pages

stats_html = f"""<div class="stat-strip">
  <div class="stat-box"><div class="stat-num">{len(entries)}</div><div class="stat-label">Total Entri</div></div>
  <div class="stat-box"><div class="stat-num">{n_lema}</div><div class="stat-label">LEMA</div></div>
  <div class="stat-box"><div class="stat-num">{n_sub}</div><div class="stat-label">SUBLEMA</div></div>
  <div class="stat-box"><div class="stat-num">{n_hal}</div><div class="stat-label">Hal. PDF</div></div>
</div>"""

stats_placeholder  = st.empty()
stats_placeholder.markdown(stats_html, unsafe_allow_html=True)

# ── Progress ─────────────────────────────────────────────────────────────────
progress_ph = st.empty()
status_ph   = st.empty()

if not st.session_state.done:
    status_ph.markdown(
        '<span class="status-pill status-idle">⬜ Belum diproses</span>',
        unsafe_allow_html=True,
    )
else:
    status_ph.markdown(
        f'<span class="status-pill status-done">✅ Selesai — {len(entries)} entri diekstrak</span>',
        unsafe_allow_html=True,
    )

st.markdown("---")

# ── Dua kolom: live feed | tabel ─────────────────────────────────────────────
lcol, rcol = st.columns([1, 2], gap="large")

with lcol:
    st.markdown('<div class="section-title">📡 Live Feed</div>', unsafe_allow_html=True)
    st.caption("Entri baru muncul satu per satu di sini")
    feed_ph = st.empty()

with rcol:
    st.markdown('<div class="section-title">📋 Tabel Semua Entri</div>', unsafe_allow_html=True)
    table_ph = st.empty()

# Render data yang sudah ada (setelah rerun)
if st.session_state.feed:
    feed_html = '<div class="feed-wrap">' + \
        ''.join(entry_card_html(e) for e in st.session_state.feed[-20:]) + \
        '</div>'
    feed_ph.markdown(feed_html, unsafe_allow_html=True)

if st.session_state.entries:
    df_cur = pd.DataFrame(st.session_state.entries)
    table_ph.dataframe(df_cur, use_container_width=True, height=420, hide_index=True)

# ── Download (muncul kalau sudah selesai) ────────────────────────────────────
dl_ph = st.empty()
if st.session_state.done and st.session_state.entries:
    with dl_ph.container():
        st.markdown("---")
        c1, c2, c3 = st.columns([1, 2, 1])
        with c2:
            excel_bytes = make_excel(st.session_state.entries)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            st.download_button(
                label="⬇️ Download Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"kamus_arab_indonesia_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        st.success(
            f"✅ **{len(st.session_state.entries)} entri** berhasil diekstrak "
            f"({n_lema} LEMA + {n_sub} SUBLEMA)"
        )

# ═══════════════════════════════════════════════════════════════════════════════
# PROCESSING LOGIC
# ═══════════════════════════════════════════════════════════════════════════════
if btn_proses and api_key and st.session_state.pdf_bytes:

    # Konfigurasi Gemini
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)

    # Tentukan range halaman (0-indexed)
    if 'use_range' in dir() and use_range:
        idx_start = p_start - 1
        idx_end   = p_end
    else:
        idx_start = 0
        idx_end   = st.session_state.n_pages

    batch_starts = list(range(idx_start, idx_end, batch_size))
    total_batches = len(batch_starts)

    # Reset jika proses ulang dari awal
    # (komentari baris ini jika ingin lanjut/append)
    st.session_state.entries = []
    st.session_state.feed    = []
    st.session_state.done    = False

    status_ph.markdown(
        '<span class="status-pill status-processing">🔄 Sedang diproses...</span>',
        unsafe_allow_html=True,
    )

    for b_idx, b_start in enumerate(batch_starts):
        b_end = min(b_start + batch_size, idx_end)

        # Update progress
        pct = b_idx / total_batches
        progress_ph.progress(
            pct,
            text=f"⏳ Halaman {b_start+1}–{b_end} / {idx_end}  "
                 f"(batch {b_idx+1}/{total_batches})"
        )

        # Konversi halaman ke gambar
        try:
            imgs = pdf_pages_to_images(
                st.session_state.pdf_bytes, b_start, b_end, dpi
            )
        except Exception as ex:
            st.warning(f"⚠️ Gagal baca hal. {b_start+1}–{b_end}: {ex}")
            continue

        if not imgs:
            continue

        # Panggil Gemini
        new_entries = call_gemini(model, imgs)

        # Tampilkan satu per satu
        for entry in new_entries:
            st.session_state.entries.append(entry)
            st.session_state.feed.append(entry)

            # ── Live feed (20 entri terakhir) ──
            feed_html = '<div class="feed-wrap">' + \
                ''.join(entry_card_html(e) for e in st.session_state.feed[-20:]) + \
                '</div>'
            feed_ph.markdown(feed_html, unsafe_allow_html=True)

            # ── Tabel ──
            df_now = pd.DataFrame(st.session_state.entries)
            table_ph.dataframe(df_now, use_container_width=True, height=420, hide_index=True)

            # ── Stats ──
            n_l = sum(1 for e in st.session_state.entries if e.get('LEMA'))
            n_s = sum(1 for e in st.session_state.entries if e.get('SUBLEMA'))
            stats_placeholder.markdown(f"""<div class="stat-strip">
  <div class="stat-box"><div class="stat-num">{len(st.session_state.entries)}</div><div class="stat-label">Total Entri</div></div>
  <div class="stat-box"><div class="stat-num">{n_l}</div><div class="stat-label">LEMA</div></div>
  <div class="stat-box"><div class="stat-num">{n_s}</div><div class="stat-label">SUBLEMA</div></div>
  <div class="stat-box"><div class="stat-num">{b_end}/{idx_end}</div><div class="stat-label">Halaman</div></div>
</div>""", unsafe_allow_html=True)

            time.sleep(0.06)   # animasi satu-per-satu

        # Jeda antar batch (hindari rate limit)
        if b_idx < total_batches - 1:
            time.sleep(delay_s)

    # ── Selesai ──
    progress_ph.progress(1.0, text="✅ Semua batch selesai diproses!")
    st.session_state.done = True
    st.rerun()
