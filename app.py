import streamlit as st
import zipfile
import io
import re
from PIL import Image
from psd_tools import PSDImage
import fitz  # PyMuPDF

# ── ページ設定 ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="File Layer Splitter",
    page_icon="🎨",
    layout="centered",
)

# ── カスタム CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Space Mono', monospace;
}

.stApp {
    background: #0a0a0f;
    background-image:
        radial-gradient(ellipse 80% 50% at 20% 20%, rgba(100, 60, 255, 0.12) 0%, transparent 60%),
        radial-gradient(ellipse 60% 40% at 80% 80%, rgba(255, 60, 120, 0.08) 0%, transparent 60%);
}

.hero {
    text-align: center;
    padding: 3rem 0 2rem;
    border-bottom: 1px solid rgba(255,255,255,0.06);
    margin-bottom: 2.5rem;
}
.hero h1 {
    font-family: Helvetica, Arial, sans-serif;
    font-weight: 800;
    font-size: 3rem;
    letter-spacing: -0.02em;
    background: linear-gradient(135deg, #ffb347 0%, #ff8c00 50%, #ff6a00 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin: 0 0 0.4rem;
    line-height: 1.1;
}
.hero p {
    color: rgba(255,255,255,0.38);
    font-size: 0.78rem;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    margin: 0;
}

/* バッジ：対応形式 */
.badge-row {
    display: flex;
    justify-content: center;
    gap: 8px;
    flex-wrap: wrap;
    margin-top: 1rem;
}
.badge {
    font-family: 'Space Mono', monospace;
    font-size: 0.65rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    padding: 3px 10px;
    border-radius: 20px;
    border: 1px solid;
}
.badge-psd  { color: #b388ff; border-color: rgba(179,136,255,0.4); background: rgba(179,136,255,0.07); }
.badge-ai   { color: #ff9e40; border-color: rgba(255,158,64,0.4);  background: rgba(255,158,64,0.07);  }
.badge-eps  { color: #40c8ff; border-color: rgba(64,200,255,0.4);  background: rgba(64,200,255,0.07);  }
.badge-pdf  { color: #ff6060; border-color: rgba(255,96,96,0.4);   background: rgba(255,96,96,0.07);   }
.badge-jpg  { color: #60ff9e; border-color: rgba(96,255,158,0.4);  background: rgba(96,255,158,0.07);  }

[data-testid="stFileUploader"] {
    background: rgba(255,255,255,0.03) !important;
    border: 1px dashed rgba(179, 136, 255, 0.35) !important;
    border-radius: 12px !important;
    padding: 1.5rem !important;
    transition: border-color 0.2s ease;
}
[data-testid="stFileUploader"]:hover {
    border-color: rgba(179, 136, 255, 0.65) !important;
}

.stButton > button,
.stDownloadButton > button {
    font-family: 'Space Mono', monospace !important;
    font-weight: 700 !important;
    font-size: 0.78rem !important;
    letter-spacing: 0.12em !important;
    text-transform: uppercase !important;
    border-radius: 6px !important;
    padding: 0.65rem 1.8rem !important;
    transition: all 0.2s ease !important;
}
.stDownloadButton > button {
    background: linear-gradient(135deg, #7c4dff, #ff4081) !important;
    border: none !important;
    color: #fff !important;
    width: 100% !important;
}
.stDownloadButton > button:hover {
    opacity: 0.88 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 8px 24px rgba(124, 77, 255, 0.4) !important;
}

[data-testid="stMetric"] {
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 10px;
    padding: 1rem 1.2rem !important;
}
[data-testid="stMetricLabel"] { color: rgba(255,255,255,0.35) !important; font-size: 0.68rem !important; }
[data-testid="stMetricValue"] { color: #b388ff !important; font-family: 'Syne', sans-serif !important; }

hr { border-color: rgba(255,255,255,0.06) !important; }
p, li { color: rgba(255,255,255,0.6) !important; }
h2, h3 { color: rgba(255,255,255,0.85) !important; font-family: 'Syne', sans-serif !important; }

[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #7c4dff, #ff4081) !important;
    border-radius: 4px !important;
}
[data-testid="stAlert"] {
    border-radius: 10px !important;
    font-size: 0.82rem !important;
}

::-webkit-scrollbar { width: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: rgba(179,136,255,0.3); border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# ── ヘッダー ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>File Layer Splitter</h1>
    <p>Upload · Extract · Download</p>
    <div class="badge-row">
        <span class="badge badge-psd">PSD</span>
        <span class="badge badge-ai">AI</span>
        <span class="badge badge-eps">EPS</span>
        <span class="badge badge-pdf">PDF</span>
        <span class="badge badge-jpg">JPG</span>
    </div>
</div>
""", unsafe_allow_html=True)


# ── ヘルパー関数 ──────────────────────────────────────────────────────────────

def sanitize_name(name: str) -> str:
    """ファイル名として安全な文字列に変換"""
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "layer"


# ── PSD 処理 ──────────────────────────────────────────────────────────────────

def collect_layers(layer, depth=0):
    """再帰的にリーフレイヤーを収集する"""
    results = []
    if isinstance(layer, PSDImage):
        for child in layer:
            results.extend(collect_layers(child, depth))
    elif layer.is_group():
        for child in layer:
            results.extend(collect_layers(child, depth + 1))
    else:
        results.append(layer)
    return results


def process_psd(file_bytes: bytes):
    """
    PSD を解析してレイヤーごとに PNG 化する。
    Returns: list of (file_name, label, png_bytes, PIL.Image), meta dict
    """
    psd = PSDImage.open(io.BytesIO(file_bytes))
    layers = collect_layers(psd)
    results = []
    name_count: dict[str, int] = {}

    for layer in layers:
        try:
            img = layer.composite()
        except Exception:
            continue
        if img is None:
            continue

        base = sanitize_name(layer.name)
        if base in name_count:
            name_count[base] += 1
            fname = f"{base}_{name_count[base]:02d}.png"
        else:
            name_count[base] = 0
            fname = f"{base}.png"

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        results.append((fname, layer.name, buf.getvalue(), img))

    meta = {"width": psd.width, "height": psd.height}
    return results, meta


# ── AI / EPS / PDF 処理（PyMuPDF） ────────────────────────────────────────────

def _fitz_filetype(name: str) -> str:
    """拡張子から PyMuPDF の filetype 文字列を返す"""
    ext = name.rsplit(".", 1)[-1].lower()
    return {"ai": "pdf", "eps": "eps", "pdf": "pdf"}.get(ext, "pdf")


def process_via_fitz(file_bytes: bytes, original_name: str):
    """
    PyMuPDF を使って各ページを高解像度 PNG に変換する。
    Returns: list of (file_name, label, png_bytes, PIL.Image), meta dict
    """
    doc = fitz.open(stream=file_bytes, filetype=_fitz_filetype(original_name))
    results = []
    matrix = fitz.Matrix(2, 2)  # 2× スケール（≈144 dpi）

    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=matrix, alpha=True)
        png_bytes = pix.tobytes("png")
        img = Image.open(io.BytesIO(png_bytes))
        fname = f"page_{i + 1:03d}.png"
        label = f"Page {i + 1}"
        results.append((fname, label, png_bytes, img))

    meta = {
        "pages":  doc.page_count,
        "width":  int(doc[0].rect.width * 2)  if doc.page_count else 0,
        "height": int(doc[0].rect.height * 2) if doc.page_count else 0,
    }
    doc.close()
    return results, meta


# ── JPG / JPEG 処理 ───────────────────────────────────────────────────────────

def process_jpg(file_bytes: bytes, original_name: str):
    """
    JPG を PNG に変換して返す。
    Returns: list of (file_name, label, png_bytes, PIL.Image), meta dict
    """
    img = Image.open(io.BytesIO(file_bytes)).convert("RGBA")
    base = sanitize_name(original_name.rsplit(".", 1)[0])
    fname = f"{base}.png"

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    png_bytes = buf.getvalue()

    meta = {"width": img.width, "height": img.height}
    return [(fname, original_name, png_bytes, img)], meta


# ── ZIP 生成 ──────────────────────────────────────────────────────────────────

def build_zip(items: list) -> bytes:
    """(file_name, label, png_bytes, img) のリストを ZIP にまとめる"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, _, png_bytes, _ in items:
            zf.writestr(fname, png_bytes)
    return buf.getvalue()


# ── メイン UI ─────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "ファイルをドロップ、またはクリックして選択",
    type=["psd", "ai", "eps", "pdf", "jpg", "jpeg"],
    help="対応形式: PSD / AI / EPS / PDF / JPG",
)

if uploaded is not None:
    st.divider()
    ext = uploaded.name.rsplit(".", 1)[-1].lower()
    file_bytes = uploaded.read()

    # ── 処理の振り分け ────────────────────────────────────────────────────────
    spinner_msg = {
        "psd":  "PSD レイヤーを解析中...",
        "ai":   "AI ファイルをレンダリング中...",
        "eps":  "EPS ファイルをレンダリング中...",
        "pdf":  "PDF ページを変換中...",
        "jpg":  "JPG を変換中...",
        "jpeg": "JPG を変換中...",
    }.get(ext, "処理中...")

    with st.spinner(spinner_msg):
        try:
            if ext == "psd":
                items, meta = process_psd(file_bytes)
            elif ext in ("ai", "eps", "pdf"):
                items, meta = process_via_fitz(file_bytes, uploaded.name)
            elif ext in ("jpg", "jpeg"):
                items, meta = process_jpg(file_bytes, uploaded.name)
            else:
                st.error(f"未対応の形式です: .{ext}")
                st.stop()
        except Exception as e:
            st.error(f"ファイルの処理に失敗しました: {e}")
            st.stop()

    if not items:
        st.warning("変換可能なコンテンツが見つかりませんでした。")
        st.stop()

    # ── メトリクス ────────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    col1.metric("形式", f".{ext.upper()}")

    if ext == "psd":
        col2.metric("レイヤー数", len(items))
        col3.metric("キャンバス", f"{meta['width']} × {meta['height']}")
    elif ext in ("ai", "eps", "pdf"):
        col2.metric("ページ数", meta.get("pages", len(items)))
        col3.metric("解像度倍率", "2×（≈144dpi）")
    else:
        col2.metric("画像数", len(items))
        col3.metric("サイズ", f"{meta['width']} × {meta['height']}")

    st.divider()

    # ── プレビュー ────────────────────────────────────────────────────────────
    st.markdown("### プレビュー")
    progress = st.progress(0, text="サムネイルを生成中...")

    cols = st.columns(min(4, len(items)))
    for i, (fname, label, png_bytes, img) in enumerate(items):
        progress.progress((i + 1) / len(items), text=f"{fname} を処理中...")
        thumb_buf = io.BytesIO()
        thumb = img.copy()
        thumb.thumbnail((120, 120))
        # RGBA → RGB（背景を暗色で合成して st.image に渡す）
        if thumb.mode == "RGBA":
            bg = Image.new("RGB", thumb.size, (30, 30, 30))
            bg.paste(thumb, mask=thumb.split()[3])
            thumb = bg
        thumb.save(thumb_buf, format="PNG")
        with cols[i % len(cols)]:
            st.image(thumb_buf.getvalue(), caption=label, use_container_width=True)

    progress.empty()

    # ── ZIP ダウンロード ───────────────────────────────────────────────────────
    st.divider()
    st.markdown("### ダウンロード")

    zip_bytes = build_zip(items)
    zip_name  = f"{uploaded.name.rsplit('.', 1)[0]}_export.zip"
    total_kb  = sum(len(it[2]) for it in items) / 1024

    st.success(f"✓ {len(items)} ファイルを ZIP にパッケージ化しました（合計 {total_kb:.1f} KB）")

    st.download_button(
        label="⬇  ZIP をダウンロード",
        data=zip_bytes,
        file_name=zip_name,
        mime="application/zip",
    )

    with st.expander("含まれるファイル一覧"):
        for fname, label, png_bytes, _ in items:
            st.markdown(
                f"- `{fname}` &nbsp;&nbsp;·&nbsp;&nbsp; "
                f"<span style='color:rgba(255,255,255,0.35)'>{label}</span> "
                f"<span style='color:rgba(255,255,255,0.2);float:right'>{len(png_bytes)//1024} KB</span>",
                unsafe_allow_html=True,
            )

else:
    # 初期状態のヒント
    st.markdown("""
    <div style="
        background: rgba(255,255,255,0.025);
        border: 1px solid rgba(255,255,255,0.06);
        border-radius: 12px;
        padding: 1.6rem 2rem;
        margin-top: 1.5rem;
        font-size: 0.78rem;
        color: rgba(255,255,255,0.35);
        line-height: 2.2;
    ">
    <strong style="color:rgba(255,255,255,0.55)">使い方</strong><br>
    1. ファイルをアップロードエリアにドロップ<br>
    2. 形式に応じて自動処理（レイヤー分割 / ページ変換 / PNG化）<br>
    3. ZIP ダウンロードボタンで一括取得<br><br>
    <strong style="color:rgba(255,255,255,0.55)">形式別の処理</strong><br>
    <code>.psd</code> &nbsp;— レイヤーごとに PNG 分割（psd-tools）<br>
    <code>.ai / .eps / .pdf</code> &nbsp;— 各ページを高解像度 PNG 化（PyMuPDF）<br>
    <code>.jpg / .jpeg</code> &nbsp;— PNG に変換して ZIP 化
    </div>
    """, unsafe_allow_html=True)
