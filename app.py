import streamlit as st
import zipfile
import io
import re
import json
from psd_tools import PSDImage

# ── ページ設定 ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="PSD Layer Splitter",
    page_icon="🎨",
    layout="centered",
)

# ── カスタム CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@700;800&display=swap');

/* ベース */
html, body, [class*="css"] {
    font-family: 'Space Mono', monospace;
}

/* 背景 */
.stApp {
    background: #0a0a0f;
    background-image:
        radial-gradient(ellipse 80% 50% at 20% 20%, rgba(100, 60, 255, 0.12) 0%, transparent 60%),
        radial-gradient(ellipse 60% 40% at 80% 80%, rgba(255, 60, 120, 0.08) 0%, transparent 60%);
}

/* ヘッダー */
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

/* アップロードゾーン */
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

/* ボタン */
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

/* レイヤーカード グリッド */
.layer-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
    gap: 12px;
    margin: 1.5rem 0;
}
.layer-card {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 10px;
    padding: 10px;
    text-align: center;
    font-size: 0.68rem;
    color: rgba(255,255,255,0.5);
    letter-spacing: 0.04em;
    word-break: break-all;
}
.layer-card img {
    max-width: 100%;
    border-radius: 6px;
    margin-bottom: 6px;
    background: repeating-conic-gradient(#444 0% 25%, #333 0% 50%) 0 0 / 12px 12px;
}

/* メトリクス */
[data-testid="stMetric"] {
    background: rgba(255,255,255,0.03);
    border: 1px solid rgba(255,255,255,0.07);
    border-radius: 10px;
    padding: 1rem 1.2rem !important;
}
[data-testid="stMetricLabel"] { color: rgba(255,255,255,0.35) !important; font-size: 0.68rem !important; }
[data-testid="stMetricValue"] { color: #b388ff !important; font-family: 'Syne', sans-serif !important; }

/* 区切り線 */
hr { border-color: rgba(255,255,255,0.06) !important; }

/* テキスト */
p, li { color: rgba(255,255,255,0.6) !important; }
h2, h3 { color: rgba(255,255,255,0.85) !important; font-family: 'Syne', sans-serif !important; }

/* プログレス */
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #7c4dff, #ff4081) !important;
    border-radius: 4px !important;
}

/* サクセスバナー */
[data-testid="stAlert"] {
    border-radius: 10px !important;
    font-size: 0.82rem !important;
}

/* スクロールバー */
::-webkit-scrollbar { width: 4px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: rgba(179,136,255,0.3); border-radius: 4px; }
</style>
""", unsafe_allow_html=True)

# ── ヘッダー ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>PSD Layer Splitter</h1>
    <p>Upload · Extract · Download</p>
</div>
""", unsafe_allow_html=True)

# ── ヘルパー関数 ──────────────────────────────────────────────────────────────

def sanitize_name(name: str) -> str:
    """ファイル名として安全な文字列に変換"""
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "layer"


def collect_layers(layer, depth=0):
    """
    再帰的にレイヤーを収集する。
    グループは子を展開し、リーフレイヤー（画像を持つもの）のみを返す。
    """
    results = []
    if hasattr(layer, "__iter__") and not isinstance(layer, PSDImage):
        for child in layer:
            results.extend(collect_layers(child, depth + 1))
    else:
        # PSDImage 自体（ルート）はスキップ
        if isinstance(layer, PSDImage):
            for child in layer:
                results.extend(collect_layers(child, depth))
        else:
            # is_group() が True のものは子を再帰展開
            if layer.is_group():
                for child in layer:
                    results.extend(collect_layers(child, depth + 1))
            else:
                results.append(layer)
    return results


def process_psd(file_bytes: bytes):
    """
    PSD を解析し、レイヤーごとの PNG bytes と名前のリストを返す。
    Returns: list of (layer_name, png_bytes)
    """
    psd = PSDImage.open(io.BytesIO(file_bytes))
    layers = collect_layers(psd)

    results = []
    name_count: dict[str, int] = {}

    for layer in layers:
        # 非表示レイヤーをスキップ（オプション：スキップしたい場合はコメントを外す）
        # if not layer.is_visible():
        #     continue

        try:
            img = layer.composite()
        except Exception:
            continue

        if img is None:
            continue

        base_name = sanitize_name(layer.name)
        # 重複名に連番を付ける
        if base_name in name_count:
            name_count[base_name] += 1
            file_name = f"{base_name}_{name_count[base_name]:02d}.png"
        else:
            name_count[base_name] = 0
            file_name = f"{base_name}.png"

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        meta = {
            "filename": file_name,
            "name": layer.name,
            "x": layer.left,
            "y": layer.top,
            "width": layer.width,
            "height": layer.height,
            "opacity": round(layer.opacity / 255, 6),
        }
        results.append((file_name, layer.name, buf.getvalue(), img, meta))

    return results, psd


def generate_image_id() -> str:
    """画像アイテム用のユニーク ID を生成する（例: i8f3a2c1d...）"""
    import uuid
    return "i" + uuid.uuid4().hex


def build_bp(layer_data: list, psd) -> str:
    """
    PSD レイヤー情報から Spinno .bp 形式の JSON 文字列（1行）を生成する。
    ルート構造: {"doc_type": "spinno", "design_data": {"doc": {...}}}
    """
    items = []
    for file_name, layer_name, _, _, meta in layer_data:
        w = meta["width"]
        h = meta["height"]
        cx = meta["x"] + w / 2          # left + width/2  → 中心 X
        cy = meta["y"] + h / 2          # top  + height/2 → 中心 Y
        opacity = meta["opacity"]        # 0.0〜1.0

        item = {
            "type": 0,                   # 画像レイヤー固定値
            "image_id": generate_image_id(),
            "name": layer_name,
            "x": cx,
            "y": cy,
            "w": w,
            "h": h,
            "ow": w,
            "oh": h,
            "angle": 0,
            "opacity": opacity,
            "flip_h": False,
            "flip_v": False,
            "gsize": 20,
            "noareaover": 0,
            "locked": False,
            "visible": True,
            "blend_mode": "normal",
        }
        items.append(item)

    doc = {
        "width": psd.width,
        "height": psd.height,
        "items": items,
    }

    bp_obj = {
        "doc_type": "spinno",
        "design_data": {
            "doc": doc,
        },
    }

    # 改行なし・ensure_ascii=False で1行出力
    return json.dumps(bp_obj, ensure_ascii=False, separators=(",", ":"))


def build_zip(layer_data: list, psd, bp_name: str) -> bytes:
    """レイヤー PNG と .bp ファイルを 1 つの ZIP にまとめる"""
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for file_name, _, png_bytes, _, __ in layer_data:
            zf.writestr(file_name, png_bytes)
        bp_str = build_bp(layer_data, psd)
        zf.writestr(bp_name, bp_str.encode("utf-8"))
    return zip_buf.getvalue()


# ── メイン UI ─────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "PSDファイルをドロップ、またはクリックして選択",
    type=["psd"],
    help="Adobe Photoshop .psd 形式のファイルに対応しています",
)

if uploaded is not None:
    st.divider()

    with st.spinner("PSDを解析中..."):
        try:
            file_bytes = uploaded.read()
            layer_data, psd = process_psd(file_bytes)
        except Exception as e:
            st.error(f"PSDの読み込みに失敗しました: {e}")
            st.stop()

    if not layer_data:
        st.warning("変換可能なレイヤーが見つかりませんでした。")
        st.stop()

    # ── メトリクス ───────────────────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    col1.metric("キャンバスサイズ", f"{psd.width} × {psd.height}")
    col2.metric("検出レイヤー数", len(layer_data))
    total_kb = sum(len(d[2]) for d in layer_data) / 1024
    col3.metric("合計サイズ", f"{total_kb:.1f} KB")

    st.divider()

    # ── レイヤープレビュー ───────────────────────────────────────────────────
    st.markdown("### レイヤープレビュー")
    progress = st.progress(0, text="サムネイルを生成中...")

    # カード HTML を構築
    cols = st.columns(min(4, len(layer_data)))
    for i, (file_name, layer_name, png_bytes, img, _) in enumerate(layer_data):
        progress.progress((i + 1) / len(layer_data), text=f"{file_name} を処理中...")
        thumb_buf = io.BytesIO()
        thumb = img.copy()
        thumb.thumbnail((120, 120))
        thumb.save(thumb_buf, format="PNG")

        with cols[i % len(cols)]:
            st.image(thumb_buf.getvalue(), caption=layer_name, use_container_width=True)

    progress.empty()

    # ── ZIP 生成 & ダウンロード ───────────────────────────────────────────────
    st.divider()
    st.markdown("### ダウンロード")

    base_name = uploaded.name.rsplit(".", 1)[0]
    bp_name = f"{base_name}.bp"
    zip_name = f"{base_name}_layers.zip"
    zip_bytes = build_zip(layer_data, psd, bp_name)

    st.success(f"✓ {len(layer_data)} レイヤーを ZIP にパッケージ化しました ({len(zip_bytes)/1024:.1f} KB)")

    st.download_button(
        label="⬇  ZIP をダウンロード",
        data=zip_bytes,
        file_name=zip_name,
        mime="application/zip",
    )

    with st.expander("含まれるファイル一覧"):
        for file_name, layer_name, png_bytes, _, __ in layer_data:
            st.markdown(
                f"- `{file_name}` &nbsp;&nbsp;·&nbsp;&nbsp; "
                f"<span style='color:rgba(255,255,255,0.35)'>{layer_name}</span> "
                f"<span style='color:rgba(255,255,255,0.2);float:right'>{len(png_bytes)//1024} KB</span>",
                unsafe_allow_html=True,
            )
        st.markdown(
            f"- `{bp_name}` &nbsp;&nbsp;·&nbsp;&nbsp; "
            f"<span style='color:rgba(179,136,255,0.7)'>Spinno レイアウトデータ</span>",
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
        line-height: 2;
    ">
    <strong style="color:rgba(255,255,255,0.55)">使い方</strong><br>
    1. 上のアップロードエリアに <code>.psd</code> ファイルをドロップ<br>
    2. レイヤーが自動的に PNG に変換・プレビュー表示<br>
    3. ZIP ダウンロードボタンで一括取得
    </div>
    """, unsafe_allow_html=True)
