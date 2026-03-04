import io
import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

# =========================
# Config
# =========================
st.set_page_config(page_title="Universal Excel Analyzer (V2)", layout="wide")
st.title("📊 Universal Excel Analyzer (V2)")
st.caption("Upload file Excel → pilih sheet → filter data → eksplorasi distribusi, numerik, crosstab, Likert, dan export.")

# =========================
# Helpers
# =========================
@st.cache_data(show_spinner=False)
def load_excel(file) -> dict:
    xls = pd.ExcelFile(file)
    sheets = {}
    for name in xls.sheet_names:
        try:
            df = xls.parse(name)
            # drop fully empty columns
            df = df.dropna(axis=1, how="all")
            sheets[name] = df
        except Exception:
            pass
    return sheets

def normalize_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_datetime64_any_dtype(s):
        return s
    if pd.api.types.is_numeric_dtype(s):
        return s
    # object-like
    s = s.astype(str).str.strip()
    s = s.replace({"nan": None, "None": None, "": None, "NaN": None})
    return s

def infer_column_types(df: pd.DataFrame):
    cat_cols, num_cols, dt_cols, other_cols = [], [], [], []
    for c in df.columns:
        s = df[c]
        if pd.api.types.is_datetime64_any_dtype(s):
            dt_cols.append(c)
        elif pd.api.types.is_numeric_dtype(s):
            num_cols.append(c)
        else:
            # try parse datetime if looks like date-ish
            s2 = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
            if s2.notna().mean() >= 0.8 and s2.notna().sum() >= 5:
                dt_cols.append(c)
                df[c] = s2
                continue

            nun = s.nunique(dropna=True)
            if nun <= 80:
                cat_cols.append(c)
            else:
                other_cols.append(c)
    return cat_cols, num_cols, dt_cols, other_cols, df

LIKERT_MAP = {
    # Indonesian common
    "sangat setuju": 5,
    "setuju": 4,
    "netral": 3,
    "ragu": 3,
    "ragu-ragu": 3,
    "tidak setuju": 2,
    "sangat tidak setuju": 1,
    # English common
    "strongly agree": 5,
    "agree": 4,
    "neutral": 3,
    "disagree": 2,
    "strongly disagree": 1,
}

def looks_like_likert(s: pd.Series) -> bool:
    s = normalize_series(s).dropna()
    if s.empty:
        return False
    # take unique values (lower)
    uniq = pd.Series(s.unique()).astype(str).str.lower().str.strip()
    hit = uniq.map(lambda x: x in LIKERT_MAP).mean()
    # if most unique values are likert-ish
    return hit >= 0.6 and len(uniq) <= 10

def apply_filters(df: pd.DataFrame, filter_cols: list[str]) -> pd.DataFrame:
    out = df.copy()
    for c in filter_cols:
        s = out[c]
        if pd.api.types.is_numeric_dtype(s):
            mn, mx = float(np.nanmin(s)), float(np.nanmax(s))
            if np.isfinite(mn) and np.isfinite(mx):
                rng = st.sidebar.slider(f"Filter {c}", mn, mx, (mn, mx))
                out = out[out[c].between(rng[0], rng[1], inclusive="both")]
        elif pd.api.types.is_datetime64_any_dtype(s):
            dmin, dmax = out[c].min(), out[c].max()
            if pd.notna(dmin) and pd.notna(dmax):
                dr = st.sidebar.date_input(
                    f"Filter tanggal {c}", value=(dmin.date(), dmax.date())
                )
                if isinstance(dr, (list, tuple)) and len(dr) == 2:
                    start, end = pd.to_datetime(dr[0]), pd.to_datetime(dr[1]) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1)
                    out = out[(out[c] >= start) & (out[c] <= end)]
        else:
            vals = normalize_series(s).dropna()
            uniq = sorted(vals.unique().tolist())
            if len(uniq) > 0:
                chosen = st.sidebar.multiselect(f"Filter {c}", options=uniq, default=[])
                if chosen:
                    out = out[normalize_series(out[c]).isin(chosen)]
    return out

def to_excel_bytes(dfs: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, data in dfs.items():
            safe = re.sub(r"[\[\]\:\*\?\/\\]", "_", name)[:31]  # excel sheet name limit
            data.to_excel(writer, sheet_name=safe, index=True)
    return buf.getvalue()

# =========================
# Upload
# =========================
uploaded = st.file_uploader("📎 Upload file Excel (.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("Silakan upload file Excel untuk mulai.")
    st.stop()

sheets = load_excel(uploaded)
if not sheets:
    st.error("Tidak bisa membaca file Excel. Pastikan format .xlsx dan tidak terenkripsi.")
    st.stop()

# =========================
# Sidebar: Select sheet
# =========================
st.sidebar.header("⚙️ Kontrol")
sheet_name = st.sidebar.selectbox("Pilih sheet", list(sheets.keys()))
df_raw = sheets[sheet_name].copy()

cat_cols, num_cols, dt_cols, other_cols, df_raw = infer_column_types(df_raw)

st.sidebar.write(f"Rows: **{len(df_raw):,}** | Columns: **{len(df_raw.columns)}**")
st.sidebar.caption(f"Kategori: {len(cat_cols)} | Numerik: {len(num_cols)} | Tanggal: {len(dt_cols)}")

# =========================
# Filters (universal)
# =========================
st.sidebar.subheader("🧰 Filter data (opsional)")
default_filter_cols = cat_cols[:2] + dt_cols[:1]
filter_cols = st.sidebar.multiselect(
    "Pilih kolom untuk filter",
    options=list(df_raw.columns),
    default=[c for c in default_filter_cols if c in df_raw.columns],
)
df = apply_filters(df_raw, filter_cols)

# =========================
# Preview
# =========================
with st.expander("🔎 Preview data (setelah filter)"):
    st.dataframe(df.head(50), use_container_width=True)

# =========================
# Summary top metrics
# =========================
m1, m2, m3, m4 = st.columns(4)
with m1:
    st.metric("Total row (raw)", f"{len(df_raw):,}")
with m2:
    st.metric("Total row (filtered)", f"{len(df):,}")
with m3:
    st.metric("Jumlah kolom", f"{df.shape[1]:,}")
with m4:
    st.metric("Missing cells", f"{int(df.isna().sum().sum()):,}")

st.divider()

# =========================
# Tabs
# =========================
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📌 Kategori", "📈 Numerik", "🧮 Crosstab", "✅ Likert", "📤 Export"])

# ---------- TAB 1: Categorical ----------
with tab1:
    st.subheader("📌 Distribusi kategori")
    if len(df.columns) == 0:
        st.warning("Sheet kosong setelah filter.")
        st.stop()

    default_col = cat_cols[0] if cat_cols else df.columns[0]
    col_choice = st.selectbox("Pilih kolom kategori", options=list(df.columns), index=list(df.columns).index(default_col))
    top_n = st.slider("Top N kategori", min_value=5, max_value=60, value=20, step=1)
    chart_type = st.radio("Tipe chart", ["Bar", "Pie"], horizontal=True, index=0)
    show_missing = st.checkbox("Tampilkan Missing/Blank", value=False)

    s = normalize_series(df[col_choice])
    if not show_missing:
        s = s.dropna()
    vc = s.value_counts(dropna=show_missing).head(top_n).reset_index()
    vc.columns = [col_choice, "count"]

    left, right = st.columns([1.2, 1])
    with left:
        if chart_type == "Pie":
            fig = px.pie(vc, names=col_choice, values="count")
        else:
            fig = px.bar(vc, x=col_choice, y="count")
            fig.update_layout(xaxis_title=col_choice, yaxis_title="Count")
        st.plotly_chart(fig, use_container_width=True)

    with right:
        total = len(df)
        non_missing = int(df[col_choice].notna().sum())
        missing = total - non_missing
        st.subheader("Ringkasan")
        st.metric("Total respon (filtered)", f"{total:,}")
        st.metric("Non-missing", f"{non_missing:,}")
        st.metric("Missing", f"{missing:,}")
        st.write("Top kategori (tabel):")
        st.dataframe(vc, use_container_width=True, height=360)

# ---------- TAB 2: Numeric ----------
with tab2:
    st.subheader("📈 Analisa numerik")
    if not num_cols:
        st.info("Tidak ada kolom numerik terdeteksi.")
    else:
        num_choice = st.selectbox("Pilih kolom numerik", options=num_cols, index=0)
        s = pd.to_numeric(df[num_choice], errors="coerce")
        st.write("Statistik ringkas:")
        stats = pd.DataFrame({
            "metric": ["count", "missing", "mean", "median", "min", "max", "std"],
            "value": [
                int(s.notna().sum()),
                int(s.isna().sum()),
                float(s.mean()) if s.notna().any() else np.nan,
                float(s.median()) if s.notna().any() else np.nan,
                float(s.min()) if s.notna().any() else np.nan,
                float(s.max()) if s.notna().any() else np.nan,
                float(s.std()) if s.notna().any() else np.nan,
            ]
        })
        st.dataframe(stats, use_container_width=True)

        bins = st.slider("Jumlah bins histogram", 5, 80, 30)
        fig = px.histogram(df, x=num_choice, nbins=bins)
        st.plotly_chart(fig, use_container_width=True)

# ---------- TAB 3: Crosstab ----------
with tab3:
    st.subheader("🧮 Crosstab (2 kolom)")
    c1, c2 = st.columns(2)
    with c1:
        col_a = st.selectbox("Kolom A", options=list(df.columns), index=0)
    with c2:
        col_b = st.selectbox("Kolom B", options=list(df.columns), index=min(1, len(df.columns)-1))

    if col_a == col_b:
        st.info("Pilih dua kolom yang berbeda.")
    else:
        a = normalize_series(df[col_a])
        b = normalize_series(df[col_b])
        dropna = st.checkbox("Drop NA (lebih bersih)", value=False)
        ct = pd.crosstab(a, b, dropna=dropna)

        st.write("Crosstab (count):")
        st.dataframe(ct, use_container_width=True)

        if ct.shape[0] <= 60 and ct.shape[1] <= 60:
            fig2 = px.imshow(ct, aspect="auto")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.caption("Crosstab terlalu besar untuk heatmap. Pilih kolom dengan kategori lebih sedikit / gunakan filter.")

# ---------- TAB 4: Likert ----------
with tab4:
    st.subheader("✅ Deteksi & ringkasan Likert")
    likert_candidates = [c for c in df.columns if looks_like_likert(df[c])]
    if not likert_candidates:
        st.info("Tidak menemukan kolom Likert secara otomatis. Anda bisa pilih kolom manual juga.")
        manual = st.selectbox("Pilih kolom (manual)", options=list(df.columns))
        likert_candidates = [manual] if manual else []

    chosen = st.multiselect("Kolom Likert untuk diringkas", options=list(df.columns), default=likert_candidates)

    summaries = []
    for c in chosen:
        s = normalize_series(df[c]).dropna().astype(str).str.lower().str.strip()
        mapped = s.map(LIKERT_MAP).dropna()
        if mapped.empty:
            st.warning(f"Kolom '{c}' tidak bisa dipetakan ke Likert map (contoh nilai: {df[c].dropna().astype(str).head(5).tolist()}).")
            continue

        avg = float(mapped.mean())
        n = int(mapped.shape[0])
        dist = mapped.value_counts().sort_index()
        # back label for display
        inv = {v: k for k, v in LIKERT_MAP.items()}
        dist_df = pd.DataFrame({
            "score": dist.index,
            "label_example": [inv.get(int(i), str(i)) for i in dist.index],
            "count": dist.values
        })

        summaries.append({"kolom": c, "n": n, "avg_score": avg})

        st.markdown(f"**{c}** — n={n:,}, avg score={avg:.2f}")
        st.dataframe(dist_df, use_container_width=True)

        fig = px.bar(dist_df, x="score", y="count", hover_data=["label_example"])
        st.plotly_chart(fig, use_container_width=True)

    if summaries:
        st.write("Ringkasan semua kolom Likert:")
        st.dataframe(pd.DataFrame(summaries).sort_values("avg_score", ascending=False), use_container_width=True)

# ---------- TAB 5: Export ----------
with tab5:
    st.subheader("📤 Export hasil ke Excel")
    st.caption("Export akan memasukkan: data filtered, ringkasan missing per kolom, dan crosstab (jika ada pilihan).")

    # Missing summary
    miss = pd.DataFrame({
        "column": df.columns,
        "missing_count": df.isna().sum().values,
        "missing_pct": (df.isna().mean().values * 100).round(2),
        "unique_non_na": [df[c].nunique(dropna=True) for c in df.columns],
        "dtype": [str(df[c].dtype) for c in df.columns],
    }).sort_values("missing_pct", ascending=False)

    st.write("Preview ringkasan missing:")
    st.dataframe(miss.head(30), use_container_width=True)

    # Optional crosstab export config (reuse from tab3-like)
    st.markdown("**(Opsional) Crosstab untuk export**")
    ea, eb = st.columns(2)
    with ea:
        exp_a = st.selectbox("Export Crosstab A", options=list(df.columns), index=0, key="exp_a")
    with eb:
        exp_b = st.selectbox("Export Crosstab B", options=list(df.columns), index=min(1, len(df.columns)-1), key="exp_b")

    ct_export = None
    if exp_a != exp_b:
        ct_export = pd.crosstab(normalize_series(df[exp_a]), normalize_series(df[exp_b]), dropna=False)

    export_name = st.text_input("Nama file export", value="excel_analysis_v2.xlsx")
    if st.button("Generate & Download Excel", type="primary"):
        payload = {
            "filtered_data": df,
            "missing_summary": miss,
        }
        if ct_export is not None:
            payload["crosstab"] = ct_export

        xbytes = to_excel_bytes(payload)
        st.download_button(
            label="⬇️ Download Excel hasil analisa",
            data=xbytes,
            file_name=export_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success("Export siap diunduh.")