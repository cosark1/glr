"""
Sektör Bazlı İşgücü Ekonomik Göstergesi - Streamlit Dashboard
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import os

from data_loader import load_all_data, NACE_MAPPING, NACE_TO_SECTOR
from export_utils import export_excel, export_word
from generate_report import create_academic_report

# ── Sayfa ayarı ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Sektörel İşgücü Analizi",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main .block-container { padding-top: 1.5rem; }
    .metric-card {
        background: linear-gradient(135deg, #1F4E79 0%, #2E86C1 100%);
        border-radius: 10px; padding: 15px 20px; color: white; text-align: center;
    }
    .metric-card h3 { font-size: 14px; margin: 0; opacity: 0.85; }
    .metric-card h1 { font-size: 28px; margin: 5px 0 0 0; }
    .kadran-k1 { background-color: #FCE4EC; border-left: 4px solid #C62828; padding: 10px; border-radius: 5px; }
    .kadran-k2 { background-color: #E8F5E9; border-left: 4px solid #2E7D32; padding: 10px; border-radius: 5px; }
    .kadran-k3 { background-color: #FFF3E0; border-left: 4px solid #E65100; padding: 10px; border-radius: 5px; }
    .kadran-k4 { background-color: #E3F2FD; border-left: 4px solid #1565C0; padding: 10px; border-radius: 5px; }
</style>
""", unsafe_allow_html=True)


# ── Veri yükleme ────────────────────────────────────────────────────────────
@st.cache_data
def load_data(ana_path: str, sgk_path: str = None):
    return load_all_data(ana_path, sgk_path)


def find_default_files():
    """Varsayılan dosya yollarını bul."""
    base = Path(__file__).parent.parent
    ana = base / "SEKTÖR BAZLI İŞGÜCÜ EKONOMİK GÖSTERGESİ.xlsx"
    sgk_dir = base / "sgk_veriler"
    sgk = None
    if sgk_dir.exists():
        for f in sgk_dir.iterdir():
            if "BÖLÜM 1" in f.name and f.suffix == ".xlsx":
                sgk = f
                break
    return str(ana) if ana.exists() else None, str(sgk) if sgk else None


# ── Sidebar ─────────────────────────────────────────────────────────────────
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/b/b4/Flag_of_Turkey.svg/200px-Flag_of_Turkey.svg.png", width=60)
    st.title("Sektörel İşgücü Analizi")
    st.caption("TÜİK + SGK Verileri")

    from datetime import date
    _months = ["Ocak","Şubat","Mart","Nisan","Mayıs","Haziran",
               "Temmuz","Ağustos","Eylül","Ekim","Kasım","Aralık"]
    _today = date.today()
    _date_str = f"{_today.day} {_months[_today.month-1]} {_today.year}"
    st.caption(f"🗓 Son güncelleme: {_date_str}  |  v2.4")

    st.divider()

    default_ana, default_sgk = find_default_files()

    st.subheader("Veri Kaynakları")

    with st.expander("Yüklenmesi Gereken Dosyalar", expanded=False):
        st.markdown("""
**1. Ana Dosya (Zorunlu):**
- **Dosya:** `SEKTÖR BAZLI İŞGÜCÜ EKONOMİK GÖSTERGESİ.xlsx`
- Bu dosya aşağıdaki sayfaları içermelidir:
  - **I.2.14** — TÜİK: Gelir Yöntemiyle GSYH, İktisadi Faaliyet Kollarına Göre (2009-2024)
  - **TABLO-1.12** — SGK: NACE Rev.2 Faaliyet Gruplarına ve İşyeri Büyüklüğüne Göre İşyeri Sayıları
  - **TABLO-1.13** — SGK: NACE Rev.2 Faaliyet Gruplarına ve İşyeri Büyüklüğüne Göre Zorunlu Sigortalı Sayıları

**2. SGK Bölüm 1 (Opsiyonel - Önerilir):**
- **Dosya:** SGK Yıllık İstatistik Yayını — Bölüm 1 (örn. `istatistik_yilligi_2024_bolum1.xlsx`)
- Bu dosya aşağıdaki sayfayı içermelidir:
  - **TABLO-1.16** — SGK: NACE Rev.2 Faaliyet Gruplarına Göre Prime Esas Ortalama Günlük Kazanç (daimi/geçici, kamu/özel, erkek/kadın)
""")

    ana_file = st.file_uploader(
        "Ana Dosya (TÜİK I.2.14 + SGK 1.12/1.13)",
        type=["xlsx"], key="ana",
        help="SEKTÖR BAZLI İŞGÜCÜ EKONOMİK GÖSTERGESİ.xlsx"
    )
    sgk_file = st.file_uploader(
        "SGK Bölüm 1 — TABLO-1.16 (Opsiyonel)",
        type=["xlsx"], key="sgk",
        help="SGK Yıllık İstatistik Bölüm 1: Günlük kazanç verileri"
    )

    use_defaults = False
    if not ana_file and default_ana:
        st.info(f"Varsayılan dosya kullanılıyor")
        use_defaults = True

    st.divider()

    page = st.radio("Sayfa", [
        "📊 Özet Dashboard",
        "📈 Trend Analizi",
        "👥 İstihdam Yapısı",
        "🎯 Kadran Analizi",
        "📋 Teşvik Kılavuzu",
        "📖 Metodoloji",
        "📥 Rapor İndir",
    ])


# ── Veri yükle ──────────────────────────────────────────────────────────────
try:
    if ana_file:
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(ana_file.read())
            ana_path = tmp.name
        sgk_path = None
        if sgk_file:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp2:
                tmp2.write(sgk_file.read())
                sgk_path = tmp2.name
        data = load_data(ana_path, sgk_path)
    elif use_defaults:
        data = load_data(default_ana, default_sgk)
    else:
        st.warning("Lütfen ana veri dosyasını yükleyin veya varsayılan dosyanın mevcut olduğundan emin olun.")
        st.stop()
except Exception as e:
    st.error(f"Veri yüklenirken hata: {e}")
    st.stop()

summary = data["summary"]
trend_df = data["trend"]
insured = data["insured"]
workplace = data["workplace"]
wages = data.get("wages", pd.DataFrame())


# ── Sayfalar ────────────────────────────────────────────────────────────────

if page == "📊 Özet Dashboard":
    st.title("Sektör Bazlı İşgücü Ekonomik Göstergesi")
    st.caption("2024 Yılı - Cari Fiyatlarla")

    # KPI kartları
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Toplam İstihdam", f"{summary['istihdam'].sum():,.0f}", help="SGK 4/1-a zorunlu sigortalı")
    with col2:
        st.metric("Toplam GKD", f"{summary['gkd_2024'].sum():,.0f} Myr TL")
    with col3:
        ort_maliyet = summary["isgucu_2024"].sum() / summary["gkd_2024"].sum() * 100
        st.metric("Ort. İşgücü Maliyet Oranı", f"%{ort_maliyet:.1f}")
    with col4:
        k1_count = len(summary[summary["kadran"].str.startswith("K1")])
        st.metric("K1 Sektör Sayısı", k1_count, help="Yüksek İstihdam + Yüksek Maliyet")
    with col5:
        st.metric("Toplam İşyeri", f"{summary['isyeri'].sum():,.0f}")

    st.divider()

    # Ana tablo
    st.subheader("Sektörel Gösterge Tablosu")

    display_cols = {
        "sektor": "Sektör",
        "gkd_2024": "GKD (Myr TL)",
        "isgucu_2024": "İşgücü (Myr TL)",
        "maliyet_orani_2024": "Maliyet Oranı (%)",
        "maliyet_etkinligi": "Maliyet Etkinliği",
        "istihdam": "İstihdam",
        "istihdam_payi": "İstihdam Payı (%)",
        "kisi_basi_gkd": "Kişi Başı GKD (TL)",
        "ort_gunluk_kazanc": "Ort. Günlük Kazanç (TL)",
        "kobi_orani": "KOBİ Oranı (%)",
        "kadran": "Kadran",
    }

    df_display = summary[list(display_cols.keys())].copy()
    df_display.columns = list(display_cols.values())

    # Renkli kadran gösterimi
    def color_kadran(val):
        if "K1" in str(val):
            return "background-color: #FCE4EC"
        elif "K2" in str(val):
            return "background-color: #E8F5E9"
        elif "K3" in str(val):
            return "background-color: #FFF3E0"
        elif "K4" in str(val):
            return "background-color: #E3F2FD"
        return ""

    styled = df_display.style.map(color_kadran, subset=["Kadran"]).format({
        "GKD (Myr TL)": "{:,.1f}",
        "İşgücü (Myr TL)": "{:,.1f}",
        "Maliyet Oranı (%)": "{:.1f}",
        "Maliyet Etkinliği": "{:.2f}",
        "İstihdam": "{:,.0f}",
        "İstihdam Payı (%)": "{:.2f}",
        "Kişi Başı GKD (TL)": "{:,.0f}",
        "Ort. Günlük Kazanç (TL)": "{:,.2f}",
        "KOBİ Oranı (%)": "{:.1f}",
    })

    st.dataframe(styled, use_container_width=True, height=600)

    # Scatter plot
    st.subheader("İstihdam Payı vs İşgücü Maliyet Oranı")
    fig = px.scatter(
        summary,
        x="istihdam_payi",
        y="maliyet_orani_2024",
        size="istihdam",
        color="kadran",
        hover_name="sektor",
        hover_data={"maliyet_etkinligi": ":.2f", "ort_gunluk_kazanc": ":,.0f", "kobi_orani": ":.1f"},
        color_discrete_map={
            "K1: Yüksek İstihdam + Yüksek Maliyet": "#E53935",
            "K2: Yüksek İstihdam + Düşük Maliyet": "#43A047",
            "K3: Düşük İstihdam + Yüksek Maliyet": "#FB8C00",
            "K4: Düşük İstihdam + Düşük Maliyet": "#1E88E5",
        },
        labels={
            "istihdam_payi": "İstihdam Payı (%)",
            "maliyet_orani_2024": "İşgücü Maliyet Oranı (%)",
            "istihdam": "İstihdam Sayısı",
        },
    )
    med_emp = summary["istihdam_payi"].median()
    med_cost = summary["maliyet_orani_2024"].median()
    fig.add_hline(y=med_cost, line_dash="dash", line_color="gray", opacity=0.5,
                  annotation_text=f"Medyan Maliyet: %{med_cost:.1f}")
    fig.add_vline(x=med_emp, line_dash="dash", line_color="gray", opacity=0.5,
                  annotation_text=f"Medyan İstihdam: %{med_emp:.1f}")
    fig.update_layout(height=550, template="plotly_white")
    st.plotly_chart(fig, use_container_width=True)

    # Bar chart - maliyet oranı
    st.subheader("Sektörel İşgücü Maliyet Oranı Sıralaması")
    sorted_df = summary.sort_values("maliyet_orani_2024", ascending=True)
    fig2 = px.bar(
        sorted_df, x="maliyet_orani_2024", y="sektor", orientation="h",
        color="kadran",
        color_discrete_map={
            "K1: Yüksek İstihdam + Yüksek Maliyet": "#E53935",
            "K2: Yüksek İstihdam + Düşük Maliyet": "#43A047",
            "K3: Düşük İstihdam + Yüksek Maliyet": "#FB8C00",
            "K4: Düşük İstihdam + Düşük Maliyet": "#1E88E5",
        },
        labels={"maliyet_orani_2024": "İşgücü Maliyet Oranı (%)", "sektor": ""},
    )
    fig2.update_layout(height=600, template="plotly_white", showlegend=False)
    st.plotly_chart(fig2, use_container_width=True)


elif page == "📈 Trend Analizi":
    st.title("Sektörel Trend Analizi (2009-2024)")

    # Sektör seçimi
    sectors = sorted(trend_df["sektor"].unique())
    selected = st.multiselect("Sektör Seçin", sectors, default=sectors[:5])

    if selected:
        filtered = trend_df[trend_df["sektor"].isin(selected)]

        tab1, tab2, tab3 = st.tabs(["İşgücü Maliyet Oranı", "GKD Trendi", "İşgücü Ödemesi"])

        with tab1:
            fig = px.line(filtered, x="yil", y="maliyet_orani", color="sektor",
                          labels={"yil": "Yıl", "maliyet_orani": "İşgücü Maliyet Oranı (%)", "sektor": "Sektör"},
                          markers=True)
            fig.update_layout(height=500, template="plotly_white")
            st.plotly_chart(fig, use_container_width=True)

        with tab2:
            fig2 = px.line(filtered, x="yil", y="gkd", color="sektor",
                           labels={"yil": "Yıl", "gkd": "Gayrisafi Katma Değer (Milyar TL)", "sektor": "Sektör"},
                           markers=True)
            fig2.update_layout(height=500, template="plotly_white")
            st.plotly_chart(fig2, use_container_width=True)

        with tab3:
            fig3 = px.line(filtered, x="yil", y="isgucu", color="sektor",
                           labels={"yil": "Yıl", "isgucu": "İşgücü Ödemesi (Milyar TL)", "sektor": "Sektör"},
                           markers=True)
            fig3.update_layout(height=500, template="plotly_white")
            st.plotly_chart(fig3, use_container_width=True)

        # Karşılaştırma tablosu
        st.subheader("Dönem Karşılaştırması")
        col1, col2 = st.columns(2)
        with col1:
            y1 = st.selectbox("Başlangıç Yılı", list(range(2009, 2024)), index=10)
        with col2:
            y2 = st.selectbox("Bitiş Yılı", list(range(2010, 2025)), index=14)

        if y1 < y2:
            comp_data = []
            for s in selected:
                s_data = filtered[filtered["sektor"] == s]
                d1 = s_data[s_data["yil"] == y1]
                d2 = s_data[s_data["yil"] == y2]
                if not d1.empty and not d2.empty:
                    comp_data.append({
                        "Sektör": s,
                        f"Maliyet Oranı {y1} (%)": d1.iloc[0]["maliyet_orani"],
                        f"Maliyet Oranı {y2} (%)": d2.iloc[0]["maliyet_orani"],
                        "Değişim (puan)": round(d2.iloc[0]["maliyet_orani"] - d1.iloc[0]["maliyet_orani"], 1),
                        f"GKD {y1} (Myr)": round(d1.iloc[0]["gkd"], 1),
                        f"GKD {y2} (Myr)": round(d2.iloc[0]["gkd"], 1),
                        "GKD Büyüme (x)": round(d2.iloc[0]["gkd"] / d1.iloc[0]["gkd"], 1) if d1.iloc[0]["gkd"] > 0 else 0,
                    })
            if comp_data:
                st.dataframe(pd.DataFrame(comp_data), use_container_width=True)


elif page == "👥 İstihdam Yapısı":
    st.title("Sektörel İstihdam Yapısı (2024)")

    tab1, tab2, tab3 = st.tabs(["Alt Sektör Detay", "İşyeri Büyüklüğü", "Günlük Kazanç"])

    with tab1:
        # Ana sektör filtresi
        sectors = sorted(insured["ana_sektor"].unique())
        sel_sector = st.selectbox("Ana Sektör Filtresi", ["Tümü"] + sectors)

        df_show = insured.copy()
        if sel_sector != "Tümü":
            df_show = df_show[df_show["ana_sektor"] == sel_sector]

        # Günlük kazanç merge
        if not wages.empty and "gunluk_kazanc_toplam" in wages.columns:
            wage_cols = wages[["nace_kodu", "gunluk_kazanc_toplam", "gunluk_kazanc_erkek",
                               "gunluk_kazanc_kadin", "sigortali_kadin", "sigortali_erkek"]].copy()
            df_show = df_show.merge(wage_cols, on="nace_kodu", how="left")

        show_cols = ["nace_kodu", "faaliyet", "ana_sektor", "toplam"]
        show_names = {"nace_kodu": "NACE", "faaliyet": "Faaliyet", "ana_sektor": "Ana Sektör", "toplam": "Sigortalı"}
        if "gunluk_kazanc_toplam" in df_show.columns:
            show_cols.append("gunluk_kazanc_toplam")
            show_names["gunluk_kazanc_toplam"] = "Ort. Günlük Kazanç (TL)"

        df_show = df_show.sort_values("toplam", ascending=False)
        st.dataframe(
            df_show[show_cols].rename(columns=show_names),
            use_container_width=True,
            height=500,
        )

        st.metric("Toplam Sigortalı (Filtre)", f"{df_show['toplam'].sum():,.0f}")

    with tab2:
        st.subheader("İşyeri Büyüklüğüne Göre İstihdam Dağılımı")

        sel_sector2 = st.selectbox("Sektör", ["Tümü"] + sectors, key="size_sector")
        df_size = insured.copy()
        if sel_sector2 != "Tümü":
            df_size = df_size[df_size["ana_sektor"] == sel_sector2]

        size_cols = [c for c in df_size.columns if c.startswith("boy_")]
        size_totals = df_size[size_cols].sum()
        size_df = pd.DataFrame({
            "Büyüklük": [c.replace("boy_", "") for c in size_cols],
            "Sigortalı Sayısı": size_totals.values,
        })

        fig = px.bar(size_df, x="Büyüklük", y="Sigortalı Sayısı",
                     labels={"Büyüklük": "İşyeri Büyüklüğü (Çalışan Sayısı)"},
                     color="Sigortalı Sayısı", color_continuous_scale="Blues")
        fig.update_layout(height=400, template="plotly_white")
        st.plotly_chart(fig, use_container_width=True)

        # KOBİ vs Büyük karşılaştırma
        kobi_cols = ["boy_1", "boy_2-3", "boy_4-6", "boy_7-9", "boy_10-19", "boy_20-29", "boy_30-49"]
        buyuk_cols = [c for c in size_cols if c not in kobi_cols]
        kobi_total = df_size[kobi_cols].sum().sum()
        buyuk_total = df_size[buyuk_cols].sum().sum()
        grand = kobi_total + buyuk_total

        col1, col2 = st.columns(2)
        with col1:
            st.metric("KOBİ İstihdamı (<50)", f"{kobi_total:,.0f}", f"%{kobi_total / grand * 100:.1f}" if grand > 0 else "")
        with col2:
            st.metric("Büyük İşletme (50+)", f"{buyuk_total:,.0f}", f"%{buyuk_total / grand * 100:.1f}" if grand > 0 else "")

    with tab3:
        if not wages.empty and "gunluk_kazanc_toplam" in wages.columns:
            st.subheader("Sektörel Ortalama Günlük Kazanç")

            wage_agg = wages.groupby("ana_sektor").apply(
                lambda x: pd.Series({
                    "ort_kazanc": np.average(x["gunluk_kazanc_toplam"], weights=x["sigortali_toplam"]) if x["sigortali_toplam"].sum() > 0 else 0,
                    "toplam_sigortali": x["sigortali_toplam"].sum(),
                })
            ).reset_index().sort_values("ort_kazanc", ascending=True)

            fig = px.bar(wage_agg, x="ort_kazanc", y="ana_sektor", orientation="h",
                         color="ort_kazanc", color_continuous_scale="RdYlGn",
                         labels={"ort_kazanc": "Ort. Günlük Kazanç (TL)", "ana_sektor": ""},
                         hover_data={"toplam_sigortali": ":,.0f"})
            fig.update_layout(height=600, template="plotly_white")
            st.plotly_chart(fig, use_container_width=True)

            # Cinsiyet farkı
            st.subheader("Cinsiyet Bazlı Günlük Kazanç Farkı")
            gender_agg = wages.groupby("ana_sektor").apply(
                lambda x: pd.Series({
                    "erkek": np.average(x["gunluk_kazanc_erkek"], weights=x["sigortali_toplam"]) if x["sigortali_toplam"].sum() > 0 and x["gunluk_kazanc_erkek"].sum() > 0 else 0,
                    "kadin": np.average(x["gunluk_kazanc_kadin"], weights=x["sigortali_toplam"]) if x["sigortali_toplam"].sum() > 0 and x["gunluk_kazanc_kadin"].sum() > 0 else 0,
                })
            ).reset_index()
            gender_agg["fark_pct"] = ((gender_agg["erkek"] - gender_agg["kadin"]) / gender_agg["erkek"] * 100).round(1)
            gender_agg = gender_agg.sort_values("fark_pct", ascending=True)

            fig2 = go.Figure()
            fig2.add_trace(go.Bar(name="Erkek", x=gender_agg["erkek"], y=gender_agg["ana_sektor"], orientation="h", marker_color="#1E88E5"))
            fig2.add_trace(go.Bar(name="Kadın", x=gender_agg["kadin"], y=gender_agg["ana_sektor"], orientation="h", marker_color="#E53935"))
            fig2.update_layout(barmode="group", height=600, template="plotly_white",
                               xaxis_title="Ortalama Günlük Kazanç (TL)")
            st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("Günlük kazanç verisi için SGK Bölüm 1 dosyasını yükleyin.")


elif page == "🎯 Kadran Analizi":
    st.title("Kadran Analizi: İstihdam Kapasitesi vs İşgücü Maliyeti")

    # Kadran açıklamaları
    col1, col2 = st.columns(2)
    with col1:
        k1 = summary[summary["kadran"].str.startswith("K1")]
        st.markdown(f"""<div class="kadran-k1">
            <strong>K1: Yüksek İstihdam + Yüksek Maliyet ({len(k1)} sektör)</strong><br>
            TEŞVİK ÖNCELİĞİ - Bu sektörler hem çok kişi istihdam ediyor hem de işgücü maliyet oranı yüksek.
        </div>""", unsafe_allow_html=True)

        k3 = summary[summary["kadran"].str.startswith("K3")]
        st.markdown(f"""<div class="kadran-k3">
            <strong>K3: Düşük İstihdam + Yüksek Maliyet ({len(k3)} sektör)</strong><br>
            YAPISAL DÖNÜŞÜM - Verimlilik artışı ve dijital dönüşüm öncelikli.
        </div>""", unsafe_allow_html=True)

    with col2:
        k2 = summary[summary["kadran"].str.startswith("K2")]
        st.markdown(f"""<div class="kadran-k2">
            <strong>K2: Yüksek İstihdam + Düşük Maliyet ({len(k2)} sektör)</strong><br>
            SÜRDÜRÜLEBİLİR YAPI - Mevcut yapı korunmalı, rekabet gücü desteklenmeli.
        </div>""", unsafe_allow_html=True)

        k4 = summary[summary["kadran"].str.startswith("K4")]
        st.markdown(f"""<div class="kadran-k4">
            <strong>K4: Düşük İstihdam + Düşük Maliyet ({len(k4)} sektör)</strong><br>
            İZLEME - Büyüme potansiyeli izlenmeli.
        </div>""", unsafe_allow_html=True)

    st.divider()

    # Kadran scatter plot (daha detaylı)
    fig = px.scatter(
        summary, x="istihdam_payi", y="maliyet_orani_2024",
        size="gkd_2024", color="kadran", text="sektor",
        hover_data={
            "maliyet_etkinligi": ":.2f",
            "kisi_basi_gkd": ":,.0f",
            "kobi_orani": ":.1f",
            "ort_gunluk_kazanc": ":,.0f",
            "maliyet_trend": ":+.1f",
        },
        color_discrete_map={
            "K1: Yüksek İstihdam + Yüksek Maliyet": "#E53935",
            "K2: Yüksek İstihdam + Düşük Maliyet": "#43A047",
            "K3: Düşük İstihdam + Yüksek Maliyet": "#FB8C00",
            "K4: Düşük İstihdam + Düşük Maliyet": "#1E88E5",
        },
        labels={
            "istihdam_payi": "İstihdam Payı (%)",
            "maliyet_orani_2024": "İşgücü Maliyet Oranı (%)",
            "gkd_2024": "GKD (Milyar TL)",
        },
    )
    med_emp = summary["istihdam_payi"].median()
    med_cost = summary["maliyet_orani_2024"].median()
    fig.add_hline(y=med_cost, line_dash="dot", line_color="gray", opacity=0.6)
    fig.add_vline(x=med_emp, line_dash="dot", line_color="gray", opacity=0.6)
    fig.update_traces(textposition="top center", textfont_size=8)
    fig.update_layout(height=650, template="plotly_white",
                      legend=dict(orientation="h", yanchor="bottom", y=-0.2))
    st.plotly_chart(fig, use_container_width=True)

    # Kadran detay tabloları
    st.divider()
    selected_kadran = st.selectbox("Kadran Detay", summary["kadran"].unique())
    k_data = summary[summary["kadran"] == selected_kadran]

    detail_cols = ["sektor", "istihdam", "istihdam_payi", "maliyet_orani_2024",
                   "maliyet_etkinligi", "kisi_basi_gkd", "ort_gunluk_kazanc", "kobi_orani", "kadin_orani", "maliyet_trend"]
    detail_names = {
        "sektor": "Sektör", "istihdam": "İstihdam", "istihdam_payi": "İstihdam Payı (%)",
        "maliyet_orani_2024": "Maliyet Oranı (%)", "maliyet_etkinligi": "Etkinlik",
        "kisi_basi_gkd": "Kişi Başı GKD (TL)", "ort_gunluk_kazanc": "Günlük Kazanç (TL)",
        "kobi_orani": "KOBİ (%)", "kadin_orani": "Kadın (%)", "maliyet_trend": "Trend (puan)",
    }
    avail_cols = [c for c in detail_cols if c in k_data.columns]
    st.dataframe(k_data[avail_cols].rename(columns=detail_names), use_container_width=True)


elif page == "📋 Teşvik Kılavuzu":
    st.title("İstihdam Teşviki Politika Kılavuzu")

    k1 = summary[summary["kadran"].str.startswith("K1")]

    st.header("A. Teşvik Öncelikli Sektörler (Kadran 1)")
    if not k1.empty:
        for _, row in k1.iterrows():
            with st.expander(f"**{row['sektor']}** — Maliyet: %{row.get('maliyet_orani_2024',0):.1f} | İstihdam: {row.get('istihdam',0):,.0f}", expanded=True):
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("İstihdam Payı", f"%{row.get('istihdam_payi',0):.1f}")
                c2.metric("Maliyet Etkinliği", f"{row.get('maliyet_etkinligi',0):.2f}x")
                c3.metric("KOBİ Oranı", f"%{row.get('kobi_orani',0):.1f}")
                c4.metric("Ort. Günlük Kazanç", f"{row.get('ort_gunluk_kazanc',0):,.0f} TL")

                c5, c6 = st.columns(2)
                c5.metric("5 Yıllık Trend", f"{row.get('maliyet_trend',0):+.1f} puan")
                c6.metric("Kadın Oranı", f"%{row.get('kadin_orani',0):.1f}")
    else:
        st.info("Kadran 1'de sektör bulunmamaktadır.")

    st.divider()

    st.header("B. Teşvik Tasarım Parametreleri")

    params = [
        ("SGK Prim İndirimi", """
- **K1 Sektörleri:** İşveren SGK prim payında %10-15 indirim
- **K3 Sektörleri:** %5-10 indirim
- Net istihdam artışı şartı uygulanmalı
- Baz yıl istihdamının altına düşmeme koşulu"""),
        ("İşyeri Büyüklüğüne Göre Kademe", """
| Büyüklük | Teşvik Oranı | Gerekçe |
|----------|-------------|---------|
| Mikro (<10) | Tam oran | Kayıt dışılık en yüksek |
| Küçük (10-49) | %75 | KOBİ desteği |
| Orta (50-249) | %50 | Verimlilik odaklı |
| Büyük (250+) | Net artış için | Deadweight loss önleme |"""),
        ("Bölgesel Farklılaştırma", """
- **6. Bölge illeri:** Ek %5 puan teşvik
- **5. Bölge:** Ek %3 puan
- Bölgesel işgücü arz-talep dengesi dikkate alınmalı
- OSB ve sanayi bölgelerinde ek teşvik"""),
        ("Cinsiyet Eşitliği", """
- Kadın oranı %30 altı sektörlere ek %5 puan
- Hedef sektörler: İnşaat, Ulaştırma, Madencilik
- Kreş ve esnek çalışma desteği ile birlikte"""),
        ("Süre ve Kademeli Çıkış", """
- **1-3. yıl:** Tam oran teşvik
- **4. yıl:** %75'e düşür
- **5. yıl:** %50'ye düşür
- **6. yıl:** Etki değerlendirmesine göre uzat/sonlandır"""),
    ]

    for title, content in params:
        with st.expander(f"**{title}**", expanded=False):
            st.markdown(content)

    st.divider()

    st.header("C. Dikkat Edilecek Hususlar")

    risks = {
        "Ölü Ağırlık Kaybı (Deadweight Loss)": "Zaten istihdam edilecek kişiler için teşvik verilmesi. **Önlem:** Net istihdam artışı şartı, baz yıl koşulu.",
        "İkame Etkisi": "Teşvikli sektörün teşviksiz sektörden işgücü çekmesi. **Önlem:** Sektör bazlı tavanlar.",
        "Kayıt Dışılık": "Teşvik + denetim birlikte yürütülmeli. Kayıt dışından geçiş için ek teşvik.",
        "Mali Sürdürülebilirlik": "Teşvik maliyeti < ek SGK + vergi geliri olmalı. Yıllık maliyet-fayda analizi zorunlu.",
        "Rekabet Bozulması": "Tüm işletmelerin eşit koşullarda yararlanması sağlanmalı.",
    }

    for risk, desc in risks.items():
        st.warning(f"**{risk}:** {desc}")


elif page == "📖 Metodoloji":
    st.title("Metodoloji ve Navigasyon Rehberi")

    st.header("Navigasyon Rehberi")
    st.markdown("""
Bu uygulama, TÜİK ve SGK verilerini kullanarak Türkiye'deki sektörlerin işgücü maliyeti yapısını analiz etmektedir.
Aşağıda her sayfanın ne sunduğu ve nasıl kullanılacağı açıklanmaktadır:

| Sayfa | Amacı | Ne Bulursunuz? |
|-------|-------|----------------|
| **Özet Dashboard** | Büyük resmi görmek | 20 ana sektörün tüm temel göstergeleri, kadran sınıflandırması, scatter plot ve sıralama grafikleri |
| **Trend Analizi** | Tarihsel değişimi izlemek | 2009-2024 arası sektörel işgücü maliyet oranı, GKD ve işgücü ödemesi trendleri; dönem karşılaştırması |
| **İstihdam Yapısı** | Alt sektör detaylarını incelemek | NACE 2 haneli alt sektör kırılımları, işyeri büyüklüğü dağılımı, günlük kazanç ve cinsiyet farkları |
| **Kadran Analizi** | Teşvik önceliklerini belirlemek | İstihdam payı × Maliyet oranı kadran matrisi, her kadrandaki sektörlerin detay tabloları |
| **Teşvik Kılavuzu** | Politika önerilerini görmek | K1 sektör profilleri, SGK prim indirimi, bölgesel farklılaştırma, cinsiyet eşitliği önerileri |
| **Metodoloji** | Analizin nasıl yapıldığını anlamak | Veri kaynakları, NACE eşleştirmesi, hesaplama formülleri, kadran sınıflandırma yöntemi |
| **Rapor İndir** | Sonuçları almak | Excel raporu, Word raporu ve akademik analiz raporu (docx) |
""")

    st.divider()
    st.header("Veri Kaynakları ve İçerikleri")

    st.subheader("1. TÜİK Verileri")
    st.markdown("""
- **Tablo I.2.14 — Gelir Yöntemiyle GSYH** (kaynak dosya: `SEKTÖR BAZLI İŞGÜCÜ EKONOMİK GÖSTERGESİ.xlsx`)
  - İktisadi faaliyet kollarına (NACE Rev.2 bölüm düzeyi, 20 ana sektör) göre:
    - Gayrisafi Katma Değer (GKD) — Milyar TL, cari fiyatlarla
    - İşgücüne Yapılan Ödemeler — Milyar TL
    - İşletme Artığı (Brüt) — Milyar TL
  - Dönem: 2009-2024 (yıllık)
""")

    st.subheader("2. SGK Verileri")
    st.markdown("""
- **TABLO-1.12 — İşyeri Sayıları** (kaynak dosya: `SEKTÖR BAZLI İŞGÜCÜ EKONOMİK GÖSTERGESİ.xlsx`)
  - NACE Rev.2 faaliyet grupları (2 haneli, ~87 alt sektör) × İşyeri büyüklüğü (1, 2-3, ..., 1000+)
- **TABLO-1.13 — Zorunlu Sigortalı Sayıları** (kaynak dosya: aynı)
  - NACE Rev.2 faaliyet grupları × İşyeri büyüklüğü dağılımı
- **TABLO-1.16 — Prime Esas Ortalama Günlük Kazanç** (kaynak dosya: SGK Bölüm 1)
  - NACE Rev.2 faaliyet grupları × Daimi/Geçici, Kamu/Özel, Erkek/Kadın kırılımları
""")

    st.divider()
    st.header("TÜİK - SGK Sektör Eşleştirmesi (NACE Mapping)")

    st.markdown("""
TÜİK'in GSYH tablosu (I.2.14) **20 ana sektör** (NACE Rev.2 bölüm düzeyi — tek harf kodu) kullanırken,
SGK tabloları **NACE 2 haneli kodları** (~87 alt sektör) kullanmaktadır.

Eşleştirme, uluslararası standart NACE Rev.2 hiyerarşisine dayalı olarak yapılmaktadır:
- Her 2 haneli NACE kodu, üst düzey NACE bölümüne (A, B, C, ... T) karşılık gelir
- Bu bölümler TÜİK'in 20 ana sektörüne birebir eşlenir

**Eşleştirme yöntemi:** `data_loader.py` içindeki `NACE_MAPPING` sözlüğü, her ana sektör için ilgili 2 haneli NACE kodlarını tanımlar. `NACE_TO_SECTOR` ters eşleştirmesi ile SGK'daki her alt sektör ilgili ana sektöre atanır.
""")

    mapping_data = []
    for sector, codes in NACE_MAPPING.items():
        code_str = ", ".join([f"{c:02d}" for c in codes])
        if len(codes) <= 3:
            code_range = code_str
        else:
            code_range = f"{codes[0]:02d}-{codes[-1]:02d}"
        mapping_data.append({"Ana Sektör (TÜİK)": sector, "NACE 2 Haneli Kodlar": code_range, "Alt Sektör Sayısı": len(codes)})

    st.dataframe(pd.DataFrame(mapping_data), use_container_width=True, hide_index=True)

    st.divider()
    st.header("Hesaplama Formülleri")

    st.markdown("""
| Gösterge | Formül | Birimi | Açıklama |
|----------|--------|--------|----------|
| **İşgücü Maliyet Oranı** | İşgücü Ödemesi / GKD × 100 | % | Katma değerin ne kadarı işgücüne gidiyor |
| **Kişi Başı Katma Değer** | GKD / Sigortalı Sayısı × 10⁹ | TL | İşgücü verimliliğinin bir göstergesi |
| **Kişi Başı İşgücü Maliyeti** | İşgücü Ödemesi / Sigortalı Sayısı × 10⁹ | TL | Ortalama işgücü maliyeti |
| **İstihdam Yoğunluğu** | Sektör Sigortalısı / Toplam × 100 | % | Sektörün toplam istihdamdaki payı |
| **Maliyet Etkinliği** | GKD / İşgücü Ödemesi | Katsayı | 1 TL işgücüne karşı ne kadar katma değer üretiliyor |
| **İşletme Artığı Payı** | İşletme Artığı / GKD × 100 | % | Sermaye getirisi payı |
| **KOBİ Yoğunluğu** | <50 çalışanlı işyeri / Toplam işyeri × 100 | % | Küçük işletme oranı |
| **Trend** | Maliyet Oranı (2024) − Maliyet Oranı (2019) | Puan | 5 yıllık maliyet baskısı değişimi |
""")

    st.divider()
    st.header("Kadran Sınıflandırma Yöntemi")

    st.markdown("""
Sektörler iki eksen üzerinde medyan değerlere göre dört kadrana ayrılmaktadır:

- **X Ekseni:** İstihdam Payı (%) — sektörün toplam istihdamdaki ağırlığı
- **Y Ekseni:** İşgücü Maliyet Oranı (%) — işgücü ödemesinin katma değere oranı

| Kadran | İstihdam Payı | Maliyet Oranı | Anlam | Politika Yanıtı |
|--------|--------------|---------------|-------|-----------------|
| **K1** | ≥ Medyan | ≥ Medyan | Yüksek İstihdam + Yüksek Maliyet | TEŞVİK ÖNCELİĞİ |
| **K2** | ≥ Medyan | < Medyan | Yüksek İstihdam + Düşük Maliyet | Sürdürülebilir yapı, koruma |
| **K3** | < Medyan | ≥ Medyan | Düşük İstihdam + Yüksek Maliyet | Yapısal dönüşüm |
| **K4** | < Medyan | < Medyan | Düşük İstihdam + Düşük Maliyet | İzleme, stratejik destek |

Medyan değerler verideki 20 sektörün ortanca değerleridir ve dinamik olarak hesaplanır.
""")

    st.divider()
    st.header("Uluslararası Karşılaştırma Kaynakları")
    st.markdown("""
Akademik raporda kullanılan uluslararası kaynaklar:

- **IMF** — World Economic Outlook (Ekim 2025, Ocak 2026): Küresel büyüme tahminleri, sektörel trendler
- **OECD** — Türkiye Ekonomik İncelemesi (2025): Verimlilik, kadın işgücü, hizmetler sektörü önerileri
- **OECD** — Taxing Wages 2025: Türkiye vergi takozu (%39,0 vs OECD ort. %34,9)
- **Dünya Bankası** — Türkiye Ülke Ekonomik Memorandumu: İstihdam ile Refah
- **ILO** — Dünya İstihdam ve Sosyal Görünüm 2025: Küresel istihdam eğilimleri
- **UNIDO** — Sanayi İstatistikleri Yıllığı 2025: İmalat sektörü çarpan etkisi
- **AB Komisyonu** — CBAM (Karbon Sınır Düzenleme Mekanizması): Türkiye etkisi
""")


elif page == "📥 Rapor İndir":
    st.title("Rapor İndir")
    st.write("Analiz sonuçlarını Excel, Word veya akademik rapor formatında indirebilirsiniz.")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📗 Excel Raporu")
        st.write("6 sayfalık detaylı Excel raporu: Dashboard, Trend, İstihdam Yapısı, Kadran Analizi, Teşvik Kılavuzu, Veri Kaynakları")
        if st.button("Excel Oluştur", type="primary"):
            with st.spinner("Excel raporu hazırlanıyor..."):
                excel_buf = export_excel(data)
            st.download_button(
                "📥 Excel İndir",
                data=excel_buf,
                file_name="Sektor_Bazli_Isgucu_Analiz_Raporu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with col2:
        st.subheader("📘 Word Raporu")
        st.write("Profesyonel Word raporu: Yönetici Özeti, Sektörel Tablo, Kadran Analizi, Politika Önerileri, Risk Analizi")
        if st.button("Word Oluştur", type="primary"):
            with st.spinner("Word raporu hazırlanıyor..."):
                word_buf = export_word(data)
            st.download_button(
                "📥 Word İndir",
                data=word_buf,
                file_name="Sektor_Bazli_Isgucu_Analiz_Raporu.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

    st.divider()

    st.subheader("📕 Akademik Analiz Raporu")
    st.write("""Küresel ekonomik gelişmeler ve uluslararası kurumların (IMF, OECD, Dünya Bankası, ILO, UNIDO)
    sektörel tahminleri ışığında Türkiye'de desteklenmesi gereken sektörleri belirleyen, atıf kurallarına
    uygun akademik rapor. İçerik: Yönetici Özeti, Küresel Görünüm, Sektörel Analiz, Kadran Analizi,
    Uluslararası Kurum Değerlendirmeleri, Stratejik Öncelikli Sektörler, Politika Önerileri, Kaynakça.""")
    if st.button("Akademik Rapor Oluştur", type="primary"):
        with st.spinner("Akademik rapor hazırlanıyor..."):
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_report:
                create_academic_report(data, tmp_report.name)
                tmp_report.seek(0)
                with open(tmp_report.name, "rb") as f:
                    report_bytes = f.read()
        st.download_button(
            "📥 Akademik Rapor İndir (.docx)",
            data=report_bytes,
            file_name="Kuresel_Gelismeler_Isiginda_Sektorel_Analiz_Raporu.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    st.divider()
    st.subheader("Veri Özeti")
    st.json({
        "Analiz Dönemi": "2024",
        "Trend Dönemi": "2009-2024",
        "Sektör Sayısı (Ana)": len(summary),
        "Alt Sektör Sayısı (NACE)": len(insured),
        "Toplam İstihdam": int(summary["istihdam"].sum()),
        "Toplam İşyeri": int(summary["isyeri"].sum()),
        "Toplam GKD (Milyar TL)": round(summary["gkd_2024"].sum(), 1),
        "K1 Sektör Sayısı": len(summary[summary["kadran"].str.startswith("K1")]),
    })
