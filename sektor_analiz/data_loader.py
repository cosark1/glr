"""
Veri yükleme ve işleme modülü.
TÜİK GSYH (I.2.14), SGK işyeri/sigortalı (TABLO-1.12/1.13/1.16) verilerini okur,
NACE eşleştirmesi yapar ve gösterge hesaplar.
"""
import pandas as pd
import numpy as np

# ── NACE 2-haneli → TÜİK Ana Sektör eşleştirmesi ──────────────────────────
NACE_MAPPING = {
    "Tarım, Ormancılık ve Balıkçılık": list(range(1, 4)),
    "Madencilik ve Taşocakçılığı": list(range(5, 10)),
    "İmalat Sanayi": list(range(10, 34)),
    "Elektrik, Gaz, Buhar ve İklimlendirme": [35],
    "Su Temini; Kanalizasyon, Atık Yönetimi": list(range(36, 40)),
    "İnşaat": list(range(41, 44)),
    "Toptan ve Perakende Ticaret": list(range(45, 48)),
    "Ulaştırma, Depolama": list(range(49, 54)),
    "Konaklama ve Yiyecek Hizmeti": [55, 56],
    "Bilgi ve İletişim": list(range(58, 64)),
    "Finans ve Sigorta Faaliyetleri": list(range(64, 67)),
    "Gayrimenkul Faaliyetleri": [68],
    "Mesleki, Bilimsel ve Teknik Faal.": list(range(69, 76)),
    "İdari ve Destek Hizmet Faaliyetleri": list(range(77, 83)),
    "Kamu Yönetimi ve Savunma": [84],
    "Eğitim": [85],
    "İnsan Sağlığı ve Sosyal Hizmet": list(range(86, 89)),
    "Kültür, Sanat, Eğlence, Spor": list(range(90, 94)),
    "Diğer Hizmet Faaliyetleri": list(range(94, 97)),
    "Hanehalkları İşverenler": [97, 98],
}

# Ters mapping: NACE kodu → sektör adı
NACE_TO_SECTOR = {}
for sector, codes in NACE_MAPPING.items():
    for code in codes:
        NACE_TO_SECTOR[code] = sector

# TÜİK Excel'deki ham sektör adları → canonical NACE_MAPPING anahtarlarına eşleştirme
# (TÜİK'te sondaki boşluklar, farklı uzunlukta isimler olabiliyor)
TUIK_NAME_NORMALIZE = {
    "Tarım, Ormancılık ve Balıkçılık": "Tarım, Ormancılık ve Balıkçılık",
    "Madencilik ve Taşocakçılığı": "Madencilik ve Taşocakçılığı",
    "İmalat Sanayi": "İmalat Sanayi",
    "Elektrik, Gaz, Buhar ve İklimlendirme Üretimi ve Dağıtımı": "Elektrik, Gaz, Buhar ve İklimlendirme",
    "Su Temini; Kanalizasyon, Atık Yönetimi ve İyileştirme Faal.": "Su Temini; Kanalizasyon, Atık Yönetimi",
    "İnşaat": "İnşaat",
    "Toptan ve Perakende Ticaret": "Toptan ve Perakende Ticaret",
    "Ulaştırma, Depolama": "Ulaştırma, Depolama",
    "Konaklama ve Yiyecek Hizmeti Faaliyetleri": "Konaklama ve Yiyecek Hizmeti",
    "Bilgi ve İletişim": "Bilgi ve İletişim",
    "Finans ve Sigorta Faaliyetleri": "Finans ve Sigorta Faaliyetleri",
    "Gayrimenkul Faaliyetleri": "Gayrimenkul Faaliyetleri",
    "Mesleki, Bilimsel ve Teknik Faaliyetler": "Mesleki, Bilimsel ve Teknik Faal.",
    "İdari ve Destek Hizmet Faaliyetleri": "İdari ve Destek Hizmet Faaliyetleri",
    "Kamu Yönetimi ve Savunma; Zorunlu Sosyal Güvenlik": "Kamu Yönetimi ve Savunma",
    "Eğitim": "Eğitim",
    "İnsan Sağlığı ve Sosyal Hizmet Faaliyetleri": "İnsan Sağlığı ve Sosyal Hizmet",
    "Kültür, Sanat, Eğlence, Dinlence ve Spor": "Kültür, Sanat, Eğlence, Spor",
    "Diğer Hizmet Faaliyetleri": "Diğer Hizmet Faaliyetleri",
    "Hanehalklarının İşverenler Olarak Faaliyetleri": "Hanehalkları İşverenler",
}

# "Sektör Toplamı" gibi toplam satırları hariç tutulacak
EXCLUDE_SECTORS = {"Sektör Toplamı"}

YEARS = list(range(2009, 2025))
YEAR_COLS = {y: i for i, y in enumerate(YEARS)}  # yıl → kolon indeks (0-based from data start)


def _normalize_sector_name(raw_name: str) -> str:
    """TÜİK ham sektör adını canonical NACE_MAPPING anahtarına dönüştürür."""
    # Önce tam eşleşme dene
    if raw_name in TUIK_NAME_NORMALIZE:
        return TUIK_NAME_NORMALIZE[raw_name]
    # Fuzzy: NACE_MAPPING anahtarlarında alt-string ara
    for canonical in NACE_MAPPING:
        if canonical in raw_name or raw_name in canonical:
            return canonical
    return raw_name


def load_tuik_gdp(path: str, sheet: str = "I.2.14") -> pd.DataFrame:
    """TÜİK I.2.14 sayfasından sektörel GSYH bileşenlerini çeker."""
    df_raw = pd.read_excel(path, sheet_name=sheet, header=None)

    sectors = []
    current_sector = None

    for i in range(df_raw.shape[0]):
        col0 = df_raw.iloc[i, 0]
        col1 = df_raw.iloc[i, 1]

        # Sektör adı col0'da, "Gayrisafi Katma Değer" col1'de
        if pd.notna(col0) and pd.notna(col1) and "Katma" in str(col1):
            raw_name = str(col0).strip()
            if raw_name in EXCLUDE_SECTORS:
                continue
            current_sector = raw_name
            gkd_row = i
        elif current_sector and pd.notna(col1):
            label = str(col1).strip()
            if "İşgücüne" in label or "cüne" in label:
                labor_row = i
            elif "İşletme" in label and "Brüt" in label:
                surplus_row = i
                # Sektör verilerini topla - canonical isim kullan
                canonical_name = _normalize_sector_name(current_sector)
                row_data = {"sektor": canonical_name}
                for y in YEARS:
                    col_idx = 2 + YEARS.index(y)
                    gkd_val = df_raw.iloc[gkd_row, col_idx]
                    labor_val = df_raw.iloc[labor_row, col_idx]
                    surplus_val = df_raw.iloc[surplus_row, col_idx]
                    row_data[f"gkd_{y}"] = _to_float(gkd_val)
                    row_data[f"isgucu_{y}"] = _to_float(labor_val)
                    row_data[f"isletme_artigi_{y}"] = _to_float(surplus_val)
                sectors.append(row_data)
                current_sector = None

    return pd.DataFrame(sectors)


def load_sgk_workplace(path: str, sheet: str = "TABLO-1.12") -> pd.DataFrame:
    """SGK TABLO-1.12: NACE 2 haneli işyeri sayıları."""
    return _load_sgk_nace_table(path, sheet)


def load_sgk_insured(path: str, sheet: str = "TABLO-1.13") -> pd.DataFrame:
    """SGK TABLO-1.13: NACE 2 haneli zorunlu sigortalı sayıları."""
    return _load_sgk_nace_table(path, sheet)


def _load_sgk_nace_table(path: str, sheet: str) -> pd.DataFrame:
    """SGK NACE tablolarını okur (1.12 ve 1.13 ortak yapı)."""
    df_raw = pd.read_excel(path, sheet_name=sheet, header=None)

    size_labels = ["1", "2-3", "4-6", "7-9", "10-19", "20-29", "30-49",
                   "50-99", "100-249", "250-499", "500-749", "750-999", "1000+"]

    rows = []
    for i in range(8, df_raw.shape[0]):
        code = df_raw.iloc[i, 0]
        desc = df_raw.iloc[i, 1]
        total = df_raw.iloc[i, 15]

        if pd.isna(code) or not str(code).strip().isdigit():
            continue

        nace_code = int(str(code).strip())
        row = {
            "nace_kodu": nace_code,
            "faaliyet": str(desc).strip() if pd.notna(desc) else "",
            "toplam": _to_float(total),
        }
        # Büyüklük dağılımı
        for j, label in enumerate(size_labels):
            val = df_raw.iloc[i, 2 + j]
            row[f"boy_{label}"] = _to_float(val)

        row["ana_sektor"] = NACE_TO_SECTOR.get(nace_code, "Diğer")
        rows.append(row)

    return pd.DataFrame(rows)


def load_sgk_wages(path: str, sheet: str = "TABLO-1.16") -> pd.DataFrame:
    """SGK TABLO-1.16: Prime esas ortalama günlük kazançlar."""
    df_raw = pd.read_excel(path, sheet_name=sheet, header=None)

    rows = []
    for i in range(8, df_raw.shape[0]):
        code = df_raw.iloc[i, 0]
        if pd.isna(code) or not str(code).strip().isdigit():
            continue

        nace_code = int(str(code).strip())
        row = {
            "nace_kodu": nace_code,
            "faaliyet": str(df_raw.iloc[i, 1]).strip() if pd.notna(df_raw.iloc[i, 1]) else "",
            "isyeri_toplam": _to_float(df_raw.iloc[i, 6]),
            "sigortali_toplam": _to_float(df_raw.iloc[i, 13]),
            "sigortali_erkek": _to_float(df_raw.iloc[i, 11]),
            "sigortali_kadin": _to_float(df_raw.iloc[i, 12]),
            "sigortali_kamu": _to_float(df_raw.iloc[i, 9]),
            "sigortali_ozel": _to_float(df_raw.iloc[i, 10]),
            "gunluk_kazanc_daimi": _to_float(df_raw.iloc[i, 14]),
            "gunluk_kazanc_gecici": _to_float(df_raw.iloc[i, 15]),
            "gunluk_kazanc_kamu": _to_float(df_raw.iloc[i, 16]),
            "gunluk_kazanc_ozel": _to_float(df_raw.iloc[i, 17]),
            "gunluk_kazanc_erkek": _to_float(df_raw.iloc[i, 18]),
            "gunluk_kazanc_kadin": _to_float(df_raw.iloc[i, 19]),
            "gunluk_kazanc_toplam": _to_float(df_raw.iloc[i, 20]),
        }
        row["ana_sektor"] = NACE_TO_SECTOR.get(nace_code, "Diğer")
        rows.append(row)

    return pd.DataFrame(rows)


def compute_sector_summary(df_tuik: pd.DataFrame, df_insured: pd.DataFrame,
                           df_workplace: pd.DataFrame, df_wages: pd.DataFrame) -> pd.DataFrame:
    """Ana sektör düzeyinde özet gösterge tablosu oluşturur."""
    # SGK verilerini ana sektör bazında topla
    ins_agg = df_insured.groupby("ana_sektor").agg(
        istihdam=("toplam", "sum"),
        isyeri=("toplam", "sum"),
    ).reset_index()

    # İşyeri büyüklüğü - KOBİ (<50 çalışan) oranı
    size_small = ["boy_1", "boy_2-3", "boy_4-6", "boy_7-9", "boy_10-19", "boy_20-29", "boy_30-49"]
    size_all = [c for c in df_workplace.columns if c.startswith("boy_")]
    wp_agg = df_workplace.groupby("ana_sektor")[size_small + size_all].sum().reset_index()
    wp_agg["kobi_isyeri"] = wp_agg[size_small].sum(axis=1)
    wp_agg["toplam_isyeri"] = wp_agg[size_all].sum(axis=1)
    wp_agg["kobi_orani"] = (wp_agg["kobi_isyeri"] / wp_agg["toplam_isyeri"] * 100).round(1)

    # Ortalama günlük kazanç (ağırlıklı ortalama)
    df_wages["kazanc_x_sigortali"] = df_wages["gunluk_kazanc_toplam"] * df_wages["sigortali_toplam"]
    wage_agg = df_wages.groupby("ana_sektor").agg(
        kazanc_x_toplam=("kazanc_x_sigortali", "sum"),
        sigortali_toplam=("sigortali_toplam", "sum"),
        kadin_sigortali=("sigortali_kadin", "sum"),
        erkek_sigortali=("sigortali_erkek", "sum"),
    ).reset_index()
    wage_agg["ort_gunluk_kazanc"] = (wage_agg["kazanc_x_toplam"] / wage_agg["sigortali_toplam"]).round(2)
    wage_agg["kadin_orani"] = (wage_agg["kadin_sigortali"] / wage_agg["sigortali_toplam"] * 100).round(1)

    # TÜİK verilerinden 2024 ve trend
    tuik_cols = {
        "sektor": "sektor",
        "gkd_2024": "gkd_2024",
        "isgucu_2024": "isgucu_2024",
        "isletme_artigi_2024": "isletme_artigi_2024",
        "gkd_2019": "gkd_2019",
        "isgucu_2019": "isgucu_2019",
    }
    df_t = df_tuik[list(tuik_cols.keys())].copy()
    df_t.columns = list(tuik_cols.values())

    # İşgücü maliyet oranı
    df_t["maliyet_orani_2024"] = (df_t["isgucu_2024"] / df_t["gkd_2024"] * 100).round(1)
    df_t["maliyet_orani_2019"] = (df_t["isgucu_2019"] / df_t["gkd_2019"] * 100).round(1)
    df_t["maliyet_trend"] = (df_t["maliyet_orani_2024"] - df_t["maliyet_orani_2019"]).round(1)
    df_t["isletme_artigi_payi"] = (df_t["isletme_artigi_2024"] / df_t["gkd_2024"] * 100).round(1)
    df_t["maliyet_etkinligi"] = (df_t["gkd_2024"] / df_t["isgucu_2024"]).round(2)

    # Birleştir
    summary = df_t.merge(ins_agg, left_on="sektor", right_on="ana_sektor", how="left")
    summary = summary.merge(wp_agg[["ana_sektor", "kobi_orani"]], left_on="sektor", right_on="ana_sektor", how="left")
    summary = summary.merge(wage_agg[["ana_sektor", "ort_gunluk_kazanc", "kadin_orani"]],
                            left_on="sektor", right_on="ana_sektor", how="left")

    # Kişi başı göstergeler (GKD Milyar TL → TL, istihdam kişi)
    summary["kisi_basi_gkd"] = (summary["gkd_2024"] * 1_000_000_000 / summary["istihdam"]).round(0)
    summary["kisi_basi_isgucu_maliyeti"] = (summary["isgucu_2024"] * 1_000_000_000 / summary["istihdam"]).round(0)

    # İstihdam payı
    toplam_istihdam = summary["istihdam"].sum()
    summary["istihdam_payi"] = (summary["istihdam"] / toplam_istihdam * 100).round(2)

    # Kadran sınıflandırması
    med_istihdam = summary["istihdam_payi"].median()
    med_maliyet = summary["maliyet_orani_2024"].median()

    def assign_quadrant(row):
        high_emp = row["istihdam_payi"] >= med_istihdam
        high_cost = row["maliyet_orani_2024"] >= med_maliyet
        if high_emp and high_cost:
            return "K1: Yüksek İstihdam + Yüksek Maliyet"
        elif high_emp and not high_cost:
            return "K2: Yüksek İstihdam + Düşük Maliyet"
        elif not high_emp and high_cost:
            return "K3: Düşük İstihdam + Yüksek Maliyet"
        else:
            return "K4: Düşük İstihdam + Düşük Maliyet"

    summary["kadran"] = summary.apply(assign_quadrant, axis=1)

    # Sütun seçimi ve sıralama
    result_cols = [
        "sektor", "gkd_2024", "isgucu_2024", "isletme_artigi_2024",
        "maliyet_orani_2024", "maliyet_orani_2019", "maliyet_trend",
        "maliyet_etkinligi", "isletme_artigi_payi",
        "istihdam", "isyeri", "istihdam_payi",
        "kisi_basi_gkd", "kisi_basi_isgucu_maliyeti",
        "ort_gunluk_kazanc", "kobi_orani", "kadin_orani", "kadran"
    ]
    summary = summary[[c for c in result_cols if c in summary.columns]]
    summary = summary.sort_values("istihdam_payi", ascending=False).reset_index(drop=True)

    return summary


def compute_trend_data(df_tuik: pd.DataFrame) -> pd.DataFrame:
    """Her sektör için yıllık işgücü maliyet oranı trendi."""
    rows = []
    for _, row in df_tuik.iterrows():
        for y in YEARS:
            gkd = row.get(f"gkd_{y}", 0)
            isgucu = row.get(f"isgucu_{y}", 0)
            if gkd and gkd > 0:
                rows.append({
                    "sektor": row["sektor"],
                    "yil": y,
                    "gkd": gkd,
                    "isgucu": isgucu,
                    "maliyet_orani": round(isgucu / gkd * 100, 2),
                })
    return pd.DataFrame(rows)


def _to_float(val) -> float:
    """Güvenli float dönüşümü."""
    if pd.isna(val):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def load_all_data(ana_dosya_path: str, sgk_bolum1_path: str = None):
    """Tüm verileri yükler ve hesaplar. Tek giriş noktası."""
    df_tuik = load_tuik_gdp(ana_dosya_path)
    df_workplace = load_sgk_workplace(ana_dosya_path)
    df_insured = load_sgk_insured(ana_dosya_path)

    # SGK Bölüm 1 varsa oradan günlük kazanç al
    if sgk_bolum1_path:
        df_wages = load_sgk_wages(sgk_bolum1_path)
    else:
        df_wages = _create_empty_wages(df_insured)

    summary = compute_sector_summary(df_tuik, df_insured, df_workplace, df_wages)
    trend = compute_trend_data(df_tuik)

    return {
        "tuik": df_tuik,
        "workplace": df_workplace,
        "insured": df_insured,
        "wages": df_wages,
        "summary": summary,
        "trend": trend,
    }


def _create_empty_wages(df_insured):
    """Günlük kazanç verisi yoksa boş DataFrame oluşturur."""
    rows = []
    for _, r in df_insured.iterrows():
        rows.append({
            "nace_kodu": r["nace_kodu"],
            "faaliyet": r["faaliyet"],
            "sigortali_toplam": r["toplam"],
            "sigortali_kadin": 0,
            "sigortali_erkek": 0,
            "gunluk_kazanc_toplam": 0,
            "ana_sektor": r["ana_sektor"],
            "kazanc_x_sigortali": 0,
        })
    return pd.DataFrame(rows)
