import io
import base64
from datetime import datetime, date
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from supabase import create_client, Client

# -------------------------------------------------------------
# 🌐 SUPABASE BAĞLANTI BİLGİLERİ
# -------------------------------------------------------------
try:
    SUPABASE_URL = st.secrets["SUPABASE_URL"]
    SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
except Exception:
    SUPABASE_URL = "https://ngokzlndzpodmjiffmjv.supabase.co"
    SUPABASE_KEY = "sb_publishable_LJldycoOPfyCh-stDwAFjg_EVpjACxQ"

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

try:
    supabase = init_supabase()
except Exception as e:
    st.error("Supabase bağlantısı kurulamadı. Lütfen bilgilerinizi kontrol edin.")

# Sayfa Yapılandırması
st.set_page_config(page_title="Yalçın Marketler Zinciri - Manav Portalı", page_icon="🥭", layout="wide")

# -------------------------------------------------------------
# 🎨 DİNAMİK TEMA UYUMLU CSS DÜZENLEMELERİ
# -------------------------------------------------------------
st.markdown("""
    <style>
        #MainMenu {visibility: hidden !important;}
        footer {visibility: hidden !important;}
        header {visibility: hidden !important;}
        [data-testid="stHeader"] {display: none !important;}
        [data-testid="stSidebar"] {display: none !important;}
        
        .block-container {
            padding-top: 2rem !important;
            padding-bottom: 2rem !important;
        }

        @keyframes fadeInZoom {
            0% { opacity: 0; transform: scale(0.9); }
            100% { opacity: 1; transform: scale(1); }
        }

        .logo-card-container {
            animation: fadeInZoom 1s ease-out forwards;
            background-color: #ffffff !important;
            border-radius: 20px;
            padding: 30px 20px;
            box-shadow: 0 10px 25px rgba(0, 0, 0, 0.15);
            max-width: 460px;
            margin: 0 auto 20px auto;
            display: flex;
            justify-content: center;
            align-items: center;
        }

        .animated-logo {
            max-width: 100%;
            height: auto;
            display: block;
        }

        .welcome-title {
            text-align: center;
            font-size: 26px;
            font-weight: 800;
            letter-spacing: 1px;
            color: var(--text-color) !important;
            margin-top: 15px;
            margin-bottom: 5px;
        }
        
        .welcome-sub {
            text-align: center;
            font-size: 15px;
            color: var(--text-color) !important;
            opacity: 0.75;
            margin-bottom: 25px;
        }
    </style>
""", unsafe_allow_html=True)

SUBE_LISTESI = [
    "Raufbey", "Metin Tamer", "Hacı Osmanlı", "Salı Yolu", "Kadiri Yolu", 
    "Nahır Yolu", "Eyup Sultan", "Bulvar", "Düziçi Çarşı", "Aşiyan", "Zeytinlik"
]

SUBE_SIFRELERI = {
    "Raufbey": "1001", "Metin Tamer": "1002", "Hacı Osmanlı": "1003",
    "Salı Yolu": "1004", "Kadiri Yolu": "1005", "Nahır Yolu": "1006",
    "Eyup Sultan": "1007", "Bulvar": "1008", "Düziçi Çarşı": "1009",
    "Aşiyan": "1010", "Zeytinlik": "1011"
}

HAL_SIFRESI = "2024"
YONETICI_SIFRESI = "1234"

URUNLER = [
    {"KODU": "053016", "ADI": "MNV.ACI DOLMALIK"}, {"KODU": "09857", "ADI": "MNV.ALA KARPUZ"},
    {"KODU": "00015264", "ADI": "MNV.ANANAS"}, {"KODU": "08385", "ADI": "MNV.ARMUT"},
    {"KODU": "84", "ADI": "MNV.AVOKADO ADET"}, {"KODU": "058418", "ADI": "MNV.BAMYA"},
    {"KODU": "09921", "ADI": "MNV.BARBUNYA"}, {"KODU": "055710", "ADI": "MNV.BEYAZ SOGAN"},
    {"KODU": "01248", "ADI": "MNV.BEYAZ UZUM"}, {"KODU": "09965", "ADI": "MNV.BURSA DOMATES"},
    {"KODU": "B.2901020", "ADI": "MNV.BURSA SEFTALI"}, {"KODU": "05695", "ADI": "MNV.CARLISTON BIBER"},
    {"KODU": "09859", "ADI": "MNV.CEKIRDEKSIZ KARPUZ"}, {"KODU": "04239", "ADI": "MNV.CEKIRDEKSIZ UZUM"},
    {"KODU": "09911", "ADI": "MNV.CERI DOMATES"}, {"KODU": "01127", "ADI": "MNV.CILEK"},
    {"KODU": "00001922", "ADI": "MNV.DERE OTU"}, {"KODU": "09949", "ADI": "MNV.DEVECI ARMUT"},
    {"KODU": "05485", "ADI": "MNV.DOLMALIK BIBER"}, {"KODU": "B.2801083", "ADI": "MNV.EJDER MEYVESI ADET"},
    {"KODU": "07704", "ADI": "MNV.ELMA ARJANTIN"}, {"KODU": "07703", "ADI": "MNV.ELMA GOLDEN"},
    {"KODU": "07701", "ADI": "MNV.ELMA STARKING"}, {"KODU": "09966", "ADI": "MNV.ERIK KG"},
    {"KODU": "06108", "ADI": "MNV.FASULYE YESIL"}, {"KODU": "120", "ADI": "MNV.FIRIK MISIR"},
    {"KODU": "81", "ADI": "MNV.GOBEKLI MARUL"}, {"KODU": "39", "ADI": "MNV.HINDISTAN CEVIZI"},
    {"KODU": "09941", "ADI": "MNV.ISPANAK"}, {"KODU": "024179", "ADI": "MNV.ITALYAN ERIK"},
    {"KODU": "053743", "ADI": "MNV.ITHAL MUZ"}, {"KODU": "01785", "ADI": "MNV.JALAPENO BIBER"},
    {"KODU": "05773", "ADI": "MNV.KABAK BEYAZ"}, {"KODU": "06374", "ADI": "MNV.KABAK SIYAH"},
    {"KODU": "07603", "ADI": "MNV.KARA LAHANA"}, {"KODU": "07451", "ADI": "MNV.KAVUN ANKARA"},
    {"KODU": "B.2901023", "ADI": "MNV.KAVURMALIK SOGAN"}, {"KODU": "053054", "ADI": "MNV.KAYISI"},
    {"KODU": "05790", "ADI": "MNV.KEMER PATLICAN"}, {"KODU": "09935", "ADI": "MNV.KIRAZ"},
    {"KODU": "01151", "ADI": "MNV.KIRMIZI KAPYA BIBER"}, {"KODU": "056063", "ADI": "MNV.KIRMIZI PANCAR"},
    {"KODU": "01153", "ADI": "MNV.KIRMIZI SILI BIBER"}, {"KODU": "06375", "ADI": "MNV.KIRMIZI SOGAN"},
    {"KODU": "09398", "ADI": "MNV.KIVI"}, {"KODU": "76", "ADI": "MNV.KIVIRCIK MARUL"},
    {"KODU": "B.2901017", "ADI": "MNV.KOZLEMELIK PATLICAN"}, {"KODU": "053197", "ADI": "MNV.KURU SARIMSAK"},
    {"KODU": "09934", "ADI": "MNV.LIMON"}, {"KODU": "09997", "ADI": "MNV.LUX SALATALIK"},
    {"KODU": "053056", "ADI": "MNV.MALATYA KAYISI"}, {"KODU": "051279", "ADI": "MNV.MANGO ADET"},
    {"KODU": "8699211220011", "ADI": "MNV.MANTAR PAKET"}, {"KODU": "69", "ADI": "MNV.MARUL"},
    {"KODU": "70", "ADI": "MNV.MAYDANOZ"}, {"KODU": "014146", "ADI": "MNV.MUZ"},
    {"KODU": "54", "ADI": "MNV.NANE"}, {"KODU": "01302", "ADI": "MNV.NAR KIRMIZI"},
    {"KODU": "052827", "ADI": "MNV.NEKTARI"}, {"KODU": "056065", "ADI": "MNV.PEMBE DOMATES"},
    {"KODU": "B.2901021", "ADI": "MNV.RED GLOBE UZUM"}, {"KODU": "05983", "ADI": "MNV.REYHAN"},
    {"KODU": "00012256", "ADI": "MNV.ROKA"}, {"KODU": "09915", "ADI": "MNV.SALKIM DOMATES"},
    {"KODU": "09399", "ADI": "MNV.SANTA MARIA ARMUT"}, {"KODU": "07604", "ADI": "MNV.SARI HAVUC"},
    {"KODU": "054069", "ADI": "MNV.SARI KAVUN"}, {"KODU": "053793", "ADI": "MNV.SEKERPARE"},
    {"KODU": "98", "ADI": "MNV.SEMIZ OTU"}, {"KODU": "B.2901018", "ADI": "MNV.SIVRI BIBER"},
    {"KODU": "MNV.017485", "ADI": "MNV.SIYAH UZUM"}, {"KODU": "017185", "ADI": "MNV.SUS BIBERI"},
    {"KODU": "2901098", "ADI": "MNV.TATLI  PATATES"}, {"KODU": "2901117", "ADI": "MNV.TAZE PATATES"},
    {"KODU": "015575", "ADI": "MNV.TOPAK PATLICAN"}, {"KODU": "08384", "ADI": "MNV.YABAN MERSINI PAKET"},
    {"KODU": "05694", "ADI": "MNV.YAYLA DOMATES"}, {"KODU": "2909808", "ADI": "MNV.YAYLA ELMASI"},
    {"KODU": "05700", "ADI": "MNV.YENI DUNYA"}, {"KODU": "09913", "ADI": "MNV.YERLI DOMATES"},
    {"KODU": "052128", "ADI": "MNV.YERLI SALATALIK"}, {"KODU": "016870", "ADI": "MNV.YESIL KAPYA BIBER"},
    {"KODU": "053742", "ADI": "MNV.YESIL SILI BIBER"}, {"KODU": "13", "ADI": "MNV.YESIL SOGAN"},
    {"KODU": "051277", "ADI": "MNV.ZENCEFIL"}
]

# Session State
if "site_giris_yapildi" not in st.session_state:
    st.session_state.site_giris_yapildi = False

if "aktif_rol" not in st.session_state:
    st.session_state.aktif_rol = "🏬 Şube Girişi"

if "giris_yapilan_sube" not in st.session_state:
    st.session_state.giris_yapilan_sube = None

if "hal_authed" not in st.session_state:
    st.session_state.hal_authed = False

if "admin_authed" not in st.session_state:
    st.session_state.admin_authed = False


# TEK ÜRÜN EXCEL ÇIKTISI
def generate_hal_excel(urun_adi, urun_kodu, hal_toplam, dagitim_dict, kalan, tarih_str):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hal_Dagitim_Listesi"

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    thin = Side(border_style="thin", color="000000")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_title = Font(name="Calibri", size=14, bold=True)
    font_normal = Font(name="Calibri", size=10)

    ws.cell(row=1, column=1, value="YALÇIN MARKETLER ZİNCİRİ - HAL MALI ŞUBE DAĞITIM LİSTESİ").font = font_title
    ws.cell(row=2, column=1, value=f"Tarih: {tarih_str} | Ürün: {urun_adi} (Kod: {urun_kodu}) | Halden Alınan: {hal_toplam:.0f} Kasa").font = font_bold

    headers = ["Şube Adı", "Dağıtılan Miktar (Kasa)"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num, value=header)
        cell.font = font_bold
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")

    row_idx = 5
    toplam_dagitilan = 0
    for sube, miktar in dagitim_dict.items():
        if miktar > 0:
            c1 = ws.cell(row=row_idx, column=1, value=sube)
            c2 = ws.cell(row=row_idx, column=2, value=f"{miktar:.0f} Kasa")
            
            c1.font = font_normal
            c2.font = font_normal
            c2.alignment = Alignment(horizontal="center")
            
            c1.border = border
            c2.border = border
            toplam_dagitilan += miktar
            row_idx += 1

    ws.cell(row=row_idx, column=1, value="TOPLAM DAĞITILAN").font = font_bold
    c_tot = ws.cell(row=row_idx, column=2, value=f"{toplam_dagitilan:.0f} Kasa")
    c_tot.font = font_bold
    c_tot.alignment = Alignment(horizontal="center")
    
    ws.cell(row=row_idx, column=1).border = border
    c_tot.border = border

    ws.column_dimensions['A'].width = 25
    ws.column_dimensions['B'].width = 25

    wb.save(output)
    return output.getvalue()


# 🚚 TOPLU HAL DAĞITIM EXCEL ÇIKTISI (YENİ TABLODAN ÇEKER: hal_dagitim)
def generate_toplu_hal_excel(tarih_sorgu_str):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    
    ws1 = wb.active
    ws1.title = "Sevkiyat_Matris_Tablosu"

    ws1.page_setup.orientation = ws1.ORIENTATION_LANDSCAPE
    ws1.page_setup.paperSize = ws1.PAPERSIZE_A4
    ws1.sheet_properties.pageSetUpPr.fitToPage = True
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 0
    
    # Artık hal_dagitim tablosundan sorgu çekiyoruz
    res = supabase.table("hal_dagitim").select("sube, urun_kodu, urun_adi, dağıtılan_miktar").eq("tarih", tarih_sorgu_str).execute()
    
    if not res.data:
        return None

    df_hal = pd.DataFrame(res.data)
    df_hal['dağıtılan_miktar'] = pd.to_numeric(df_hal['dağıtılan_miktar'], errors='coerce').fillna(0)
    df_hal = df_hal[df_hal['dağıtılan_miktar'] > 0]

    if df_hal.empty:
        return None

    thin = Side(border_style="thin", color="D3D3D3")
    border = Border(top=thin, left=thin, right=thin, bottom=thin)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_title = Font(name="Calibri", size=13, bold=True)
    font_normal = Font(name="Calibri", size=9)

    ws1.cell(row=1, column=1, value=f"YALÇIN MARKETLER ZİNCİRİ - SEVKİYAT DAĞITIM MATRİSİ ({tarih_sorgu_str})").font = font_title

    pivot_hal = pd.pivot_table(
        df_hal, 
        values='dağıtılan_miktar', 
        index=['urun_kodu', 'urun_adi'], 
        columns=['sube'], 
        aggfunc='sum', 
        fill_value=0
    )
    pivot_hal['TOPLAM SEVK'] = pivot_hal.sum(axis=1)

    ws1.cell(row=3, column=1, value="Ürün Kodu").font = font_bold
    ws1.cell(row=3, column=2, value="Ürün Adı").font = font_bold
    ws1.cell(row=3, column=1).fill = header_fill
    ws1.cell(row=3, column=2).fill = header_fill
    ws1.cell(row=3, column=1).border = border
    ws1.cell(row=3, column=2).border = border

    col_idx = 3
    sube_cols = [c for c in pivot_hal.columns if c != 'TOPLAM SEVK'] + ['TOPLAM SEVK']
    for sube in sube_cols:
        cell = ws1.cell(row=3, column=col_idx, value=sube)
        cell.font = font_bold
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        col_idx += 1

    row_idx = 4
    for (kodu, adi), r_data in pivot_hal.iterrows():
        ws1.cell(row=row_idx, column=1, value=str(kodu)).font = font_normal
        ws1.cell(row=row_idx, column=2, value=str(adi)).font = font_normal
        ws1.cell(row=row_idx, column=1).border = border
        ws1.cell(row=row_idx, column=2).border = border

        c_idx = 3
        for sube in sube_cols:
            val = r_data[sube]
            val_str = f"{int(val)} Kasa" if val > 0 else ""
            c = ws1.cell(row=row_idx, column=c_idx, value=val_str)
            c.font = font_bold if sube == 'TOPLAM SEVK' else font_normal
            c.alignment = Alignment(horizontal="center")
            c.border = border
            c_idx += 1
        row_idx += 1

    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 25
    for c in range(3, col_idx):
        ws1.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 13

    # DETAY LİSTE SAYFASI
    ws2 = wb.create_sheet(title="Urun_Bazli_Liste")
    ws2.page_setup.orientation = ws2.ORIENTATION_LANDSCAPE
    ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
    ws2.sheet_properties.pageSetUpPr.fitToPage = True
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 0

    ws2.cell(row=1, column=1, value=f"YALÇIN MARKETLER ZİNCİRİ - SEVKİYAT DETAY LİSTESİ ({tarih_sorgu_str})").font = font_title

    headers2 = ["Ürün Kodu", "Ürün Adı", "Şube Adı", "Verilecek Miktar (Kasa)"]
    for c_i, h in enumerate(headers2, 1):
        cell = ws2.cell(row=3, column=c_i, value=h)
        cell.font = font_bold
        cell.fill = header_fill
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    r2_idx = 4
    for _, row_item in df_hal.sort_values(by=['urun_adi', 'sube']).iterrows():
        ws2.cell(row=r2_idx, column=1, value=str(row_item['urun_kodu'])).font = font_normal
        ws2.cell(row=r2_idx, column=2, value=str(row_item['urun_adi'])).font = font_normal
        ws2.cell(row=r2_idx, column=3, value=str(row_item['sube'])).font = font_normal
        c4 = ws2.cell(row=r2_idx, column=4, value=f"{int(row_item['dağıtılan_miktar'])} Kasa")
        c4.font = font_bold
        c4.alignment = Alignment(horizontal="center")

        for ci in range(1, 5):
            ws2.cell(row=r2_idx, column=ci).border = border
        r2_idx += 1

    ws2.column_dimensions['A'].width = 12
    ws2.column_dimensions['B'].width = 25
    ws2.column_dimensions['C'].width = 18
    ws2.column_dimensions['D'].width = 22

    wb.save(output)
    return output.getvalue()


# KARŞILAMA EKRANI
if not st.session_state.site_giris_yapildi:
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        try:
            with open("logo.png", "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode()
            
            st.markdown(f'''
                <div class="logo-card-container">
                    <img src="data:image/png;base64,{encoded_string}" class="animated-logo">
                </div>
            ''', unsafe_allow_html=True)
        except Exception:
            st.warning("Logo yüklenemedi.")

        st.markdown('<div class="welcome-title">YALÇIN MARKETLER ZİNCİRİ</div>', unsafe_allow_html=True)
        st.markdown('<div class="welcome-sub">Manav Sipariş ve Stok Yönetim Portalı</div>', unsafe_allow_html=True)
        
        if st.button("🚀 SİSTEME GİRİŞ YAP", type="primary", use_container_width=True):
            st.session_state.site_giris_yapildi = True
            st.rerun()

else:
    st.markdown("### 📌 Sayfa Geçişi")
    m_col1, m_col2, m_col3, m_col4 = st.columns([1, 1, 1, 0.8])
    
    with m_col1:
        if st.button("🏬 Şube Girişi", type="primary" if st.session_state.aktif_rol == "🏬 Şube Girişi" else "secondary", use_container_width=True):
            st.session_state.aktif_rol = "🏬 Şube Girişi"
            st.rerun()

    with m_col2:
        if st.button("🚛 Hal Dağıtım Paneli", type="primary" if st.session_state.aktif_rol == "🚛 Hal Dağıtım Paneli" else "secondary", use_container_width=True):
            st.session_state.aktif_rol = "🚛 Hal Dağıtım Paneli"
            st.rerun()
            
    with m_col3:
        if st.button("👑 Merkez Panel", type="primary" if st.session_state.aktif_rol == "👑 Merkez Panel" else "secondary", use_container_width=True):
            st.session_state.aktif_rol = "👑 Merkez Panel"
            st.rerun()

    with m_col4:
        if st.button("🚪 Çıkış", use_container_width=True):
            st.session_state.site_giris_yapildi = False
            st.session_state.giris_yapilan_sube = None
            st.rerun()

    st.divider()

    rol = st.session_state.aktif_rol

    # 1. ŞUBE SİPARİŞ GİRİŞİ
    if rol == "🏬 Şube Girişi":
        st.markdown("<h2 style='text-align: center;'>🥭 Şube Manav Sipariş Portalı</h2>", unsafe_allow_html=True)

        bugun_str = datetime.now().strftime('%Y-%m-%d')
        st.caption(f"Tarih: {datetime.now().strftime('%d.%m.%Y')}")

        subeler = ["-- Seçiniz --"] + SUBE_LISTESI
        secilen_sube = st.selectbox("📍 **Lütfen Şubenizi Seçin:**", subeler)

        if secilen_sube != "-- Seçiniz --":
            if st.session_state.giris_yapilan_sube != secilen_sube:
                st.info(f"🔒 **{secilen_sube}** şubesinin sipariş ekranına erişmek için lütfen şube şifrenizi giriniz.")
                s_col1, s_col2 = st.columns([2, 1])
                with s_col1:
                    girilen_pin = st.text_input(f"🔑 {secilen_sube} Şube Şifresi:", type="password", key=f"pin_input_{secilen_sube}")
                with s_col2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("Giriş Yap", type="primary", use_container_width=True):
                        if girilen_pin == SUBE_SIFRELERI.get(secilen_sube):
                            st.session_state.giris_yapilan_sube = secilen_sube
                            st.success("✅ Şifre Doğrulandı!")
                            st.rerun()
                        else:
                            st.error("❌ Hatalı Şube Şifresi!")
            else:
                st.success(f"🔓 **{secilen_sube}** Şubesi Girişi Aktif")
                if st.button("🔒 Şube Oturumunu Kapat", type="secondary"):
                    st.session_state.giris_yapilan_sube = None
                    st.rerun()

                st.divider()

                # Hal Dağıtım Verisini Artık 'hal_dagitim' Tablosundan Çekiyor
                with st.expander(f"🚛 **{secilen_sube} - Halden Şubemize Ayrılan/Gelen Mal Miktarları (Bugün)**", expanded=True):
                    hal_res = supabase.table("hal_dagitim").select("urun_kodu, urun_adi, dağıtılan_miktar").eq("sube", secilen_sube).eq("tarih", bugun_str).execute()
                    
                    if hal_res.data:
                        hal_df = pd.DataFrame(hal_res.data)
                        hal_df['dağıtılan_miktar'] = pd.to_numeric(hal_df['dağıtılan_miktar'], errors='coerce').fillna(0)
                        hal_df = hal_df[hal_df['dağıtılan_miktar'] > 0]
                        
                        if not hal_df.empty:
                            hal_df = hal_df.rename(columns={
                                'urun_kodu': 'Ürün Kodu',
                                'urun_adi': 'Ürün Adı',
                                'dağıtılan_miktar': 'Gelen / Ayrılan Miktar (Kasa)'
                            })
                            st.dataframe(hal_df, use_container_width=True, hide_index=True)
                        else:
                            st.info("ℹ️ Bugün için şubenize henüz halden yüklenen veya dağıtılan mal girişi yapılmadı.")
                    else:
                        st.info("ℹ️ Bugün için şubenize henüz halden yüklenen veya dağıtılan mal girişi yapılmadı.")

                st.divider()

                # Şube Kendi Siparişini 'siparisler' Tablosundan Sorgular
                res = supabase.table("siparisler").select("urun_kodu, mevcut_stok, siparis_miktari").eq("sube", secilen_sube).eq("tarih", bugun_str).execute()
                
                kayitli_dict = {}
                for r in res.data:
                    try:
                        sip_val = float(r['siparis_miktari']) if r['siparis_miktari'] is not None else 0.0
                    except (ValueError, TypeError):
                        sip_val = 0.0

                    kayitli_dict[r['urun_kodu']] = {
                        'stok': str(r['mevcut_stok']) if r['mevcut_stok'] is not None else "0",
                        'siparis': sip_val
                    }

                df = pd.DataFrame(URUNLER)
                arama = st.text_input("🔍 **Ürün Ara (Adı veya Kodu):**", "")
                filtre_df = df[df['ADI'].str.contains(arama, case=False) | df['KODU'].str.contains(arama, case=False)] if arama else df

                kaydedilecek_veriler = []
                st.subheader("📦 Stok ve Sipariş Girişi (Kasa)")

                for index, row in filtre_df.iterrows():
                    kod = row['KODU']
                    varsayilan_stok_str = kayitli_dict.get(kod, {}).get('stok', "0")
                    varsayilan_siparis = kayitli_dict.get(kod, {}).get('siparis', 0.0)

                    with st.expander(f"**{row['ADI']}** *(Kod: {kod})*"):
                        col1, col2 = st.columns([1.5, 1])
                        
                        with col1:
                            stok_dolu = st.checkbox("🟢 Reyon Dolu (Depo Boş)", value=(varsayilan_stok_str == "Reyon Dolu"), key=f"dolu_{kod}")
                            if not stok_dolu:
                                try:
                                    def_val = float(varsayilan_stok_str)
                                except ValueError:
                                    def_val = 0.0
                                stok_val = st.number_input("Mevcut Stok (Kasa)", min_value=0.0, step=1.0, value=def_val, key=f"stok_{kod}")
                                stok_kayit = str(int(stok_val))
                            else:
                                stok_kayit = "Reyon Dolu"
                                st.caption("📌 *Stok 'Reyon Dolu' olarak kaydedilecek.*")

                        with col2:
                            siparis = st.number_input("Sipariş (Kasa)", min_value=0.0, step=1.0, value=float(varsayilan_siparis), key=f"sip_{kod}")
                            
                        if (stok_kayit != "0" and stok_kayit != "0.0") or siparis > 0:
                            kaydedilecek_veriler.append({
                                "sube": secilen_sube,
                                "tarih": bugun_str,
                                "urun_kodu": kod,
                                "urun_adi": row['ADI'],
                                "mevcut_stok": stok_kayit,
                                "siparis_miktari": float(siparis)
                            })

                st.divider()

                btn_col1, btn_col2 = st.columns([2, 1])

                with btn_col1:
                    if st.button("💾 Siparişleri Güncelle / Kaydet", type="primary", use_container_width=True):
                        supabase.table("siparisler").delete().eq("sube", secilen_sube).eq("tarih", bugun_str).execute()
                        if len(kaydedilecek_veriler) > 0:
                            supabase.table("siparisler").insert(kaydedilecek_veriler).execute()
                            st.success(f"✅ **{secilen_sube}** şubesinin siparişi buluta başarıyla kaydedildi!")
                        else:
                            st.warning("⚠️ Tüm değerler 0 yapıldığı için bugünkü siparişiniz temizlendi.")
                        st.rerun()

                with btn_col2:
                    if st.button("🗑️ Bugünkü Siparişi İptal Et", type="secondary", use_container_width=True):
                        supabase.table("siparisler").delete().eq("sube", secilen_sube).eq("tarih", bugun_str).execute()
                        st.error("🗑️ Bugünkü siparişiniz tamamen silindi!")
                        st.rerun()

    # 2. HAL DAĞITIM PANELİ
    elif rol == "🚛 Hal Dağıtım Paneli":
        st.markdown("<h2 style='text-align: center;'>🚛 Hal Satınalma ve Dağıtım Paneli</h2>", unsafe_allow_html=True)

        if not st.session_state.hal_authed:
            hal_pin = st.text_input("🔑 Lütfen Satınalma/Hal Yetkili Şifresini Giriniz:", type="password")
            if st.button("Giriş Yap", type="primary"):
                if hal_pin == HAL_SIFRESI or hal_pin == YONETICI_SIFRESI:
                    st.session_state.hal_authed = True
                    st.rerun()
                else:
                    st.error("❌ Hatalı Satınalma Şifresi!")
        else:
            st.success("🔓 Satınalma Yetkili Girişi Aktif")
            if st.button("🔒 Oturumu Kapat"):
                st.session_state.hal_authed = False
                st.rerun()

            st.divider()

            st.markdown("#### 📅 Sevkiyat ve Dağıtım Tarihi Seçimi")
            
            t_col1, t_col2 = st.columns([2, 5])
            with t_col1:
                secilen_hal_tarihi = st.date_input("İşlem Yapmak İstediğiniz Tarih:", value=date.today())
                hal_tarih_str = secilen_hal_tarihi.strftime('%Y-%m-%d')
            
            with t_col2:
                st.write("") 
                if secilen_hal_tarihi == date.today():
                    st.info("🟢 **Bugünün** verileri ve dağıtım listesi görüntüleniyor.")
                else:
                    st.warning(f"🟡 **{secilen_hal_tarihi.strftime('%d.%m.%Y')}** tarihine ait dağıtım verileri görüntüleniyor.")

            st.divider()

            is_today = (secilen_hal_tarihi == date.today())
            tarih_label = "BUGÜNÜN" if is_today else f"{secilen_hal_tarihi.strftime('%d.%m.%Y')} TARİHLİ"

            st.info(f"📦 **Sevkiyatçılar İçin Toplu Dağıtım Çıktısı:** {tarih_label} halden girilen tüm ürünlerin ve şube dağıtımlarının olduğu tek Excel dosyasını (A4 Yatay Baskıya Hazır) indirebilirsiniz.")
            
            toplu_excel_bytes = generate_toplu_hal_excel(hal_tarih_str)
            if toplu_excel_bytes:
                st.download_button(
                    label=f"🚚 {tarih_label} TÜM SEVKİYAT DAĞITIM LİSTESİNİ İNDİR (YAZDIRMAYA HAZIR EXCEL)",
                    data=toplu_excel_bytes,
                    file_name=f"Toplu_Hal_Sevkiyat_Listesi_{hal_tarih_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
            else:
                st.caption(f"ℹ️ *{secilen_hal_tarihi.strftime('%d.%m.%Y')} tarihi için henüz halden girilmiş herhangi bir dağıtım bulunmuyor.*")

            st.divider()

            urun_listesi_adlar = [f"{u['ADI']} ({u['KODU']})" for u in URUNLER]
            secilen_urun_combo = st.selectbox("🛒 **Halden Alınan Ürünü Seçin:**", urun_listesi_adlar)
            
            secilen_urun_kod = secilen_urun_combo.split("(")[-1].replace(")", "").strip()
            secilen_urun_ad = secilen_urun_combo.split("(")[0].strip()

            hal_toplam_kasa = st.number_input(f"📦 **Halden Alınan Toplam Miktar ({secilen_urun_ad}):**", min_value=0.0, step=1.0, value=0.0)

            st.subheader("🏬 Şubelere Dağıtım Tablosu")
            st.caption("Lütfen halden yükleme yapılan şubelere verilen kasa miktarlarını girin:")

            dagitim_dict = {}
            toplam_dagitilan = 0.0

            d_col1, d_col2 = st.columns(2)
            
            for i, sube_adi in enumerate(SUBE_LISTESI):
                target_col = d_col1 if i % 2 == 0 else d_col2
                with target_col:
                    val = st.number_input(f"📍 {sube_adi}:", min_value=0.0, step=1.0, value=0.0, key=f"hal_dag_{sube_adi}_{secilen_urun_kod}")
                    dagitim_dict[sube_adi] = val
                    toplam_dagitilan += val

            kalan_kasa = hal_toplam_kasa - toplam_dagitilan

            st.divider()
            
            m1, m2, m3 = st.columns(3)
            m1.metric("Halden Alınan", f"{hal_toplam_kasa:.0f} Kasa")
            m2.metric("Şubelere Dağıtılan", f"{toplam_dagitilan:.0f} Kasa")
            
            if kalan_kasa < 0:
                m3.metric("⚠️ Fazla Dağıtılan", f"{abs(kalan_kasa):.0f} Kasa", delta_color="inverse")
                st.error("⚠️ Halden aldığınız miktardan daha fazla dağıtım yaptınız!")
            else:
                m3.metric("Kalan (Depo/Yedek)", f"{kalan_kasa:.0f} Kasa")

            st.divider()

            h_btn1, h_btn2 = st.columns(2)

            with h_btn1:
                if st.button("💾 Hal Dağıtımını Kaydet ve Şubelere Bildir", type="primary", use_container_width=True):
                    if hal_toplam_kasa == 0:
                        st.warning("⚠️ Halden alınan miktar 0 olamaz.")
                    elif kalan_kasa < 0:
                        st.error("❌ Hata: Alınan miktardan fazlası dağıtılamaz!")
                    else:
                        kayit_listesi = []
                        for sube, miktar in dagitim_dict.items():
                            if miktar > 0:
                                kayit_listesi.append({
                                    "sube": sube,
                                    "tarih": hal_tarih_str,
                                    "urun_kodu": secilen_urun_kod,
                                    "urun_adi": secilen_urun_ad,
                                    "dağıtılan_miktar": float(miktar)
                                })
                        
                        if len(kayit_listesi) > 0:
                            # Kayıtlar Doğrudan 'hal_dagitim' Tablosuna Yapılıyor
                            supabase.table("hal_dagitim").insert(kayit_listesi).execute()
                            st.success(f"✅ **{secilen_urun_ad}** dağıtımı ({hal_tarih_str} tarihi için) başarıyla kaydedildi!")
                            st.rerun()
                        else:
                            st.warning("⚠️ Şubelere herhangi bir miktar girilmedi.")

            with h_btn2:
                if toplam_dagitilan > 0 and kalan_kasa >= 0:
                    hal_excel_bytes = generate_hal_excel(secilen_urun_ad, secilen_urun_kod, hal_toplam_kasa, dagitim_dict, kalan_kasa, hal_tarih_str)
                    st.download_button(
                        label="📄 Sadece Bu Ürünün Excel Listesini İndir",
                        data=hal_excel_bytes,
                        file_name=f"Hal_Dagitim_{secilen_urun_kod}_{hal_tarih_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

    # 3. MERKEZ YÖNETİM PANELİ
    elif rol == "👑 Merkez Panel":
        st.markdown("<h2 style='text-align: center;'>🔒 Merkez Yönetim Paneli</h2>", unsafe_allow_html=True)

        if not st.session_state.admin_authed:
            sifre_giris = st.text_input("🔑 Lütfen Yönetim Şifresini Giriniz:", type="password")
            if st.button("Giriş Yap", type="primary"):
                if sifre_giris == YONETICI_SIFRESI:
                    st.session_state.admin_authed = True
                    st.rerun()
                else:
                    st.error("❌ Hatalı şifre!")
        else:
            st.success("🔓 Yetkili Girişi Başarılı")
            if st.button("🚪 Oturumu Kapat"):
                st.session_state.admin_authed = False
                st.rerun()

            f_col1, f_col2, f_col3 = st.columns([1.2, 1, 2])
            
            with f_col1:
                secilen_tarih = st.date_input("📅 Tarih Seçin", value=date.today())
                tarih_str = secilen_tarih.strftime('%Y-%m-%d')

            with f_col2:
                filtre_sube = st.selectbox("🏬 Şube Filtresi", ["Tümü"] + SUBE_LISTESI)

            with f_col3:
                arama_admin = st.text_input("🔍 Ürün Arama (Ad / Kod)", "")

            st.divider()

            # MERKEZ PANELİNDE İKİ AYRI SEKME SIFIR KARIŞIKLIK SAĞLAR:
            tab_sip, tab_hal = st.tabs(["🛒 Şube Sipariş ve Stok Talepleri", "🚛 Hal Sevkiyat ve Dağıtım Verileri"])

            # SEKME 1: ŞUBELERİN İSTEDİĞİ GERÇEK SİPARİŞLER (siparisler tablosu)
            with tab_sip:
                query = supabase.table("siparisler").select("*").eq("tarih", tarih_str)
                if filtre_sube != "Tümü":
                    query = query.eq("sube", filtre_sube)

                res = query.execute()

                if res.data:
                    df_res = pd.DataFrame(res.data)
                    df_res['siparis_miktari'] = pd.to_numeric(df_res['siparis_miktari'], errors='coerce').fillna(0)
                    
                    if arama_admin:
                        df_res = df_res[
                            df_res['urun_adi'].str.contains(arama_admin, case=False, na=False) | 
                            df_res['urun_kodu'].str.contains(arama_admin, case=False, na=False)
                        ]

                    if not df_res.empty:
                        m1, m2, m3 = st.columns(3)
                        m1.metric("📦 Toplam Sipariş Veren Şube", f"{df_res['sube'].nunique()} Şube")
                        m2.metric("🍉 Toplam Sipariş Kalemi", f"{len(df_res)} Kalem")
                        m3.metric("📊 Toplam Şube Sipariş Miktarı", f"{int(df_res['siparis_miktari'].sum())} Kasa")

                        st.subheader("📋 Şube Sipariş Detay Tablosu")
                        
                        df_display = df_res.rename(columns={
                            "sube": "Şube",
                            "urun_kodu": "Ürün Kodu",
                            "urun_adi": "Ürün Adı",
                            "mevcut_stok": "Mevcut Stok",
                            "siparis_miktari": "Sipariş Miktarı (Kasa)"
                        })[['Şube', 'Ürün Kodu', 'Ürün Adı', 'Mevcut Stok', 'Sipariş Miktarı (Kasa)']]
                        
                        st.dataframe(df_display, use_container_width=True, hide_index=True)

                        st.divider()

                        st.subheader("📊 Şube - Sipariş Dağıtım Matrisi")
                        pivot_df = pd.pivot_table(
                            df_res,
                            values='siparis_miktari',
                            index=['urun_kodu', 'urun_adi'],
                            columns=['sube'],
                            aggfunc='sum',
                            fill_value=0
                        )
                        pivot_df['TOPLAM SİPARİŞ'] = pivot_df.sum(axis=1)
                        st.dataframe(pivot_df, use_container_width=True)
                    else:
                        st.info("ℹ️ Aranan kriterlere uygun şube sipariş verisi bulunamadı.")
                else:
                    st.info(f"ℹ️ {secilen_tarih.strftime('%d.%m.%Y')} tarihi için henüz kaydedilmiş bir şube siparişi bulunmuyor.")

            # SEKME 2: HAL YETKİLİSİNİN ŞUBELERE DAĞITTIĞI MALLAR (hal_dagitim tablosu)
            with tab_hal:
                query_h = supabase.table("hal_dagitim").select("*").eq("tarih", tarih_str)
                if filtre_sube != "Tümü":
                    query_h = query_h.eq("sube", filtre_sube)

                res_h = query_h.execute()

                if res_h.data:
                    df_h = pd.DataFrame(res_h.data)
                    df_h['dağıtılan_miktar'] = pd.to_numeric(df_h['dağıtılan_miktar'], errors='coerce').fillna(0)

                    if arama_admin:
                        df_h = df_h[
                            df_h['urun_adi'].str.contains(arama_admin, case=False, na=False) | 
                            df_h['urun_kodu'].str.contains(arama_admin, case=False, na=False)
                        ]

                    if not df_h.empty:
                        st.subheader("🚚 Halden Şubelere Sevk Edilen Mal Dağıtım Matrisi")
                        pivot_h = pd.pivot_table(
                            df_h,
                            values='dağıtılan_miktar',
                            index=['urun_kodu', 'urun_adi'],
                            columns=['sube'],
                            aggfunc='sum',
                            fill_value=0
                        )
                        pivot_h['TOPLAM SEVK'] = pivot_h.sum(axis=1)
                        st.dataframe(pivot_h, use_container_width=True)

                        excel_bytes = generate_toplu_hal_excel(tarih_str)
                        if excel_bytes:
                            st.download_button(
                                label="📥 Seçilen Tarihin Hal Sevkiyat Raporunu Excel Olarak İndir (A4 Baskıya Hazır)",
                                data=excel_bytes,
                                file_name=f"Hal_Sevkiyat_Raporu_{tarih_str}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                                use_container_width=True
                            )
                    else:
                        st.info("ℹ️ Aranan kriterlere uygun hal sevkiyat verisi bulunamadı.")
                else:
                    st.info(f"ℹ️ {secilen_tarih.strftime('%d.%m.%Y')} tarihi için halden yapılmış bir dağıtım/sevkiyat kaydı bulunmuyor.")
