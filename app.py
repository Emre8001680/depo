import pandas as pd
import streamlit as st
from datetime import datetime, date
import sqlite3
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Sayfa Yapılandırması
st.set_page_config(page_title="Yalçın Marketler Zinciri - Manav Portalı", page_icon="🥭", layout="wide")

# CSS DÜZENLEMELERİ (Gereksiz Üst Bar, Fork, GitHub ve Yan Menü Kalıntılarını Gizler)
st.markdown("""
    <style>
        /* Sağ üst menü, footer ve Streamlit üst bar/Fork alanlarını tamamen gizle */
        #MainMenu {visibility: hidden !important;}
        footer {visibility: hidden !important;}
        header {visibility: hidden !important;}
        [data-testid="stHeader"] {display: none !important;}
        [data-testid="stSidebar"] {display: none !important;}
        
        /* Sayfa üst boşluğunu sıfırla */
        .block-container {
            padding-top: 1.5rem !important;
            padding-bottom: 2rem !important;
        }

        /* LOGO ANİMASYONU */
        @keyframes fadeInZoom {
            0% { opacity: 0; transform: scale(0.85); }
            100% { opacity: 1; transform: scale(1); }
        }

        .animated-logo {
            animation: fadeInZoom 1.2s ease-out forwards;
            display: block;
            margin-left: auto;
            margin-right: auto;
            max-width: 450px;
            width: 100%;
            height: auto;
            padding: 20px 0;
        }

        .welcome-title {
            text-align: center;
            font-size: 28px;
            font-weight: 700;
            color: #1E293B;
            margin-top: 10px;
            margin-bottom: 5px;
        }
        
        .welcome-sub {
            text-align: center;
            font-size: 16px;
            color: #64748B;
            margin-bottom: 30px;
        }
    </style>
""", unsafe_allow_html=True)

# Session State Tanımlamaları
if "site_giris_yapildi" not in st.session_state:
    st.session_state.site_giris_yapildi = False

if "aktif_rol" not in st.session_state:
    st.session_state.aktif_rol = "🏬 Şube Sipariş Girişi"

# -------------------------------------------------------------
# 🌟 ANİMASYONLU İLK GİRİŞ / KARŞILAMA EKRANI (SPLASH SCREEN)
# -------------------------------------------------------------
if not st.session_state.site_giris_yapildi:
    st.markdown("<br><br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        import base64
        try:
            with open("logo.png", "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode()
            st.markdown(f'<img src="data:image/png;base64,{encoded_string}" class="animated-logo">', unsafe_allow_html=True)
        except Exception:
            st.image("logo.png", use_container_width=True)

        st.markdown('<div class="welcome-title">YALÇIN MARKETLER ZİNCİRİ</div>', unsafe_allow_html=True)
        st.markdown('<div class="welcome-sub">Manav Sipariş ve Stok Yönetim Portalı</div>', unsafe_allow_html=True)
        
        if st.button("🚀 SİSTEME GİRİŞ YAP", type="primary", use_container_width=True):
            st.session_state.site_giris_yapildi = True
            st.rerun()

# -------------------------------------------------------------
# 🏢 ANA UYGULAMA (GİRİŞ YAPILDIKTAN SONRA AÇILAN EKRAN)
# -------------------------------------------------------------
else:
    # 🔒 Yönetici Giriş Şifresi
    YONETICI_SIFRESI = "1234"

    # Veritabanı Bağlantısı
    conn = sqlite3.connect('manav_siparisleri.db', check_same_thread=False)
    c = conn.cursor()

    c.execute('''
        CREATE TABLE IF NOT EXISTS siparisler (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sube TEXT,
            tarih TEXT,
            urun_kodu TEXT,
            urun_adi TEXT,
            mevcut_stok REAL,
            siparis_miktari REAL
        )
    ''')
    conn.commit()

    # --- MOBİL VE MASAÜSTÜ İÇİN ÜST HIZLI GEÇİŞ MENÜSÜ ---
    st.markdown("### 📌 Sayfa Geçişi")
    m_col1, m_col2, m_col3 = st.columns([1, 1, 1])
    
    with m_col1:
        if st.button("🏬 Şube Girişi", type="primary" if st.session_state.aktif_rol == "🏬 Şube Sipariş Girişi" else "secondary", use_container_width=True):
            st.session_state.aktif_rol = "🏬 Şube Sipariş Girişi"
            st.rerun()
            
    with m_col2:
        if st.button("👑 Merkez Panel", type="primary" if st.session_state.aktif_rol == "👑 Merkez Yönetim Paneli" else "secondary", use_container_width=True):
            st.session_state.aktif_rol = "👑 Merkez Yönetim Paneli"
            st.rerun()

    with m_col3:
        if st.button("🚪 Çıkış", use_container_width=True):
            st.session_state.site_giris_yapildi = False
            st.rerun()

    st.divider()

    rol = st.session_state.aktif_rol

    # -------------------------------------------------------------
    # 1. ŞUBE SİPARİŞ GİRİŞ VE GÜNCELLEME EKRANI
    # -------------------------------------------------------------
    if rol == "🏬 Şube Sipariş Girişi":
        st.markdown("<h2 style='text-align: center;'>🥭 Şube Manav Sipariş Portalı</h2>", unsafe_allow_html=True)

        bugun_str = datetime.now().strftime('%Y-%m-%d')
        st.caption(f"Tarih: {datetime.now().strftime('%d.%m.%Y')}")

        subeler = [
            "-- Seçiniz --", "Raufbey", "Metin Tamer", "Hacı Osmanlı", "Salı Yolu", 
            "Kadiri Yolu", "Nahır Yolu", "Eyup Sultan", "Bulvar", "Düziçi Çarşı", "Aşiyan", "Zeytinlik"
        ]
        secilen_sube = st.selectbox("📍 **Lütfen Şubenizi Seçin:**", subeler)

        if secilen_sube != "-- Seçiniz --":
            mevcut_kayitlar = pd.read_sql_query(
                "SELECT urun_kodu, mevcut_stok, siparis_miktari FROM siparisler WHERE sube = ? AND tarih = ?",
                conn,
                params=(secilen_sube, bugun_str)
            )
            
            kayitli_dict = {}
            for _, r in mevcut_kayitlar.iterrows():
                kayitli_dict[r['urun_kodu']] = {
                    'stok': float(r['mevcut_stok']),
                    'siparis': float(r['siparis_miktari'])
                }

            if len(kayitli_dict) > 0:
                st.info(f"ℹ️ **{secilen_sube}** şubesinin bugün girilmiş siparişleri yüklendi. Değişiklik yapıp tekrar kaydedebilirsiniz.")

            urunler = [
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

            df = pd.DataFrame(urunler)
            st.divider()

            arama = st.text_input("🔍 **Ürün Ara (Adı veya Kodu):**", "")
            filtre_df = df[df['ADI'].str.contains(arama, case=False) | df['KODU'].str.contains(arama, case=False)] if arama else df

            kaydedilecek_veriler = []
            st.subheader("🛒 Stok ve Sipariş Girişi")

            for index, row in filtre_df.iterrows():
                kod = row['KODU']
                varsayilan_stok = kayitli_dict.get(kod, {}).get('stok', 0.0)
                varsayilan_siparis = kayitli_dict.get(kod, {}).get('siparis', 0.0)

                with st.expander(f"**{row['ADI']}** *(Kod: {kod})*"):
                    col1, col2 = st.columns(2)
                    with col1:
                        stok = st.number_input("Mevcut Stok", min_value=0.0, step=0.5, value=varsayilan_stok, key=f"stok_{kod}")
                    with col2:
                        siparis = st.number_input("Sipariş Miktarı", min_value=0.0, step=0.5, value=varsayilan_siparis, key=f"sip_{kod}")
                        
                    if stok > 0 or siparis > 0:
                        kaydedilecek_veriler.append((
                            secilen_sube, 
                            bugun_str, 
                            kod, 
                            row['ADI'], 
                            stok, 
                            siparis
                        ))

            st.divider()

            btn_col1, btn_col2 = st.columns([2, 1])

            with btn_col1:
                if st.button("💾 Siparişleri Güncelle / Kaydet", type="primary", use_container_width=True):
                    c.execute("DELETE FROM siparisler WHERE sube = ? AND tarih = ?", (secilen_sube, bugun_str))
                    
                    if len(kaydedilecek_veriler) > 0:
                        c.executemany('''
                            INSERT INTO siparisler (sube, tarih, urun_kodu, urun_adi, mevcut_stok, siparis_miktari)
                            VALUES (?, ?, ?, ?, ?, ?)
                        ''', kaydedilecek_veriler)
                        conn.commit()
                        st.success(f"✅ **{secilen_sube}** şubesinin siparişi başarıyla güncellendi!")
                    else:
                        conn.commit()
                        st.warning("⚠️ Tüm değerler 0 yapıldığı için bugünkü siparişiniz temizlendi.")
                    st.rerun()

            with btn_col2:
                if st.button("🗑️ Bugünkü Siparişi İptal Et", type="secondary", use_container_width=True):
                    c.execute("DELETE FROM siparisler WHERE sube = ? AND tarih = ?", (secilen_sube, bugun_str))
                    conn.commit()
                    st.error("🗑️ Bugünkü siparişiniz tamamen silindi!")
                    st.rerun()

    # -------------------------------------------------------------
    # 2. ŞİFRELİ MERKEZ YÖNETİM PANELİ
    # -------------------------------------------------------------
    elif rol == "👑 Merkez Yönetim Paneli":
        st.markdown("<h2 style='text-align: center;'>🔒 Merkez Yönetim Paneli</h2>", unsafe_allow_html=True)

        if "admin_authed" not in st.session_state:
            st.session_state.admin_authed = False

        if not st.session_state.admin_authed:
            sifre_giris = st.text_input("🔑 Lütfen Yönetim Şifresini Giriniz:", type="password")
            if st.button("Giriş Yap", type="primary"):
                if sifre_giris == YONETICI_SIFRESI:
                    st.session_state.admin_authed = True
                    st.rerun()
                else:
                    st.error("❌ Hatalı şifre! Lütfen tekrar deneyin.")
        else:
            st.success("🔓 Yetkili Girişi Başarılı")
            if st.button("🚪 Oturumu Kapat"):
                st.session_state.admin_authed = False
                st.rerun()

            st.markdown("""
                <style>
                    div[data-testid="stDataFrame"] table {
                        font-size: 11px !important;
                    }
                    div[data-testid="stDataFrame"] td, div[data-testid="stDataFrame"] th {
                        padding: 2px 4px !important;
                        white-space: nowrap !important;
                    }
                </style>
            """, unsafe_allow_html=True)

            st.subheader("📊 Sipariş ve Stok Çizelgesi")

            with st.container():
                f_col1, f_col2, f_col3 = st.columns([1.2, 1, 2])
                
                with f_col1:
                    tum_gecmis = st.checkbox("Tüm Geçmiş Verileri Göster", value=False)
                
                with f_col2:
                    if not tum_gecmis:
                        secili_tarih = st.date_input("Tarih Seçin:", value=date.today())
                        secili_tarih_str = secili_tarih.strftime('%Y-%m-%d')
                        tarih_etiket = secili_tarih.strftime('%d.%m.%Y')
                    else:
                        st.info("🗓️ Tüm Geçmiş Seçili")
                        tarih_etiket = "Tüm_Gecmis"

            if tum_gecmis:
                df_siparisler = pd.read_sql_query(
                    "SELECT sube AS 'Şube', tarih AS 'Tarih', urun_kodu AS 'Ürün Kodu', urun_adi AS 'Ürün Adı', mevcut_stok AS 'Mevcut Stok', siparis_miktari AS 'Sipariş Miktarı' FROM siparisler ORDER BY id DESC", 
                    conn
                )
            else:
                df_siparisler = pd.read_sql_query(
                    "SELECT sube AS 'Şube', tarih AS 'Tarih', urun_kodu AS 'Ürün Kodu', urun_adi AS 'Ürün Adı', mevcut_stok AS 'Mevcut Stok', siparis_miktari AS 'Sipariş Miktarı' FROM siparisler WHERE tarih LIKE ? ORDER BY id DESC", 
                    conn, 
                    params=(f"{secili_tarih_str}%",)
                )

            if df_siparisler.empty:
                st.warning(f"ℹ️ Seçilen tarihte ({tarih_etiket}) veritabanında kayıtlı sipariş/stok bulunmamaktadır.")
            else:
                with f_col3:
                    secili_subeler = st.multiselect("Şube Filtresi:", df_siparisler['Şube'].unique(), default=df_siparisler['Şube'].unique())
                
                filtreli_df = df_siparisler[df_siparisler['Şube'].isin(secili_subeler)]

                st.divider()

                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Toplam Kalem", len(filtreli_df))
                col2.metric("Sipariş Veren Şube", filtreli_df['Şube'].nunique())
                col3.metric("Toplam Stok", f"{filtreli_df['Mevcut Stok'].sum():,.1f} Kg")
                col4.metric("Toplam Sipariş", f"{filtreli_df['Sipariş Miktarı'].sum():,.1f} Kg")

                st.divider()

                pivot_genel = pd.pivot_table(
                    filtreli_df, 
                    values=['Mevcut Stok', 'Sipariş Miktarı'], 
                    index=['Ürün Kodu', 'Ürün Adı'], 
                    columns=['Şube'], 
                    aggfunc='sum', 
                    fill_value=0
                )

                pivot_genel = pivot_genel.swaplevel(0, 1, axis=1)
                pivot_genel = pivot_genel.sort_index(axis=1, level=0)
                pivot_genel = pivot_genel.rename(columns={'Mevcut Stok': 'Stok', 'Sipariş Miktarı': 'Sip.'})

                toplam_stok = filtreli_df.groupby(['Ürün Kodu', 'Ürün Adı'])['Mevcut Stok'].sum()
                toplam_siparis = filtreli_df.groupby(['Ürün Kodu', 'Ürün Adı'])['Sipariş Miktarı'].sum()

                pivot_genel[('GENEL TOPLAM', 'Top. Stok')] = toplam_stok
                pivot_genel[('GENEL TOPLAM', 'Top. Sip.')] = toplam_siparis

                mask = (pivot_genel[('GENEL TOPLAM', 'Top. Stok')] > 0) | (pivot_genel[('GENEL TOPLAM', 'Top. Sip.')] > 0)
                pivot_genel = pivot_genel[mask]

                st.dataframe(pivot_genel, use_container_width=True, height=550)

                st.divider()

                def generate_excel(df_pivot, etiket):
                    output = io.BytesIO()
                    wb = openpyxl.Workbook()
                    ws = wb.active
                    ws.title = "Siparis_Cizelgesi"

                    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
                    ws.page_setup.paperSize = ws.PAPERSIZE_A4
                    ws.sheet_properties.pageSetUpPr.fitToPage = True
                    ws.page_setup.fitToWidth = 1
                    ws.page_setup.fitToHeight = 0

                    thin = Side(border_style="thin", color="D3D3D3")
                    border = Border(top=thin, left=thin, right=thin, bottom=thin)
                    header_fill = PatternFill(start_color="F0F2F6", end_color="F0F2F6", fill_type="solid")
                    font_bold = Font(name="Calibri", size=9, bold=True)
                    font_normal = Font(name="Calibri", size=8.5)

                    c_title = ws.cell(row=1, column=1, value=f"YALÇIN MARKETLER ZİNCİRİ MANAV SİPARİŞ ÇİZELGESİ ({etiket})")
                    c_title.font = Font(size=12, bold=True)
                    
                    c_kodu = ws.cell(row=3, column=1, value="Ürün Kodu")
                    c_kodu.font = font_bold
                    c_adi = ws.cell(row=3, column=2, value="Ürün Adı")
                    c_adi.font = font_bold
                    
                    col_idx = 3
                    for col in df_pivot.columns:
                        sube_adi = str(col[0]) if isinstance(col, tuple) else str(col)
                        metrik_adi = str(col[1]) if isinstance(col, tuple) and len(col) > 1 else ""
                        
                        cell_sube = ws.cell(row=3, column=col_idx, value=sube_adi)
                        cell_sube.font = font_bold
                        
                        cell_metrik = ws.cell(row=4, column=col_idx, value=metrik_adi)
                        cell_metrik.font = font_bold
                        
                        col_idx += 1

                    for c in range(3, col_idx, 2):
                        if c + 1 < col_idx:
                            ws.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+1)

                    row_idx = 5
                    for (kodu, adi), row_data in df_pivot.iterrows():
                        c_k = ws.cell(row=row_idx, column=1, value=str(kodu))
                        c_k.font = font_normal
                        
                        c_a = ws.cell(row=row_idx, column=2, value=str(adi))
                        c_a.font = font_normal
                        
                        c_idx = 3
                        for val in row_data:
                            val_num = float(val) if val != 0 else ""
                            cell_val = ws.cell(row=row_idx, column=c_idx, value=val_num)
                            cell_val.font = font_normal
                            cell_val.alignment = Alignment(horizontal="center")
                            c_idx += 1
                        row_idx += 1

                    for r in ws.iter_rows(min_row=3, max_row=row_idx-1, min_col=1, max_col=col_idx-1):
                        for cell in r:
                            cell.border = border
                            if cell.row in (3, 4):
                                cell.fill = header_fill
                                cell.alignment = Alignment(horizontal="center", vertical="center")

                    ws.column_dimensions['A'].width = 10
                    ws.column_dimensions['B'].width = 24
                    for c in range(3, col_idx):
                        col_letter = openpyxl.utils.get_column_letter(c)
                        ws.column_dimensions[col_letter].width = 6.5

                    wb.save(output)
                    return output.getvalue()

                excel_bytes = generate_excel(pivot_genel, tarih_etiket)

                st.download_button(
                    label="🖨️ A4 Yatay Çıktı İçin Excel Dosyasını İndir",
                    data=excel_bytes,
                    file_name=f"Yalcin_Market_Manav_Siparis_{tarih_etiket}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
