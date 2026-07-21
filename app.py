import pandas as pd
import streamlit as st
from datetime import datetime, date
import sqlite3
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Sayfa Yapılandırması
st.set_page_config(page_title="Manav Yönetim & Sipariş Portalı", page_icon="🥭", layout="wide")

# 🔒 Merkez Yönetim Paneli Giriş Şifresi
YONETICI_SIFRESI = "1234"

# Veritabanı Bağlantısı
conn = sqlite3.connect('manav_siparisleri.db', check_same_thread=False)
c = conn.cursor()

# Tablo Oluşturma
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

# Sol Menü
st.sidebar.title("📌 Menü")
rol = st.sidebar.radio("Erişim Türü:", ["🏬 Şube Sipariş Girişi", "👑 Merkez Yönetim Paneli"])

# -------------------------------------------------------------
# 1. ŞUBE SİPARİŞ GİRİŞ EKRANI
# -------------------------------------------------------------
if rol == "🏬 Şube Sipariş Girişi":
    st.title("🥭 Şube Manav Sipariş Portalı")
    bugun_str = datetime.now().strftime('%Y-%m-%d')
    st.caption(f"Tarih: {datetime.now().strftime('%d.%m.%Y')}")

    subeler = [
        "-- Seçiniz --", "Raufbey", "Metin Tamer", "Hacı Osmanlı", "Salı Yolu", 
        "Kadiri Yolu", "Nahır Yolu", "Eyup Sultan", "Bulvar", "Düziçi Çarşı", "Aşiyan", "Zeytinlik"
    ]
    secilen_sube = st.selectbox("📍 **Lütfen Şubenizi Seçin:**", subeler)

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
        with st.expander(f"**{row['ADI']}** *(Kod: {row['KODU']})*"):
            col1, col2 = st.columns(2)
            with col1:
                stok = st.number_input("Mevcut Stok", min_value=0.0, step=0.5, key=f"stok_{row['KODU']}")
            with col2:
                siparis = st.number_input("Sipariş Miktarı", min_value=0.0, step=0.5, key=f"sip_{row['KODU']}")
                
            if stok > 0 or siparis > 0:
                kaydedilecek_veriler.append((
                    secilen_sube, 
                    bugun_str, 
                    row['KODU'], 
                    row['ADI'], 
                    stok, 
                    siparis
                ))

    st.divider()

    if st.button("💾 Siparişi Sisteme Kaydet & Merkeze Gönder", type="primary"):
        if secilen_sube == "-- Seçiniz --":
            st.error("Lütfen önce yukarıdan şubenizi seçiniz!")
        elif len(kaydedilecek_veriler) == 0:
            st.warning("Lütfen en az bir ürün için stok veya sipariş miktarı giriniz!")
        else:
            c.executemany('''
                INSERT INTO siparisler (sube, tarih, urun_kodu, urun_adi, mevcut_stok, siparis_miktari)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', kaydedilecek_veriler)
            conn.commit()
            st.success("✅ Siparişiniz başarıyla veritabanına kaydedildi! Merkez birimi anında ekranında görebilir.")

# -------------------------------------------------------------
# 2. ŞİFRELİ MERKEZ YÖNETİM PANİLİ (TARİH FİLTRELİ & GEÇMİŞ DAHİL)
# -------------------------------------------------------------
elif rol == "👑 Merkez Yönetim Paneli":
    st.title("🔒 Merkez Yönetim Paneli")

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
        if st.button("🚪 Çıkış Yap"):
            st.session_state.admin_authed = False
            st.rerun()

        # Custom CSS - Tabloyu Dar (Kompakt) Yapma
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

        st.sidebar.subheader("📅 Tarih Filtresi")
        tum_gecmis = st.sidebar.checkbox("Tüm Geçmiş Verileri Göster", value=True)

        if tum_gecmis:
            df_siparisler = pd.read_sql_query(
                "SELECT sube AS 'Şube', tarih AS 'Tarih', urun_kodu AS 'Ürün Kodu', urun_adi AS 'Ürün Adı', mevcut_stok AS 'Mevcut Stok', siparis_miktari AS 'Sipariş Miktarı' FROM siparisler ORDER BY id DESC", 
                conn
            )
            st.subheader("📋 Genel Sipariş Çizelgesi (Tüm Geçmiş)")
            tarih_etiket = "Tüm_Gecmis"
        else:
            secili_tarih = st.sidebar.date_input("İncelenecek Günü Seçin:", value=date.today())
            secili_tarih_str = secili_tarih.strftime('%Y-%m-%d')
            df_siparisler = pd.read_sql_query(
                "SELECT sube AS 'Şube', tarih AS 'Tarih', urun_kodu AS 'Ürün Kodu', urun_adi AS 'Ürün Adı', mevcut_stok AS 'Mevcut Stok', siparis_miktari AS 'Sipariş Miktarı' FROM siparisler WHERE tarih LIKE ? ORDER BY id DESC", 
                conn, 
                params=(f"{secili_tarih_str}%",)
            )
            st.subheader(f"📋 Günlük Çizelge: {secili_tarih.strftime('%d.%m.%Y')}")
            tarih_etiket = secili_tarih.strftime('%d.%m.%Y')

        if df_siparisler.empty:
            st.info("ℹ️ Seçilen tarih kriterine uygun veritabanında kayıtlı sipariş/stok bulunmamaktadır.")
        else:
            st.sidebar.subheader("🎯 Şube Filtresi")
            secili_subeler = st.sidebar.multiselect("Şubeleri Seçin:", df_siparisler['Şube'].unique(), default=df_siparisler['Şube'].unique())
            filtreli_df = df_siparisler[df_siparisler['Şube'].isin(secili_subeler)]

            # Özet Kartlar
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Toplam Kalem", len(filtreli_df))
            col2.metric("Sipariş Veren Şube", filtreli_df['Şube'].nunique())
            col3.metric("Toplam Stok", f"{filtreli_df['Mevcut Stok'].sum():,.1f} Kg")
            col4.metric("Toplam Sipariş", f"{filtreli_df['Sipariş Miktarı'].sum():,.1f} Kg")

            st.divider()

            # PIVOT TABLO
            pivot_genel = pd.pivot_table(
                filtreli_df, 
                values=['Mevcut Stok', 'Sipariş Miktarı'], 
                index=['Ürün Kodu', 'Ürün Adı'], 
                columns=['Şube'], 
                aggfunc='sum', 
                fill_value=0
            )

            # Şubeleri üst başlık yapma
            pivot_genel = pivot_genel.swaplevel(0, 1, axis=1)
            pivot_genel = pivot_genel.sort_index(axis=1, level=0)

            # Sütun isimlerini kısaltarak genişliği minimuma indirme
            pivot_genel = pivot_genel.rename(columns={'Mevcut Stok': 'Stok', 'Sipariş Miktarı': 'Sip.'})

            # GENEL TOPLAM SÜTUNLARI
            toplam_stok = filtreli_df.groupby(['Ürün Kodu', 'Ürün Adı'])['Mevcut Stok'].sum()
            toplam_siparis = filtreli_df.groupby(['Ürün Kodu', 'Ürün Adı'])['Sipariş Miktarı'].sum()

            pivot_genel[('GENEL TOPLAM', 'Top. Stok')] = toplam_stok
            pivot_genel[('GENEL TOPLAM', 'Top. Sip.')] = toplam_siparis

            # Sadece en az 1 stok veya siparişi olan ürünleri göster
            mask = (pivot_genel[('GENEL TOPLAM', 'Top. Stok')] > 0) | (pivot_genel[('GENEL TOPLAM', 'Top. Sip.')] > 0)
            pivot_genel = pivot_genel[mask]

            st.dataframe(pivot_genel, use_container_width=True, height=550)

            st.divider()

            # EXCEL A4 YATAY ÇIKTI OLUŞTURMA
            def generate_excel(df_pivot, etiket):
                output = io.BytesIO()
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Siparis_Cizelgesi"

                # A4 Yatay ve Sayfaya Sığdır Ayarları
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

                ws.cell(row=1, column=1, value=f"MANAV SİPARİŞ ÇİZELGESİ ({etiket})").font = Font(size=12, bold=True)
                
                ws.cell(row=3, column=1, value="Ürün Kodu").font = font_bold
                ws.cell(row=3, column=2, value="Ürün Adı").font = font_bold
                
                col_idx = 3
                for col in df_pivot.columns:
                    sube, metrik = col
                    ws.cell(row=3, column=col_idx, value=sube).font = font_bold
                    ws.cell(row=4, column=col_idx, value=metrik).font = font_bold
                    
                    if col_idx % 2 == 1 and col_idx < len(df_pivot.columns) + 2:
                        ws.merge_cells(start_row=3, start_column=col_idx, end_row=3, end_column=col_idx+1)
                    col_idx += 1

                row_idx = 5
                for (kodu, adi), row_data in df_pivot.iterrows():
                    ws.cell(row=row_idx, column=1, value=str(kodu)).font = font_normal
                    ws.cell(row=row_idx, column=2, value=str(adi)).font = font_normal
                    
                    c_idx = 3
                    for val in row_data:
                        cell = ws.cell(row=row_idx, column=c_idx, value=val if val != 0 else "")
                        cell.font = font_normal
                        cell.alignment = Alignment(horizontal="center")
                        c_idx += 1
                    row_idx += 1

                for r in ws.iter_rows(min_row=3, max_row=row_idx-1, min_col=1, max_col=len(df_pivot.columns)+2):
                    for cell in r:
                        cell.border = border
                        if cell.row in (3, 4):
                            cell.fill = header_fill
                            cell.alignment = Alignment(horizontal="center", vertical="center")

                ws.column_dimensions['A'].width = 10
                ws.column_dimensions['B'].width = 24
                for c in range(3, len(df_pivot.columns) + 3):
                    col_letter = openpyxl.utils.get_column_letter(c)
                    ws.column_dimensions[col_letter].width = 6.5

                wb.save(output)
                return output.getvalue()

            excel_bytes = generate_excel(pivot_genel, tarih_etiket)

            st.download_button(
                label="🖨️ A4 Yatay Çıktı İçin Excel Dosyasını İndir",
                data=excel_bytes,
                file_name=f"Manav_Siparis_Cizelgesi_{tarih_etiket}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
