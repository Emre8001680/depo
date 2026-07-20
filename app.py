import pandas as pd
import streamlit as st
from datetime import datetime
import sqlite3

# Sayfa Yapılandırması
st.set_page_config(page_title="Manav Yönetim & Sipariş Portalı", page_icon="🥭", layout="wide")

# Veritabanı Bağlantısı (Siparişleri Uygulama İçinde Saklar)
conn = sqlite3.connect('manav_siparisleri.db', check_same_thread=False)
c = conn.cursor()

# Veritabanı Tablosu Oluşturma
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

# Sol Menü: Kullanıcı Rolü Seçimi
st.sidebar.title("📌 Menü")
rol = st.sidebar.radio("Erişim Türü:", ["🏬 Şube Sipariş Girişi", "👑 Merkez Yönetim Paneli"])

# -------------------------------------------------------------
# 1. ŞUBE SİPARİŞ GİRİŞ EKRANI
# -------------------------------------------------------------
if rol == "🏬 Şube Sipariş Girişi":
    st.title("🥭 Şube Manav Sipariş Portalı")
    st.caption(f"Tarih: {datetime.now().strftime('%d.%m.%Y')}")

    subeler = [
        "-- Seçiniz --", "Raufbey", "Metin Tamer", "Hacı Osmanlı", "Salı Yolu", 
        "Kadiri Yolu", "Nahır Yolu", "Eyup Sultan", "Bulvar", "Düziçi Çarşı", "Aşiyan", "Zeytinlik"
    ]
    secilen_sube = st.selectbox("📍 **Lütfen Şubenizi Seçin:**", subeler)

    # Ürün Listesi
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
                    datetime.now().strftime('%Y-%m-%d %H:%M'), 
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
# 2. MERKEZ YÖNETİM PANİLİ (DİJİTAL SİPARİŞ TAKİBİ)
# -------------------------------------------------------------
elif rol == "👑 Merkez Yönetim Paneli":
    st.title("📊 Merkez Sipariş & Stok Takip Paneli")
    st.caption("Şubeler tarafından girilen anlık veriler aşağıda listelenmektedir.")

    # Verileri Çek
    df_siparisler = pd.read_sql_query("SELECT sube AS 'Şube', tarih AS 'Tarih/Saat', urun_kodu AS 'Ürün Kodu', urun_adi AS 'Ürün Adı', mevcut_stok AS 'Mevcut Stok', siparis_miktari AS 'Sipariş Miktarı' FROM siparisler ORDER BY id DESC", conn)

    if df_siparisler.empty:
        st.info("Henüz sistemde kaydedilmiş bir sipariş bulunmamaktadır.")
    else:
        # Filtreleme Seçenekleri
        st.sidebar.subheader("🎯 Yönetim Filtreleri")
        secili_subeler = st.sidebar.multiselect("Şubeye Göre Filtrele:", df_siparisler['Şube'].unique(), default=df_siparisler['Şube'].unique())
        
        filtreli_df = df_siparisler[df_siparisler['Şube'].isin(secili_subeler)]

        # Özet Kartları (KPI)
        col1, col2, col3 = st.columns(3)
        col1.metric("Toplam Sipariş Kalemi", len(filtreli_df))
        col2.metric("Sipariş Veren Şube Sayısı", filtreli_df['Şube'].nunique())
        col3.metric("Toplam Talep Edilen Miktar", f"{filtreli_df['Sipariş Miktarı'].sum():,.1f} Kg/Adet")

        st.divider()

        # Sekmeli Görünüm
        tab1, tab2 = st.tabs(["📋 Şube Bazlı Detaylı Liste", "📦 Konsolide Toplam Siparişler (Depo/Hal İçin)"])

        with tab1:
            st.subheader("Şubelerin Anlık Giriş Detayları")
            st.dataframe(filtreli_df, use_container_width=True)

        with tab2:
            st.subheader("Ürün Bazında Toplam Sipariş Miktarları")
            st.caption("Tüm şubelerden gelen toplam ihtiyacı gösterir:")
            
            toplam_df = filtreli_df.groupby(['Ürün Kodu', 'Ürün Adı'])['Sipariş Miktarı'].sum().reset_index()
            toplam_df = toplam_df[toplam_df['Sipariş Miktarı'] > 0] # Sadece siparişi olanlar
            st.dataframe(toplam_df, use_container_width=True)

        st.divider()
        
        # Excel İndirme Butonu
        csv_data = filtreli_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 Günlük Sipariş Raporunu Excel/CSV Olarak İndir",
            data=csv_data,
            file_name=f"Manav_Siparis_Raporu_{datetime.now().strftime('%d_%m_%Y')}.csv",
            mime="text/csv",
            type="primary"
        )
