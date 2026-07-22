import pandas as pd
import streamlit as st
from datetime import datetime, date
import sqlite3
import io
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# Sayfa Yapılandırması
st.set_page_config(page_title="Yalçın Marketler Zinciri - Manav Portalı", page_icon="🥭", layout="wide")

# CSS DÜZENLEMELERİ VE OTO-SEÇİM JAVASCRIPT'I
st.markdown("""
    <style>
        #MainMenu {visibility: hidden !important;}
        footer {visibility: hidden !important;}
        header {visibility: hidden !important;}
        [data-testid="stHeader"] {display: none !important;}
        [data-testid="stSidebar"] {display: none !important;}
        
        .block-container {
            padding-top: 1.5rem !important;
            padding-bottom: 2rem !important;
        }

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

    <script>
        // Tıklanan input kutusundaki metni otomatik seçme fonksiyonu
        document.addEventListener('focusin', function(e) {
            if (e.target.tagName === 'INPUT') {
                e.target.select();
            }
        });
    </script>
""", unsafe_allow_html=True)

# Session State Tanımlamaları
if "site_giris_yapildi" not in st.session_state:
    st.session_state.site_giris_yapildi = False

if "aktif_rol" not in st.session_state:
    st.session_state.aktif_rol = "🏬 Şube Sipariş Girişi"

# -------------------------------------------------------------
# 🌟 KARŞILAMA EKRANI
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
# 🏢 ANA UYGULAMA
# -------------------------------------------------------------
else:
    YONETICI_SIFRESI = "1234"

    conn = sqlite3.connect('manav_siparisleri.db', check_same_thread=False)
    c = conn.cursor()

    c.execute('''
        CREATE TABLE IF NOT EXISTS siparisler (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sube TEXT,
            tarih TEXT,
            urun_kodu TEXT,
            urun_adi TEXT,
            mevcut_stok TEXT,
            siparis_miktari REAL
        )
    ''')
    conn.commit()

    # --- GEÇİŞ MENÜSÜ ---
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
    # 1. ŞUBE SİPARİŞ GİRİŞİ
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
                    'stok': str(r['mevcut_stok']),
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
            st.subheader("📦 Stok ve Sipariş Girişi (Kasa)")

            for index, row in filtre_df.iterrows():
                kod = row['KODU']
                varsayilan_stok_str = kayitli_dict.get(kod, {}).get('stok', "0")
                varsayilan_siparis = kayitli_dict.get(kod, {}).get('siparis', 0.0)

                with st.expander(f"**{row['ADI']}** *(Kod: {kod})*"):
                    col1, col2, col3 = st.columns([1.5, 1, 1.5])
                    
                    with col1:
                        # Stok Durumu Seçeneği
                        stok_dolu = st.checkbox("🟢 Reyon Dolu (Depo Boş)", value=(varsayilan_stok_str == "Reyon Dolu"), key=f"dolu_{kod}")
                        
                        if not stok_dolu:
                            stok_val = st.number_input("Mevcut Stok (Kasa)", min_value=0.0, step=1.0, value=float(varsayilan_stok_str) if varsayilan_stok_str.replace('.','',1).isdigit() else 0.0, key=f"stok_{kod}")
                            stok_kayit = str(stok_val)
                        else:
                            stok_kayit = "Reyon Dolu"
                            st.caption("📌 *Stok 'Reyon Dolu' olarak kaydedilecek.*")

                    with col2:
                        siparis = st.number_input("Sipariş (Kasa)", min_value=0.0, step=1.0, value=varsayilan_siparis, key=f"sip_{kod}")
                        
                    if (stok_kayit != "0.0" and stok_kayit != "0") or siparis > 0:
                        kaydedilecek_veriler.append((
                            secilen_sube, 
                            bugun_str, 
                            kod, 
                            row['ADI'], 
                            stok_kayit, 
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
    # 2. MERKEZ YÖNETİM PANELİ
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
                    st.error("❌ Hatalı şifre!")
        else:
            st.success("🔓 Yetkili Girişi Başarılı")
            if st.button("🚪 Oturumu Kapat"):
                st.session_state.admin_authed = False
                st.rerun()

            st.subheader("📊 Sipariş ve Stok Çizelgesi (Kasa)")

            f_col1, f_col2, f_col3 = st.columns([1.2, 1, 2])
            
            with f_col1:
                tum_gecmis = st.checkbox("Tüm Geçmiş Verileri Göster", value=False)
            
            with f_col2:
                if not tum_gecmis:
                    secili_tarih = st.date_input("Tarih Seçin:", value=date.today())
                    secili_tarih_str = secili_tarih.strftime('%Y-%m-%d')
                    tarih_etiket = secili_tarih.strftime('%d.%m.%Y')
                else:
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
                st.warning(f"ℹ️ Seçilen tarihte ({tarih_etiket}) kayıtlı sipariş bulunmamaktadır.")
            else:
                with f_col3:
                    secili_subeler = st.multiselect("Şube Filtresi:", df_siparisler['Şube'].unique(), default=df_siparisler['Şube'].unique())
                
                filtreli_df = df_siparisler[df_siparisler['Şube'].isin(secili_subeler)]

                st.divider()

                col1, col2, col3 = st.columns(3)
                col1.metric("Toplam Kalem", len(filtreli_df))
                col2.metric("Sipariş Veren Şube", filtreli_df['Şube'].nunique())
                col3.metric("Toplam Sipariş", f"{filtreli_df['Sipariş Miktarı'].sum():,.0f} Kasa")

                st.divider()

                pivot_genel = pd.pivot_table(
                    filtreli_df, 
                    values=['Mevcut Stok', 'Sipariş Miktarı'], 
                    index=['Ürün Kodu', 'Ürün Adı'], 
                    columns=['Şube'], 
                    aggfunc='first', 
                    fill_value="0"
                )

                st.dataframe(pivot_genel, use_container_width=True, height=550)
