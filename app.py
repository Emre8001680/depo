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
SUPABASE_URL = "https://ngokzlndzpodmjiffmjv.supabase.co"
SUPABASE_KEY = "sb_publishable_LJldycoOPfyCh-stDwAFjg_EVpjACxQ"

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

try:
    supabase = init_supabase()
except Exception as e:
    st.error("Supabase bağlantısı kurulamadı. Lütfen internet bağlantınızı ve bilgileri kontrol edin.")

# Sayfa Yapılandırması
st.set_page_config(page_title="Yalçın Marketler Zinciri - Manav Portalı", page_icon="🥭", layout="wide")

# CSS DÜZENLEMELERİ
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

if "giris_yapilan_sube" not in st.session_state:
    st.session_state.giris_yapilan_sube = None

# -------------------------------------------------------------
# 🌟 KARŞILAMA EKRANI
# -------------------------------------------------------------
if not st.session_state.site_giris_yapildi:
    st.markdown("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
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

    SUBE_SIFRELERI = {
        "Raufbey": "1001",
        "Metin Tamer": "1002",
        "Hacı Osmanlı": "1003",
        "Salı Yolu": "1004",
        "Kadiri Yolu": "1005",
        "Nahır Yolu": "1006",
        "Eyup Sultan": "1007",
        "Bulvar": "1008",
        "Düziçi Çarşı": "1009",
        "Aşiyan": "1010",
        "Zeytinlik": "1011"
    }

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
            st.session_state.giris_yapilan_sube = None
            st.rerun()

    st.divider()

    rol = st.session_state.aktif_rol

    # -------------------------------------------------------------
    # 1. ŞUBE SİPARİŞ GİRİŞİ (SUPABASE BULUT DESTEKLİ)
    # -------------------------------------------------------------
    if rol == "🏬 Şube Sipariş Girişi":
        st.markdown("<h2 style='text-align: center;'>🥭 Şube Manav Sipariş Portalı</h2>", unsafe_allow_html=True)

        bugun_str = datetime.now().strftime('%Y-%m-%d')
        st.caption(f"Tarih: {datetime.now().strftime('%d.%m.%Y')}")

        subeler = ["-- Seçiniz --"] + list(SUBE_SIFRELERI.keys())
        secilen_sube = st.selectbox("📍 **Lütfen Şubenizi Seçin:**", subeler)

        if secilen_sube != "-- Seçiniz --":
            
            # 🔒 ŞÜBE ŞİFRE DOĞRULAMA KONTROLÜ
            if st.session_state.giris_yapilan_sube != secilen_sube:
                st.info(f"🔒 **{secilen_sube}** şubesinin sipariş ekranına erişmek için lütfen şube şifrenizi giriniz.")
                
                s_col1, s_col2 = st.columns([2, 1])
                with s_col1:
                    girilen_pin = st.text_input(f"🔑 {secilen_sube} Şube Şifresi:", type="password", key=f"pin_input_{secilen_sube}")
                with s_col2:
                    st.markdown("<br>", unsafe_allow_html=True)
                    if st.button("Giriş Yap", type="primary", use_container_width=True):
                        dogru_pin = SUBE_SIFRELERI.get(secilen_sube)
                        if girilen_pin == dogru_pin:
                            st.session_state.giris_yapilan_sube = secilen_sube
                            st.success("✅ Şifre Doğrulandı!")
                            st.rerun()
                        else:
                            st.error("❌ Hatalı Şube Şifresi!")
            
            # ✅ ŞİFRE DOĞRULANDI - SİPARİŞ FORMU AÇILIYOR
            else:
                st.success(f"🔓 **{secilen_sube}** Şubesi Girişi Aktif")
                
                if st.button("🔒 Şube Oturumunu Kapat", type="secondary"):
                    st.session_state.giris_yapilan_sube = None
                    st.rerun()

                st.divider()

                # --- SUPABASE'DEN VERİ ÇEKME ---
                res = supabase.table("siparisler") \
                    .select("urun_kodu, mevcut_stok, siparis_miktari") \
                    .eq("sube", secilen_sube) \
                    .eq("tarih", bugun_str) \
                    .execute()
                
                kayitli_dict = {}
                for r in res.data:
                    kayitli_dict[r['urun_kodu']] = {
                        'stok': str(r['mevcut_stok']),
                        'siparis': float(r['siparis_miktari'])
                    }

                if len(kayitli_dict) > 0:
                    st.info(f"ℹ️ **{secilen_sube}** şubesinin bugün girilmiş siparişleri buluttan yüklendi.")

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
                        col1, col2 = st.columns([1.5, 1])
                        
                        with col1:
                            stok_dolu = st.checkbox("🟢 Reyon Dolu (Depo Boş)", value=(varsayilan_stok_str == "Reyon Dolu"), key=f"dolu_{kod}")
                            
                            if not stok_dolu:
                                stok_val = st.number_input("Mevcut Stok (Kasa)", min_value=0.0, step=1.0, value=float(varsayilan_stok_str) if varsayilan_stok_str.replace('.','',1).isdigit() else 0.0, key=f"stok_{kod}")
                                stok_kayit = str(int(stok_val))
                            else:
                                stok_kayit = "Reyon Dolu"
                                st.caption("📌 *Stok 'Reyon Dolu' olarak kaydedilecek.*")

                        with col2:
                            siparis = st.number_input("Sipariş (Kasa)", min_value=0.0, step=1.0, value=varsayilan_siparis, key=f"sip_{kod}")
                            
                        if (stok_kayit != "0" and stok_kayit != "0.0") or siparis > 0:
                            kaydedilecek_veriler.append({
                                "sube": secilen_sube,
                                "tarih": bugun_str,
                                "urun_kodu": kod,
                                "urun_adi": row['ADI'],
                                "mevcut_stok": stok_kayit,
                                "siparis_miktari": siparis
                            })

                st.divider()

                btn_col1, btn_col2 = st.columns([2, 1])

                with btn_col1:
                    if st.button("💾 Siparişleri Güncelle / Kaydet", type="primary", use_container_width=True):
                        # SUPABASE KAYIT GÜNCELLEME
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
                        st.error("🗑️ Bugünkü siparişiniz buluttan tamamen silindi!")
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

            # SUPABASE'DEN VERİ SORGULAMA
            if tum_gecmis:
                res = supabase.table("siparisler").select("sube, tarih, urun_kodu, urun_adi, mevcut_stok, siparis_miktari").order("id", desc=True).execute()
            else:
                res = supabase.table("siparisler").select("sube, tarih, urun_kodu, urun_adi, mevcut_stok, siparis_miktari").eq("tarih", secili_tarih_str).order("id", desc=True).execute()

            df_siparisler = pd.DataFrame(res.data)

            if df_siparisler.empty:
                st.warning(f"ℹ️ Seçilen tarihte ({tarih_etiket}) kayıtlı sipariş bulunmamaktadır.")
            else:
                df_siparisler = df_siparisler.rename(columns={
                    'sube': 'Şube',
                    'tarih': 'Tarih',
                    'urun_kodu': 'Ürün Kodu',
                    'urun_adi': 'Ürün Adı',
                    'mevcut_stok': 'Mevcut Stok',
                    'siparis_miktari': 'Sipariş Miktarı'
                })

                with f_col3:
                    secili_subeler = st.multiselect("Şube Filtresi:", df_siparisler['Şube'].unique(), default=df_siparisler['Şube'].unique())
                
                filtreli_df = df_siparisler[df_siparisler['Şube'].isin(secili_subeler)].copy()

                st.divider()

                col1, col2, col3 = st.columns(3)
                col1.metric("Toplam Kalem", len(filtreli_df))
                col2.metric("Sipariş Veren Şube", filtreli_df['Şube'].nunique())
                col3.metric("Toplam Sipariş", f"{filtreli_df['Sipariş Miktarı'].sum():,.0f} Kasa")

                st.divider()

                # Pivot Tablo Oluşturma
                pivot_genel = pd.pivot_table(
                    filtreli_df, 
                    values=['Mevcut Stok', 'Sipariş Miktarı'], 
                    index=['Ürün Kodu', 'Ürün Adı'], 
                    columns=['Şube'], 
                    aggfunc='first', 
                    fill_value="0"
                )

                pivot_genel = pivot_genel.swaplevel(0, 1, axis=1)
                pivot_genel = pivot_genel.sort_index(axis=1, level=0)
                pivot_genel = pivot_genel.rename(columns={'Mevcut Stok': 'Stok', 'Sipariş Miktarı': 'Sip.'})

                # GENEL TOPLAM HESAPLAMA
                toplam_siparis = filtreli_df.groupby(['Ürün Kodu', 'Ürün Adı'])['Sipariş Miktarı'].sum()
                
                def calc_stok_toplam(g):
                    num_sum = 0
                    rd_cnt = 0
                    for val in g['Mevcut Stok']:
                        val_str = str(val).strip()
                        if val_str == "Reyon Dolu":
                            rd_cnt += 1
                        else:
                            try:
                                num_sum += float(val_str)
                            except:
                                pass
                    res = []
                    if num_sum > 0:
                        res.append(f"{int(num_sum)} Kasa")
                    if rd_cnt > 0:
                        res.append(f"{rd_cnt} RD")
                    return " + ".join(res) if res else "0"

                toplam_stok_str = filtreli_df.groupby(['Ürün Kodu', 'Ürün Adı']).apply(calc_stok_toplam)

                pivot_genel[('GENEL TOPLAM', 'Top. Stok / RD')] = toplam_stok_str
                pivot_genel[('GENEL TOPLAM', 'Top. Sipariş')] = toplam_siparis

                st.dataframe(pivot_genel, use_container_width=True, height=550)

                st.divider()

                # --- EXCEL İNDİRME HAZIRLIĞI ---
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

                    ws.cell(row=1, column=1, value=f"YALÇIN MARKETLER ZİNCİRİ MANAV SİPARİŞ ÇİZELGESİ ({etiket})").font = Font(size=12, bold=True)
                    
                    ws.cell(row=3, column=1, value="Ürün Kodu").font = font_bold
                    ws.cell(row=3, column=2, value="Ürün Adı").font = font_bold
                    
                    col_idx = 3
                    for col in df_pivot.columns:
                        sube_adi = str(col[0]) if isinstance(col, tuple) else str(col)
                        metrik_adi = str(col[1]) if isinstance(col, tuple) and len(col) > 1 else ""
                        
                        ws.cell(row=3, column=col_idx, value=sube_adi).font = font_bold
                        ws.cell(row=4, column=col_idx, value=metrik_adi).font = font_bold
                        col_idx += 1

                    for c in range(3, col_idx, 2):
                        if c + 1 < col_idx:
                            ws.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+1)

                    row_idx = 5
                    for (kodu, adi), row_data in df_pivot.iterrows():
                        ws.cell(row=row_idx, column=1, value=str(kodu)).font = font_normal
                        ws.cell(row=row_idx, column=2, value=str(adi)).font = font_normal
                        
                        c_idx = 3
                        for val in row_data:
                            cell_val = ws.cell(row=row_idx, column=c_idx, value=str(val) if val != 0 else "")
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
                        ws.column_dimensions[col_letter].width = 7.5

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
