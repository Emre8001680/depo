import pandas as pd
import streamlit as st
from datetime import datetime
import urllib.parse

# Sayfa Yapılandırması (Mobil Uyumlu)
st.set_page_config(page_title="Manav Sipariş Portalı", page_icon="🥭", layout="centered")

# CSS ile Mobil Tasarımı Özelleştirme
st.markdown("""
    <style>
        .stButton>button {
            width: 100%;
            background-color: #25D366;
            color: white;
            font-size: 18px;
            font-weight: bold;
            border-radius: 10px;
            padding: 10px;
        }
        .stButton>button:hover {
            background-color: #128C7E;
            color: white;
        }
    </style>
""", unsafe_allow_html=True)

# Başlık
st.title("🥭 Manav Sipariş Portalı")
st.caption(f"Tarih: {datetime.now().strftime('%d.%m.%Y')}")

# 1. Şube Seçimi (Güncellendi)
subeler = [
    "-- Seçiniz --", 
    "Raufbey",
    "Metin Tamer",
    "Hacı Osmanlı",
    "Salı Yolu",
    "Kadiri Yolu",
    "Nahır Yolu",
    "Eyup Sultan",
    "Bulvar",
    "Düziçi Çarşı",
    "Aşiyan",
    "Zeytinlik"
]
secilen_sube = st.selectbox("📍 **Lütfen Şubenizi Seçin:**", subeler)

# 2. Ürün Listesi
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

# 3. Arama Barı
arama = st.text_input("🔍 **Ürün Ara (Adı veya Kodu):**", "")

if arama:
    filtre_df = df[df['ADI'].str.contains(arama, case=False) | df['KODU'].str.contains(arama, case=False)]
else:
    filtre_df = df

# 4. Ürün Listesi Ekranı
siparis_verileri = []

st.subheader("🛒 Stok ve Sipariş Girişi")

for index, row in filtre_df.iterrows():
    with st.expander(f"**{row['ADI']}** *(Kod: {row['KODU']})*"):
        col1, col2 = st.columns(2)
        with col1:
            stok = st.number_input("Mevcut Stok", min_value=0.0, step=0.5, key=f"stok_{row['KODU']}")
        with col2:
            siparis = st.number_input("Sipariş Miktarı", min_value=0.0, step=0.5, key=f"sip_{row['KODU']}")
            
        if stok > 0 or siparis > 0:
            siparis_verileri.append({
                "urun": row['ADI'],
                "stok": stok,
                "siparis": siparis
            })

st.divider()

# 5. WhatsApp Gönderme Butonu
if st.button("📲 WhatsApp ile Gönder"):
    if secilen_sube == "-- Seçiniz --":
        st.error("Lütfen önce yukarıdan bir şube seçiniz!")
    elif len(siparis_verileri) == 0:
        st.warning("Lütfen en az bir ürün için stok veya sipariş miktarı giriniz!")
    else:
        # Mesajı Oluştur
        mesaj = f"🛒 *MANAV GÜNLÜK SİPARİŞ RAPORU*\n"
        mesaj += f"📍 *Şube:* {secilen_sube}\n"
        mesaj += f"📅 *Tarih:* {datetime.now().strftime('%d.%m.%Y')}\n"
        mesaj += "-----------------------------------\n\n"
        
        for item in siparis_verileri:
            mesaj += f"🔹 *{item['urun']}*\n"
            mesaj += f"   • Mevcut Stok: {item['stok']}\n"
            mesaj += f"   • Sipariş: {item['siparis']}\n\n"
        
        # WhatsApp Linkini Hazırla
        encoded_msg = urllib.parse.quote(mesaj)
        wa_url = f"https://wa.me/?text={encoded_msg}"
        
        st.success("Sipariş metni hazırlandı! Aşağıdaki bağlantıya tıklayarak gruba gönderebilirsiniz:")
        st.markdown(f"[👉 **WhatsApp'a Aktarmak İçin Tıklayın**]({wa_url})", unsafe_allow_html=True)