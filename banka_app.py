
import streamlit as st
import pdfplumber
import extract_msg
from bs4 import BeautifulSoup
import re
import pandas as pd
from io import BytesIO
import xlrd  # Ziraat .xls dosyaları için gerekli

# --- ORTAK YARDIMCI FONKSİYONLAR ---

def temizle_sayi(raw: str) -> str:
    """İçinden sadece sayı, nokta, virgül ve +/- işaretlerini bırak."""
    if not raw: return ""
    return re.sub(r"[^\d\.,+-]", "", str(raw))

def sec_en_mantikli_sayi(candidates):
    """Bulunan sayı adayları arasından en mantıklı (genelde sonuncu) olanı seç."""
    temiz_liste = [c for c in candidates if len(c) > 2]
    if not temiz_liste:
        return "Bulunamadı"
    return temiz_liste[-1]

# --- ZİRAAT XLS BAKİYE OKUMA ---

def bakiye_oku_ziraat_xls(file_obj):
    try:
        # Ziraat .xls dosyaları genelde xlrd motoru ile okunur
        df = pd.read_excel(file_obj, header=None, engine='xlrd')
        for row_idx, row in df.iterrows():
            row_list = [str(val) for val in row.values if str(val) != 'nan']
            row_text = " ".join(row_list).upper()
            if "BAKIYE" in row_text:
                candidates = []
                for cell in row_list:
                    nums = re.findall(r'[\d\.]+\,\d{2}|[\d\,]+\.\d{2}|\d+', cell)
                    candidates.extend(nums)
                if candidates:
                    return candidates[-1]
    except Exception as e:
        return f"Hata: {str(e)}"
    return "Bulunamadı"

# --- PDF BAKİYE OKUMA ---

def bakiye_oku_pdf(file_obj):
    keywords = ["KAPANIŞ BAKİYESİ", "CLOSING BALANCE", "BAKİYE", "DEVR. BAKİYE", "MEVCUT BAKİYE", "NET BAKİYE"]
    try:
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                # Önce metin içinden ara
                words = page.extract_words() or []
                for i, word in enumerate(words):
                    clean_word = word['text'].upper().replace('İ', 'I')
                    if any(k.replace('İ', 'I') in clean_word for k in keywords):
                        for j in range(i + 1, min(i + 20, len(words))):
                            val = words[j]['text']
                            if "/" in val or any(c.isalpha() for c in val): continue
                            if re.search(r"\d", val):
                                res = temizle_sayi(val)
                                if len(res) > 2: return res
                
                # Sonra tablo içinden ara
                tables = page.extract_tables() or []
                for table in tables:
                    for row in reversed(table):
                        row_text = " ".join([str(c) for c in row if c]).replace('\n', ' ')
                        if any(k in row_text.upper() for k in keywords):
                            parts = row_text.split()
                            candidates = [temizle_sayi(p) for p in parts if "/" not in p and re.search(r"[\d\.,]+", p)]
                            if candidates: return sec_en_mantikli_sayi(candidates)
    except: pass
    return "Bulunamadı"

# --- MSG BAKİYE OKUMA ---

def oku_msg_icerik(file_obj):
    try:
        raw_bytes = file_obj.read()
        file_obj.seek(0)
        msg = extract_msg.Message(BytesIO(raw_bytes))
        html = msg.htmlBody if hasattr(msg, "htmlBody") and msg.htmlBody else None
        if html:
            return BeautifulSoup(html, 'html.parser').get_text(separator=' ')
        return msg.body if hasattr(msg, "body") and msg.body else ""
    except: return ""

def bakiye_oku_msg_genel(file_obj):
    text = oku_msg_icerik(file_obj)
    if not text: return "Bulunamadı"
    desenler = [r"MEVCUT BAKIYE[:\s]*([\d\.\,]+)", r"NET BAKIYE[:\s]*([\d\.\,]+)", r"BAKIYE[:\s]*([\d\.\,]+)"]
    adaylar = []
    for pattern in desenler:
        m = re.search(pattern, text.upper())
        if m: adaylar.append(temizle_sayi(m.group(1)))
    if not adaylar:
        adaylar = re.findall(r'[\d\.]+\,\d{2}', text)
    return sec_en_mantikli_sayi(adaylar)

# --- İŞBANK XLS BAKİYE OKUMA ---

def bakiye_oku_isbank_xls(file_obj):
    try:
        content = file_obj.getvalue().decode('latin-1', errors='ignore')
        for line in content.splitlines():
            if any(k in line for k in ["Mevcut Bakiye", "Net Bakiye"]):
                nums = re.findall(r'[\d\.,]+', line.replace('"', ''))
                if nums: return nums[-1]
    except: pass
    return "Bulunamadı"

# --- STREAMLIT ARAYÜZÜ ---

st.set_page_config(page_title="PMI Banka Mutabakat", layout="wide")
st.title("🏦 Banka Ekstre Bakiye Ayıklayıcı")

uploaded_files = st.file_uploader("Dosyaları Seçin (PDF, MSG, XLS)", accept_multiple_files=True, type=['pdf', 'msg', 'xls'])

if uploaded_files:
    results = []
    for uploaded_file in uploaded_files:
        fname = uploaded_file.name
        ext = fname.split('.')[-1].lower()
        
        if ext == "pdf":
            bakiye = bakiye_oku_pdf(uploaded_file)
        elif ext == "msg":
            bakiye = bakiye_oku_msg_genel(uploaded_file)
        elif ext == "xls":
            if "HESAPOZETI" in fname.upper() or "ZIRAAT" in fname.upper():
                bakiye = bakiye_oku_ziraat_xls(uploaded_file)
            else:
                bakiye = bakiye_oku_isbank_xls(uploaded_file)
        else:
            bakiye = "Format Desteklenmiyor"
        
        results.append({"Dosya Adı": fname, "Ayıklanan Bakiye": bakiye})

    df = pd.DataFrame(results)
    st.divider()
    st.subheader("📊 Ayıklanan Bakiyeler")
    st.dataframe(df, use_container_width=True)
 
 
 
 
 
