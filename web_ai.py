import io, json, os, re, time
import pandas as pd
import streamlit as st
from google import genai
from google.genai import types
from PIL import Image, ImageOps
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Kebab King AI", page_icon="🧾", layout="wide")

LANG_DICT = {
    "Türkçe": {"title":"🧾 King Akıllı Muhasebe (JPK)","subtitle":"Faturaları yükleyin; yapay zeka JPK_V7 uyumlu resmi Excel raporunuzu hazırlasın.","sidebar_title":"⚙️ Sistem Durumu","api_label":"Gemini API Anahtarı:","api_success":"✅ Anahtar hazır.","api_link":"Ücretsiz API Anahtarı Alın","upload_label":"Faturaları seçin veya sürükleyin (PNG, JPG)","warning_api":"Gemini API anahtarı bulunamadı.","warning_yellow":"Sarı ile işaretlenen satırlar bir uyarı içerir (eksik NIP, matematiksel hata) - elle kontrol edin.","progress":"İşlenen fatura","success_msg":"Tüm faturalar başarıyla okundu!","preview":"Kurumsal Veri Önizlemesi","download":"📊 Resmi Muhasebe Excel Raporunu İndir","timing":"Ortalama işlem süresi"},
    "English": {"title":"🧾 King Smart Accounting (JPK)","subtitle":"Upload invoices; AI will generate your JPK_V7 compliant official Excel report.","sidebar_title":"⚙️ System Status","api_label":"Gemini API Key:","api_success":"✅ Key is ready.","api_link":"Get a Free API Key","upload_label":"Drag and drop or select invoices (PNG, JPG)","warning_api":"Gemini API key was not found.","warning_yellow":"Rows highlighted in yellow contain warnings (missing NIP, math errors) - verify manually.","progress":"Processed invoice","success_msg":"All invoices read successfully!","preview":"Corporate Data Preview","download":"📊 Download Official Excel Report","timing":"Average processing time"},
    "Polski": {"title":"🧾 Inteligentna Księgowość King (JPK)","subtitle":"Prześlij faktury; AI wygeneruje oficjalny raport Excel zgodny z JPK_V7.","sidebar_title":"⚙️ Status Systemu","api_label":"Klucz API Gemini:","api_success":"✅ Klucz jest gotowy.","api_link":"Pobierz darmowy klucz API","upload_label":"Wybierz lub przeciągnij faktury (PNG, JPG)","warning_api":"Nie znaleziono klucza API Gemini.","warning_yellow":"Wiersze zaznaczone na żółto zawierają ostrzeżenia (brak NIP, błędy matematyczne) - sprawdź ręcznie.","progress":"Przetworzona faktura","success_msg":"Wszystkie faktury zostały odczytane!","preview":"Podgląd Danych Firmowych","download":"📊 Pobierz Oficjalny Raport Excel","timing":"Średni czas przetwarzania"},
}

MODEL_NAME, FALLBACK_MODEL = "gemini-3.8-flash", "gemini-3.7-flash"
MAX_REASONABLE_AMOUNT, MAX_IMAGE_SIZE = 100_000, 1800

PROMPT = """Sen Polonya resmi muhasebe ve vergi (JPK_V7) mevzuatında uzmanlaşmış
kıdemli bir yapay zeka asistanısın.

Ekli fatura görselini dikkatlice incele ve yalnızca aşağıdaki JSON alanlarını doldur.

Kurallar:
- Tarihleri YYYY-MM-DD formatında döndür.
- NIP numaralarında tire/boşluk kullanma; yalnızca 10 rakam döndür.
- Brutto = Netto + VAT matematiksel kontrolünü yap.
- Bulamadığın veya okuyamadığın alanları tahmin etme; boş bırak.
- "Not" alanına yalnızca kısa bir muhasebe uyarısı gerekiyorsa yaz.

{
  "Data Wystawienia": "YYYY-MM-DD",
  "Data Sprzedaży": "YYYY-MM-DD",
  "Nr Faktury": "Örnek: FV 1/09/2019",
  "Satıcı Unvanı": "Firma Adı Sp. z o.o.",
  "NIP Sprzedawcy": "1234567890",
  "NIP Nabywcy": "1234567890",
  "Netto (PLN)": 100.50,
  "Stawka VAT": "23%",
  "Kwota VAT (PLN)": 23.11,
  "Wartość Brutto (PLN)": 123.61,
  "Sposób Płatności": "Przelew / Gotówka / Za pobraniem",
  "Not": ""
}"""

def get_api_key():
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
        if key: return str(key).strip()
    except Exception: pass
    return os.getenv("GEMINI_API_KEY", "").strip()

@st.cache_resource
def get_gemini_client(api_key):
    return genai.Client(api_key=api_key, http_options=types.HttpOptions(
        timeout=90_000,
        retry_options=types.HttpRetryOptions(
            attempts=4, initial_delay=1.0, max_delay=12.0, exp_base=2.0,
            jitter=0.2, http_status_codes=[408,429,500,502,503,504]
        )
    ))

def extract_json(text):
    text = (text or "").strip()
    try: return json.loads(text)
    except Exception: pass
    fenced = re.sub(r"```(?:json)?\s*(.*?)\s*```", r"\1", text, flags=re.DOTALL).strip()
    try: return json.loads(fenced)
    except Exception: pass
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end > start: return json.loads(text[start:end+1])
    raise ValueError("Yanıt içinden geçerli bir JSON çıkarılamadı.")

def validate_and_flag(data):
    notes = []
    existing = str(data.get("Not","") or "").strip()
    if existing: notes.append(existing)
    netto, vat, brutto = data.get("Netto (PLN)"), data.get("Kwota VAT (PLN)"), data.get("Wartość Brutto (PLN)")
    if all(isinstance(x,(int,float)) for x in (netto,vat,brutto)):
        if abs(netto + vat - brutto) > 0.05:
            notes.append(f"Matematiksel tutarsızlık: Netto+VAT ({netto+vat:.2f}) ≠ Brutto ({brutto:.2f})")
        if brutto > MAX_REASONABLE_AMOUNT: notes.append("Anormal yüksek tutar")
    for field,label in [("NIP Sprzedawcy","Satıcı NIP"),("NIP Nabywcy","Alıcı NIP")]:
        nip = str(data.get(field,"") or "")
        digits = re.sub(r"\D","",nip)
        if nip: data[field] = digits
        if nip and len(digits) != 10: notes.append(f"{label} 10 hane değil")
    data["Not"] = " | ".join(notes)
    return data

def prepare_image(file_bytes):
    image = Image.open(io.BytesIO(file_bytes))
    image = ImageOps.exif_transpose(image).convert("RGB")
    image.thumbnail((MAX_IMAGE_SIZE,MAX_IMAGE_SIZE), Image.Resampling.LANCZOS)
    return image

def error_code(error):
    return getattr(error,"status_code",None) or getattr(error,"code",None)

def friendly_api_error(error):
    code = error_code(error)
    if code in (401,403): return "API anahtarı geçersiz veya yetkilendirilmemiş. Streamlit Secrets içindeki GEMINI_API_KEY değerini kontrol edin."
    if code == 429: return "Gemini kullanım limiti geçici olarak aşıldı. Biraz sonra tekrar deneyin."
    if code == 503: return "Gemini servisi şu anda yoğun. Otomatik tekrar denemeler başarısız oldu."
    if code in (408,500,502,504): return "Gemini servisine şu anda ulaşılamıyor. Otomatik tekrar denemeler başarısız oldu."
    if code == 400: return "Gemini isteği geçersiz. Model veya istek yapılandırmasını kontrol edin."
    return f"Gemini API hatası: {error}"

def parse_with_gemini(file, api_key):
    started = time.perf_counter()
    try: image = prepare_image(file.getvalue())
    except Exception as e: return {"_dosya":file.name,"_time":0.0,"Nr Faktury":"Hata","Not":f"Görsel okunamadı: {e}"}
    client = get_gemini_client(api_key)
    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        max_output_tokens=500,
        thinking_config=types.ThinkingConfig(thinking_level="low")
    )
    try:
        response = client.models.generate_content(model=MODEL_NAME, contents=[PROMPT,image], config=config)
        parsed = extract_json(response.text)
    except Exception as primary_error:
        if error_code(primary_error) != 503:
            return {"_dosya":file.name,"_time":round(time.perf_counter()-started,2),"Nr Faktury":"Hata","Not":friendly_api_error(primary_error)}
        try:
            response = client.models.generate_content(model=FALLBACK_MODEL, contents=[PROMPT,image], config=config)
            parsed = extract_json(response.text)
        except Exception as fallback_error:
            return {"_dosya":file.name,"_time":round(time.perf_counter()-started,2),"Nr Faktury":"Hata","Not":friendly_api_error(fallback_error)}
    parsed["_dosya"], parsed["_time"] = file.name, round(time.perf_counter()-started,2)
    return validate_and_flag(parsed)

def create_professional_excel(data_list):
    wb = Workbook(); ws = wb.active; ws.title = "Resmi_Muhasebe_Defteri"
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    warning_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    header_font = Font(color="FFFFFF",bold=True)
    thin = Border(left=Side(style="thin"),right=Side(style="thin"),top=Side(style="thin"),bottom=Side(style="thin"))
    center = Alignment(horizontal="center",vertical="center",wrap_text=True)
    headers = ["Lp.","Data Wystawienia","Data Sprzedaży","Nr Faktury","Satıcı Unvanı","NIP Sprzedawcy","NIP Nabywcy","Netto (PLN)","Stawka VAT","Kwota VAT (PLN)","Wartość Brutto (PLN)","Sposób Płatności","Uwaga / Kontrola"]
    ws.append(headers)
    for c,h in enumerate(headers,1):
        cell=ws.cell(1,c); cell.fill=header_fill; cell.font=header_font; cell.alignment=center; cell.border=thin
        ws.column_dimensions[get_column_letter(c)].width=22
    ws.column_dimensions["A"].width=6; ws.column_dimensions["E"].width=30; ws.column_dimensions["M"].width=30; ws.freeze_panes="A2"
    for idx,data in enumerate(data_list,1):
        ws.append([idx,data.get("Data Wystawienia",""),data.get("Data Sprzedaży",""),data.get("Nr Faktury",""),data.get("Satıcı Unvanı",""),data.get("NIP Sprzedawcy",""),data.get("NIP Nabywcy",""),data.get("Netto (PLN)",0.0),data.get("Stawka VAT",""),data.get("Kwota VAT (PLN)",0.0),data.get("Wartość Brutto (PLN)",0.0),data.get("Sposób Płatności",""),data.get("Not","")])
        for c in range(1,len(headers)+1):
            cell=ws.cell(idx+1,c); cell.border=thin; cell.alignment=center
            if c in (8,10,11): cell.number_format='#,##0.00 "zł"'
            if str(data.get("Not","")).strip(): cell.fill=warning_fill
    output=io.BytesIO(); wb.save(output); return output.getvalue()

with st.sidebar:
    selected_lang = st.radio("🌍 Language / Język / Dil",["Türkçe","English","Polski"],index=0)
    ui = LANG_DICT[selected_lang]; st.divider(); st.header(ui["sidebar_title"])
    api_key_input = get_api_key()
    if api_key_input: st.success(ui["api_success"])
    else:
        api_key_input = st.text_input(ui["api_label"],type="password")
        if api_key_input: st.success(ui["api_success"])
        st.markdown(f"[{ui['api_link']}](https://aistudio.google.com/app/apikey)")

st.title(ui["title"]); st.markdown(f"**{ui['subtitle']}**"); st.info(ui["warning_yellow"])
uploaded_files = st.file_uploader(ui["upload_label"],type=["png","jpg","jpeg"],accept_multiple_files=True)

if uploaded_files:
    if not api_key_input: st.warning(ui["warning_api"])
    else:
        all_parsed_data=[]; progress_bar=st.progress(0); status_text=st.empty(); total=len(uploaded_files)
        for i,f in enumerate(uploaded_files):
            status_text.text(f"{ui['progress']}: {i+1} / {total} ({f.name})")
            all_parsed_data.append(parse_with_gemini(f,api_key_input)); progress_bar.progress((i+1)/total)
        status_text.empty(); st.success(ui["success_msg"])
        timings=[d["_time"] for d in all_parsed_data if d.get("_time",0)>0]
        if timings: st.caption(f"⏱️ {ui['timing']}: {sum(timings)/len(timings):.2f} s")
        st.subheader(ui["preview"])
        df_preview=pd.DataFrame([{k:v for k,v in d.items() if not k.startswith("_")} for d in all_parsed_data])
        df_preview.insert(0,"Dosya",[d["_dosya"] for d in all_parsed_data])
        def highlight_notes(row):
            return ["background-color: #FFF2CC; color: black"]*len(row) if str(row.get("Not","")).strip() else [""]*len(row)
        st.dataframe(df_preview.style.apply(highlight_notes,axis=1),use_container_width=True)
        st.download_button(label=ui["download"],data=create_professional_excel(all_parsed_data),file_name="invoice_accounting_report.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
