import streamlit as st
from google import genai
import pandas as pd
from PIL import Image
import json
import re
import io
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Kebab King AI", page_icon="🧾", layout="wide")

LANG_DICT = {
    "Türkçe": {
        "title": "🧾 King Akıllı Muhasebe (JPK)",
        "subtitle": "Faturaları yükleyin; yapay zeka JPK_V7 uyumlu resmi Excel raporunuzu hazırlasın.",
        "sidebar_title": "⚙️ Sistem Durumu",
        "api_label": "Gemini API Anahtarı:",
        "api_success": "✅ Anahtar oturum için hazır.",
        "api_link": "Ücretsiz API Anahtarı Alın",
        "upload_label": "Faturaları Seçin veya Sürükleyin (PNG, JPG)",
        "warning_api": "Sistem bağlantısı bekleniyor...",
        "warning_yellow": "Sarı ile işaretlenen satırlar bir uyarı içerir (eksik NIP, matematiksel hata) - elle kontrol edin.",
        "progress": "İşlenen fatura",
        "success_msg": "Tüm faturalar başarıyla okundu!",
        "preview": "Kurumsal Veri Önizlemesi",
        "download": "📊 Resmi Muhasebe Excel Raporunu İndir"
    },
    "English": {
        "title": "🧾 King Smart Accounting (JPK)",
        "subtitle": "Upload invoices; AI will generate your JPK_V7 compliant official Excel report.",
        "sidebar_title": "⚙️ System Status",
        "api_label": "Gemini API Key:",
        "api_success": "✅ Key is ready for this session.",
        "api_link": "Get a Free API Key",
        "upload_label": "Drag and drop or select invoices (PNG, JPG)",
        "warning_api": "Awaiting system connection...",
        "warning_yellow": "Rows highlighted in yellow contain warnings (missing NIP, math errors) - verify manually.",
        "progress": "Processed invoice",
        "success_msg": "All invoices read successfully!",
        "preview": "Corporate Data Preview",
        "download": "📊 Download Official Excel Report"
    },
    "Polski": {
        "title": "🧾 Inteligentna Księgowość King (JPK)",
        "subtitle": "Prześlij faktury; AI wygeneruje oficjalny raport Excel zgodny z JPK_V7.",
        "sidebar_title": "⚙️ Status Systemu",
        "api_label": "Klucz API Gemini:",
        "api_success": "✅ Klucz jest gotowy do sesji.",
        "api_link": "Pobierz darmowy klucz API",
        "upload_label": "Wybierz lub przeciągnij faktury (PNG, JPG)",
        "warning_api": "Oczekiwanie na połączenie z systemem...",
        "warning_yellow": "Wiersze zaznaczone na żółto zawierają ostrzeżenia (brak NIP, błędy matematyczne) - sprawdź ręcznie.",
        "progress": "Przetworzona faktura",
        "success_msg": "Wszystkie faktury zostały odczytane!",
        "preview": "Podgląd Danych Firmowych",
        "download": "📊 Pobierz Oficjalny Raport Excel"
    }
}

MAX_REASONABLE_AMOUNT = 100_000

PROMPT = """
Sen Polonya resmi muhasebe ve vergi (JPK_V7) mevzuatında uzmanlaşmış kıdemli bir yapay zeka asistanısın.
Ekli fatura görselini titizlikle incele ve tüm resmi muhasebe parametrelerini SADECE aşağıdaki JSON formatında döndür.
Matematiksel tutarlılığı kontrol et (Brutto = Netto + VAT). NIP numaralarında tire/boşluk bırakma, sadece rakamları al.

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
  "Not": "Tutar uyuşmazlığı veya belirsiz alan varsa kısa muhasebesel uyarı yaz, yoksa boş bırak"
}
"""

def extract_json(text):
    text = text.strip()
    try: return json.loads(text)
    except: pass
    fenced = re.sub(r'```(?:json)?\n?(.*?)\n?```', r'\1', text, flags=re.DOTALL).strip()
    try: return json.loads(fenced)
    except: pass
    start, end = text.find('{'), text.rfind('}')
    if start != -1 and end != -1 and end > start:
        return json.loads(text[start:end + 1])
    raise ValueError("Yanıt içinden geçerli bir JSON çıkarılamadı")

def validate_and_flag(data):
    notes = []
    existing = str(data.get("Not", "") or "").strip()
    if existing: notes.append(existing)

    netto, vat, brutto = data.get("Netto (PLN)"), data.get("Kwota VAT (PLN)"), data.get("Wartość Brutto (PLN)")
    if all(isinstance(x, (int, float)) for x in (netto, vat, brutto)):
        if abs((netto + vat) - brutto) > 0.05:
            notes.append(f"Matematiksel tutarsızlık: Netto+VAT ({netto + vat:.2f}) ≠ Brutto ({brutto:.2f})")
        if brutto > MAX_REASONABLE_AMOUNT:
            notes.append("Anormal yüksek tutar")

    for field, label in [("NIP Sprzedawcy", "Satıcı NIP"), ("NIP Nabywcy", "Alıcı NIP")]:
        nip = str(data.get(field, "") or "")
        digits = re.sub(r'\D', '', nip)
        if nip and len(digits) != 10: notes.append(f"{label} 10 hane değil")

    data["Not"] = " | ".join(n for n in notes if n)
    return data

def create_professional_excel(data_list):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resmi_Muhasebe_Defteri"
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_aligned = Alignment(horizontal="center", vertical="center", wrap_text=True)

    headers = ["Lp.", "Data Wystawienia", "Data Sprzedaży", "Nr Faktury", "Satıcı Unvanı", "NIP Sprzedawcy", "NIP Nabywcy", "Netto (PLN)", "Stawka VAT", "Kwota VAT (PLN)", "Wartość Brutto (PLN)", "Sposób Płatności", "Uwaga / Kontrola"]
    ws.append(headers)

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, center_aligned, thin_border
        ws.column_dimensions[get_column_letter(col_num)].width = 22

    ws.column_dimensions['A'].width, ws.column_dimensions['E'].width, ws.column_dimensions['M'].width = 6, 30, 30
    ws.freeze_panes = "A2"

    for idx, data in enumerate(data_list, start=1):
        ws.append([
            idx, data.get("Data Wystawienia", ""), data.get("Data Sprzedaży", ""), data.get("Nr Faktury", ""),
            data.get("Satıcı Unvanı", ""), data.get("NIP Sprzedawcy", ""), data.get("NIP Nabywcy", ""),
            data.get("Netto (PLN)", 0.0), data.get("Stawka VAT", ""), data.get("Kwota VAT (PLN)", 0.0),
            data.get("Wartość Brutto (PLN)", 0.0), data.get("Sposób Płatności", ""), data.get("Not", "")
        ])
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=idx + 1, column=col_num)
            cell.border, cell.alignment = thin_border, center_aligned
            if col_num in [8, 10, 11]: cell.number_format = '#,##0.00 "zł"'
            if str(data.get("Not", "")).strip():
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

def parse_with_gemini(file, api_key, max_retries=2):
    last_error = None
    for attempt in range(max_retries + 1):
        try:
            image = Image.open(file)
            client = genai.Client(api_key=api_key)
            response = client.models.generate_content(model='gemini-3.6-flash', contents=[PROMPT, image])
            parsed = extract_json(response.text)
            parsed["_dosya"] = file.name
            return validate_and_flag(parsed)
        except Exception as e:
            last_error = e
            if attempt < max_retries:
                time.sleep(2 * (attempt + 1))
                continue
    return {"_dosya": file.name, "Nr Faktury": "Hata", "Not": f"API Hatası: {last_error}"}

with st.sidebar:
    selected_lang = st.radio("🌍 Language / Język / Dil", ["Türkçe", "English", "Polski"], index=0)
    ui = LANG_DICT[selected_lang]
    st.divider()
    st.header(ui["sidebar_title"])
    
    # API Anahtarını otomatik olarak Bulut'tan (Secrets) çeker
    if "GEMINI_API_KEY" in st.secrets:
        api_key_input = st.secrets["GEMINI_API_KEY"]
        st.success("✅ Kurumsal Sistem Bağlantısı Aktif.")
    else:
        api_key_input = st.text_input(ui["api_label"], type="password")
        if api_key_input:
            st.success(ui["api_success"])
        st.markdown(f"[{ui['api_link']}](https://aistudio.google.com/app/apikey)")

st.title(ui["title"])
st.markdown(f"**{ui['subtitle']}**")
st.info(ui["warning_yellow"])

uploaded_files = st.file_uploader(ui["upload_label"], type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files:
    if not api_key_input:
        st.warning(ui["warning_api"])
    else:
        all_parsed_data = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        total_files, completed = len(uploaded_files), 0

        with ThreadPoolExecutor(max_workers=2) as executor:
            futures = {executor.submit(parse_with_gemini, f, api_key_input): f for f in uploaded_files}
            for future in as_completed(futures):
                result = future.result()
                all_parsed_data.append(result)
                completed += 1
                progress_bar.progress(completed / total_files)
                status_text.text(f"{ui['progress']}: {completed} / {total_files} ({result['_dosya']})")

        status_text.empty()
        st.success(ui["success_msg"])

        st.subheader(ui["preview"])
        df_preview = pd.DataFrame([{k: v for k, v in d.items() if not k.startswith("_")} for d in all_parsed_data])
        df_preview.insert(0, "Dosya", [d["_dosya"] for d in all_parsed_data])

        def highlight_notes(row):
            note = str(row.get("Not", ""))
            return ['background-color: #FFF2CC; color: black'] * len(row) if note.strip() else [''] * len(row)

        st.dataframe(df_preview.style.apply(highlight_notes, axis=1), use_container_width=True)

        st.download_button(
            label=ui["download"],
            data=create_professional_excel(all_parsed_data),
            file_name="invoice_accounting_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
