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
    "Türkçe": {
        "title": "🧾 King Akıllı Muhasebe (JPK)",
        "subtitle": "Faturaları yükleyin; yapay zeka JPK_V7 uyumlu resmi Excel raporunuzu hazırlasın.",
        "sidebar_title": "⚙️ Sistem Durumu",
        "api_label": "Gemini API Anahtarı:",
        "api_success": "✅ Anahtar hazır.",
        "api_link": "Ücretsiz API Anahtarı Alın",
        "upload_label": "Faturaları seçin veya sürükleyin (PNG, JPG)",
        "warning_api": "Gemini API anahtarı bulunamadı.",
        "warning_yellow": "Sarı ile işaretlenen satırlar bir uyarı içerir (eksik NIP, matematiksel hata) - elle kontrol edin.",
        "progress": "İşlenen fatura",
        "success_msg": "Tüm faturalar başarıyla okundu!",
        "preview": "Kurumsal Veri Önizlemesi",
        "download": "📊 Resmi Muhasebe Excel Raporunu İndir",
        "timing": "Ortalama işlem süresi",
    },
    "English": {
        "title": "🧾 King Smart Accounting (JPK)",
        "subtitle": "Upload invoices; AI will generate your JPK_V7 compliant official Excel report.",
        "sidebar_title": "⚙️ System Status",
        "api_label": "Gemini API Key:",
        "api_success": "✅ Key is ready.",
        "api_link": "Get a Free API Key",
        "upload_label": "Drag and drop or select invoices (PNG, JPG)",
        "warning_api": "Gemini API key was not found.",
        "warning_yellow": "Rows highlighted in yellow contain warnings (missing NIP, math errors) - verify manually.",
        "progress": "Processed invoice",
        "success_msg": "All invoices read successfully!",
        "preview": "Corporate Data Preview",
        "download": "📊 Download Official Excel Report",
        "timing": "Average processing time",
    },
    "Polski": {
        "title": "🧾 Inteligentna Księgowość King (JPK)",
        "subtitle": "Prześlij faktury; AI wygeneruje oficjalny raport Excel zgodny z JPK_V7.",
        "sidebar_title": "⚙️ Status Systemu",
        "api_label": "Klucz API Gemini:",
        "api_success": "✅ Klucz jest gotowy.",
        "api_link": "Pobierz darmowy klucz API",
        "upload_label": "Wybierz lub przeciągnij faktury (PNG, JPG)",
        "warning_api": "Nie znaleziono klucza API Gemini.",
        "warning_yellow": "Wiersze zaznaczone na żółto zawierają ostrzeżenia (brak NIP, błędy matematyczne) - sprawdź ręcznie.",
        "progress": "Przetworzona faktura",
        "success_msg": "Wszystkie faktury zostały odczytane!",
        "preview": "Podgląd Danych Firmowych",
        "download": "📊 Pobierz Oficjalny Raport Excel",
        "timing": "Średni czas przetwarzania",
    },
}

MODEL_NAME = "gemini-3.8-flash"
FALLBACK_MODEL = "gemini-3.7-flash"
MAX_REASONABLE_AMOUNT = 100_000
MAX_IMAGE_SIZE = 1800

PROMPT = """
Bu görseldeki Polonya faturasını muhasebe veri çıkarımı için incele.

Kurallar:
- Yalnızca fatura üzerinde açıkça görülen bilgileri çıkar.
- Okunamayan veya bulunmayan alanları boş bırak.
- Tarihleri YYYY-MM-DD olarak döndür.
- NIP numaralarından boşluk, tire ve ülke kodunu çıkar; 10 rakam döndür.
- Netto, VAT ve Brutto tutarlarını faturadaki değerlerden al.
- Brutto = Netto + VAT kontrolünü yap; uyuşmazlık varsa Not alanına kısa bir uyarı yaz.
- Resmi fatura numarası ile el yazısı notları birbirine karıştırma.
"""

INVOICE_SCHEMA = {
    "type": "object",
    "properties": {
        "Data Wystawienia": {"type": "string"},
        "Data Sprzedaży": {"type": "string"},
        "Nr Faktury": {"type": "string"},
        "Satıcı Unvanı": {"type": "string"},
        "NIP Sprzedawcy": {"type": "string"},
        "NIP Nabywcy": {"type": "string"},
        "Netto (PLN)": {"type": ["number", "null"]},
        "Stawka VAT": {"type": "string"},
        "Kwota VAT (PLN)": {"type": ["number", "null"]},
        "Wartość Brutto (PLN)": {"type": ["number", "null"]},
        "Sposób Płatności": {"type": "string"},
        "Not": {"type": "string"},
    },
    "required": [
        "Data Wystawienia", "Data Sprzedaży", "Nr Faktury", "Satıcı Unvanı",
        "NIP Sprzedawcy", "NIP Nabywcy", "Netto (PLN)", "Stawka VAT",
        "Kwota VAT (PLN)", "Wartość Brutto (PLN)", "Sposób Płatności", "Not",
    ],
}

def get_api_key():
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
        if key:
            return str(key).strip()
    except Exception:
        pass
    return os.getenv("GEMINI_API_KEY", "").strip()

@st.cache_resource
def get_gemini_client(api_key):
    return genai.Client(
        api_key=api_key,
        http_options=types.HttpOptions(
            timeout=45_000,
            retry_options=types.HttpRetryOptions(
                attempts=3,
                initial_delay=1.0,
                max_delay=6.0,
                exp_base=2.0,
                jitter=0.2,
                http_status_codes=[408, 429, 500, 502, 503, 504],
            ),
        ),
    )

def validate_and_flag(data):
    notes = []
    existing = str(data.get("Not", "") or "").strip()
    if existing:
        notes.append(existing)

    netto = data.get("Netto (PLN)")
    vat = data.get("Kwota VAT (PLN)")
    brutto = data.get("Wartość Brutto (PLN)")

    if all(isinstance(x, (int, float)) for x in (netto, vat, brutto)):
        if abs(netto + vat - brutto) > 0.05:
            notes.append(
                f"Matematiksel tutarsızlık: Netto+VAT ({netto + vat:.2f}) "
                f"≠ Brutto ({brutto:.2f})"
            )
        if brutto > MAX_REASONABLE_AMOUNT:
            notes.append("Anormal yüksek tutar")

    for field, label in [("NIP Sprzedawcy", "Satıcı NIP"), ("NIP Nabywcy", "Alıcı NIP")]:
        nip = str(data.get(field, "") or "")
        digits = re.sub(r"\D", "", nip)
        if nip:
            data[field] = digits
        if nip and len(digits) != 10:
            notes.append(f"{label} 10 hane değil")

    data["Not"] = " | ".join(notes)
    return data

def prepare_image(file_bytes):
    image = Image.open(io.BytesIO(file_bytes))
    image = ImageOps.exif_transpose(image).convert("RGB")
    image.thumbnail((MAX_IMAGE_SIZE, MAX_IMAGE_SIZE), Image.Resampling.LANCZOS)
    return image

def error_code(error):
    return getattr(error, "status_code", None) or getattr(error, "code", None)

def friendly_api_error(error):
    code = error_code(error)
    if code in (401, 403):
        return "API anahtarı geçersiz veya yetkisiz. Streamlit Secrets içindeki GEMINI_API_KEY değerini kontrol edin."
    if code == 429:
        return "Gemini kullanım limiti aşıldı. Biraz sonra tekrar deneyin."
    if code == 503:
        return "Gemini servisi şu anda yoğun. Otomatik tekrar denemeleri başarısız oldu."
    if code in (408, 500, 502, 504):
        return "Gemini servisine geçici olarak ulaşılamıyor. Lütfen tekrar deneyin."
    if code == 400:
        return "Gemini isteği geçersiz. İstek/model yapılandırmasını kontrol edin."
    return f"Gemini API hatası: {error}"

def request_invoice(client, model, image):
    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_json_schema=INVOICE_SCHEMA,
        max_output_tokens=300,
        thinking_config=types.ThinkingConfig(thinking_level="low"),
    )
    response = client.models.generate_content(
        model=model,
        contents=[PROMPT, image],
        config=config,
    )
    parsed = response.parsed
    if parsed is None:
        raise ValueError("Gemini yapılandırılmış bir sonuç döndürmedi.")
    return dict(parsed) if not isinstance(parsed, dict) else parsed

def parse_with_gemini(file, api_key):
    started = time.perf_counter()
    try:
        image = prepare_image(file.getvalue())
    except Exception as e:
        return {"_dosya": file.name, "_time": 0.0, "Nr Faktury": "Hata", "Not": f"Görsel okunamadı: {e}"}

    client = get_gemini_client(api_key)

    try:
        parsed = request_invoice(client, MODEL_NAME, image)
    except Exception as primary_error:
        # SDK has already retried transient errors; fallback only for 503.
        if error_code(primary_error) != 503:
            return {"_dosya": file.name, "_time": round(time.perf_counter() - started, 2), "Nr Faktury": "Hata", "Not": friendly_api_error(primary_error)}
        try:
            parsed = request_invoice(client, FALLBACK_MODEL, image)
        except Exception as fallback_error:
            return {"_dosya": file.name, "_time": round(time.perf_counter() - started, 2), "Nr Faktury": "Hata", "Not": friendly_api_error(fallback_error)}

    parsed["_dosya"] = file.name
    parsed["_time"] = round(time.perf_counter() - started, 2)
    return validate_and_flag(parsed)

def create_professional_excel(data_list):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resmi_Muhasebe_Defteri"
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    warning_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    headers = ["Lp.", "Data Wystawienia", "Data Sprzedaży", "Nr Faktury", "Satıcı Unvanı", "NIP Sprzedawcy", "NIP Nabywcy", "Netto (PLN)", "Stawka VAT", "Kwota VAT (PLN)", "Wartość Brutto (PLN)", "Sposób Płatności", "Uwaga / Kontrola"]
    ws.append(headers)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(1, c)
        cell.fill, cell.font, cell.alignment, cell.border = header_fill, header_font, center, thin
        ws.column_dimensions[get_column_letter(c)].width = 22
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["E"].width = 30
    ws.column_dimensions["M"].width = 30
    ws.freeze_panes = "A2"

    for idx, data in enumerate(data_list, 1):
        ws.append([
            idx, data.get("Data Wystawienia", ""), data.get("Data Sprzedaży", ""), data.get("Nr Faktury", ""),
            data.get("Satıcı Unvanı", ""), data.get("NIP Sprzedawcy", ""), data.get("NIP Nabywcy", ""),
            data.get("Netto (PLN)") or 0.0, data.get("Stawka VAT", ""), data.get("Kwota VAT (PLN)") or 0.0,
            data.get("Wartość Brutto (PLN)") or 0.0, data.get("Sposób Płatności", ""), data.get("Not", "")
        ])
        for c in range(1, len(headers) + 1):
            cell = ws.cell(idx + 1, c)
            cell.border, cell.alignment = thin, center
            if c in (8, 10, 11):
                cell.number_format = '#,##0.00 "zł"'
            if str(data.get("Not", "")).strip():
                cell.fill = warning_fill

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

with st.sidebar:
    selected_lang = st.radio("🌍 Language / Język / Dil", ["Türkçe", "English", "Polski"], index=0)
    ui = LANG_DICT[selected_lang]
    st.divider()
    st.header(ui["sidebar_title"])
    api_key_input = get_api_key()
    if api_key_input:
        st.success(ui["api_success"])
    else:
        api_key_input = st.text_input(ui["api_label"], type="password")
        if api_key_input:
            st.success(ui["api_success"])
        st.markdown(f"[{ui['api_link']}](https://aistudio.google.com/app/apikey)")

st.title(ui["title"])
st.markdown(f"**{ui['subtitle']}**")
st.info(ui["warning_yellow"])

uploaded_files = st.file_uploader(ui["upload_label"], type=["png", "jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files:
    if not api_key_input:
        st.warning(ui["warning_api"])
    else:
        all_parsed_data = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        total = len(uploaded_files)

        for i, f in enumerate(uploaded_files):
            status_text.text(f"{ui['progress']}: {i + 1} / {total} ({f.name})")
            all_parsed_data.append(parse_with_gemini(f, api_key_input))
            progress_bar.progress((i + 1) / total)

        status_text.empty()
        st.success(ui["success_msg"])

        timings = [d["_time"] for d in all_parsed_data if d.get("_time", 0) > 0]
        if timings:
            st.caption(f"⏱️ {ui['timing']}: {sum(timings) / len(timings):.2f} s")

        st.subheader(ui["preview"])
        df_preview = pd.DataFrame([{k: v for k, v in d.items() if not k.startswith("_")} for d in all_parsed_data])
        df_preview.insert(0, "Dosya", [d["_dosya"] for d in all_parsed_data])

        def highlight_notes(row):
            note = str(row.get("Not", "") or "")
            return ["background-color: #FFF2CC; color: black"] * len(row) if note.strip() else [""] * len(row)

        st.dataframe(df_preview.style.apply(highlight_notes, axis=1), use_container_width=True)
        st.download_button(
            label=ui["download"],
            data=create_professional_excel(all_parsed_data),
            file_name="invoice_accounting_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
