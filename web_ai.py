import streamlit as st
from google import genai
import pandas as pd
from PIL import Image
import json
import re
import io
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# STREAMLIT CONFIG
# =========================================================

st.set_page_config(
    page_title="Kebab King AI",
    page_icon="🧾",
    layout="wide"
)


# =========================================================
# LANGUAGE
# =========================================================

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
        "download": "📊 Resmi Muhasebe Excel Raporunu İndir",
        "processing": "Faturalar paralel olarak işleniyor...",
        "done": "Tamamlanan"
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
        "download": "📊 Download Official Excel Report",
        "processing": "Invoices are being processed in parallel...",
        "done": "Completed"
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
        "download": "📊 Pobierz Oficjalny Raport Excel",
        "processing": "Faktury są przetwarzane równolegle...",
        "done": "Ukończono"
    }
}


# =========================================================
# CONSTANTS
# =========================================================

MAX_REASONABLE_AMOUNT = 100_000

# Çok büyük parallel değerleri rate-limit'e sokmamak için
MAX_WORKERS = 4

# OCR için yeterli, gereksiz büyük görselleri küçült
MAX_IMAGE_SIZE = 2200

MODEL_NAME = "gemini-3.6-flash"


# =========================================================
# PROMPT
# =========================================================

PROMPT = """
Sen Polonya resmi muhasebe ve vergi (JPK_V7) mevzuatında uzmanlaşmış kıdemli
bir yapay zeka asistanısın.

Ekli fatura görselini titizlikle incele ve tüm resmi muhasebe parametrelerini
SADECE aşağıdaki JSON formatında döndür.

Matematiksel tutarlılığı kontrol et:
Brutto = Netto + VAT

NIP numaralarında tire/boşluk bırakma, sadece rakamları al.

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


# =========================================================
# GEMINI CLIENT
# =========================================================

@st.cache_resource
def get_gemini_client(api_key: str):
    """
    Client sadece bir kez oluşturulur.
    Streamlit rerun olduğunda tekrar initialize edilmez.
    """
    return genai.Client(api_key=api_key)


# =========================================================
# JSON PARSER
# =========================================================

def extract_json(text):
    if not text:
        raise ValueError("Gemini boş yanıt döndürdü.")

    text = text.strip()

    # Direkt JSON
    try:
        return json.loads(text)
    except Exception:
        pass

    # Markdown fenced JSON
    fenced = re.sub(
        r"```(?:json)?\s*(.*?)\s*```",
        r"\1",
        text,
        flags=re.DOTALL
    ).strip()

    try:
        return json.loads(fenced)
    except Exception:
        pass

    # JSON object bul
    start = text.find("{")
    end = text.rfind("}")

    if start != -1 and end != -1 and end > start:
        try:
            return json.loads(text[start:end + 1])
        except Exception:
            pass

    raise ValueError("Yanıt içinden geçerli bir JSON çıkarılamadı.")


# =========================================================
# VALIDATION
# =========================================================

def validate_and_flag(data):
    notes = []

    existing = str(data.get("Not", "") or "").strip()

    if existing:
        notes.append(existing)

    netto = data.get("Netto (PLN)")
    vat = data.get("Kwota VAT (PLN)")
    brutto = data.get("Wartość Brutto (PLN)")

    if all(isinstance(x, (int, float)) for x in (netto, vat, brutto)):

        if abs((netto + vat) - brutto) > 0.05:
            notes.append(
                f"Matematiksel tutarsızlık: "
                f"Netto+VAT ({netto + vat:.2f}) ≠ "
                f"Brutto ({brutto:.2f})"
            )

        if brutto > MAX_REASONABLE_AMOUNT:
            notes.append("Anormal yüksek tutar")

    # NIP validation
    for field, label in [
        ("NIP Sprzedawcy", "Satıcı NIP"),
        ("NIP Nabywcy", "Alıcı NIP")
    ]:

        nip = str(data.get(field, "") or "")
        digits = re.sub(r"\D", "", nip)

        if nip and len(digits) != 10:
            notes.append(f"{label} 10 hane değil")

        # Temiz NIP'i kaydet
        if nip:
            data[field] = digits

    data["Not"] = " | ".join(n for n in notes if n)

    return data


# =========================================================
# IMAGE OPTIMIZATION
# =========================================================

def prepare_image(file_bytes):
    """
    Çok büyük invoice görsellerini küçültür.
    Böylece Gemini'ye gereksiz büyük image gönderilmez.
    """

    image = Image.open(io.BytesIO(file_bytes))

    # EXIF rotation düzeltmeye yardımcı olur
    try:
        from PIL import ImageOps
        image = ImageOps.exif_transpose(image)
    except Exception:
        pass

    # RGB
    if image.mode != "RGB":
        image = image.convert("RGB")

    # Resize
    image.thumbnail(
        (MAX_IMAGE_SIZE, MAX_IMAGE_SIZE),
        Image.Resampling.LANCZOS
    )

    return image


# =========================================================
# SINGLE INVOICE PARSER
# =========================================================

def parse_with_gemini(file_name, file_bytes, client, max_retries=2):

    last_error = None

    image = None

    try:
        image = prepare_image(file_bytes)
    except Exception as e:
        return {
            "_dosya": file_name,
            "Nr Faktury": "Hata",
            "Not": f"Görsel okunamadı: {e}"
        }

    for attempt in range(max_retries + 1):

        try:

            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=[
                    PROMPT,
                    image
                ]
            )

            parsed = extract_json(response.text)

            parsed["_dosya"] = file_name

            return validate_and_flag(parsed)

        except Exception as e:

            last_error = e

            # Son deneme değilse bekle
            if attempt < max_retries:

                # Exponential-ish backoff
                wait_time = 2 ** attempt

                import time
                time.sleep(wait_time)

    return {
        "_dosya": file_name,
        "Nr Faktury": "Hata",
        "Not": f"API Hatası: {last_error}"
    }


# =========================================================
# FILE HASH
# =========================================================

def get_file_hash(file_bytes):

    return hashlib.md5(file_bytes).hexdigest()


# =========================================================
# EXCEL
# =========================================================

def create_professional_excel(data_list):

    wb = Workbook()

    ws = wb.active
    ws.title = "Resmi_Muhasebe_Defteri"

    header_fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    warning_fill = PatternFill(
        start_color="FFF2CC",
        end_color="FFF2CC",
        fill_type="solid"
    )

    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    center_aligned = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    headers = [
        "Lp.",
        "Data Wystawienia",
        "Data Sprzedaży",
        "Nr Faktury",
        "Satıcı Unvanı",
        "NIP Sprzedawcy",
        "NIP Nabywcy",
        "Netto (PLN)",
        "Stawka VAT",
        "Kwota VAT (PLN)",
        "Wartość Brutto (PLN)",
        "Sposób Płatności",
        "Uwaga / Kontrola"
    ]

    ws.append(headers)

    for col_num, header in enumerate(headers, 1):

        cell = ws.cell(
            row=1,
            column=col_num
        )

        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_aligned
        cell.border = thin_border

        ws.column_dimensions[
            get_column_letter(col_num)
        ].width = 22

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["E"].width = 30
    ws.column_dimensions["M"].width = 30

    ws.freeze_panes = "A2"

    for idx, data in enumerate(data_list, start=1):

        ws.append([
            idx,
            data.get("Data Wystawienia", ""),
            data.get("Data Sprzedaży", ""),
            data.get("Nr Faktury", ""),
            data.get("Satıcı Unvanı", ""),
            data.get("NIP Sprzedawcy", ""),
            data.get("NIP Nabywcy", ""),
            data.get("Netto (PLN)", 0.0),
            data.get("Stawka VAT", ""),
            data.get("Kwota VAT (PLN)", 0.0),
            data.get("Wartość Brutto (PLN)", 0.0),
            data.get("Sposób Płatności", ""),
            data.get("Not", "")
        ])

        for col_num in range(1, len(headers) + 1):

            cell = ws.cell(
                row=idx + 1,
                column=col_num
            )

            cell.border = thin_border
            cell.alignment = center_aligned

            if col_num in [8, 10, 11]:

                cell.number_format = '#,##0.00 "zł"'

            if str(data.get("Not", "")).strip():

                cell.fill = warning_fill

    output = io.BytesIO()

    wb.save(output)

    return output.getvalue()


# =========================================================
# UI - SIDEBAR
# =========================================================

with st.sidebar:

    selected_lang = st.radio(
        "🌍 Language / Język / Dil",
        ["Türkçe", "English", "Polski"],
        index=0
    )

    ui = LANG_DICT[selected_lang]

    st.divider()

    st.header(ui["sidebar_title"])

    if "GEMINI_API_KEY" in st.secrets:

        api_key_input = st.secrets["GEMINI_API_KEY"]

        st.success(
            "✅ Kurumsal Sistem Bağlantısı Aktif."
        )

    else:

        api_key_input = st.text_input(
            ui["api_label"],
            type="password"
        )

        if api_key_input:

            st.success(
                ui["api_success"]
            )

        st.markdown(
            f"[{ui['api_link']}]"
            "(https://aistudio.google.com/app/apikey)"
        )


# =========================================================
# MAIN UI
# =========================================================

st.title(ui["title"])

st.markdown(
    f"**{ui['subtitle']}**"
)

st.info(
    ui["warning_yellow"]
)


uploaded_files = st.file_uploader(
    ui["upload_label"],
    type=["png", "jpg", "jpeg"],
    accept_multiple_files=True
)


# =========================================================
# PROCESSING
# =========================================================

if uploaded_files:

    if not api_key_input:

        st.warning(
            ui["warning_api"]
        )

    else:

        # -------------------------------------------------
        # Session state
        # -------------------------------------------------

        if "invoice_cache" not in st.session_state:
            st.session_state.invoice_cache = {}

        client = get_gemini_client(api_key_input)

        # -------------------------------------------------
        # Prepare files
        # -------------------------------------------------

        files_to_process = []

        all_parsed_data = []

        for file in uploaded_files:

            file_bytes = file.getvalue()

            file_hash = get_file_hash(file_bytes)

            cache_key = f"{file.name}_{file_hash}"

            # Önceden işlendi mi?
            if cache_key in st.session_state.invoice_cache:

                all_parsed_data.append(
                    st.session_state.invoice_cache[cache_key]
                )

            else:

                files_to_process.append(
                    (
                        file.name,
                        file_bytes,
                        cache_key
                    )
                )

        # -------------------------------------------------
        # Parallel processing
        # -------------------------------------------------

        total_new = len(files_to_process)

        if total_new > 0:

            st.subheader(ui["processing"])

            progress_bar = st.progress(0)

            status_text = st.empty()

            completed = 0

            # Worker sayısını dosya sayısından büyük yapma
            worker_count = min(
                MAX_WORKERS,
                total_new
            )

            with ThreadPoolExecutor(
                max_workers=worker_count
            ) as executor:

                future_map = {}

                for file_name, file_bytes, cache_key in files_to_process:

                    future = executor.submit(
                        parse_with_gemini,
                        file_name,
                        file_bytes,
                        client
                    )

                    future_map[future] = (
                        file_name,
                        cache_key
                    )

                for future in as_completed(future_map):

                    file_name, cache_key = future_map[future]

                    try:

                        result = future.result()

                    except Exception as e:

                        result = {
                            "_dosya": file_name,
                            "Nr Faktury": "Hata",
                            "Not": f"Beklenmeyen hata: {e}"
                        }

                    # Cache
                    st.session_state.invoice_cache[
                        cache_key
                    ] = result

                    all_parsed_data.append(result)

                    completed += 1

                    progress_bar.progress(
                        completed / total_new
                    )

                    status_text.text(
                        f"{ui['done']}: "
                        f"{completed} / {total_new} "
                        f"({file_name})"
                    )

            status_text.empty()

        # -------------------------------------------------
        # SUCCESS
        # -------------------------------------------------

        st.success(
            ui["success_msg"]
        )

        # -------------------------------------------------
        # Sort according to uploaded file order
        # -------------------------------------------------

        result_map = {
            d["_dosya"]: d
            for d in all_parsed_data
        }

        ordered_data = []

        for file in uploaded_files:

            if file.name in result_map:

                ordered_data.append(
                    result_map[file.name]
                )

        all_parsed_data = ordered_data

        # -------------------------------------------------
        # PREVIEW
        # -------------------------------------------------

        st.subheader(
            ui["preview"]
        )

        df_preview = pd.DataFrame([
            {
                k: v
                for k, v in d.items()
                if not k.startswith("_")
            }
            for d in all_parsed_data
        ])

        if not df_preview.empty:

            df_preview.insert(
                0,
                "Dosya",
                [
                    d["_dosya"]
                    for d in all_parsed_data
                ]
            )

            def highlight_notes(row):

                note = str(
                    row.get("Not", "")
                )

                if note.strip():

                    return [
                        "background-color: #FFF2CC; color: black"
                    ] * len(row)

                return [""] * len(row)

            st.dataframe(
                df_preview.style.apply(
                    highlight_notes,
                    axis=1
                ),
                use_container_width=True
            )

        # -------------------------------------------------
        # DOWNLOAD
        # -------------------------------------------------

        excel_data = create_professional_excel(
            all_parsed_data
        )

        st.download_button(
            label=ui["download"],
            data=excel_data,
            file_name="invoice_accounting_report.xlsx",
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument"
                ".spreadsheetml.sheet"
            )
        )
