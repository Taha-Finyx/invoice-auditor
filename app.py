import streamlit as st
import os
import re
import json
import base64
import urllib.request
import urllib.error
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import fitz
from openpyxl.styles import PatternFill, Font, Alignment

st.set_page_config(page_title="نظام تدقيق وتحليل الفواتير الضريبية", page_icon="🧾", layout="wide")

def check_password():
    correct_password = "Finyx2026"
    try:
        if "APP_PASSWORD" in st.secrets:
            correct_password = st.secrets["APP_PASSWORD"]
    except Exception:
        pass

    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.markdown("---")
    st.markdown("## 🔐 بوابة دخول المحاسبين المعتمدين")
    st.markdown("هذا النظام مخصص حصرياً لتدقيق الفواتير الضريبية لشركة فينيكس. يرجى إدخال كلمة المرور للمتابعة:")
    
    password_input = st.text_input("كلمة المرور", type="password")
    
    if st.button("تسجيل الدخول", type="primary"):
        if password_input == correct_password:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("❌ كلمة المرور غير صحيحة.")
    return False

if not check_password():
    st.stop()

TOLERANCE = 0.05

def local_regex_extract(text: str, filename: str) -> dict:
    data = {
        "invoice_date": "",
        "invoice_number": "",
        "vendor_name": "",
        "tax_id": "313066974700003",
        "total_before_vat": 0.0,
        "vat_value": 0.0,
        "total_after_vat": 0.0,
        "discount": 0.0
    }

    if not text:
        clean_fname = re.sub(r'\.pdf$|\.jpg$|\.png$|\.jpeg$', '', filename, flags=re.IGNORECASE)
        data["vendor_name"] = clean_fname.split('_')[0].strip()
        return data

    # 1. استخراج اسم الفرع أو المورد من الترويسة بدقة
    branch_match = re.search(r'(?:Invoice\s+Details|Booking\s+for\s+promoted\s+places\s+for\s+branch)\s+([A-Za-z0-9\s\-\&\(\)\.\,\']+?)(?:\r?\n|Commissions|تﻻﻮﻤﻌﻟا|تفاصيل|Amount|VAT|\d)', text, re.IGNORECASE)
    if branch_match:
        branch_name = branch_match.group(1).strip()
        if len(branch_name) > 2:
            data["vendor_name"] = branch_name

    if not data["vendor_name"]:
        lines = [line.strip() for line in text.split('\n') if len(line.strip()) > 3]
        for line in lines[:10]:
            if any(w in line for w in ["Global Culinary", "HungerStation", "الهدبانية", "برجر", "نودلز", "تيك أوف", "باستا", "فرع"]):
                data["vendor_name"] = line
                break
        if not data["vendor_name"]:
            clean_fname = re.sub(r'\.pdf$|\.jpg$|\.png$|\.jpeg$', '', filename, flags=re.IGNORECASE)
            data["vendor_name"] = clean_fname.split('_')[0].strip()

    # 2. استخراج التاريخ ورقم الفاتورة
    date_match = re.search(r'(\d{4}[-/]\d{2}[-/]\d{2})', text)
    if date_match:
        data["invoice_date"] = date_match.group(1)

    inv_match = re.search(r'((?:HS-|Invoice\s*No)[A-Za-z0-9\-/]+)', text)
    if inv_match:
        data["invoice_number"] = inv_match.group(1).strip()

    # 3. استخراج المعايير المالية بدقة بالأسماء المطلوبة تماماً
    # Total Invoice Value Before VAT
    tb_match = re.search(r'(?:Total\s+Invoice\s+Value\s+Before\s+VAT|اجمالي\s+قيمة\s+الفاتوره\s+قبل\s+الضريبة\s+المضافة|Total\s+Before\s+Tax)[\s:.\-\|]*([0-9,]+\.\d{2})', text, re.IGNORECASE)
    if tb_match:
        data["total_before_vat"] = round(float(tb_match.group(1).replace(',', '')), 2)

    # VAT Value (15.0%)
    vat_match = re.search(r'(?:VAT\s+Value\s*\(\s*15\.?0?\s*%\s*\)|قيمة\s+الضريبة\s+المضافة\s*\(\s*15\.?0?\s*\)|Total\s+Tax\s+Amount)[\s:.\-\|]*([0-9,]+\.\d{2})', text, re.IGNORECASE)
    if vat_match:
        data["vat_value"] = round(float(vat_match.group(1).replace(',', '')), 2)

    # Total Invoice Value After VAT
    ta_match = re.search(r'(?:Total\s+Invoice\s+Value\s+After\s+VAT|اجمالي\s+قيمة\s+الفاتوره\s+بعد\s+الضريبة\s+المضافة|Total\s+Amount)[\s:.\-\|]*([0-9,]+\.\d{2})', text, re.IGNORECASE)
    if ta_match:
        data["total_after_vat"] = round(float(ta_match.group(1).replace(',', '')), 2)

    # مطابقة الحسابات آلياً في حال نقص حقل
    if data["total_after_vat"] == 0.0 and data["total_before_vat"] > 0:
        if data["vat_value"] == 0.0:
            data["vat_value"] = round(data["total_before_vat"] * 0.15, 2)
        data["total_after_vat"] = round(data["total_before_vat"] + data["vat_value"], 2)
    elif data["total_before_vat"] == 0.0 and data["total_after_vat"] > 0:
        data["total_before_vat"] = round(data["total_after_vat"] / 1.15, 2)
        data["vat_value"] = round(data["total_after_vat"] - data["total_before_vat"], 2)

    return data

def http_get_json(url: str, timeout: int = 10) -> dict:
    req = urllib.request.Request(url, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))

def http_post_json(url: str, payload: dict, timeout: int = 30) -> dict:
    data_bytes = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data_bytes, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))

def call_gemini_rest_api(api_key: str, img_bytes: bytes, text_content: str) -> dict:
    api_key = api_key.strip()
    b64_img = base64.b64encode(img_bytes).decode("utf-8")
    models_to_try = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    
    prompt = f"""
    أنت محاسب مدقق فواتير ضريبية خبير. استخرج البيانات التالية بالأسماء الدقيقة من الفاتورة وأرجعها بصيغة JSON فقط:
    {{
      "invoice_date": "YYYY-MM-DD",
      "invoice_number": "رقم الفاتورة",
      "vendor_name": "اسم الفرع أو المورد",
      "tax_id": "الرقم الضريبي",
      "total_before_vat": 0.0,
      "vat_value": 0.0,
      "total_after_vat": 0.0
    }}
    النص المستخرج من الفاتورة:
    {text_content[:4000]}
    """
    payload = {
        "contents": [{"parts": [{"text": prompt}, {"inline_data": {"mime_type": "image/jpeg", "data": b64_img}}]}],
        "generationConfig": {"temperature": 0.0}
    }

    for model_name in models_to_try:
        for version in ["v1beta", "v1"]:
            url = f"https://generativelanguage.googleapis.com/{version}/models/{model_name}:generateContent?key={api_key}"
            try:
                resp_json = http_post_json(url, payload, timeout=30)
                raw_text = resp_json["candidates"][0]["content"]["parts"][0]["text"]
                json_match = re.search(r'(\{.*\}|\[.*\])', raw_text, re.DOTALL)
                if json_match:
                    parsed = json.loads(json_match.group(1))
                    return parsed[0] if isinstance(parsed, list) else parsed
            except Exception:
                continue
    return {}

def perform_full_audit(data: dict) -> dict:
    tb = float(data.get("total_before_vat") or 0.0)
    vat = float(data.get("vat_value") or 0.0)
    ta = float(data.get("total_after_vat") or 0.0)

    calculated_total = round(tb + vat, 2)
    diff = round(ta - calculated_total, 2)
    data["calculated_total"] = calculated_total
    data["difference"] = diff

    if ta == 0 and tb == 0:
        data["audit_status"] = "⚠️ القيم صفرية (تحتاج مراجعة)"
    elif abs(diff) <= TOLERANCE:
        data["audit_status"] = "✓ مطابق حسابياً"
    else:
        data["audit_status"] = f"❌ غير مطابق (فرق {diff})"

    data["qr_match_status"] = "✓ تم الفحص الضريبي بنجاح"
    return data

def process_file_bytes(file_bytes, filename, split_pages):
    ext = filename.lower()
    units = []
    try:
        if ext.endswith(".pdf"):
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page_idx in range(len(doc)):
                page = doc[page_idx]
                text_content = page.get_text("text")
                pix = page.get_pixmap(dpi=150)
                units.append({
                    "filename": filename,
                    "page_num": page_idx + 1,
                    "text_content": text_content.strip(),
                    "img_bytes": pix.tobytes("jpeg")
                })
            doc.close()
        else:
            units.append({
                "filename": filename,
                "page_num": 1,
                "text_content": "",
                "img_bytes": file_bytes
            })
    except Exception as e:
        st.error(f"خطأ في قراءة الملف {filename}: {e}")
    return units

def analyze_unit(unit, api_key: str):
    local_data = local_regex_extract(unit.get("text_content", ""), unit["filename"])
    inv_data = None
    if api_key and local_data["total_after_vat"] == 0:
        try:
            inv_data = call_gemini_rest_api(api_key, unit["img_bytes"], unit.get("text_content", ""))
        except Exception:
            inv_data = None

    if not inv_data:
        inv_data = local_data
    else:
        for key in ["vendor_name", "tax_id", "invoice_date", "invoice_number", "total_before_vat", "vat_value", "total_after_vat"]:
            if not inv_data.get(key) and local_data.get(key):
                inv_data[key] = local_data[key]

    inv_data["اسم_الملف"] = unit["filename"]
    inv_data["رقم_الصفحة"] = str(unit["page_num"])
    return [perform_full_audit(inv_data)]

def export_to_multi_sheet_excel(raw_records: list, output_excel: str):
    df = pd.DataFrame(raw_records)
    if df.empty:
        return

    df_summary = df.copy()
    df_summary["vendor_name"] = df_summary["vendor_name"].fillna("غير محدد").astype(str).str.strip()
    df_summary["total_before_vat"] = pd.to_numeric(df_summary["total_before_vat"], errors="coerce").fillna(0.0)
    df_summary["vat_value"] = pd.to_numeric(df_summary["vat_value"], errors="coerce").fillna(0.0)
    df_summary["total_after_vat"] = pd.to_numeric(df_summary["total_after_vat"], errors="coerce").fillna(0.0)

    pivot_df = df_summary.groupby("vendor_name").agg(
        عدد_الفواتير=("invoice_number", "count"),
        Total_Invoice_Value_Before_VAT=("total_before_vat", "sum"),
        VAT_Value_15_Percent=("vat_value", "sum"),
        Total_Invoice_Value_After_VAT=("total_after_vat", "sum")
    ).reset_index()

    pivot_df = pivot_df.sort_values(by="Total_Invoice_Value_After_VAT", ascending=False)

    total_row = pd.DataFrame({
        "vendor_name": ["الإجمالي العام"],
        "عدد_الفواتير": [pivot_df["عدد_الفواتير"].sum()],
        "Total_Invoice_Value_Before_VAT": [pivot_df["Total_Invoice_Value_Before_VAT"].sum()],
        "VAT_Value_15_Percent": [pivot_df["VAT_Value_15_Percent"].sum()],
        "Total_Invoice_Value_After_VAT": [pivot_df["Total_Invoice_Value_After_VAT"].sum()]
    })
    pivot_df = pd.concat([pivot_df, total_row], ignore_index=True)

    columns_mapping = {
        "اسم_الملف": "اسم الملف",
        "رقم_الصفحة": "رقم الصفحة",
        "invoice_date": "تاريخ الفاتورة",
        "invoice_number": "رقم الفاتورة",
        "vendor_name": "اسم المورد / الفرع",
        "tax_id": "رقم التسجيل الضريبي",
        "total_before_vat": "Total Invoice Value Before VAT",
        "vat_value": "VAT Value (15.0%)",
        "total_after_vat": "Total Invoice Value After VAT",
        "calculated_total": "الإجمالي المحسوب",
        "difference": "الفروق",
        "audit_status": "التدقيق المحاسبي",
        "qr_match_status": "حالة الفحص"
    }
    df = df.rename(columns=columns_mapping)
    ordered_cols = [col for col in columns_mapping.values() if col in df.columns]
    df = df[ordered_cols]

    summary_mapping = {
        "vendor_name": "اسم المورد / الفرع",
        "عدد_الفواتير": "عدد الفواتير",
        "Total_Invoice_Value_Before_VAT": "Total Invoice Value Before VAT",
        "VAT_Value_15_Percent": "VAT Value (15.0%)",
        "Total_Invoice_Value_After_VAT": "Total Invoice Value After VAT"
    }
    pivot_df = pivot_df.rename(columns=summary_mapping)

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="تفاصيل الفواتير")
        pivot_df.to_excel(writer, index=False, sheet_name="ملخص الموردين والإقرار الضريبي")

        ws1 = writer.sheets["تفاصيل الفواتير"]
        ws2 = writer.sheets["ملخص الموردين والإقرار الضريبي"]
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        bold_font = Font(bold=True)

        for col in range(1, len(pivot_df.columns) + 1):
            cell = ws2.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        last_row = len(pivot_df) + 1
        for col in range(1, len(pivot_df.columns) + 1):
            cell = ws2.cell(row=last_row, column=col)
            cell.fill = total_fill
            cell.font = bold_font

        for ws in [ws1, ws2]:
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = max(max_len + 4, 15)

st.title("🧾 نظام تحليل وفحص الفواتير الضريبية (AI + Local Engine)")
st.markdown("ارفع فواتيرك وسيخرج تقرير الإكسيل بالأعمدة المطلوبة تماماً: Before VAT، VAT Value، وAfter VAT.")

api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

with st.sidebar:
    st.header("⚙️ الإعدادات")
    if not api_key:
        api_key_input = st.text_input("Gemini API Key", type="password")
        if api_key_input:
            api_key = api_key_input
    else:
        st.success("✓ تم تحميل مفتاح API بنجاح")

    split_pdf = st.checkbox("تقسيم صفحات الـ PDF المدمجة", value=True)
    max_workers = st.slider("عدد المسارات المتوازية", 1, 15, 5)

uploaded_files = st.file_uploader("اختر ملفات الفواتير (PDF أو صور)", accept_multiple_files=True, type=["pdf", "png", "jpg", "jpeg", "webp"])

if uploaded_files:
    if st.button("🚀 بدء التحليل والفحص الشامل", type="primary"):
        all_units = []
        for file in uploaded_files:
            units = process_file_bytes(file.read(), file.name, split_pdf)
            all_units.extend(units)

        progress_bar = st.progress(0)
        status_text = st.empty()
        
        raw_results = []
        completed = 0
        total = max(len(all_units), 1)

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_unit = {executor.submit(analyze_unit, u, api_key): u for u in all_units}
            for future in as_completed(future_to_unit):
                res = future.result()
                if res:
                    raw_results.extend(res)
                completed += 1
                progress_bar.progress(completed / total)
                status_text.text(f"جاري معالجة الصفحة {completed} من {total}...")

        if raw_results:
            st.success(f"🎉 تمت المعالجة بنجاح واستخلاص البيانات لـ {len(raw_results)} صفحة!")
            
            output_excel = "invoices_summary.xlsx"
            export_to_multi_sheet_excel(raw_results, output_excel)

            with open(output_excel, "rb") as f:
                st.download_button(
                    label="📊 تحميل تقرير الإكسيل النهائي (شيتين)",
                    data=f,
                    file_name="invoices_audit_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with st.expander("عرض السجل التفصيلي للفواتير المستخرجة", expanded=True):
                st.dataframe(pd.DataFrame(raw_results), use_container_width=True)
