import os
import re
import json
import base64
import fitz  # PyMuPDF
import cv2
import numpy as np
import pandas as pd
from typing import List, Optional
import google.generativeai as genai
from concurrent.futures import ThreadPoolExecutor, as_completed
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment

st.set_page_config(page_title="نظام تدقيق وتحليل الفواتير الضريبية", page_icon="🧾", layout="wide")

# ==========================================
# نظام حماية بكلمة المرور للمحاسبين
# ==========================================
def check_password():
    correct_password = ""
    try:
        correct_password = st.secrets.get("APP_PASSWORD", "Finyx2026")
    except Exception:
        correct_password = "Finyx2026"

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

# ==========================================
# الكشف التلقائي عن أحدث موديل متاح في حسابك
# ==========================================
@st.cache_resource
def get_working_model_name(api_key_str: str) -> str:
    genai.configure(api_key=api_key_str)
    try:
        available_models = [
            m.name for m in genai.list_models() 
            if "generateContent" in m.supported_generation_methods
        ]
        # ترتيب الأولوية لأحدث وأسرع النماذج التي تدعم الصور والنصوص
        preferred_keywords = ["2.5-flash", "2.0-flash", "flash", "pro"]
        for keyword in preferred_keywords:
            for model_name in available_models:
                if keyword in model_name.lower() and "vision" not in model_name.lower():
                    return model_name
        if available_models:
            return available_models[0]
    except Exception:
        pass
    return "models/gemini-2.5-flash"

# ==========================================
# دوال تحليل QR والتدقيق المحاسبي
# ==========================================
TOLERANCE = 0.05

def decode_tlv_base64(b64_string: str) -> dict:
    try:
        raw_bytes = base64.b64decode(b64_string.strip())
        idx = 0
        tlv_data = {}
        while idx < len(raw_bytes):
            tag = raw_bytes[idx]
            length = raw_bytes[idx + 1]
            value_bytes = raw_bytes[idx + 2 : idx + 2 + length]
            if tag in (1, 2, 3, 4, 5):
                tlv_data[tag] = value_bytes.decode("utf-8", errors="ignore")
            idx += 2 + length
        if not tlv_data:
            return None
        return {
            "qr_vendor": tlv_data.get(1, ""),
            "qr_tax_id": tlv_data.get(2, ""),
            "qr_date": tlv_data.get(3, ""),
            "qr_total": float(tlv_data.get(4, 0.0) or 0.0),
            "qr_tax": float(tlv_data.get(5, 0.0) or 0.0),
        }
    except Exception:
        return None

def extract_qr_from_image(img_bgr) -> dict:
    if img_bgr is None:
        return None
    detector = cv2.QRCodeDetector()
    for scale in [1.0, 1.5, 0.75]:
        resized = cv2.resize(img_bgr, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC) if scale != 1.0 else img_bgr
        data, _, _ = detector.detectAndDecode(resized)
        if data:
            decoded_tlv = decode_tlv_base64(data)
            if decoded_tlv:
                return decoded_tlv
            return {"qr_raw": data, "qr_total": None, "qr_tax": None, "qr_tax_id": ""}
    return None

def perform_full_audit(data: dict, qr_info: dict = None) -> dict:
    if qr_info is None:
        qr_info = data.get("_qr_info")

    subtotal = float(data.get("subtotal") or 0.0)
    discount = float(data.get("discount") or 0.0)
    tax = float(data.get("tax_amount") or 0.0)
    extracted_total = float(data.get("total_amount") or 0.0)

    calculated_total = round(subtotal - discount + tax, 2)
    diff = round(extracted_total - calculated_total, 2)
    data["calculated_total"] = calculated_total
    data["difference"] = diff

    if extracted_total == 0 and subtotal == 0:
        data["audit_status"] = "⚠️ خطأ: القيم صفرية"
    elif abs(diff) <= TOLERANCE:
        data["audit_status"] = "✓ مطابق حسابياً"
    else:
        data["audit_status"] = f"❌ غير مطابق (فرق {diff})"

    if not qr_info:
        data["qr_total"] = ""
        data["qr_tax"] = ""
        data["qr_match_status"] = "⚪ بدون QR"
    elif qr_info.get("qr_total") is not None:
        qr_total = round(qr_info["qr_total"], 2)
        qr_tax = round(qr_info["qr_tax"], 2)
        data["qr_total"] = qr_total
        data["qr_tax"] = qr_tax

        total_match = abs(extracted_total - qr_total) <= TOLERANCE
        tax_match = abs(tax - qr_tax) <= TOLERANCE

        if total_match and tax_match:
            data["qr_match_status"] = "✓ مطابق للـ QR (100%)"
        else:
            data["qr_match_status"] = "❌ اختلاف مع الـ QR"
    else:
        data["qr_total"] = ""
        data["qr_tax"] = ""
        data["qr_match_status"] = "ℹ️ QR غير مشفر"

    return data

def process_file_bytes(file_bytes, filename, split_pages):
    ext = filename.lower()
    units = []
    try:
        if ext.endswith(".pdf"):
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            for page_idx in range(len(doc)):
                page = doc[page_idx]
                text_content = page.get_text()
                
                pix = page.get_pixmap(dpi=150)
                img_np = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                img_bgr = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR if pix.n == 3 else cv2.COLOR_RGBA2BGR)
                _, buffer = cv2.imencode(".jpg", img_bgr, [int(cv2.IMWRITE_JPEG_QUALITY), 85])
                
                units.append({
                    "filename": filename,
                    "page_num": page_idx + 1,
                    "text_content": text_content.strip(),
                    "img_bytes": buffer.tobytes(),
                    "img_bgr": img_bgr
                })
            doc.close()
        else:
            img_array = np.frombuffer(file_bytes, np.uint8)
            img_bgr = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
            units.append({
                "filename": filename,
                "page_num": 1,
                "text_content": "",
                "img_bytes": file_bytes,
                "img_bgr": img_bgr
            })
    except Exception as e:
        st.error(f"خطأ في قراءة الملف {filename}: {e}")
    return units

def analyze_unit(unit, active_model_name):
    qr_info = extract_qr_from_image(unit["img_bgr"])
    model = genai.GenerativeModel(active_model_name)
    
    text_hint = f"\nالنص المستخرج من الصفحة للاستعانة به:\n{unit['text_content']}" if unit.get("text_content") else ""
    
    prompt = f"""
    أنت محاسب خبير ومدقق فواتير ضريبية. قم بتحليل صورة الفاتورة المرفقة (والنص المرفق إن وجد) واستخرج البيانات بدقة على شكل كائن JSON صحيح يحتوي على المفاتيح الآتية فقط:
    {{
      "invoice_date": "تاريخ الفاتورة بصيغة YYYY-MM-DD",
      "invoice_number": "رقم الفاتورة",
      "vendor_name": "اسم المورد أو الشركة",
      "tax_id": "رقم التسجيل الضريبي إن وجد",
      "subtotal": 0.0,
      "discount": 0.0,
      "tax_amount": 0.0,
      "total_amount": 0.0
    }}
    إذا لم تجد قيمة لحقل رقمي ضع 0.0، وإذا لم تجد نصاً ضع "". أجب بـ JSON فقط بدون أي نص إضافي.
    {text_hint}
    """

    contents = [{"mime_type": "image/jpeg", "data": unit["img_bytes"]}, prompt]

    try:
        response = model.generate_content(contents)
        
        if not response.text:
            return []

        text_res = response.text.strip()
        # استخراج كائن JSON بنجاح حتى لو وجد نص حوله
        json_match = re.search(r'(\{.*\}|\[.*\])', text_res, re.DOTALL)
        if json_match:
            text_res = json_match.group(1)
            
        parsed_data = json.loads(text_res.strip())
        invoices_list = parsed_data if isinstance(parsed_data, list) else [parsed_data]
        
        results = []
        for inv in invoices_list:
            inv["اسم_الملف"] = unit["filename"]
            inv["رقم_الصفحة"] = str(unit["page_num"])
            audited = perform_full_audit(inv, qr_info if len(invoices_list) == 1 else None)
            results.append(audited)
        return results
    except Exception as e:
        return [{
            "اسم_الملف": unit["filename"],
            "رقم_الصفحة": str(unit["page_num"]),
            "invoice_date": "-",
            "invoice_number": "-",
            "vendor_name": f"خطأ معالجة: {str(e)}",
            "tax_id": "-",
            "subtotal": 0.0,
            "discount": 0.0,
            "tax_amount": 0.0,
            "total_amount": 0.0,
            "calculated_total": 0.0,
            "difference": 0.0,
            "audit_status": "❌ خطأ في القراءة",
            "qr_match_status": "⚪ بدون QR"
        }]

def export_to_multi_sheet_excel(raw_records: list, output_excel: str):
    df = pd.DataFrame(raw_records)
    if df.empty:
        return

    df_summary = df.copy()
    df_summary["vendor_name"] = df_summary["vendor_name"].fillna("غير محدد").astype(str).str.strip()
    df_summary["total_amount"] = pd.to_numeric(df_summary["total_amount"], errors="coerce").fillna(0.0)
    df_summary["tax_amount"] = pd.to_numeric(df_summary["tax_amount"], errors="coerce").fillna(0.0)
    df_summary["subtotal"] = pd.to_numeric(df_summary["subtotal"], errors="coerce").fillna(0.0)
    df_summary["discount"] = pd.to_numeric(df_summary.get("discount", 0.0), errors="coerce").fillna(0.0)

    pivot_df = df_summary.groupby("vendor_name").agg(
        عدد_الفواتير=("invoice_number", "count"),
        إجمالي_قبل_الضريبة=("subtotal", "sum"),
        إجمالي_الخصم=("discount", "sum"),
        إجمالي_الضريبة=("tax_amount", "sum"),
        الإجمالي_الشامل=("total_amount", "sum")
    ).reset_index()

    pivot_df = pivot_df.sort_values(by="الإجمالي_الشامل", ascending=False)

    total_row = pd.DataFrame({
        "vendor_name": ["الإجمالي العام"],
        "عدد_الفواتير": [pivot_df["عدد_الفواتير"].sum()],
        "إجمالي_قبل_الضريبة": [pivot_df["إجمالي_قبل_الضريبة"].sum()],
        "إجمالي_الخصم": [pivot_df["إجمالي_الخصم"].sum()],
        "إجمالي_الضريبة": [pivot_df["إجمالي_الضريبة"].sum()],
        "الإجمالي_الشامل": [pivot_df["الإجمالي_الشامل"].sum()]
    })
    pivot_df = pd.concat([pivot_df, total_row], ignore_index=True)

    columns_mapping = {
        "اسم_الملف": "اسم الملف",
        "رقم_الصفحة": "رقم الصفحة",
        "invoice_date": "تاريخ الفاتورة",
        "invoice_number": "رقم الفاتورة",
        "vendor_name": "اسم المورد",
        "tax_id": "رقم التسجيل الضريبي",
        "subtotal": "القيمة قبل الخصم",
        "discount": "الخصم",
        "tax_amount": "الضريبة",
        "total_amount": "الإجمالي",
        "calculated_total": "الإجمالي المحسوب",
        "difference": "الفروق",
        "audit_status": "التدقيق المحاسبي",
        "qr_match_status": "مطابقة QR"
    }
    df = df.rename(columns=columns_mapping)
    ordered_cols = [col for col in columns_mapping.values() if col in df.columns]
    df = df[ordered_cols]

    summary_mapping = {
        "vendor_name": "اسم المورد",
        "عدد_الفواتير": "عدد الفواتير",
        "إجمالي_قبل_الضريبة": "إجمالي قبل الضريبة",
        "إجمالي_الخصم": "إجمالي الخصم",
        "إجمالي_الضريبة": "إجمالي ضريبة القيمة المضافة",
        "الإجمالي_الشامل": "الإجمالي النهائي شامل الضريبة"
    }
    pivot_df = pivot_df.rename(columns=summary_mapping)

    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="تفاصيل الفواتير")
        ws1 = writer.sheets["تفاصيل الفواتير"]

        pivot_df.to_excel(writer, index=False, sheet_name="ملخص الموردين والإقرار الضريبي")
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

# ==========================================
# واجهة المستخدم عبر Streamlit
# ==========================================
st.title("🧾 نظام تحليل وفحص الفواتير الضريبية (AI + QR)")
st.markdown("ارفع فواتيرك (PDF أو صور) وسيقوم النظام باستخراج البيانات وتدقيقها محاسبياً وتصدير تقرير إكسيل احترافي بشيتين.")

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
        st.success("✓ تم تحميل مفتاح API بنجاح من أسرار المنصة")

    split_pdf = st.checkbox("تقسيم صفحات الـ PDF المدمجة", value=True)
    max_workers = st.slider("عدد المسارات المتوازية", 1, 15, 5)

uploaded_files = st.file_uploader("اختر ملفات الفواتير (PDF أو صور)", accept_multiple_files=True, type=["pdf", "png", "jpg", "jpeg", "webp"])

if uploaded_files and api_key:
    genai.configure(api_key=api_key)
    active_model = get_working_model_name(api_key)
    st.sidebar.info(f"🤖 الموديل النشط: `{active_model}`")

    if st.button("🚀 بدء التحليل والفحص الشامل", type="primary"):
        all_units = []
        for file in uploaded_files:
            units = process_file_bytes(file.read(), file.name, split_pdf)
            all_units.extend(units)

        progress_bar = st.progress(0)
        status_text = st.empty()
        
        raw_results = []
        completed = 0
        total = len(all_units)

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_unit = {executor.submit(analyze_unit, u, active_model): u for u in all_units}
            for future in as_completed(future_to_unit):
                res = future.result()
                if res:
                    raw_results.extend(res)
                completed += 1
                progress_bar.progress(completed / total)
                status_text.text(f"جاري معالجة الصفحة {completed} من {total}...")

        if raw_results:
            st.success(f"🎉 تمت المعالجة بنجاح واستخلاص البيانات لـ {len(raw_results)} وحدة!")
            
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
        else:
            st.warning("لم يتم استخراج أي بيانات. يرجى التأكد من وضوح الفواتير المرفوعة.")
elif uploaded_files and not api_key:
    st.error("يرجى إدخال مفتاح الـ API في الشريط الجانبي أو إعداده في الـ Secrets.")
