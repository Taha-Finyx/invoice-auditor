import os
import re
import json
import base64
import urllib.request
import urllib.error
import fitz  # PyMuPDF
import cv2
import numpy as np
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import streamlit as st
from openpyxl.styles import PatternFill, Font, Alignment

st.set_page_config(page_title="نظام تدقيق وتحليل الفواتير الضريبية", page_icon="🧾", layout="wide")

# ==========================================
# 1. نظام حماية بكلمة المرور للمحاسبين
# ==========================================
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

# ==========================================
# 2. محرك قراءة QR الفاتورة الإلكترونية (ZATCA)
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
            "qr_date": tlv_data.get(3, "")[:10],
            "qr_total": float(re.sub(r"[^\d.]", "", tlv_data.get(4, "0")) or 0.0),
            "qr_tax": float(re.sub(r"[^\d.]", "", tlv_data.get(5, "0")) or 0.0),
        }
    except Exception:
        return None

def extract_qr_from_image(img_bgr) -> dict:
    if img_bgr is None:
        return None
    detector = cv2.QRCodeDetector()
    for scale in [1.0, 1.5, 2.0, 0.75]:
        try:
            resized = cv2.resize(img_bgr, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC) if scale != 1.0 else img_bgr
            data, _, _ = detector.detectAndDecode(resized)
            if not data:
                gray = cv2.cvtColor(resized, cv2.COLOR_BGR2GRAY)
                _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)
                data, _, _ = detector.detectAndDecode(thresh)
            if data:
                decoded_tlv = decode_tlv_base64(data)
                if decoded_tlv:
                    return decoded_tlv
                return {"qr_raw": data, "qr_total": None, "qr_tax": None, "qr_tax_id": ""}
        except Exception:
            continue
    return None

# ==========================================
# 3. محرك الاستخراج المحلي الذكي (Regex + QR)
# ==========================================
def local_regex_extract(text: str, filename: str, qr_info: dict = None) -> dict:
    data = {
        "invoice_date": "",
        "invoice_number": "",
        "vendor_name": "",
        "tax_id": "",
        "subtotal": 0.0,
        "discount": 0.0,
        "tax_amount": 0.0,
        "total_amount": 0.0
    }

    if qr_info and qr_info.get("qr_total") is not None:
        data["vendor_name"] = qr_info.get("qr_vendor", "")
        data["tax_id"] = qr_info.get("qr_tax_id", "")
        data["invoice_date"] = qr_info.get("qr_date", "")
        data["total_amount"] = round(float(qr_info.get("qr_total", 0.0)), 2)
        data["tax_amount"] = round(float(qr_info.get("qr_tax", 0.0)), 2)
        if data["total_amount"] > 0:
            data["subtotal"] = round(data["total_amount"] - data["tax_amount"], 2)

    if not text:
        return data

    if not data["tax_id"]:
        tax_match = re.search(r'\b(3\d{13}3|\d{3}-\d{3}-\d{3}|\d{15})\b', text)
        if tax_match:
            data["tax_id"] = tax_match.group(1)

    if not data["invoice_date"]:
        date_match = re.search(r'(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/]\d{2}[-/]\d{4})', text)
        if date_match:
            data["invoice_date"] = date_match.group(1)
        else:
            fname_date = re.search(r'(202\d)(\d{2})(\d{2})', filename)
            if fname_date:
                data["invoice_date"] = f"{fname_date.group(1)}-{fname_date.group(2)}-{fname_date.group(3)}"

    inv_match = re.search(r'(?:Invoice\s*(?:No|Number|#)|رقم\s*الفاتورة|فاتورة\s*رقم)[\s:.\-#]*([A-Za-z0-9\-/]+)', text, re.IGNORECASE)
    if inv_match:
        data["invoice_number"] = inv_match.group(1).strip()

    if not data["vendor_name"]:
        lines = [line.strip() for line in text.split('\n') if len(line.strip()) > 3]
        for line in lines[:6]:
            if any(w in line for w in ["شركة", "مؤسسة", "Company", "LLC", "Co.", "Burger", "برجر"]):
                data["vendor_name"] = line
                break
        if not data["vendor_name"]:
            clean_fname = re.sub(r'\.pdf$|\.jpg$|\.png$', '', filename, flags=re.IGNORECASE)
            data["vendor_name"] = clean_fname.split('_')[0].strip()

    if data["total_amount"] == 0.0:
        amounts = re.findall(r'\b\d{1,3}(?:,\d{3})*\.\d{2}\b', text)
        float_amounts = sorted(list(set([float(a.replace(',', '')) for a in amounts if float(a.replace(',', '')) > 0])))
        if len(float_amounts) >= 2:
            for total_cand in reversed(float_amounts):
                for sub_cand in float_amounts:
                    if sub_cand < total_cand:
                        diff_tax = round(total_cand - sub_cand, 2)
                        if abs(diff_tax - round(sub_cand * 0.15, 2)) <= 0.5 or abs(diff_tax - round(sub_cand * 0.14, 2)) <= 0.5 or diff_tax in float_amounts:
                            data["total_amount"] = total_cand
                            data["subtotal"] = sub_cand
                            data["tax_amount"] = diff_tax
                            break
                if data["total_amount"] > 0:
                    break
        elif len(float_amounts) == 1:
            data["total_amount"] = float_amounts[0]
            data["subtotal"] = round(float_amounts[0] / 1.15, 2)
            data["tax_amount"] = round(data["total_amount"] - data["subtotal"], 2)

    return data

# ==========================================
# 4. اتصال مباشر بـ Gemini باستخدام urllib المدمجة
# ==========================================
def http_get_json(url: str, timeout: int = 10) -> dict:
    req = urllib.request.Request(url, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))

def http_post_json(url: str, payload: dict, timeout: int = 25) -> dict:
    data_bytes = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data_bytes, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))

def call_gemini_rest_api(api_key: str, img_bytes: bytes, text_content: str) -> dict:
    api_key = api_key.strip()
    b64_img = base64.b64encode(img_bytes).decode("utf-8")

    models_to_try = []
    try:
        list_url = f"https://generativelanguage.googleapis.com/v1beta/models?key={api_key}"
        res_json = http_get_json(list_url, timeout=10)
        for m in res_json.get("models", []):
            if "generateContent" in m.get("supportedGenerationMethods", []):
                m_name = m["name"].replace("models/", "")
                if any(k in m_name for k in ["2.5-flash", "2.0-flash", "flash", "pro"]) and "vision" not in m_name:
                    models_to_try.append(m_name)
    except Exception:
        pass

    for fallback_m in ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]:
        if fallback_m not in models_to_try:
            models_to_try.append(fallback_m)

    prompt = f"""
    أنت مدقق فواتير ضريبية خبير. استخرج بيانات الفاتورة التالية وأرجع النتيجة بصيغة JSON فقط يحتوي على هذه المفاتيح:
    {{
      "invoice_date": "YYYY-MM-DD",
      "invoice_number": "رقم الفاتورة",
      "vendor_name": "اسم الشركة أو المورد",
      "tax_id": "الرقم الضريبي",
      "subtotal": 0.0,
      "discount": 0.0,
      "tax_amount": 0.0,
      "total_amount": 0.0
    }}
    النص المساعد المستخرج من الملف:
    {text_content[:3000] if text_content else "لا يوجد نص، اقرأ من الصورة المرفقة"}
    """

    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": "image/jpeg", "data": b64_img}}
            ]
        }],
        "generationConfig": {"temperature": 0.0}
    }

    last_error = ""
    for model_name in models_to_try:
        for version in ["v1beta", "v1"]:
            url = f"https://generativelanguage.googleapis.com/{version}/models/{model_name}:generateContent?key={api_key}"
            try:
                resp_json = http_post_json(url, payload, timeout=25)
                raw_text = resp_json["candidates"][0]["content"]["parts"][0]["text"]
                json_match = re.search(r'(\{.*\}|\[.*\])', raw_text, re.DOTALL)
                if json_match:
                    parsed = json.loads(json_match.group(1))
                    return parsed[0] if isinstance(parsed, list) else parsed
            except Exception as e:
                last_error = str(e)

    raise RuntimeError(last_error)

# ==========================================
# 5. التدقيق المحاسبي ومعالجة الملفات
# ==========================================
def perform_full_audit(data: dict, qr_info: dict = None) -> dict:
    subtotal = float(data.get("subtotal") or 0.0)
    discount = float(data.get("discount") or 0.0)
    tax = float(data.get("tax_amount") or 0.0)
    extracted_total = float(data.get("total_amount") or 0.0)

    calculated_total = round(subtotal - discount + tax, 2)
    diff = round(extracted_total - calculated_total, 2)
    data["calculated_total"] = calculated_total
    data["difference"] = diff

    if extracted_total == 0 and subtotal == 0:
        data["audit_status"] = "⚠️ القيم صفرية (تحتاج مراجعة)"
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
            data["qr_match_status"] = f"❌ اختلاف مع QR ({qr_total})"
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
                text_content = page.get_text("text")
                
                pix = page.get_pixmap(dpi=180)
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

def analyze_unit(unit, api_key: str):
    qr_info = extract_qr_from_image(unit["img_bgr"])
    
    inv_data = None
    if api_key:
        try:
            inv_data = call_gemini_rest_api(api_key, unit["img_bytes"], unit.get("text_content", ""))
        except Exception:
            inv_data = None

    local_data = local_regex_extract(unit.get("text_content", ""), unit["filename"], qr_info)
    
    if not inv_data or (float(inv_data.get("total_amount") or 0) == 0 and local_data["total_amount"] > 0):
        inv_data = local_data
    else:
        for key in ["vendor_name", "tax_id", "invoice_date", "invoice_number"]:
            if not inv_data.get(key) and local_data.get(key):
                inv_data[key] = local_data[key]

    inv_data["اسم_الملف"] = unit["filename"]
    inv_data["رقم_الصفحة"] = str(unit["page_num"])
    audited = perform_full_audit(inv_data, qr_info)
    return [audited]

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
# 6. واجهة المستخدم عبر Streamlit
# ==========================================
st.title("🧾 نظام تحليل وفحص الفواتير الضريبية (AI + QR + Local Engine)")
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
        api_key_input = st.text_input("Gemini API Key (اختياري مع المحرك المحلي)", type="password")
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
        total = len(all_units)

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
