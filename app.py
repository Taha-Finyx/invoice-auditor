import streamlit as st
import os
import re
import io
import gc
import tempfile
import zipfile
import json
import base64
import urllib.request
import urllib.error
from datetime import datetime
import pandas as pd
import numpy as np
from concurrent.futures import ThreadPoolExecutor, as_completed
import fitz  # PyMuPDF
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

try:
    import cv2
    HAS_CV2 = True
except Exception:
    HAS_CV2 = False

st.set_page_config(page_title="نظام تدقيق وتحليل المشتريات (يدعم +3,000 مستند)", page_icon="🧾", layout="wide")

# ==========================================
# 1. بوابة حماية كلمة المرور
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
    st.markdown("هذا النظام مخصص لتدقيق وتحليل فواتير المشتريات الضريبية. يرجى إدخال كلمة المرور:")
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
# 2. ثوابت وقواميس الموردين والعميل
# ==========================================
TOLERANCE = 0.05
CUSTOMER_VAT_NUMBERS = {"313066974700003"}

KNOWN_VENDORS_BY_VAT = {
    "300067077700003": "شركة مطاعم ركن هاشم",
    "301317075900003": "شركة الفطائر الاولي المحدودة",
    "300914908810003": "دار الضيافة",
    "300056457900003": "Herfy Food Services Co",
    "311131637400003": "مطابع جوهرة الخليج للطباعة",
    "311233377500003": "شركة سلاته لتقديم الوجبات",
    "300048882800003": "شركة باب العالمية للخدمات المتخصصة",
    "314127719200003": "شركة عرين التميز التجارية",
    "310131935400003": "شركة افضل الطهاة لتقديم الوجبات",
    "310069655110003": "شركة هنجرستيشن المحدودة",
    "310069655100003": "شركة هنجرستيشن المحدودة",
    "310191627410003": "شركة جاهز الدولية لتقنية نظم المعلومات",
    "311024805300003": "شركة اختصار الزمن لنقل الطرود",
    "310228833400003": "شركة تطبيق مرسول للاتصال وتقنية المعلومات",
    "310480075110003": "شركة نينجا فود",
    "302008852600003": "ارام ميم ذ م م",
    "300417884800003": "شركة مجموعة عبد العزيز محمد النملة القابضة",
    "300430327600003": "شركة مجموعة عبد العزيز محمد النملة القابضة",
}

# ==========================================
# 3. محرك فك تشفير وقراءة الـ QR Code (ZATCA TLV)
# ==========================================
def decode_tlv_base64(b64_string: str) -> dict:
    try:
        cleaned = b64_string.strip()
        raw_bytes = base64.b64decode(cleaned)
        idx = 0
        tlv_data = {}
        while idx + 1 < len(raw_bytes):
            tag = raw_bytes[idx]
            length = raw_bytes[idx + 1]
            if idx + 2 + length > len(raw_bytes):
                break
            value_bytes = raw_bytes[idx + 2 : idx + 2 + length]
            if tag in (1, 2, 3, 4, 5):
                val_str = value_bytes.decode("utf-8", errors="ignore").strip()
                val_str = val_str.replace("\u202b", "").replace("\u202c", "").replace("\u200f", "")
                tlv_data[tag] = val_str
            idx += 2 + length
        if not tlv_data or (2 not in tlv_data and 4 not in tlv_data):
            return None
        
        qr_date_raw = tlv_data.get(3, "")
        qr_date = qr_date_raw[:10] if len(qr_date_raw) >= 10 else qr_date_raw
        
        return {
            "qr_vendor": tlv_data.get(1, ""),
            "qr_tax_id": re.sub(r"[^\d]", "", tlv_data.get(2, "")),
            "qr_date": qr_date,
            "qr_total": float(re.sub(r"[^\d.]", "", tlv_data.get(4, "0")) or 0.0),
            "qr_tax": float(re.sub(r"[^\d.]", "", tlv_data.get(5, "0")) or 0.0),
            "qr_raw": cleaned
        }
    except Exception:
        return None

def try_decode_cv2_qr(img_bgr, fast_mode=False) -> dict:
    if not HAS_CV2 or img_bgr is None:
        return None
    try:
        detector = cv2.QRCodeDetector()
        bordered = cv2.copyMakeBorder(img_bgr, 30, 30, 30, 30, cv2.BORDER_CONSTANT, value=[255, 255, 255])
        gray = cv2.cvtColor(bordered, cv2.COLOR_BGR2GRAY)

        scales = [1.0, 1.5] if fast_mode else [1.0, 1.5, 2.0]
        thresholds = [None, 150] if fast_mode else [None, 150, 128]

        for scale in scales:
            try:
                resized = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC) if scale != 1.0 else gray
                for thresh_val in thresholds:
                    test_img = resized if thresh_val is None else cv2.threshold(resized, thresh_val, 255, cv2.THRESH_BINARY)[1]
                    data, _, _ = detector.detectAndDecode(test_img)
                    if data:
                        tlv = decode_tlv_base64(data)
                        if tlv:
                            return tlv
                        return {"qr_raw": data, "qr_total": None, "qr_tax": None, "qr_tax_id": ""}
            except Exception:
                continue
    except Exception:
        pass
    return None

def scan_pdf_for_qr(doc, max_pages_to_scan=2, fast_mode=False) -> dict:
    if not HAS_CV2:
        return None
    
    pages_idx = list(range(min(len(doc), max_pages_to_scan)))
    if len(doc) > max_pages_to_scan and (len(doc) - 1) not in pages_idx:
        pages_idx.append(len(doc) - 1)

    # 1. فحص الصور المدمجة داخل الـ PDF
    for p_idx in pages_idx:
        page = doc[p_idx]
        try:
            for img_info in page.get_images(full=True):
                xref = img_info[0]
                base_img = doc.extract_image(xref)
                img_bytes = base_img["image"]
                img_np = np.frombuffer(img_bytes, np.uint8)
                img_bgr = cv2.imdecode(img_np, cv2.IMREAD_COLOR)
                if img_bgr is not None and img_bgr.shape[0] > 20 and img_bgr.shape[1] > 20:
                    res = try_decode_cv2_qr(img_bgr, fast_mode=fast_mode)
                    if res:
                        return res
        except Exception:
            pass

    # 2. تصوير الصفحة
    for p_idx in pages_idx:
        page = doc[p_idx]
        try:
            pix = page.get_pixmap(dpi=150)
            img_np = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
            img_bgr = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR if pix.n == 3 else cv2.COLOR_RGBA2BGR)
            res = try_decode_cv2_qr(img_bgr, fast_mode=fast_mode)
            if res:
                return res
        except Exception:
            continue
    return None

# ==========================================
# 4. دوال استخراج البيانات والتواريخ
# ==========================================
def detect_document_type(text: str, filename: str) -> str:
    combined = f"{filename}\n{text[:3000]}".lower()
    if any(w in combined for w in ["credit note", "إشعار دائن", "اشعار دائن", "إشعار خصم", "اشعار خصم", "مرتجع", "credit memo"]):
        return "إشعار دائن"
    if any(w in combined for w in ["debit note", "إشعار مدين", "اشعار مدين", "إشعار مدينة", "اشعار مدينة", "debit memo"]):
        return "إشعار مدين"
    return "فاتورة ضريبية"

def normalize_date_str(date_str: str) -> str:
    if not date_str:
        return ""
    date_str = str(date_str).strip()
    m1 = re.search(r'(202\d)[-/](\d{1,2})[-/](\d{1,2})', date_str)
    if m1:
        return f"{m1.group(1)}-{int(m1.group(2)):02d}-{int(m1.group(3)):02d}"
    m2 = re.search(r'(\d{1,2})[-/\.](\d{1,2})[-/\.](202\d)', date_str)
    if m2:
        p1, p2, y = int(m2.group(1)), int(m2.group(2)), m2.group(3)
        d, m = (p2, p1) if p1 > 12 else (p1, p2)
        return f"{y}-{m:02d}-{d:02d}"
    return date_str

def extract_invoice_date(text: str, filename: str, qr_info: dict = None) -> str:
    if qr_info and qr_info.get("qr_date"):
        qd = normalize_date_str(qr_info["qr_date"])
        if qd.startswith("202"):
            return qd

    explicit_patterns = [
        r'(?:Invoice\s+Issue\s+Date|Invoice\s+Date|Issue\s+Date|Document\s+Date|تاريخ\s+اصدار\s+الفاتورة|تاريخ\s+الفاتورة)[\s\S]{1,45}?(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/\.]\d{2}[-/\.]\d{4})',
    ]
    for pat in explicit_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return normalize_date_str(m.group(1).strip())

    all_d = re.findall(r'(202[56][-/]\d{2}[-/]\d{2}|\d{2}[-/\.]\d{2}[-/\.]202[56])', text)
    if all_d:
        return normalize_date_str(all_d[0])

    fm = re.search(r'(202\d[-_]?\d{2}[-_]?\d{2})', filename)
    if fm:
        raw_fd = re.sub(r'[^\d]', '', fm.group(1))
        if len(raw_fd) == 8:
            return f"{raw_fd[:4]}-{raw_fd[4:6]}-{raw_fd[6:8]}"
    return ""

def extract_invoice_number(text: str, filename: str) -> str:
    inv_patterns = [
        r'(PSI/\d{4}/\d+)',
        r'(SA-B2B-INV-\d+)',
        r'(SINV\d+)',
        r'(HS-(?:CPC-|ADS-)?\d{4}/\d+-\d+)',
        r'(INV/\d{4}/\d{2}/\d+)',
        r'(INV/\d{4}/\d+)',
        r'Invoice\s*Number\s*[:\|]?\s*([A-Za-z0-9\-/]+)',
        r'INVOICE\s+(\d{3,10})\b',
    ]
    for pat in inv_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            cand = m.group(1).strip()
            if cand not in CUSTOMER_VAT_NUMBERS:
                return cand

    fname_patterns = [
        r'(PSI[_/]\d{4}[_/]\d+)',
        r'(INV[_/]\d{4}[_/]\d+)',
        r'(SA-B2B-INV-\d+)',
        r'(SINV\d+)',
        r'-(\d{7,10})\.pdf$',
    ]
    for pat in fname_patterns:
        m = re.search(pat, filename, re.IGNORECASE)
        if m:
            val = m.group(1).replace("_", "/")
            return val
    return ""

def extract_amounts(text: str, qr_info: dict = None) -> tuple:
    tb, vat, ta = 0.0, 0.0, 0.0

    if qr_info and qr_info.get("qr_total") is not None and qr_info.get("qr_total") > 0:
        ta = round(float(qr_info["qr_total"]), 2)
        vat = round(float(qr_info["qr_tax"]), 2)
        tb = round(ta - vat, 2)
        return (tb, vat, ta)

    tb_m = re.search(r'(?:Total\s+Invoice\s+(?:Value\s+)?Before\s+VAT|Total\s+Before\s+Tax|Total\s+SAR\s+Excl\.\s+VAT|Total\s+Amount\s+Excl\.\s+VAT|Total\s+Taxable\s+Amount|الإجمالي\s+الفرعي)[\s:.\-\|SRSAR]*([\d,]+\.\d{1,2})', text, re.IGNORECASE)
    vat_m = re.search(r'(?:VAT\s+Value\s*\(\s*15\.?0?\s*%\s*\)|VAT\s*15\.?0?%|Total\s+Tax\s+Amount|Total\s+VAT\s+Amount|VAT\s+Total|Total\s+VAT)[\s:.\-\|SRSAR]*([\d,]+\.\d{1,2})', text, re.IGNORECASE)
    ta_m = re.search(r'(?:Total\s+Invoice\s+(?:Value\s+)?After\s+VAT|Total\s+SAR\s+Incl\.\s+VAT|Total\s+Amount\s+Incl\.\s+VAT|Total\s+Amount\s+Due|Grand\s+Total|\bTOTAL\b)[\s:.\-\|SRSAR]*([\d,]+\.\d{1,2})', text, re.IGNORECASE)

    if tb_m:
        tb = float(tb_m.group(1).replace(',', ''))
    if vat_m:
        vat = float(vat_m.group(1).replace(',', ''))
    if ta_m:
        ta = float(ta_m.group(1).replace(',', ''))

    if ta > 0 and tb > 0 and abs((tb + vat) - ta) <= 0.5:
        return (round(tb, 2), round(vat, 2), round(ta, 2))

    amounts = re.findall(r'\b\d{1,3}(?:,\d{3})*\.\d{1,2}\b', text)
    floats = sorted(list(set([round(float(a.replace(',', '')), 2) for a in amounts if float(a.replace(',', '')) > 0])))
    for total_c in reversed(floats):
        for sub_c in reversed(floats):
            if sub_c < total_c:
                tax_c = round(total_c - sub_c, 2)
                if abs(tax_c - round(sub_c * 0.15, 2)) <= 0.15:
                    return (sub_c, tax_c, total_c)

    return (round(tb, 2), round(vat, 2), round(ta, 2))

def local_smart_extract(text: str, filename: str, qr_info: dict = None) -> dict:
    doc_type = detect_document_type(text, filename)
    
    tax_id = ""
    if qr_info and qr_info.get("qr_tax_id") and qr_info["qr_tax_id"] not in CUSTOMER_VAT_NUMBERS:
        tax_id = qr_info["qr_tax_id"]
    if not tax_id:
        all_vats = re.findall(r'\b(3\d{14})\b', text)
        vendor_vats = [v for v in all_vats if v not in CUSTOMER_VAT_NUMBERS]
        if vendor_vats:
            tax_id = vendor_vats[0]

    vendor_name = KNOWN_VENDORS_BY_VAT.get(tax_id, "")
    if not vendor_name and qr_info and qr_info.get("qr_vendor") and not qr_info["qr_vendor"].isdigit():
        vendor_name = qr_info["qr_vendor"]
    if not vendor_name:
        for line in [l.strip() for l in text.split('\n') if len(l.strip()) > 3][:10]:
            if any(k in line for k in ["شركة", "مؤسسة", "مطابع", "Co", "Company"]):
                vendor_name = line
                break
        if not vendor_name:
            vendor_name = re.sub(r'\.pdf$|\.jpg$|\.png$', '', filename, flags=re.IGNORECASE)

    inv_date = extract_invoice_date(text, filename, qr_info)
    inv_no = extract_invoice_number(text, filename)
    tb, vat, ta = extract_amounts(text, qr_info)

    return {
        "invoice_date": inv_date,
        "invoice_number": inv_no,
        "doc_type": doc_type,
        "vendor_name": vendor_name,
        "tax_id": tax_id,
        "total_before_vat": tb,
        "vat_value": vat,
        "total_after_vat": ta
    }

# ==========================================
# 5. دوال الذكاء الاصطناعي والتدقيق المحاسبي
# ==========================================
def http_post_json(url: str, payload: dict, timeout: int = 20) -> dict:
    data_bytes = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data_bytes, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))

def call_gemini_rest_api(api_key: str, img_bytes: bytes, text_content: str) -> dict:
    api_key = api_key.strip()
    b64_img = base64.b64encode(img_bytes).decode("utf-8") if img_bytes else ""
    
    prompt = """
    استخرج بيانات الفاتورة بصيغة JSON فقط:
    {
      "invoice_date": "YYYY-MM-DD",
      "invoice_number": "رقم الفاتورة",
      "doc_type": "فاتورة ضريبية أو إشعار دائن",
      "vendor_name": "اسم المورد",
      "tax_id": "الرقم الضريبي للمورد (15 رقماً)",
      "total_before_vat": 0.0,
      "vat_value": 0.0,
      "total_after_vat": 0.0
    }
    """
    parts = [{"text": prompt}]
    if b64_img:
        parts.append({"inline_data": {"mime_type": "image/jpeg", "data": b64_img}})
    parts.append({"text": text_content[:3000]})

    payload = {"contents": [{"parts": parts}], "generationConfig": {"temperature": 0.0}}

    for model_name in ["gemini-2.5-flash", "gemini-2.0-flash"]:
        for version in ["v1beta", "v1"]:
            url = f"https://generativelanguage.googleapis.com/{version}/models/{model_name}:generateContent?key={api_key}"
            try:
                resp_json = http_post_json(url, payload, timeout=15)
                raw_text = resp_json["candidates"][0]["content"]["parts"][0]["text"]
                json_match = re.search(r'(\{.*\}|\[.*\])', raw_text, re.DOTALL)
                if json_match:
                    parsed = json.loads(json_match.group(1))
                    return parsed[0] if isinstance(parsed, list) else parsed
            except Exception:
                continue
    return {}

def perform_audit_and_qr_match(data: dict, qr_info: dict = None) -> dict:
    tb = round(float(data.get("total_before_vat") or 0.0), 2)
    vat = round(float(data.get("vat_value") or 0.0), 2)
    ta = round(float(data.get("total_after_vat") or 0.0), 2)

    data["total_before_vat"] = tb
    data["vat_value"] = vat
    data["total_after_vat"] = ta

    calc_total = round(tb + vat, 2)
    diff = round(ta - calc_total, 2)
    data["calculated_total"] = calc_total
    data["difference"] = diff

    if ta == 0 and tb == 0:
        data["audit_status"] = "⚠️ القيم صفرية"
    elif abs(diff) <= TOLERANCE:
        data["audit_status"] = "✓ مطابق حسابياً"
    else:
        data["audit_status"] = f"❌ فرق ({diff})"

    if not qr_info:
        data["qr_status"] = "❌ غير موجود"
        data["qr_total"] = None
        data["qr_tax"] = None
    elif qr_info.get("qr_total") is not None:
        qr_total = round(float(qr_info["qr_total"]), 2)
        qr_tax = round(float(qr_info["qr_tax"]), 2)
        data["qr_total"] = qr_total
        data["qr_tax"] = qr_tax

        if abs(ta - qr_total) <= TOLERANCE and abs(vat - qr_tax) <= TOLERANCE:
            data["qr_status"] = "✓ مطابق 100%"
        else:
            data["qr_status"] = "⚠️ موجود ومختلف"
    else:
        data["qr_status"] = "ℹ️ موجود (غير مشفر)"
    return data

# ==========================================
# 6. البث المباشر ومعالجة الملفات من القرص
# ==========================================
def process_task_from_disk(task_info: tuple, split_branches: bool, api_key: str, fast_mode: bool) -> list:
    task_type, path_or_zip, internal_name, display_name = task_info
    try:
        if task_type == "zip":
            with zipfile.ZipFile(path_or_zip, "r") as zf:
                file_bytes = zf.read(internal_name)
        else:
            with open(path_or_zip, "rb") as fp:
                file_bytes = fp.read()

        return process_and_analyze_bytes(file_bytes, display_name, split_branches, api_key, fast_mode)
    except Exception as e:
        return [{
            "invoice_date": "", "invoice_number": "", "doc_type": "خطأ قراءة",
            "vendor_name": f"خطأ: {str(e)[:40]}", "tax_id": "",
            "total_before_vat": 0.0, "vat_value": 0.0, "total_after_vat": 0.0,
            "qr_status": "❌ خطأ", "audit_status": "❌ تلف", "اسم_الملف": display_name
        }]

def process_and_analyze_bytes(file_bytes: bytes, filename: str, split_branches: bool, api_key: str, fast_mode: bool) -> list:
    if filename.lower().endswith(".pdf"):
        pdf_start = file_bytes.find(b"%PDF-")
        if pdf_start > 0:
            file_bytes = file_bytes[pdf_start:]

    ext = filename.lower()
    results = []
    try:
        if ext.endswith(".pdf"):
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            qr_info = scan_pdf_for_qr(doc, max_pages_to_scan=2, fast_mode=fast_mode)

            combined_text = "\n".join([page.get_text("text") for page in doc]).strip()
            local_data = local_smart_extract(combined_text, filename, qr_info)

            img_bytes = doc[0].get_pixmap(dpi=110).tobytes("jpeg") if (api_key and local_data["total_after_vat"] == 0.0) else b""
            unit = {
                "filename": filename,
                "text_content": combined_text,
                "img_bytes": img_bytes,
                "qr_info": qr_info
            }
            results.extend(analyze_unit(unit, api_key))
            doc.close()
        else:
            qr_info = None
            if HAS_CV2:
                img_np = np.frombuffer(file_bytes, np.uint8)
                img_bgr = cv2.imdecode(img_np, cv2.IMREAD_COLOR)
                qr_info = try_decode_cv2_qr(img_bgr, fast_mode=fast_mode)
            unit = {
                "filename": filename,
                "text_content": "",
                "img_bytes": file_bytes if api_key else b"",
                "qr_info": qr_info
            }
            results.extend(analyze_unit(unit, api_key))
    except Exception as e:
        results.append({
            "invoice_date": "", "invoice_number": "", "doc_type": "خطأ قراءة",
            "vendor_name": f"خطأ: {str(e)[:40]}", "tax_id": "",
            "total_before_vat": 0.0, "vat_value": 0.0, "total_after_vat": 0.0,
            "qr_status": "❌ خطأ", "audit_status": "❌ تلف", "اسم_الملف": filename
        })
    return results

def analyze_unit(unit: dict, api_key: str) -> list:
    qr_info = unit.get("qr_info")
    local_data = local_smart_extract(unit.get("text_content", ""), unit["filename"], qr_info)

    ai_data = {}
    if api_key and local_data["total_after_vat"] == 0.0:
        try:
            ai_data = call_gemini_rest_api(api_key, unit["img_bytes"], unit.get("text_content", ""))
        except Exception:
            pass

    final_data = local_data.copy()
    if ai_data:
        for k in ["invoice_date", "invoice_number", "doc_type", "vendor_name", "tax_id"]:
            if not final_data.get(k) and ai_data.get(k):
                final_data[k] = str(ai_data[k]).strip()
        if final_data["total_after_vat"] == 0.0 and float(ai_data.get("total_after_vat") or 0.0) > 0:
            final_data["total_before_vat"] = float(ai_data.get("total_before_vat") or 0.0)
            final_data["vat_value"] = float(ai_data.get("vat_value") or 0.0)
            final_data["total_after_vat"] = float(ai_data.get("total_after_vat") or 0.0)

    if final_data.get("tax_id") in KNOWN_VENDORS_BY_VAT:
        final_data["vendor_name"] = KNOWN_VENDORS_BY_VAT[final_data["tax_id"]]

    final_data["اسم_الملف"] = unit["filename"]
    audited = perform_audit_and_qr_match(final_data, qr_info)
    return [audited]

# ==========================================
# 7. التصدير لملف إكسيل احترافي
# ==========================================
def export_to_excel(records: list, output_path: str):
    df = pd.DataFrame(records)
    if df.empty:
        return

    df_out = pd.DataFrame({
        "التاريخ": df["invoice_date"],
        "رقم الفاتورة": df["invoice_number"],
        "نوع المستند": df["doc_type"],
        "اسم المورد": df["vendor_name"],
        "رقم التسجيل الضريبي": df["tax_id"],
        "الاجمالي قبل الضريبية": pd.to_numeric(df["total_before_vat"], errors="coerce").fillna(0.0),
        "ضريبة القيمة المضافة (15%)": pd.to_numeric(df["vat_value"], errors="coerce").fillna(0.0),
        "الاجمالي بعد الضريبة": pd.to_numeric(df["total_after_vat"], errors="coerce").fillna(0.0),
        "حالة الـ QR Code": df["qr_status"],
        "التدقيق المحاسبي": df["audit_status"],
        "اسم الملف المرفوع": df["اسم_الملف"]
    })

    total_row = pd.DataFrame({
        "التاريخ": [""], "رقم الفاتورة": [""], "نوع المستند": [""],
        "اسم المورد": ["الإجمالي العام"], "رقم التسجيل الضريبي": [""],
        "الاجمالي قبل الضريبية": [round(df_out["الاجمالي قبل الضريبية"].sum(), 2)],
        "ضريبة القيمة المضافة (15%)": [round(df_out["ضريبة القيمة المضافة (15%)"].sum(), 2)],
        "الاجمالي بعد الضريبة": [round(df_out["الاجمالي بعد الضريبة"].sum(), 2)],
        "حالة الـ QR Code": [""], "التدقيق المحاسبي": [""], "اسم الملف المرفوع": [""]
    })
    df_sheet1 = pd.concat([df_out, total_row], ignore_index=True)

    df_summary = df_out.groupby(["اسم المورد", "رقم التسجيل الضريبي"], dropna=False).agg(
        عدد_المستندات=("رقم الفاتورة", "count"),
        الاجمالي_قبل_الضريبية=("الاجمالي قبل الضريبية", "sum"),
        ضريبة_القيمة_المضافة=("ضريبة القيمة المضافة (15%)", "sum"),
        الاجمالي_بعد_الضريبة=("الاجمالي بعد الضريبة", "sum")
    ).reset_index().sort_values(by="الاجمالي_بعد_الضريبة", ascending=False)

    df_summary = df_summary.rename(columns={
        "عدد_المستندات": "عدد الفواتير",
        "الاجمالي_قبل_الضريبية": "الاجمالي قبل الضريبية",
        "ضريبة_القيمة_المضافة": "ضريبة القيمة المضافة (15%)",
        "الاجمالي_بعد_الضريبة": "الاجمالي بعد الضريبة"
    })

    sum_total_row = pd.DataFrame({
        "اسم المورد": ["الإجمالي العام"], "رقم التسجيل الضريبي": [""],
        "عدد الفواتير": [df_summary["عدد الفواتير"].sum()],
        "الاجمالي قبل الضريبية": [round(df_summary["الاجمالي قبل الضريبية"].sum(), 2)],
        "ضريبة القيمة المضافة (15%)": [round(df_summary["ضريبة القيمة المضافة (15%)"].sum(), 2)],
        "الاجمالي بعد الضريبة": [round(df_summary["الاجمالي بعد الضريبة"].sum(), 2)]
    })
    df_sheet2 = pd.concat([df_summary, sum_total_row], ignore_index=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_sheet1.to_excel(writer, index=False, sheet_name="تحليل المشتريات")
        df_sheet2.to_excel(writer, index=False, sheet_name="ملخص الموردين")

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        bold_font = Font(bold=True, size=11)

        for sheet_name, data_df in [("تحليل المشتريات", df_sheet1), ("ملخص الموردين", df_sheet2)]:
            ws = writer.sheets[sheet_name]
            ws.views.sheetView[0].rightToLeft = True

            for col_idx in range(1, len(data_df.columns) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            last_row = len(data_df) + 1
            for c_idx in range(1, len(data_df.columns) + 1):
                c = ws.cell(row=last_row, column=c_idx)
                c.fill = total_fill
                c.font = bold_font

            for col in ws.columns:
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = 22

# ==========================================
# 8. واجهة المستخدم في Streamlit
# ==========================================
st.title("🧾 نظام تحليل المشتريات وتدقيق الفواتير الضريبية (مجهز لـ +3,000 مستند)")
st.markdown("ارفع الفواتير (سواء ملفات فردية أو **ملف مضغوط ZIP** يحتوي على الآلاف) لتبدأ المعالجة الفورية والتدقيق الضريبي.")

api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

with st.sidebar:
    st.header("⚙️ إعدادات المعالجة الضخمة")
    if not api_key:
        api_key_input = st.text_input("Gemini API Key (اختياري)", type="password")
        if api_key_input:
            api_key = api_key_input
    else:
        st.success("✓ مفتاح الذكاء الاصطناعي متصل")

    split_branches = st.checkbox("تقسيم صفحات الفروع الداخلية", value=False)
    fast_mode_toggle = st.checkbox("⚡ الوضع الفائق (Turbo)", value=True)
    max_workers = st.slider("عدد المسارات المتوازية (Workers)", 1, 32, 12)

uploaded_files = st.file_uploader(
    "ارفع ملف مضغوط (ZIP) أو ملفات (PDF / صور)",
    accept_multiple_files=True,
    type=["zip", "pdf", "png", "jpg", "jpeg", "webp"]
)

if uploaded_files:
    if st.button("🚀 بدء التحليل الشامل", type="primary"):
        temp_dir = tempfile.mkdtemp()
        task_queue = []
        valid_exts = (".pdf", ".png", ".jpg", ".jpeg", ".webp")

        with st.spinner("جاري فهرسة الملفات وتجهيزها للمعالجة..."):
            for idx, uf in enumerate(uploaded_files):
                fname = uf.name
                temp_path = os.path.join(temp_dir, f"upload_{idx}_{os.path.basename(fname)}")
                with open(temp_path, "wb") as out_f:
                    out_f.write(uf.getbuffer())

                if fname.lower().endswith(".zip"):
                    try:
                        with zipfile.ZipFile(temp_path, "r") as zf:
                            for member in zf.infolist():
                                if not member.is_dir() and member.filename.lower().endswith(valid_exts):
                                    base_m = os.path.basename(member.filename)
                                    if not base_m.startswith("._") and not base_m.startswith("."):
                                        task_queue.append(("zip", temp_path, member.filename, base_m))
                    except Exception as e:
                        st.error(f"خطأ في أرشيف {fname}: {e}")
                else:
                    task_queue.append(("file", temp_path, "", fname))

        total_files = len(task_queue)
        st.info(f"📂 تم تجهيز **{total_files}** مستند. جاري بدء التحليل...")

        progress_bar = st.progress(0)
        status_text = st.empty()
        
        raw_results = []
        completed = 0
        use_fast = fast_mode_toggle or (total_files > 150)

        BATCH_SIZE = 100
        for batch_start in range(0, total_files, BATCH_SIZE):
            batch = task_queue[batch_start : batch_start + BATCH_SIZE]
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = [
                    executor.submit(process_task_from_disk, t_info, split_branches, api_key, use_fast)
                    for t_info in batch
                ]
                for future in as_completed(futures):
                    res = future.result()
                    if res:
                        raw_results.extend(res)
                    completed += 1
                    if completed % 5 == 0 or completed == total_files:
                        progress_bar.progress(completed / max(total_files, 1))
                        status_text.text(f"⏳ تم فحص {completed} من أصل {total_files} مستند...")
            gc.collect()

        st.session_state["last_audit_results"] = raw_results

if "last_audit_results" in st.session_state and st.session_state["last_audit_results"]:
    raw_results = st.session_state["last_audit_results"]
    st.success(f"🎉 تم بنجاح تحليل وتدقيق **{len(raw_results)}** مستند!")
    
    output_excel = "تحليل_المشتريات_الضريبي.xlsx"
    export_to_excel(raw_results, output_excel)

    with open(output_excel, "rb") as f:
        st.download_button(
            label=f"📥 تحميل شيت تحليل المشتريات الكامل ({len(raw_results)} مستند - Excel)",
            data=f,
            file_name="تحليل_المشتريات.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    display_df = pd.DataFrame({
        "التاريخ": [r["invoice_date"] for r in raw_results],
        "رقم الفاتورة": [r["invoice_number"] for r in raw_results],
        "نوع المستند": [r["doc_type"] for r in raw_results],
        "اسم المورد": [r["vendor_name"] for r in raw_results],
        "رقم التسجيل الضريبي": [r["tax_id"] for r in raw_results],
        "الاجمالي قبل الضريبية": [r["total_before_vat"] for r in raw_results],
        "ضريبة القيمة المضافة": [r["vat_value"] for r in raw_results],
        "الاجمالي بعد الضريبة": [r["total_after_vat"] for r in raw_results],
        "حالة الـ QR Code": [r["qr_status"] for r in raw_results],
        "التدقيق المحاسبي": [r["audit_status"] for r in raw_results],
    })
    st.dataframe(display_df, use_container_width=True)
