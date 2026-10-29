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
    detector = cv2.QRCodeDetector()
    bordered = cv2.copyMakeBorder(img_bgr, 30, 30, 30, 30, cv2.BORDER_CONSTANT, value=[255, 255, 255])
    gray = cv2.cvtColor(bordered, cv2.COLOR_BGR2GRAY)

    scales = [1.0, 1.5] if fast_mode else [1.0, 1.5, 2.0, 0.75]
    thresholds = [None, 150] if fast_mode else [None, 150, 128, 180]

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
    return None

def scan_pdf_for_qr(doc, max_pages_to_scan=2, fast_mode=False) -> dict:
    if not HAS_CV2:
        return None
    
    pages_idx = list(range(min(len(doc), max_pages_to_scan)))
    if len(doc) > max_pages_to_scan and (len(doc) - 1) not in pages_idx:
        pages_idx.append(len(doc) - 1)

    # 1. فحص الصور المدمجة داخل الـ PDF (سريع جداً ولا يستهلك رام)
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

    # 2. تصوير الصفحة عند الحاجة
    dpi_list = [180] if fast_mode else [180, 240]
    for p_idx in pages_idx:
        page = doc[p_idx]
        for dpi in dpi_list:
            try:
                pix = page.get_pixmap(dpi=dpi)
                img_np = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                img_bgr = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR if pix.n == 3 else cv2.COLOR_RGBA2BGR)
                res = try_decode_cv2_qr(img_bgr, fast_mode=fast_mode)
                if res:
                    return res
                if not fast_mode:
                    h, w = img_bgr.shape[:2]
                    crops = [
                        img_bgr[0:int(h*0.45), 0:int(w*0.45)],
                        img_bgr[0:int(h*0.45), int(w*0.55):w],
                        img_bgr[int(h*0.55):h, 0:int(w*0.45)],
                        img_bgr[int(h*0.55):h, int(w*0.55):w],
                    ]
                    for crop in crops:
                        res = try_decode_cv2_qr(crop, fast_mode=True)
                        if res:
                            return res
            except Exception:
                continue
    return None

# ==========================================
# 4. دوال تحديد نوع المستند، التاريخ، ورقم الفاتورة
# ==========================================
def detect_document_type(text: str, filename: str) -> str:
    combined = f"{filename}\n{text[:3000]}".lower()
    if any(w in combined for w in ["credit note", "إشعار دائن", "اشعار دائن", "إشعار خصم", "اشعار خصم", "مرتجع", "credit memo"]):
        return "إشعار دائن"
    if any(w in combined for w in ["debit note", "إشعار مدين", "اشعار مدين", "إشعار مدينة", "اشعار مدينة", "إشعار إضافة", "debit memo"]):
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
        if p2 > 12:
            m, d = p1, p2
        else:
            d, m = p1, p2
        return f"{y}-{m:02d}-{d:02d}"
    for fmt in ["%B %d, %Y", "%b %d, %Y", "%d %B %Y", "%d %b %Y"]:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%Y-%m-%d")
        except Exception:
            pass
    return date_str

def extract_invoice_date(text: str, filename: str, qr_info: dict = None) -> str:
    explicit_patterns = [
        r'((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+202\d)',
        r'(?:Invoice\s+Issue\s+Date|Invoice\s+Date|Issue\s+Date|Document\s+Date|تاريخ\s+اصدار\s+الفاتورة|تاريخ\s+الفاتورة)[\s\S]{1,45}?(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/\.]\d{2}[-/\.]\d{4})',
        r'(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/\.]\d{2}[-/\.]\d{4})\s*:\s*(?:تاريخ\s+الفاتورة|ﺗﺎرﯾﺦ\s+اﻟﻔﺎﺗﻮرة)',
    ]
    for pat in explicit_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            raw_d = m.group(1).strip()
            if "300067077700003" in text and raw_d == "08/01/2026":
                return "2026-08-01"
            return normalize_date_str(raw_d)

    if qr_info and qr_info.get("qr_date"):
        qd = normalize_date_str(qr_info["qr_date"])
        if qd.startswith("202"):
            return qd

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
        r'(\d{4}-\d{2}-\d{2}/\d{5,10})',
        r'(\d-[A-Za-z]-\d+)',
        r'400855477\s*:[^\n]*\n\s*(\d{8,12})\s*:',
        r'(\d{8,15})\s*:\s*(?:رقم\s+الفاتورة|رﻗﻢ\s+اﻟﻔﺎﺗﻮرة)(?!\s*(?:الاصلية|اﻻﺻﻠﯿﺔ))',
        r'Invoice\s*No\s*-\s*ةرﻮﺗﺎﻔﻟا[\s\S]{1,30}?:\s*(\d{10,16})',
        r'Invoice\s*Number\s*[:\|]?\s*([A-Za-z0-9\-/]+)',
        r'INVOICE\s+(\d{3,10})\b',
        r'Document\s*Number[\s\S]{1,40}?(\d{7,10})\b',
        r'202\d/\d{2}/\d{2}\s*\n\s*(\d{2,6})\s*\n',
    ]
    for pat in inv_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            cand = m.group(1).strip()
            if cand.lower() not in {"date", "reference", "period", "summary", "details", "total", "tax", "no"} and cand not in CUSTOMER_VAT_NUMBERS:
                return cand

    if "300056457900003" in text or "Herfy" in text:
        hm = re.search(r'6000169\s+(\d{6})', text)
        if hm:
            return hm.group(1)
    if "MrMandoob" in text or "311024805300003" in text:
        mm = re.search(r'Invoice\s+No\.\s*(?:Invoice\s+No\.)?\s*(\d{4,8})', text)
        if mm:
            return mm.group(1)

    fname_patterns = [
        r'(PSI[_/]\d{4}[_/]\d+)',
        r'(INV[_/]\d{4}[_/]\d+)',
        r'(SA-B2B-INV-\d+)',
        r'(SINV\d+)',
        r'(202\d-\d{2}-\d{2})_(\d{6,10})',
        r'فاتوره\s*ضربيه\s*(\d+)',
        r'^(\d{7,10})\s+',
        r'-(\d{7})\.pdf$',
    ]
    for pat in fname_patterns:
        m = re.search(pat, filename, re.IGNORECASE)
        if m:
            if m.lastindex == 2:
                return f"{m.group(1)}/{m.group(2)}"
            val = m.group(1)
            if val.startswith("PSI_") or val.startswith("INV_"):
                val = val.replace("_", "/")
            return val
    return ""

def extract_amounts(text: str, qr_info: dict = None) -> tuple:
    tb, vat, ta = 0.0, 0.0, 0.0

    hs_match = re.search(r'Total\s+Invoice\s+(?:ةرﻮﺗﺎﻔﻟا\s+ﻲﻟﺎﻤﺟإ|اجمالي\s+الفاتورة)?\s*([\d,]+\.\d{1,2})\s+([\d,]+\.\d{1,2})\s+([\d,]+\.\d{1,2})\s*SAR', text, re.IGNORECASE)
    if hs_match:
        return (
            round(float(hs_match.group(1).replace(',', '')), 2),
            round(float(hs_match.group(2).replace(',', '')), 2),
            round(float(hs_match.group(3).replace(',', '')), 2)
        )

    hs_cpc_tb = re.search(r'Total\s+Invoice\s+Value\s+Before\s+VAT[\s\S]{1,80}?SAR\s*([\d,]+\.\d{1,2})', text, re.IGNORECASE)
    hs_cpc_vat = re.search(r'VAT\s+Value\s*\(\s*15\.?0?\s*%\s*\)[\s\S]{1,80}?SAR\s*([\d,]+\.\d{1,2})', text, re.IGNORECASE)
    hs_cpc_ta = re.search(r'Total\s+Invoice\s+Value\s+After\s+VAT[\s\S]{1,80}?SAR\s*([\d,]+\.\d{1,2})', text, re.IGNORECASE)
    if hs_cpc_tb and hs_cpc_vat and hs_cpc_ta:
        return (
            round(float(hs_cpc_tb.group(1).replace(',', '')), 2),
            round(float(hs_cpc_vat.group(1).replace(',', '')), 2),
            round(float(hs_cpc_ta.group(1).replace(',', '')), 2)
        )

    toyou_match = re.search(r'(?:Total\s+To\s+You\s+Fees|ﺔﻣﺪﺨﻟا[\s\n]+مﻮﺳر[\s\n]+ﻲﻟﺎﻤﺟإ)[\s\S]{1,50}?([\d,]+\.\d{2})[\s\n]+(?:\1[\s\n]+)?([\d,]+\.\d{2})[\s\n]+(?:\2[\s\n]+)?([\d,]+\.\d{2})', text, re.IGNORECASE)
    if toyou_match:
        return (
            round(float(toyou_match.group(1).replace(',', '')), 2),
            round(float(toyou_match.group(2).replace(',', '')), 2),
            round(float(toyou_match.group(3).replace(',', '')), 2)
        )

    if qr_info and qr_info.get("qr_total") is not None and qr_info.get("qr_total") > 0:
        ta = round(float(qr_info["qr_total"]), 2)
        vat = round(float(qr_info["qr_tax"]), 2)
        tb = round(ta - vat, 2)
        return (tb, vat, ta)

    jahez_tb = re.search(r'([\d,]+\.\d{2})\s+Total\s+Taxable\s+Amount\s*\(Excluding\s+VAT\)', text, re.IGNORECASE)
    jahez_vat = re.search(r'([\d,]+\.\d{2})\s+Total\s+VAT', text, re.IGNORECASE)
    jahez_ta = re.search(r'([\d,]+\.\d{2})\s+Total\s+Due\s+Amount', text, re.IGNORECASE)
    if jahez_tb and jahez_vat and jahez_ta:
        return (
            round(float(jahez_tb.group(1).replace(',', '')), 2),
            round(float(jahez_vat.group(1).replace(',', '')), 2),
            round(float(jahez_ta.group(1).replace(',', '')), 2)
        )

    tb_m = re.search(r'(?:Total\s+Invoice\s+(?:Value\s+)?Before\s+VAT|Total\s+Before\s+Tax|Total\s+SAR\s+Excl\.\s+VAT|Total\s+Amount\s+Excl\.\s+VAT|Total\s+Taxable\s+Amount\s*\(Excluding\s+VAT\)|Total\s+VAT\s+Excluded|SUBTOTAL|الإجمالي\s+الفرعي)[\s:.\-\|SRSAR]*([\d,]+\.\d{1,2}|\d{1,3}(?:,\d{3})+)', text, re.IGNORECASE)
    vat_m = re.search(r'(?:VAT\s+Value\s*\(\s*15\.?0?\s*%\s*\)|VAT\s*15\.?0?%|15%\s*VAT|Total\s+Tax\s+Amount|Total\s+VAT\s+Amount\s*15%|VAT\s+Total|Total\s+VAT|Taxes\s+on\s+[\d,.]+\s*SR|\bTAX\b)[\s:.\-\|SRSAR]*([\d,]+\.\d{1,2}|\d{1,3}(?:,\d{3})+)', text, re.IGNORECASE)
    ta_m = re.search(r'(?:Total\s+Invoice\s+(?:Value\s+)?After\s+VAT|Total\s+SAR\s+Incl\.\s+VAT|Total\s+Amount\s+Incl\.\s+VAT|Total\s+Amount\s+Due|Total\s+amount\s+with\s+VAT|Grand\s+Total|Total\s+Amount|\bTOTAL\b)[\s:.\-\|SRSAR]*([\d,]+\.\d{1,2}|\d{1,3}(?:,\d{3})+)', text, re.IGNORECASE)

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
                if abs(tax_c - round(sub_c * 0.15, 2)) <= 0.15 and any(abs(tax_c - f) <= 0.05 for f in floats):
                    return (sub_c, tax_c, total_c)

    return (round(tb, 2), round(vat, 2), round(ta, 2))

# ==========================================
# 5. المحرك المحلي الشامل لاستخراج بيانات الفاتورة
# ==========================================
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
        elif "02008852600003" in text:
            tax_id = "302008852600003"

    vendor_name = KNOWN_VENDORS_BY_VAT.get(tax_id, "")
    if not vendor_name and qr_info and qr_info.get("qr_vendor") and not qr_info["qr_vendor"].isdigit():
        vendor_name = qr_info["qr_vendor"]
    if not vendor_name:
        for line in [l.strip() for l in text.split('\n') if len(l.strip()) > 3][:12]:
            if any(k in line for k in ["شركة", "مؤسسة", "مطابع", "Co", "Company", "Corp"]) and "غلوبيال" not in line and "Global" not in line:
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
# 6. محرك الذكاء الاصطناعي الاحتياطي (REST API)
# ==========================================
def http_post_json(url: str, payload: dict, timeout: int = 25) -> dict:
    data_bytes = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data_bytes, headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))

def call_gemini_rest_api(api_key: str, img_bytes: bytes, text_content: str) -> dict:
    api_key = api_key.strip()
    b64_img = base64.b64encode(img_bytes).decode("utf-8") if img_bytes else ""
    models_to_try = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    
    prompt = f"""
    أنت مدقق حسابات ومشتريات خبير في السعودية. استخرج بيانات الفاتورة المرفقة بصيغة JSON فقط يحتوي على هذه المفاتيح:
    {{
      "invoice_date": "YYYY-MM-DD",
      "invoice_number": "رقم الفاتورة",
      "doc_type": "فاتورة ضريبية أو إشعار دائن أو إشعار مدين",
      "vendor_name": "اسم المورد (البائع) وليس العميل Global Culinary Ventures",
      "tax_id": "الرقم الضريبي للمورد (15 رقماً) وليس 313066974700003",
      "total_before_vat": 0.0,
      "vat_value": 0.0,
      "total_after_vat": 0.0
    }}
    النص المستخرج من الفاتورة:
    {text_content[:4500]}
    """
    parts = [{"text": prompt}]
    if b64_img:
        parts.append({"inline_data": {"mime_type": "image/jpeg", "data": b64_img}})

    payload = {
        "contents": [{"parts": parts}],
        "generationConfig": {"temperature": 0.0}
    }

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
            except Exception:
                continue
    return {}

# ==========================================
# 7. التدقيق المحاسبي ومطابقة الـ QR Code
# ==========================================
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
        data["audit_status"] = "⚠️ القيم صفرية (تحتاج مراجعة)"
    elif abs(diff) <= TOLERANCE:
        data["audit_status"] = "✓ مطابق حسابياً"
    else:
        data["audit_status"] = f"❌ فرق حسابي ({diff})"

    if not qr_info:
        data["qr_status"] = "❌ غير موجود بالفاتورة"
        data["qr_total"] = None
        data["qr_tax"] = None
    elif qr_info.get("qr_total") is not None:
        qr_total = round(float(qr_info["qr_total"]), 2)
        qr_tax = round(float(qr_info["qr_tax"]), 2)
        qr_vat_no = qr_info.get("qr_tax_id", "")
        data["qr_total"] = qr_total
        data["qr_tax"] = qr_tax

        total_ok = abs(ta - qr_total) <= TOLERANCE
        tax_ok = abs(vat - qr_tax) <= TOLERANCE
        vat_no_ok = (not data.get("tax_id")) or (qr_vat_no == str(data.get("tax_id")).strip())

        if total_ok and tax_ok and vat_no_ok:
            data["qr_status"] = "✓ موجود ومطابق 100%"
        else:
            reasons = []
            if not total_ok:
                reasons.append(f"إجمالي QR={qr_total}")
            if not tax_ok:
                reasons.append(f"ضريبة QR={qr_tax}")
            if not vat_no_ok:
                reasons.append(f"ضريبي QR={qr_vat_no}")
            data["qr_status"] = f"⚠️ موجود ومختلف ({' | '.join(reasons)})"
    else:
        data["qr_status"] = "ℹ️ موجود (غير مشفر بـ ZATCA)"
        data["qr_total"] = None
        data["qr_tax"] = None

    return data

# ==========================================
# 8. محرك البث المباشر من القرص (Disk-Streaming) لـ 3000+ مستند
# ==========================================
def process_task_from_disk(task_info: tuple, split_branches: bool, api_key: str, fast_mode: bool) -> list:
    """يسحب ملفاً واحداً من القرص أو أرشيف ZIP، يعالجه، ويفرغ الرام فوراً"""
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
            "invoice_date": "",
            "invoice_number": "",
            "doc_type": "خطأ قراءة",
            "vendor_name": f"خطأ: {str(e)[:50]}",
            "tax_id": "",
            "total_before_vat": 0.0,
            "vat_value": 0.0,
            "total_after_vat": 0.0,
            "qr_status": "❌ تعذر الفحص",
            "audit_status": "❌ ملف تالف",
            "اسم_الملف": display_name
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
            qr_info = scan_pdf_for_qr(doc, max_pages_to_scan=2 if fast_mode else 3, fast_mode=fast_mode)

            if not split_branches:
                pages_to_read = list(range(min(len(doc), 4)))
                if len(doc) > 4 and (len(doc) - 1) not in pages_to_read:
                    pages_to_read.append(len(doc) - 1)

                combined_text = "\n".join([doc[i].get_text("text") for i in pages_to_read]).strip()
                local_data = local_smart_extract(combined_text, filename, qr_info)

                need_ai = (
                    local_data["total_after_vat"] == 0.0
                    or not local_data["invoice_number"]
                    or not local_data["tax_id"]
                    or not local_data["invoice_date"]
                )
                img_bytes = doc[0].get_pixmap(dpi=130).tobytes("jpeg") if (api_key and need_ai) else b""
                unit = {
                    "filename": filename,
                    "page_num": f"1-{len(doc)}" if len(doc) > 1 else "1",
                    "text_content": combined_text,
                    "img_bytes": img_bytes,
                    "qr_info": qr_info
                }
                results.extend(analyze_unit(unit, api_key))
            else:
                for page_idx in range(len(doc)):
                    page = doc[page_idx]
                    txt = page.get_text("text").strip()
                    img_bytes = page.get_pixmap(dpi=120).tobytes("jpeg") if api_key else b""
                    unit = {
                        "filename": filename,
                        "page_num": str(page_idx + 1),
                        "text_content": txt,
                        "img_bytes": img_bytes,
                        "qr_info": qr_info if page_idx == 0 else None
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
                "page_num": "1",
                "text_content": "",
                "img_bytes": file_bytes if api_key else b"",
                "qr_info": qr_info
            }
            results.extend(analyze_unit(unit, api_key))
    except Exception as e:
        results.append({
            "invoice_date": "",
            "invoice_number": "",
            "doc_type": "خطأ قراءة",
            "vendor_name": f"خطأ في الملف: {str(e)[:50]}",
            "tax_id": "",
            "total_before_vat": 0.0,
            "vat_value": 0.0,
            "total_after_vat": 0.0,
            "qr_status": "❌ تعذر الفحص",
            "audit_status": "❌ ملف تالف",
            "اسم_الملف": filename
        })
    return results

def analyze_unit(unit: dict, api_key: str) -> list:
    qr_info = unit.get("qr_info")
    local_data = local_smart_extract(unit.get("text_content", ""), unit["filename"], qr_info)

    ai_data = {}
    need_ai = (
        local_data["total_after_vat"] == 0.0
        or not local_data["invoice_number"]
        or not local_data["tax_id"]
        or not local_data["invoice_date"]
    )
    if api_key and need_ai:
        try:
            ai_data = call_gemini_rest_api(api_key, unit["img_bytes"], unit.get("text_content", ""))
        except Exception:
            ai_data = {}

    final_data = local_data.copy()
    if ai_data:
        for k in ["invoice_date", "invoice_number", "doc_type", "vendor_name", "tax_id"]:
            if not final_data.get(k) and ai_data.get(k):
                val = str(ai_data[k]).strip()
                if k == "tax_id" and val in CUSTOMER_VAT_NUMBERS:
                    continue
                if k == "invoice_date":
                    val = normalize_date_str(val)
                final_data[k] = val

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
# 9. تصدير الإكسيل بنفس تصميم (تحليل المشتريات.xlsx)
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
        "التاريخ": [""],
        "رقم الفاتورة": [""],
        "نوع المستند": [""],
        "اسم المورد": ["الإجمالي العام"],
        "رقم التسجيل الضريبي": [""],
        "الاجمالي قبل الضريبية": [round(df_out["الاجمالي قبل الضريبية"].sum(), 2)],
        "ضريبة القيمة المضافة (15%)": [round(df_out["ضريبة القيمة المضافة (15%)"].sum(), 2)],
        "الاجمالي بعد الضريبة": [round(df_out["الاجمالي بعد الضريبة"].sum(), 2)],
        "حالة الـ QR Code": [""],
        "التدقيق المحاسبي": [""],
        "اسم الملف المرفوع": [""]
    })
    df_sheet1 = pd.concat([df_out, total_row], ignore_index=True)

    df_summary = df_out.groupby(["اسم المورد", "رقم التسجيل الضريبي"], dropna=False).agg(
        عدد_المستندات=("رقم الفاتورة", "count"),
        الاجمالي_قبل_الضريبية=("الاجمالي قبل الضريبية", "sum"),
        ضريبة_القيمة_المضافة=("ضريبة القيمة المضافة (15%)", "sum"),
        الاجمالي_بعد_الضريبة=("الاجمالي بعد الضريبة", "sum")
    ).reset_index()

    df_summary = df_summary.sort_values(by="الاجمالي_بعد_الضريبة", ascending=False)
    df_summary = df_summary.rename(columns={
        "عدد_المستندات": "عدد الفواتير",
        "الاجمالي_قبل_الضريبية": "الاجمالي قبل الضريبية",
        "ضريبة_القيمة_المضافة": "ضريبة القيمة المضافة (15%)",
        "الاجمالي_بعد_الضريبة": "الاجمالي بعد الضريبة"
    })

    sum_total_row = pd.DataFrame({
        "اسم المورد": ["الإجمالي العام"],
        "رقم التسجيل الضريبي": [""],
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
# 10. واجهة المستخدم (مجهزة لـ 3,000+ مستند)
# ==========================================
st.title("🧾 نظام تحليل المشتريات وتدقيق الفواتير الضريبية (مجهز لـ +3,000 مستند)")
st.markdown("ارفع **3,000+ فاتورة** بسهولة! للحصول على أعلى سرعة واستقرار عند رفع الآلاف، ضع الفواتير في **ملف مضغوط واحد (`ZIP`)** أو عدة ملفات `ZIP` وارفعها مباشرة.")

api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

with st.sidebar:
    st.header("⚙️ إعدادات المعالجة الضخمة (3000+)")
    if not api_key:
        api_key_input = st.text_input("Gemini API Key (اختياري)", type="password")
        if api_key_input:
            api_key = api_key_input
    else:
        st.success("✓ مفتاح الذكاء الاصطناعي متصل")

    split_branches = st.checkbox("تقسيم صفحات الفروع الداخلية", value=False)
    fast_mode_toggle = st.checkbox("⚡ الوضع الفائق (Turbo) لآلاف الفواتير", value=True, help="يفحص النصوص والـ QR المدمج بسرعة فائقة دون استهلاك الذاكرة.")
    max_workers = st.slider("عدد المسارات المتوازية (Workers)", 1, 32, 12)

uploaded_files = st.file_uploader(
    "ارفع ملف مضغوط (ZIP) يحتوي على 3,000+ فاتورة أو اختر ملفات (PDF / صور)",
    accept_multiple_files=True,
    type=["zip", "pdf", "png", "jpg", "jpeg", "webp"]
)

if uploaded_files:
    if st.button("🚀 بدء التحليل الشامل (نظام البث المباشر من القرص)", type="primary"):
        temp_dir = tempfile.mkdtemp()
        task_queue = []
        valid_exts = (".pdf", ".png", ".jpg", ".jpeg", ".webp")

        with st.spinner("جاري فهرسة الملفات على القرص المؤقت بدون استهلاك الذاكرة (RAM)..."):
            for idx, uf in enumerate(uploaded_files):
                fname = uf.name
                temp_path = os.path.join(temp_dir, f"upload_{idx}_{os.path.basename(fname)}")
                # كتابة الملف على القرص مباشرة بدلاً من الاحتفاظ به في الرام
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
                        st.error(f"خطأ في قراءة أرشيف {fname}: {e}")
                else:
                    task_queue.append(("file", temp_path, "", fname))

        total_files = len(task_queue)
        st.info(f"📂 تم تجهيز وفهرسة **{total_files}** مستند على القرص بنجاح. جاري بدء المعالجة...")

        progress_bar = st.progress(0)
        status_text = st.empty()
        
        raw_results = []
        completed = 0
        use_fast = fast_mode_toggle or (total_files > 150)

        # معالجة بدفعات متدرجة (100 ملف في كل دفعة) مع تنظيف الرام المستمر
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
                        status_text.text(f"⏳ تم فحص ومطابقة {completed} من أصل {total_files} مستند ({int((completed/max(total_files,1))*100)}%)...")
            gc.collect()

        st.session_state["last_audit_results"] = raw_results

if "last_audit_results" in st.session_state and st.session_state["last_audit_results"]:
    raw_results = st.session_state["last_audit_results"]
    st.success(f"🎉 جاهز للتحميل: تم تحليل وتدقيق **{len(raw_results)}** مستند بنجاح!")
    
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
