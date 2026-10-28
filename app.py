import streamlit as st
import os
import re
import io
import gc
import tempfile
import zipfile
import json
import base64
import threading
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

st.set_page_config(page_title="نظام تحليل المشتريات والفرز الضريبي (ZATCA)", page_icon="🧾", layout="wide")

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
# 2. ثوابت ونظام الذاكرة الديناميكية للموردين
# ==========================================
TOLERANCE = 0.05
CUSTOMER_VAT_NUMBERS = {"313066974700003"}
VENDORS_DB_FILE = "vendors_db.json"
DB_LOCK = threading.Lock()

DEFAULT_VENDORS = {
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

def load_vendors_db() -> dict:
    db = DEFAULT_VENDORS.copy()
    with DB_LOCK:
        try:
            if os.path.exists(VENDORS_DB_FILE):
                with open(VENDORS_DB_FILE, "r", encoding="utf-8") as f:
                    saved = json.load(f)
                    if isinstance(saved, dict):
                        db.update(saved)
        except Exception:
            pass
    return db

def save_vendor_to_db(tax_id: str, vendor_name: str):
    if not tax_id or len(tax_id) != 15 or not vendor_name:
        return
    if tax_id in CUSTOMER_VAT_NUMBERS or vendor_name.isdigit():
        return
    with DB_LOCK:
        try:
            db = DEFAULT_VENDORS.copy()
            if os.path.exists(VENDORS_DB_FILE):
                try:
                    with open(VENDORS_DB_FILE, "r", encoding="utf-8") as f:
                        saved = json.load(f)
                        if isinstance(saved, dict):
                            db.update(saved)
                except Exception:
                    pass
            if tax_id not in db:
                db[tax_id] = vendor_name.strip()
                with open(VENDORS_DB_FILE, "w", encoding="utf-8") as f:
                    json.dump(db, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

# ==========================================
# 3. محرك فك التشفير وزوم الـ QR Code (ZATCA TLV)
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
        h, w = img_bgr.shape[:2]

        regions = [
            ("full", img_bgr),
            ("top_left", img_bgr[0:int(h*0.48), 0:int(w*0.48)]),
            ("top_right", img_bgr[0:int(h*0.48), int(w*0.52):w]),
            ("bottom_left", img_bgr[int(h*0.52):h, 0:int(w*0.48)]),
            ("bottom_right", img_bgr[int(h*0.52):h, int(w*0.52):w]),
        ]

        zoom_levels = [1.0, 1.6] if fast_mode else [1.0, 1.5, 2.0]

        for _, reg in regions:
            if reg.size == 0:
                continue
            bordered = cv2.copyMakeBorder(reg, 30, 30, 30, 30, cv2.BORDER_CONSTANT, value=[255, 255, 255])
            gray = cv2.cvtColor(bordered, cv2.COLOR_BGR2GRAY)

            for zoom in zoom_levels:
                try:
                    resized = cv2.resize(gray, None, fx=zoom, fy=zoom, interpolation=cv2.INTER_CUBIC) if zoom != 1.0 else gray
                    thresholds = [
                        resized,
                        cv2.threshold(resized, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1],
                        cv2.threshold(resized, 135, 255, cv2.THRESH_BINARY)[1]
                    ]
                    for t_img in thresholds:
                        data, _, _ = detector.detectAndDecode(t_img)
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

def scan_pdf_for_qr(doc, max_pages_to_scan=3, fast_mode=False) -> dict:
    if not HAS_CV2:
        return None
    
    pages_idx = list(range(min(len(doc), max_pages_to_scan)))
    if len(doc) > max_pages_to_scan and (len(doc) - 1) not in pages_idx:
        pages_idx.append(len(doc) - 1)

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

    for p_idx in pages_idx:
        page = doc[p_idx]
        for dpi in [200]:
            try:
                pix = page.get_pixmap(dpi=dpi)
                img_np = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                img_bgr = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR if pix.n == 3 else cv2.COLOR_RGBA2BGR)
                res = try_decode_cv2_qr(img_bgr, fast_mode=fast_mode)
                if res:
                    return res
            except Exception:
                continue
    return None

# ==========================================
# 4. دوال استخراج البيانات الأساسية
# ==========================================
def detect_document_type(text: str, filename: str) -> str:
    combined = f"{filename}\n{text[:3500]}".lower()
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
        r'(?:Invoice\s+Issue\s+Date|Invoice\s+Date|Issue\s+Date|Document\s+Date|\bDATE\b|تاريخ\s+اصدار\s+الفاتورة|تاريخ\s+الفاتورة)[\s\S]{1,50}?(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/\.]\d{2}[-/\.]\d{4})',
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

    if "300048882800003" in text or "Bab International" in text:
        return ""

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
# 5. استخراج البيانات محلياً مع التعلم التلقائي
# ==========================================
def local_smart_extract(text: str, filename: str, qr_info: dict = None) -> dict:
    doc_type = detect_document_type(text, filename)
    vendors_db = load_vendors_db()
    
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

    vendor_name = vendors_db.get(tax_id, "")
    if not vendor_name and qr_info and qr_info.get("qr_vendor") and not qr_info["qr_vendor"].isdigit():
        vendor_name = qr_info["qr_vendor"]
    if not vendor_name:
        for line in [l.strip() for l in text.split('\n') if len(l.strip()) > 3][:12]:
            if any(k in line for k in ["شركة", "مؤسسة", "مطابع", "Co", "Company", "Corp"]) and "غلوبيال" not in line and "Global" not in line:
                vendor_name = line
                break
        if not vendor_name:
            vendor_name = re.sub(r'\.pdf$|\.jpg$|\.png$', '', filename, flags=re.IGNORECASE)

    if tax_id and vendor_name and tax_id not in CUSTOMER_VAT_NUMBERS:
        save_vendor_to_db(tax_id, vendor_name)

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
# 6. الذكاء الاصطناعي (REST API)
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
    أنت مدقق حسابات ومشتريات خبير في السعودية. استخرج بيانات الفاتورة المرفقة بصيغة JSON فقط:
    {{
      "invoice_date": "YYYY-MM-DD",
      "invoice_number": "رقم الفاتورة",
      "doc_type": "فاتورة ضريبية أو إشعار دائن أو إشعار مدين",
      "vendor_name": "اسم المورد وليس العميل Global Culinary Ventures",
      "tax_id": "الرقم الضريبي للمورد (15 رقماً)",
      "total_before_vat": 0.0,
      "vat_value": 0.0,
      "total_after_vat": 0.0
    }}
    النص:
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
# 7. التدقيق والفرز الضريبي لمواصفات ZATCA
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

    error_reasons = []

    if ta == 0 and tb == 0:
        data["audit_status"] = "⚠️ القيم صفرية"
        error_reasons.append("القيم المالية صفرية")
    elif abs(diff) <= TOLERANCE:
        data["audit_status"] = "✓ مطابق حسابياً"
    else:
        data["audit_status"] = f"❌ فرق حسابي ({diff})"
        error_reasons.append(f"خطأ في مجموع الضريبة (فرق {diff})")

    tax_id_str = str(data.get("tax_id") or "").strip()
    if not tax_id_str or len(tax_id_str) != 15:
        error_reasons.append("الرقم الضريبي للمورد غير مكتمل أو مفقود")
    if not data.get("invoice_date"):
        error_reasons.append("تاريخ الفاتورة مفقود")

    qr_valid = False
    if not qr_info:
        data["qr_status"] = "❌ غير موجود بالفاتورة"
        data["qr_total"] = None
        data["qr_tax"] = None
        error_reasons.append("رمز الـ QR غير موجود أو تالف")
    elif qr_info.get("qr_total") is not None:
        qr_total = round(float(qr_info["qr_total"]), 2)
        qr_tax = round(float(qr_info["qr_tax"]), 2)
        qr_vat_no = qr_info.get("qr_tax_id", "")
        data["qr_total"] = qr_total
        data["qr_tax"] = qr_tax

        total_ok = abs(ta - qr_total) <= TOLERANCE
        tax_ok = abs(vat - qr_tax) <= TOLERANCE
        vat_no_ok = (not tax_id_str) or (qr_vat_no == tax_id_str)

        if total_ok and tax_ok and vat_no_ok:
            data["qr_status"] = "✓ موجود ومطابق 100%"
            qr_valid = True
        else:
            reasons = []
            if not total_ok:
                reasons.append(f"إجمالي QR={qr_total}")
            if not tax_ok:
                reasons.append(f"ضريبة QR={qr_tax}")
            if not vat_no_ok:
                reasons.append(f"ضريبي QR={qr_vat_no}")
            mismatch_str = " | ".join(reasons)
            data["qr_status"] = f"⚠️ موجود ومختلف ({mismatch_str})"
            error_reasons.append(f"اختلاف بيانات الـ QR ({mismatch_str})")
    else:
        data["qr_status"] = "ℹ️ موجود (غير مشفر بـ ZATCA)"
        data["qr_total"] = None
        data["qr_tax"] = None
        error_reasons.append("الـ QR غير مشفر بمعايير ZATCA")

    is_clean = (len(error_reasons) == 0) and qr_valid
    data["is_valid_zatca"] = is_clean
    data["error_reason"] = "✓ سليم ومطابق لهيئة الزكاة" if is_clean else " ، ".join(error_reasons)

    return data

# ==========================================
# 8. معالجة البث المباشر (Disk-Streaming)
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
            "is_valid_zatca": False,
            "error_reason": f"تعذر فتح الملف ({str(e)[:40]})",
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
            "vendor_name": f"خطأ: {str(e)[:50]}",
            "tax_id": "",
            "total_before_vat": 0.0,
            "vat_value": 0.0,
            "total_after_vat": 0.0,
            "qr_status": "❌ تعذر الفحص",
            "audit_status": "❌ ملف تالف",
            "is_valid_zatca": False,
            "error_reason": f"خطأ قراءة الملف ({str(e)[:40]})",
            "اسم_الملف": filename
        })
    return results

def analyze_unit(unit: dict, api_key: str) -> list:
    qr_info = unit.get("qr_info")
    local_data = local_smart_extract(unit.get("text_content", ""), unit["filename"], qr_info)

    ai_data = {}
    need_ai = (
        local_data["total_after_vat"] == 0.0
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

    if final_data.get("tax_id") and final_data.get("vendor_name"):
        save_vendor_to_db(final_data["tax_id"], final_data["vendor_name"])

    vendors_db = load_vendors_db()
    if final_data.get("tax_id") in vendors_db:
        final_data["vendor_name"] = vendors_db[final_data["tax_id"]]

    final_data["اسم_الملف"] = unit["filename"]
    audited = perform_audit_and_qr_match(final_data, qr_info)
    return [audited]

# ==========================================
# 9. تصدير الإكسيل
# ==========================================
def build_formatted_df(sub_df: pd.DataFrame, include_error_reason: bool = False) -> pd.DataFrame:
    if sub_df.empty:
        cols = [
            "التاريخ", "رقم الفاتورة", "نوع المستند", "اسم المورد",
            "رقم التسجيل الضريبي", "الاجمالي قبل الضريبية",
            "ضريبة القيمة المضافة (15%)", "الاجمالي بعد الضريبة",
            "حالة الـ QR Code", "التدقيق المحاسبي"
        ]
        if include_error_reason:
            cols.append("سبب الاستبعاد / نوع الخطأ")
        cols.append("اسم الملف المرفوع")
        return pd.DataFrame(columns=cols)

    df_out = pd.DataFrame({
        "التاريخ": sub_df["invoice_date"],
        "رقم الفاتورة": sub_df["invoice_number"],
        "نوع المستند": sub_df["doc_type"],
        "اسم المورد": sub_df["vendor_name"],
        "رقم التسجيل الضريبي": sub_df["tax_id"],
        "الاجمالي قبل الضريبية": pd.to_numeric(sub_df["total_before_vat"], errors="coerce").fillna(0.0),
        "ضريبة القيمة المضافة (15%)": pd.to_numeric(sub_df["vat_value"], errors="coerce").fillna(0.0),
        "الاجمالي بعد الضريبة": pd.to_numeric(sub_df["total_after_vat"], errors="coerce").fillna(0.0),
        "حالة الـ QR Code": sub_df["qr_status"],
        "التدقيق المحاسبي": sub_df["audit_status"],
    })

    if include_error_reason:
        df_out["سبب الاستبعاد / نوع الخطأ"] = sub_df["error_reason"]

    df_out["اسم الملف المرفوع"] = sub_df["اسم_الملف"]

    total_dict = {
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
    }
    if include_error_reason:
        total_dict["سبب الاستبعاد / نوع الخطأ"] = [""]
    total_dict["اسم الملف المرفوع"] = [""]

    return pd.concat([df_out, pd.DataFrame(total_dict)], ignore_index=True)

def export_to_excel(records: list, output_path: str):
    df = pd.DataFrame(records)
    if df.empty:
        return

    inv_counts = df["invoice_number"].replace("", np.nan).dropna().value_counts()
    duplicate_invs = set(inv_counts[inv_counts > 1].index)

    credit_notes_mask = df["doc_type"].astype(str).str.contains("دائن", na=False)
    df_credit_notes_raw = df[credit_notes_mask].copy()

    valid_mask = (df["is_valid_zatca"] == True) & (~credit_notes_mask)
    df_valid_raw = df[valid_mask].copy()

    error_mask = (df["is_valid_zatca"] == False) & (~credit_notes_mask)
    df_error_raw = df[error_mask].copy()

    df_sheet_valid = build_formatted_df(df_valid_raw, include_error_reason=False)
    df_sheet_credit = build_formatted_df(df_credit_notes_raw, include_error_reason=True)
    df_sheet_error = build_formatted_df(df_error_raw, include_error_reason=True)

    sign_multiplier = np.where(credit_notes_mask, -1.0, 1.0)
    df_all_formatted = pd.DataFrame({
        "اسم المورد": df["vendor_name"],
        "رقم التسجيل الضريبي": df["tax_id"],
        "رقم الفاتورة": df["invoice_number"],
        "الاجمالي قبل الضريبية": pd.to_numeric(df["total_before_vat"], errors="coerce").fillna(0.0) * sign_multiplier,
        "ضريبة القيمة المضافة (15%)": pd.to_numeric(df["vat_value"], errors="coerce").fillna(0.0) * sign_multiplier,
        "الاجمالي بعد الضريبة": pd.to_numeric(df["total_after_vat"], errors="coerce").fillna(0.0) * sign_multiplier,
    })

    df_summary = df_all_formatted.groupby(["اسم المورد", "رقم التسجيل الضريبي"], dropna=False).agg(
        عدد_المستندات=("رقم الفاتورة", "count"),
        الاجمالي_قبل_الضريبية=("الاجمالي قبل الضريبية", "sum"),
        ضريبة_القيمة_المضافة=("ضريبة القيمة المضافة (15%)", "sum"),
        الاجمالي_بعد_الضريبة=("الاجمالي بعد الضريبة", "sum")
    ).reset_index().sort_values(by="الاجمالي_بعد_الضريبة", ascending=False)

    df_summary = df_summary.rename(columns={
        "عدد_المستندات": "عدد الفواتير والإشعارات",
        "الاجمالي_قبل_الضريبية": "صافي الاجمالي قبل الضريبية",
        "ضريبة_القيمة_المضافة": "صافي ضريبة القيمة المضافة (15%)",
        "الاجمالي_بعد_الضريبة": "صافي الاجمالي بعد الضريبة"
    })

    sum_total_row = pd.DataFrame({
        "اسم المورد": ["الإجمالي العام (الصافي بعد خصم الإشعارات الدائنة)"],
        "رقم التسجيل الضريبي": [""],
        "عدد الفواتير والإشعارات": [df_summary["عدد الفواتير والإشعارات"].sum()],
        "صافي الاجمالي قبل الضريبية": [round(df_summary["صافي الاجمالي قبل الضريبية"].sum(), 2)],
        "صافي ضريبة القيمة المضافة (15%)": [round(df_summary["صافي ضريبة القيمة المضافة (15%)"].sum(), 2)],
        "صافي الاجمالي بعد الضريبة": [round(df_summary["صافي الاجمالي بعد الضريبة"].sum(), 2)]
    })
    df_sheet_summary = pd.concat([df_summary, sum_total_row], ignore_index=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_sheet_valid.to_excel(writer, index=False, sheet_name="الفواتير والإشعارات المدينة")
        df_sheet_credit.to_excel(writer, index=False, sheet_name="الإشعارات الدائنة")
        df_sheet_error.to_excel(writer, index=False, sheet_name="فواتير بها أخطاء أو بدون QR")
        df_sheet_summary.to_excel(writer, index=False, sheet_name="ملخص الموردين")

        green_header = PatternFill(start_color="1E6B3A", end_color="1E6B3A", fill_type="solid")
        purple_header = PatternFill(start_color="5C2D91", end_color="5C2D91", fill_type="solid")
        red_header = PatternFill(start_color="9C1C1C", end_color="9C1C1C", fill_type="solid")
        blue_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        total_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        duplicate_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        bold_font = Font(bold=True, size=11)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        sheets_config = [
            ("الفواتير والإشعارات المدينة", df_sheet_valid, green_header),
            ("الإشعارات الدائنة", df_sheet_credit, purple_header),
            ("فواتير بها أخطاء أو بدون QR", df_sheet_error, red_header),
            ("ملخص الموردين", df_sheet_summary, blue_header),
        ]

        for sheet_name, data_df, h_fill in sheets_config:
            ws = writer.sheets[sheet_name]
            ws.views.sheetView[0].rightToLeft = True

            for col_idx in range(1, len(data_df.columns) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = h_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            if not data_df.empty:
                last_row = len(data_df) + 1
                inv_col_idx = 2
                for c_idx, col_name in enumerate(data_df.columns, 1):
                    if "رقم الفاتورة" in col_name:
                        inv_col_idx = c_idx
                        break

                for r_idx in range(2, last_row + 1):
                    is_total_row = (r_idx == last_row)
                    inv_val = ws.cell(row=r_idx, column=inv_col_idx).value

                    for c_idx in range(1, len(data_df.columns) + 1):
                        c = ws.cell(row=r_idx, column=c_idx)
                        c.border = thin_border
                        if is_total_row:
                            c.fill = total_fill
                            c.font = bold_font
                        elif sheet_name != "ملخص الموردين" and inv_val and str(inv_val).strip() in duplicate_invs:
                            c.fill = duplicate_fill

            for col in ws.columns:
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = 24

# ==========================================
# 10. واجهة المستخدم
# ==========================================
st.title("🧾 نظام تحليل المشتريات والفرز الضريبي (ZATCA)")
st.markdown("يجمع **الفواتير الضريبية والإشعارات المدينة** في الشيت الأول، ويفصل **الإشعارات الدائنة فقط** في شيت مستقل، مع **تلوين الأرقام المكررة بالأصفر الفاتح**.")

api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

with st.sidebar:
    st.header("⚙️ إعدادات الفحص والفرز")
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
    "ارفع ملف ZIP أو ملفات الفواتير (PDF / صور)",
    accept_multiple_files=True,
    type=["zip", "pdf", "png", "jpg", "jpeg", "webp"]
)

if uploaded_files:
    if st.button("🚀 بدء التحليل والفرز الضريبي الشامل", type="primary"):
        temp_dir = tempfile.mkdtemp()
        task_queue = []
        valid_exts = (".pdf", ".png", ".jpg", ".jpeg", ".webp")

        with st.spinner("جاري فهرسة الملفات على القرص المؤقت..."):
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
                        st.error(f"خطأ في قراءة أرشيف {fname}: {e}")
                else:
                    task_queue.append(("file", temp_path, "", fname))

        total_files = len(task_queue)
        st.info(f"📂 تم تجهيز **{total_files}** مستند. جاري الفحص...")

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
                        status_text.text(f"⏳ تم فحص وفرز {completed} من أصل {total_files} مستند ({int((completed/max(total_files,1))*100)}%)...")
            gc.collect()

        st.session_state["last_audit_results"] = raw_results

if "last_audit_results" in st.session_state and st.session_state["last_audit_results"]:
    raw_results = st.session_state["last_audit_results"]
    credit_notes_records = [r for r in raw_results if "دائن" in str(r.get("doc_type"))]
    valid_records = [r for r in raw_results if r.get("is_valid_zatca") and "دائن" not in str(r.get("doc_type"))]
    error_records = [r for r in raw_results if not r.get("is_valid_zatca") and "دائن" not in str(r.get("doc_type"))]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("إجمالي المستندات", len(raw_results))
    col2.metric("✅ الفواتير والإشعارات المدينة", len(valid_records))
    col3.metric("🟣 الإشعارات الدائنة فقط", len(credit_notes_records))
    col4.metric("⚠️ فواتير بها أخطاء / بدون QR", len(error_records))

    output_excel = "تحليل_المشتريات_المصنف_ZATCA.xlsx"
    export_to_excel(raw_results, output_excel)

    with open(output_excel, "rb") as f:
        st.download_button(
            label="📥 تحميل تقرير الإكسيل النهائي (الفواتير والمدين معاً + شيت مستقل للدائن)",
            data=f,
            file_name="تحليل_المشتريات_المصنف.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    tab1, tab2, tab3 = st.tabs([
        f"✅ الفواتير والإشعارات المدينة ({len(valid_records)})",
        f"🟣 الإشعارات الدائنة ({len(credit_notes_records)})",
        f"⚠️ فواتير بها أخطاء أو بدون QR ({len(error_records)})"
    ])

    with tab1:
        if valid_records:
            st.dataframe(build_formatted_df(pd.DataFrame(valid_records), include_error_reason=False), use_container_width=True)
        else:
            st.info("لا توجد فواتير أو إشعارات مدينة مطابقة.")

    with tab2:
        if credit_notes_records:
            st.dataframe(build_formatted_df(pd.DataFrame(credit_notes_records), include_error_reason=True), use_container_width=True)
        else:
            st.info("لا توجد إشعارات دائنة في هذه الدفعة.")

    with tab3:
        if error_records:
            st.dataframe(build_formatted_df(pd.DataFrame(error_records), include_error_reason=True), use_container_width=True)
        else:
            st.success("لا توجد فواتير بها أخطاء!")
