import streamlit as st
import os
import re
import io
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

st.set_page_config(page_title="نظام تدقيق وتحليل المشتريات والفواتير الضريبية", page_icon="🧾", layout="wide")

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
# الرقم الضريبي للعميل (شركة غلوبيال كوليناري فينتشرز) لاستبعاده عند البحث عن الرقم الضريبي للمورد
CUSTOMER_VAT_NUMBERS = {"313066974700003"}

# قاموس لتوحيد وتعريب أسماء الموردين المشهورين كما في شيت تحليل المشتريات
KNOWN_VENDORS_BY_VAT = {
    "300067077700003": "شركة مطاعم ركن هاشم",
    "301317075900003": "شركة الفطائر الأولى المحدودة",
    "300914908810003": "دار الضيافة (Hospitality Board Co)",
    "300056457900003": "Herfy Food Services Co (شركة هرفي للخدمات الغذائية)",
    "311131637400003": "مطابع جوهرة الخليج للطباعة",
    "311233377500003": "شركة سلاتة لتقديم الوجبات",
    "300048882800003": "شركة باب العالمية للخدمات المتخصصة",
    "314127719200003": "شركة عرين التميز التجارية",
    "310131935400003": "شركة أفضل الطهاة لتقديم الوجبات",
    "310069655110003": "شركة هنجرستيشن المحدودة",
    "310069655100003": "شركة هنجرستيشن المحدودة",
    "310191627410003": "شركة جاهز الدولية لتقنية نظم المعلومات",
    "311024805300003": "شركة اختصار الزمن لنقل الطرود (مستر مندوب)",
    "310228833400003": "شركة تطبيق مرسول للاتصالات وتقنية المعلومات",
    "310480075110003": "شركة نينجا فود",
    "302008852600003": "أرام ميم ذ م م",
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
                tlv_data[tag] = value_bytes.decode("utf-8", errors="ignore").strip()
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

def try_decode_cv2_qr(img_bgr) -> dict:
    if not HAS_CV2 or img_bgr is None:
        return None
    detector = cv2.QRCodeDetector()
    # إضافة إطار أبيض حول الصورة لمساعدة القارئ في حال كان الـ QR ملاصقاً للحواف
    bordered = cv2.copyMakeBorder(img_bgr, 30, 30, 30, 30, cv2.BORDER_CONSTANT, value=[255, 255, 255])
    gray = cv2.cvtColor(bordered, cv2.COLOR_BGR2GRAY)

    for scale in [1.0, 1.5, 2.0, 0.75]:
        try:
            resized = cv2.resize(gray, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC) if scale != 1.0 else gray
            for thresh_val in [None, 150, 128, 180]:
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

def scan_pdf_for_qr(doc, max_pages_to_scan=4) -> dict:
    """يبحث عن رمز الـ QR في الصور المدمجة أولاً ثم في المسح الضوئي لصفحات الفاتورة"""
    if not HAS_CV2:
        return None
    
    pages_idx = list(range(min(len(doc), max_pages_to_scan)))
    if len(doc) > max_pages_to_scan and (len(doc) - 1) not in pages_idx:
        pages_idx.append(len(doc) - 1)

    # 1. فحص الصور المدمجة داخل صفحات الـ PDF مباشرة (دقة 100% وسرعة فائقة)
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
                    res = try_decode_cv2_qr(img_bgr)
                    if res:
                        return res
        except Exception:
            pass

    # 2. إذا كان الـ QR مرسوماً كخطوط أو شرائط (مثل فاتورة الفطائر الأولى)، نقوم بتصوير الصفحة بدقة عالية وفحصها
    for p_idx in pages_idx:
        page = doc[p_idx]
        for dpi in [200, 280]:
            try:
                pix = page.get_pixmap(dpi=dpi)
                img_np = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                img_bgr = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR if pix.n == 3 else cv2.COLOR_RGBA2BGR)
                res = try_decode_cv2_qr(img_bgr)
                if res:
                    return res
                # فحص الأركان الأربعة للصفحة بشكل مكبر
                h, w = img_bgr.shape[:2]
                crops = [
                    img_bgr[0:int(h*0.45), 0:int(w*0.45)],
                    img_bgr[0:int(h*0.45), int(w*0.55):w],
                    img_bgr[int(h*0.55):h, 0:int(w*0.45)],
                    img_bgr[int(h*0.55):h, int(w*0.55):w],
                ]
                for crop in crops:
                    res = try_decode_cv2_qr(crop)
                    if res:
                        return res
            except Exception:
                continue
    return None

# ==========================================
# 4. دوال تحديد نوع المستند وتنسيق التواريخ
# ==========================================
def detect_document_type(text: str, filename: str) -> str:
    combined = f"{filename}\n{text[:3000]}".lower()
    if any(w in combined for w in ["credit note", "إشعار دائن", "اشعار دائن", "مرتجع", "credit memo"]):
        return "إشعار دائن"
    if any(w in combined for w in ["debit note", "إشعار مدين", "اشعار مدين", "debit memo"]):
        return "إشعار مدين"
    return "فاتورة ضريبية"

def normalize_date_str(date_str: str) -> str:
    if not date_str:
        return ""
    date_str = str(date_str).strip()
    # إذا كان بصيغة YYYY-MM-DD أو YYYY/MM/DD
    m1 = re.search(r'(202\d)[-/](\d{1,2})[-/](\d{1,2})', date_str)
    if m1:
        return f"{m1.group(1)}-{int(m1.group(2)):02d}-{int(m1.group(3)):02d}"
    # إذا كان بصيغة DD/MM/YYYY أو DD.MM.YYYY أو DD-MM-YYYY
    m2 = re.search(r'(\d{1,2})[-/\.](\d{1,2})[-/\.](202\d)', date_str)
    if m2:
        d, m, y = int(m2.group(1)), int(m2.group(2)), m2.group(3)
        if m > 12:
            d, m = m, d
        return f"{y}-{m:02d}-{d:02d}"
    # إذا كان بصيغة مثل October 8, 2026
    for fmt in ["%B %d, %Y", "%b %d, %Y", "%d %B %Y", "%d %b %Y"]:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime("%Y-%m-%d")
        except Exception:
            pass
    return date_str

# ==========================================
# 5. المحرك المحلي الذكي لاستخراج بيانات الفواتير
# ==========================================
def local_smart_extract(text: str, filename: str, qr_info: dict = None) -> dict:
    doc_type = detect_document_type(text, filename)
    data = {
        "invoice_date": "",
        "invoice_number": "",
        "doc_type": doc_type,
        "vendor_name": "",
        "tax_id": "",
        "total_before_vat": 0.0,
        "vat_value": 0.0,
        "total_after_vat": 0.0
    }

    # 1. الاستفادة من بيانات الـ QR المؤكدة (إن وجدت)
    if qr_info and qr_info.get("qr_total") is not None:
        qr_tax_id = qr_info.get("qr_tax_id", "")
        if qr_tax_id and qr_tax_id not in CUSTOMER_VAT_NUMBERS:
            data["tax_id"] = qr_tax_id
        data["vendor_name"] = qr_info.get("qr_vendor", "")
        data["invoice_date"] = normalize_date_str(qr_info.get("qr_date", ""))
        data["total_after_vat"] = round(float(qr_info.get("qr_total", 0.0)), 2)
        data["vat_value"] = round(float(qr_info.get("qr_tax", 0.0)), 2)
        if data["total_after_vat"] > 0:
            data["total_before_vat"] = round(data["total_after_vat"] - data["vat_value"], 2)

    # 2. استخراج الرقم الضريبي للمورد (15 رقماً يبدأ بـ 3) مع استبعاد الرقم الضريبي للعميل
    if not data["tax_id"]:
        all_vats = re.findall(r'\b(3\d{14})\b', text)
        vendor_vats = [v for v in all_vats if v not in CUSTOMER_VAT_NUMBERS]
        if vendor_vats:
            data["tax_id"] = vendor_vats[0]

    # 3. مطابقة اسم المورد من القاموس المعتمد إذا توفر رقمه الضريبي
    if data["tax_id"] in KNOWN_VENDORS_BY_VAT:
        data["vendor_name"] = KNOWN_VENDORS_BY_VAT[data["tax_id"]]

    # 4. استخراج تاريخ الفاتورة من النص إن لم يوجد أو لتأكيده
    if not data["invoice_date"]:
        # بحث عن صيغة October 8, 2026
        eng_date = re.search(r'((?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+202\d)', text, re.IGNORECASE)
        if eng_date:
            data["invoice_date"] = normalize_date_str(eng_date.group(1))
        else:
            date_matches = re.findall(r'(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/\.]\d{2}[-/\.]\d{4})', text)
            for dm in date_matches:
                norm_d = normalize_date_str(dm)
                if norm_d.startswith("202"):
                    data["invoice_date"] = norm_d
                    break
        if not data["invoice_date"]:
            fname_date = re.search(r'(202\d[-_]?\d{2}[-_]?\d{2})', filename)
            if fname_date:
                raw_fd = re.sub(r'[^\d]', '', fname_date.group(1))
                if len(raw_fd) == 8:
                    data["invoice_date"] = f"{raw_fd[:4]}-{raw_fd[4:6]}-{raw_fd[6:8]}"

    # 5. استخراج رقم الفاتورة بدقة لجميع النماذج المرفقة
    inv_patterns = [
        r'(PSI/\d{4}/\d+)',
        r'(SA-B2B-INV-\d+)',
        r'(SINV\d+)',
        r'(HS-(?:CPC-|ADS-)?\d{4}/\d+-\d+)',
        r'(INV/\d{4}/[\d/]+)',
        r'(\d-\w-\d+)',
        r'Invoice\s*Number\s*[:\|]?\s*([A-Za-z0-9\-/]+)',
        r'INVOICE\s+(\d{3,10})\b',
        r'Document\s*Number[\s\S]{1,40}?(\d{7,10})\b',
        r'Invoice\s*No\.?\s*[:\|]?\s*([A-Za-z0-9\-/]{3,25})',
    ]
    for pat in inv_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            cand = m.group(1).strip()
            if cand.lower() not in {"date", "reference", "period", "summary", "details", "total", "tax"} and cand not in CUSTOMER_VAT_NUMBERS:
                data["invoice_number"] = cand
                break

    # حالة خاصة لفاتورة هرفي (600182) ومستر مندوب (7028)
    if not data["invoice_number"]:
        if "300056457900003" in text or "Herfy" in text:
            hm = re.search(r'6000169\s+(\d{6})', text)
            if hm:
                data["invoice_number"] = hm.group(1)
        elif "MrMandoob" in text or "311024805300003" in text:
            mm = re.search(r'Invoice\s+No\.\s*(?:Invoice\s+No\.)?\s*(\d{4,8})', text)
            if mm:
                data["invoice_number"] = mm.group(1)

    # 6. استخراج المبالغ المالية بدقة (قبل الضريبة، الضريبة 15%، بعد الضريبة) إذا لم تكن مستخرجة من الـ QR
    if data["total_after_vat"] == 0.0:
        # أ) فواتير هنجرستيشن المجمعة (Total Invoice: 87,396.75 | 13,109.51 | 100,506.26 SAR)
        hs_match = re.search(r'Total\s+Invoice\s+(?:ةرﻮﺗﺎﻔﻟا\s+ﻲﻟﺎﻤﺟإ|اجمالي\s+الفاتورة)?\s*([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*SAR', text, re.IGNORECASE)
        if hs_match:
            data["total_before_vat"] = float(hs_match.group(1).replace(',', ''))
            data["vat_value"] = float(hs_match.group(2).replace(',', ''))
            data["total_after_vat"] = float(hs_match.group(3).replace(',', ''))

        # ب) فواتير مستر مندوب أو إعلانات هنجرستيشن (Total Invoice Before VAT / After VAT)
        if data["total_after_vat"] == 0.0:
            tb_m = re.search(r'(?:Total\s+Invoice\s+(?:Value\s+)?Before\s+VAT|Total\s+Before\s+Tax|Total\s+SAR\s+Excl\.\s+VAT|Total\s+Amount\s+Excl\.\s+VAT|SUBTOTAL)[\s:.\-\|SAR]*([\d,]+\.\d{2})', text, re.IGNORECASE)
            vat_m = re.search(r'(?:VAT\s+Value\s*\(\s*15\.?0?\s*%\s*\)|VAT\s*15\.?0?%|15%\s*VAT|Total\s+Tax\s+Amount|Total\s+VAT\s+Amount\s*15%|VAT\s+Total|\bTAX\b)[\s:.\-\|SAR]*([\d,]+\.\d{2})', text, re.IGNORECASE)
            ta_m = re.search(r'(?:Total\s+Invoice\s+(?:Value\s+)?After\s+VAT|Total\s+SAR\s+Incl\.\s+VAT|Total\s+Amount\s+Incl\.\s+VAT|Total\s+Amount\s+Due|Grand\s+Total|Total\s+Amount|\bTOTAL\b)[\s:.\-\|SAR]*([\d,]+\.\d{2})', text, re.IGNORECASE)

            if tb_m:
                data["total_before_vat"] = float(tb_m.group(1).replace(',', ''))
            if vat_m:
                data["vat_value"] = float(vat_m.group(1).replace(',', ''))
            if ta_m:
                data["total_after_vat"] = float(ta_m.group(1).replace(',', ''))

        # ج) البحث الذكي عن أي ثلاثية أرقام تحقق معادلة (قبل الضريبة + ضريبة 15% = الإجمالي بعد الضريبة)
        if data["total_after_vat"] == 0.0:
            amounts = re.findall(r'\b\d{1,3}(?:,\d{3})*\.\d{2}\b', text)
            floats = sorted(list(set([round(float(a.replace(',', '')), 2) for a in amounts if float(a.replace(',', '')) > 0])))
            for total_c in reversed(floats):
                for sub_c in reversed(floats):
                    if sub_c < total_c:
                        tax_c = round(total_c - sub_c, 2)
                        if abs(tax_c - round(sub_c * 0.15, 2)) <= 0.15 and any(abs(tax_c - f) <= 0.05 for f in floats):
                            data["total_before_vat"] = sub_c
                            data["vat_value"] = tax_c
                            data["total_after_vat"] = total_c
                            break
                if data["total_after_vat"] > 0:
                    break

    # استكمال أي طرف ناقص في المعادلة المحاسبية
    if data["total_after_vat"] > 0 and data["total_before_vat"] == 0:
        if data["vat_value"] > 0:
            data["total_before_vat"] = round(data["total_after_vat"] - data["vat_value"], 2)
        else:
            data["total_before_vat"] = round(data["total_after_vat"] / 1.15, 2)
            data["vat_value"] = round(data["total_after_vat"] - data["total_before_vat"], 2)
    elif data["total_before_vat"] > 0 and data["total_after_vat"] == 0:
        if data["vat_value"] == 0:
            data["vat_value"] = round(data["total_before_vat"] * 0.15, 2)
        data["total_after_vat"] = round(data["total_before_vat"] + data["vat_value"], 2)

    # 7. اسم المورد الاحتياطي إن لم يتحدد بعد
    if not data["vendor_name"]:
        for line in [l.strip() for l in text.split('\n') if len(l.strip()) > 3][:12]:
            if any(k in line for k in ["شركة", "مؤسسة", "مطابع", "Co", "Company", "Corp"]) and "غلوبيال" not in line and "Global" not in line:
                data["vendor_name"] = line
                break
        if not data["vendor_name"]:
            data["vendor_name"] = re.sub(r'\.pdf$|\.jpg$|\.png$', '', filename, flags=re.IGNORECASE)

    return data

# ==========================================
# 6. محرك الذكاء الاصطناعي عبر REST API المباشر
# ==========================================
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
    أنت مدقق حسابات ومشتريات خبير في السعودية. قم باستخراج بيانات الفاتورة المرفقة وأرجع النتيجة بصيغة JSON فقط يحتوي على هذه المفاتيح:
    {{
      "invoice_date": "تاريخ الفاتورة بصيغة YYYY-MM-DD",
      "invoice_number": "رقم الفاتورة بدقة",
      "doc_type": "فاتورة ضريبية أو إشعار دائن أو إشعار مدين",
      "vendor_name": "اسم الشركة البائعة (المورد) وليس العميل Global Culinary Ventures",
      "tax_id": "الرقم الضريبي للمورد (15 رقماً) وليس الرقم الضريبي للعميل 313066974700003",
      "total_before_vat": 0.0,
      "vat_value": 0.0,
      "total_after_vat": 0.0
    }}
    ملاحظة هامة: إذا كانت الفاتورة لشركة توصيل مثل هنجرستيشن أو مستر مندوب، فاستخرج إجمالي عمولات الفاتورة الضريبية (Total Invoice / Total Commission) وليس مبيعات المطعم (Online Sales).
    النص المستخرج من الفاتورة:
    {text_content[:5000]}
    """
    parts = [{"text": prompt}]
    if img_bytes:
        parts.append({"inline_data": {"mime_type": "image/jpeg", "data": b64_img}})

    payload = {
        "contents": [{"parts": parts}],
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

    # فحص ومطابقة الـ QR Code بالتفصيل
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
# 8. معالجة الملفات المرفوعة
# ==========================================
def process_uploaded_file(file_bytes: bytes, filename: str, split_branches: bool):
    # تنظيف أي ترويسة HTTP زائدة في بداية ملف الـ PDF (مثل ملف مستر مندوب)
    if filename.lower().endswith(".pdf"):
        pdf_start = file_bytes.find(b"%PDF-")
        if pdf_start > 0:
            file_bytes = file_bytes[pdf_start:]

    ext = filename.lower()
    units = []
    try:
        if ext.endswith(".pdf"):
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            qr_info = scan_pdf_for_qr(doc, max_pages_to_scan=4)

            if not split_branches:
                # الوضع الافتراضي: قراءة الفاتورة كملف واحد مجمع (كما في شيت تحليل المشتريات)
                pages_to_read = list(range(min(len(doc), 5)))
                if len(doc) > 5 and (len(doc) - 1) not in pages_to_read:
                    pages_to_read.append(len(doc) - 1)

                combined_text = "\n".join([doc[i].get_text("text") for i in pages_to_read])
                pix = doc[0].get_pixmap(dpi=160)
                units.append({
                    "filename": filename,
                    "page_num": f"1-{len(doc)}" if len(doc) > 1 else "1",
                    "text_content": combined_text.strip(),
                    "img_bytes": pix.tobytes("jpeg"),
                    "qr_info": qr_info
                })
            else:
                for page_idx in range(len(doc)):
                    page = doc[page_idx]
                    pix = page.get_pixmap(dpi=150)
                    units.append({
                        "filename": filename,
                        "page_num": str(page_idx + 1),
                        "text_content": page.get_text("text").strip(),
                        "img_bytes": pix.tobytes("jpeg"),
                        "qr_info": qr_info if page_idx == 0 else None
                    })
            doc.close()
        else:
            qr_info = None
            if HAS_CV2:
                img_np = np.frombuffer(file_bytes, np.uint8)
                img_bgr = cv2.imdecode(img_np, cv2.IMREAD_COLOR)
                qr_info = try_decode_cv2_qr(img_bgr)
            units.append({
                "filename": filename,
                "page_num": "1",
                "text_content": "",
                "img_bytes": file_bytes,
                "qr_info": qr_info
            })
    except Exception as e:
        st.error(f"خطأ في فتح الملف {filename}: {e}")
    return units

def analyze_unit(unit: dict, api_key: str) -> list:
    qr_info = unit.get("qr_info")
    local_data = local_smart_extract(unit.get("text_content", ""), unit["filename"], qr_info)

    ai_data = {}
    # استدعاء الـ AI إذا كان هناك مفتاح API لاستكمال أو تأكيد الحقول غير الموجودة في الـ QR
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

    # ترتيب الأعمدة المطابق لملف تحليل المشتريات + أعمدة مطابقة الـ QR والتدقيق
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

    # إضافة صف الإجمالي في أسفل شيت تحليل المشتريات
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

    # الشيت الثاني: ملخص الموردين والإقرار الضريبي
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
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        for sheet_name, data_df in [("تحليل المشتريات", df_sheet1), ("ملخص الموردين", df_sheet2)]:
            ws = writer.sheets[sheet_name]
            ws.views.sheetView[0].rightToLeft = True

            for col_idx in range(1, len(data_df.columns) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            last_row = len(data_df) + 1
            for r_idx in range(2, last_row + 1):
                for c_idx in range(1, len(data_df.columns) + 1):
                    c = ws.cell(row=r_idx, column=c_idx)
                    c.border = thin_border
                    if r_idx == last_row:
                        c.fill = total_fill
                        c.font = bold_font

            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = max(max_len + 5, 16)

# ==========================================
# 10. واجهة المستخدم (Streamlit UI)
# ==========================================
st.title("🧾 نظام تحليل المشتريات وتدقيق الفواتير الضريبية (ZATCA QR + AI)")
st.markdown("يستخرج **التاريخ، رقم الفاتورة، نوع المستند، اسم المورد، الرقم الضريبي، الإجمالي قبل الضريبة، الضريبة (15%)، والإجمالي بعد الضريبة** مع **مطابقة الـ QR Code**.")

api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

with st.sidebar:
    st.header("⚙️ إعدادات الفحص")
    if not api_key:
        api_key_input = st.text_input("Gemini API Key (اختياري)", type="password")
        if api_key_input:
            api_key = api_key_input
    else:
        st.success("✓ مفتاح الذكاء الاصطناعي متصل")

    split_branches = st.checkbox("تقسيم صفحات الفروع الداخلية (مثل تفاصيل فروع هنجرستيشن)", value=False, help="اتركه غير مفعل للحصول على ملخص الفاتورة الإجمالية في سطر واحد كما في شيت تحليل المشتريات.")
    max_workers = st.slider("عدد المسارات المتوازية", 1, 10, 4)

uploaded_files = st.file_uploader("ارفع ملفات الفواتير (PDF أو صور)", accept_multiple_files=True, type=["pdf", "png", "jpg", "jpeg", "webp"])

if uploaded_files:
    if st.button("🚀 بدء التحليل ومطابقة الـ QR Code", type="primary"):
        all_units = []
        for file in uploaded_files:
            units = process_uploaded_file(file.read(), file.name, split_branches)
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
                status_text.text(f"جاري فحص ومطابقة المستند {completed} من {total}...")

        if raw_results:
            st.success(f"🎉 تم تحليل وتدقيق {len(raw_results)} مستند بنجاح!")
            
            output_excel = "تحليل_المشتريات_الضريبي.xlsx"
            export_to_excel(raw_results, output_excel)

            with open(output_excel, "rb") as f:
                st.download_button(
                    label="📥 تحميل شيت تحليل المشتريات (Excel)",
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
