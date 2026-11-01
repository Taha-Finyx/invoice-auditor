import streamlit as st
import pandas as pd
import numpy as np
import io
import zipfile
import base64
import gc
import re
import unicodedata
from datetime import datetime, date
from PIL import Image
import cv2
from pyzbar.pyzbar import decode
from pdf2image import convert_from_bytes, pdfinfo_from_bytes
from pypdf import PdfReader
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

# ==========================================
# إعدادات الصفحة وتهيئة ذاكرة الجلسة (Session State)
# ==========================================
st.set_page_config(page_title="منظومة التدقيق المحاسبي والضريبي الشاملة", layout="wide")

if 'processed_files' not in st.session_state:
    st.session_state.processed_files = set()
if 'audit_results' not in st.session_state:
    st.session_state.audit_results = []

# ==========================================
# 1. الأنماط القياسية (Pre-compiled Regex) للسرعة الفائقة
# ==========================================
INVOICE_NO_PATTERNS = [
    re.compile(r'(?:رقم الفاتورة|فاتورة رقم|رقم الفاتوره|اشعار مدين|إشعار دائن|رقم المستند)[\s:：\-]*([A-Za-z0-9\-\/]+)', re.IGNORECASE),
    re.compile(r'(?:Invoice\s*No|Invoice\s*Number|Inv\s*No|Invoice\s*#|Document\s*No|Credit\s*Note\s*No)[\s.:：\-]*([A-Za-z0-9\-\/]+)', re.IGNORECASE),
    re.compile(r'\b((?:INV|CN|DN|CRN|TAX)[\-\/][0-9A-Za-z\-\/]{2,})\b', re.IGNORECASE)
]

CLIENT_PATTERNS = [
    re.compile(r'(?:اسم العميل|العميل|السادة|المشتري|مطلوب من)[\s:：\-]*([^\n\r]{3,60})', re.IGNORECASE),
    re.compile(r'(?:Customer\s*Name|Client\s*Name|Bill\s*To|Billed\s*To|Sold\s*To)[\s.:：\-]*([^\n\r]{3,60})', re.IGNORECASE)
]

CREDIT_NOTE_PATTERN = re.compile(
    r'(?:إشعار\s*دائن|اشعار\s*دائن|اشعار\s*خصم|إشعار\s*خصم|مرتجع\s*مشتريات|فاتورة\s*مرتجع|'
    r'Credit\s*Note|Tax\s*Credit\s*Note|Credit\s*Memo|Return\s*Invoice|\bCN[\-\/]\d+)',
    re.IGNORECASE
)

DEBIT_NOTE_PATTERN = re.compile(
    r'(?:إشعار\s*مدين|اشعار\s*مدين|إشعار\s*إضافة|Debit\s*Note|Tax\s*Debit\s*Note|\bDN[\-\/]\d+)',
    re.IGNORECASE
)

# ==========================================
# 2. دوال معالجة النصوص العربية المعكوسة واستخراج حقول الـ PDF
# ==========================================
def normalize_and_fix_arabic_line(line):
    """تطبيع الحروف العربية وإصلاح الانعكاس مع الحفاظ على ترتيب الأرقام والرموز الإنجليزية"""
    if not line:
        return ""
    norm_line = unicodedata.normalize('NFKC', line)
    reversed_line = norm_line[::-1]
    fixed_reversed = re.sub(
        r'[A-Za-z0-9\-\/\.]+',
        lambda m: m.group(0)[::-1],
        reversed_line
    )
    return fixed_reversed

def prepare_searchable_pdf_text(raw_text):
    """تجهيز نص مزدوج (الأصلي + المصحح من الانعكاس) لضمان التقاط البيانات بأي اتجاه"""
    if not raw_text:
        return ""
    clean_original = unicodedata.normalize('NFKC', raw_text)
    reversed_indicators = ('ةروتاف', 'ليمعلا', 'ةكرش', 'ةبيرض', 'دروما', 'مقر', 'ةداسلا', 'نئاد')
    fixed_lines = [normalize_and_fix_arabic_line(line) for line in clean_original.splitlines()]
    fixed_text = "\n".join(fixed_lines)

    if any(ind in clean_original for ind in reversed_indicators):
        return fixed_text + "\n" + clean_original
    return clean_original + "\n" + fixed_text

def extract_invoice_text_fields(file_bytes, filename, expected_company_name="", expected_company_vat=""):
    """استخراج رقم الفاتورة واسم العميل ونص البحث بسرعة 0.01 ثانية من طبقة نص الـ PDF"""
    default_inv_no = filename.rsplit('/', 1)[-1].rsplit('.', 1)[0]

    if not filename.lower().endswith('.pdf'):
        return default_inv_no, "صورة (يتطلب فحص بصري)", ""

    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        if len(reader.pages) == 0:
            return default_inv_no, "ملف فارغ", ""

        raw_page_text = reader.pages[0].extract_text() or ""
        if len(raw_page_text.strip()) < 15:
            return default_inv_no, "PDF ممسوح ضوئياً (Scanned)", ""

        searchable_text = prepare_searchable_pdf_text(raw_page_text)
        norm_company = unicodedata.normalize('NFKC', expected_company_name.strip()) if expected_company_name else ""

        # 1. استخراج رقم الفاتورة
        extracted_inv_no = None
        for pattern in INVOICE_NO_PATTERNS:
            match = pattern.search(searchable_text)
            if match:
                candidate = match.group(1).strip()
                if len(candidate) >= 2 and candidate.lower() not in ['tax', 'date', 'vat', 'ضريبية', 'ةيبيرض', 'invoice']:
                    extracted_inv_no = candidate
                    break
        final_inv_no = extracted_inv_no if extracted_inv_no else default_inv_no

        # 2. استخراج وفحص اسم العميل
        extracted_client = None
        if norm_company and norm_company in searchable_text:
            extracted_client = f"{expected_company_name} (تطابق نصي مؤكد ✓)"
        elif expected_company_vat and expected_company_vat.strip() in searchable_text:
            extracted_client = f"{expected_company_name} (تطابق بالرقم الضريبي ✓)"
        else:
            for pattern in CLIENT_PATTERNS:
                match = pattern.search(searchable_text)
                if match:
                    extracted_client = match.group(1).strip()
                    break

        final_client = extracted_client if extracted_client else "غير واضح في النص"
        return final_inv_no, final_client, searchable_text

    except Exception:
        return default_inv_no, "خطأ في قراءة النص", ""

def detect_document_type(searchable_text, filename, inv_num, qr_gross_val):
    """تحديد نوع المستند (فاتورة مشتريات أم إشعار دائن أم إشعار مدين) وتحديد الإشارة المحاسبية"""
    combined_check = f"{filename} {inv_num} {searchable_text}"
    if qr_gross_val < 0 or CREDIT_NOTE_PATTERN.search(combined_check):
        return "إشعار دائن (Credit Note)", -1
    if DEBIT_NOTE_PATTERN.search(combined_check):
        return "إشعار مدين (Debit Note)", 1
    return "فاتورة مشتريات", 1

# ==========================================
# 3. دوال فك تشفير الـ QR وفلاتر OpenCV والضوابط الضريبية
# ==========================================
def decode_tlv_qr(base64_string):
    """فك تشفير الكيو آر كود القياسي (ZATCA TLV Base64)"""
    try:
        raw_data = base64.b64decode(base64_string.strip())
        idx = 0
        result = {}
        tag_names = {
            1: "qr_vendor_name",
            2: "qr_tax_number",
            3: "qr_date",
            4: "qr_total_with_vat",
            5: "qr_vat_amount"
        }
        while idx < len(raw_data):
            tag = raw_data[idx]
            length = raw_data[idx + 1]
            value = raw_data[idx + 2 : idx + 2 + length].decode('utf-8', errors='ignore')
            if tag in tag_names:
                result[tag_names[tag]] = value
            idx += 2 + length
        if result and ("qr_total_with_vat" in result or "qr_tax_number" in result):
            return result, "صالح (Valid TLV)"
        return {}, "غير متوافق"
    except Exception:
        return {}, "نص عادي / غير مشفر"

def smart_qr_decode(pil_img):
    """محرك قراءة QR متعدد المراحل باستخدام OpenCV لمعالجة البهتان والميلان"""
    # المرحلة 1: الفحص المباشر السريع
    decoded = decode(pil_img)
    if decoded:
        return decoded[0].data.decode('utf-8', errors='ignore'), "مباشر"

    gray = cv2.cvtColor(np.array(pil_img.convert('RGB')), cv2.COLOR_RGB2GRAY)

    # المرحلة 2: تحسين التباين (Otsu)
    _, otsu = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    decoded = decode(otsu)
    if decoded:
        del otsu, gray
        return decoded[0].data.decode('utf-8', errors='ignore'), "معالجة تباين (Otsu)"

    # المرحلة 3: التعتيب التكيفي للحبر الباهت (Adaptive) + ترميم النقاط (Erosion)
    adapt = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 21, 10)
    decoded = decode(adapt)
    if decoded:
        del adapt, otsu, gray
        return decoded[0].data.decode('utf-8', errors='ignore'), "توضيح حبر باهت (Adaptive)"

    kernel = np.ones((2, 2), np.uint8)
    eroded = cv2.erode(otsu, kernel, iterations=1)
    decoded = decode(eroded)
    if decoded:
        del eroded, adapt, otsu, gray
        return decoded[0].data.decode('utf-8', errors='ignore'), "ترميم نقاط (Morph)"

    # المرحلة 4: التكبير الفائق (Upscaling 2x) للكود الصغير
    upscaled = cv2.resize(gray, None, fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)
    _, up_otsu = cv2.threshold(upscaled, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    decoded = decode(up_otsu)
    if decoded:
        del up_otsu, upscaled, eroded, adapt, otsu, gray
        return decoded[0].data.decode('utf-8', errors='ignore'), "تكبير ذكي (2x)"

    # المرحلة 5: مصحح الميلان المدمج في OpenCV
    try:
        qr_detector = cv2.QRCodeDetector()
        val, _, _ = qr_detector.detectAndDecode(gray)
        if val:
            del up_otsu, upscaled, eroded, adapt, otsu, gray
            return val, "تصحيح ميلان (OpenCV)"
    except Exception:
        pass

    del up_otsu, upscaled, eroded, adapt, otsu, gray
    return None, "غير موجود"

def validate_saudi_vat(vat_number):
    """التحقق من أن الرقم الضريبي للمورد 15 رقماً ويبدأ وينتهي بـ 3"""
    clean_vat = str(vat_number or "").strip().replace(" ", "")
    if not clean_vat or clean_vat == "غير محدد":
        return clean_vat, False, "الرقم الضريبي للمورد مفقود"

    if re.match(r'^3\d{13}3$', clean_vat):
        return clean_vat, True, "صحيح (15 رقماً)"
    else:
        reasons = []
        if not clean_vat.isdigit():
            reasons.append("يحتوي رموز/حروف")
        if len(clean_vat) != 15:
            reasons.append(f"طوله {len(clean_vat)} بدلاً من 15")
        if not clean_vat.startswith('3') or not clean_vat.endswith('3'):
            reasons.append("لا يبدأ أو ينتهي بـ 3")
        return clean_vat, False, f"غير نظامي ({'، '.join(reasons)})"

def check_invoice_period(inv_date_str, start_date, end_date, check_enabled=True):
    """فحص توافق تاريخ الفاتورة مع الفترة الضريبية المحددة"""
    if not check_enabled:
        return "غير مفعل", None

    clean_date_str = str(inv_date_str or "").strip()[:10]
    if not clean_date_str:
        return "تاريخ مفقود", "تاريخ الفاتورة مفقود للتحقق من الفترة"

    parsed_date = None
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%Y"):
        try:
            parsed_date = datetime.strptime(clean_date_str, fmt).date()
            break
        except ValueError:
            continue

    if not parsed_date:
        return "صيغة غير معروفة", f"تعذر قراءة التاريخ ({clean_date_str})"

    if parsed_date < start_date:
        days_diff = (start_date - parsed_date).days
        return "فترة سابقة", f"فاتورة من فترة سابقة (قبل البداية بـ {days_diff} يوم - تأكد من عدم خصمها سابقاً)"
    elif parsed_date > end_date:
        days_diff = (parsed_date - end_date).days
        return "فترة لاحقة (مستبعدة)", f"خطأ قطع زمني (Cut-off): بعد نهاية الفترة بـ {days_diff} يوم (تُرحّل للإقرار القادم)"
    else:
        return "ضمن الفترة", None

# ==========================================
# 4. المعالج الرئيسي للفاتورة الواحدة (الصفحة الأولى والأخيرة + تفريغ الذاكرة)
# ==========================================
def process_single_invoice(file_bytes, filename, expected_company_name="", expected_company_vat=""):
    """معالجة شاملة لفاتورة واحدة مع حماية ذاكرة السيرفر"""
    qr_raw_text = None
    qr_page_location = "غير موجود"

    try:
        if filename.lower().endswith('.pdf'):
            pdf_info = pdfinfo_from_bytes(file_bytes)
            total_pages = int(pdf_info.get("Pages", 1))
            pages_to_check = [1]
            if total_pages > 1:
                pages_to_check.append(total_pages)

            # المحاولة بدقة 150 DPI السريعة، وإذا لم يُعثر على الكود نرفع الدقة إلى 300 DPI
            for current_dpi in [150, 300]:
                for page_num in pages_to_check:
                    imgs = convert_from_bytes(file_bytes, dpi=current_dpi, first_page=page_num, last_page=page_num)
                    if imgs:
                        img = imgs[0]
                        qr_raw_text, method_used = smart_qr_decode(img)
                        img.close()
                        del imgs
                        if qr_raw_text:
                            page_label = "الأولى" if page_num == 1 else f"الأخيرة ({page_num})"
                            qr_page_location = f"ص {page_label} [{method_used} - {current_dpi}dpi]"
                            break
                if qr_raw_text:
                    break
        else:
            img = Image.open(io.BytesIO(file_bytes))
            qr_raw_text, method_used = smart_qr_decode(img)
            img.close()
            del img
            if qr_raw_text:
                qr_page_location = f"صورة [{method_used}]"

    except Exception as e:
        return {
            "اسم الملف": filename.rsplit('/', 1)[-1], "نوع المستند": "خطأ قراءة",
            "رقم الفاتورة": "خطأ", "تاريخ الفاتورة": "", "اسم المورد": "",
            "الرقم الضريبي للمورد": "", "حالة الرقم الضريبي": "خطأ", "اسم العميل": "",
            "الإجمالي قبل الضريبة": 0.0, "ضريبة القيمة المضافة": 0.0, "الإجمالي شامل الضريبة": 0.0,
            "حالة الـ QR": f"ملف تالف ({str(e)[:25]})", "موقع الـ QR": "خطأ",
            "إجمالي الـ QR": 0.0, "ضريبة الـ QR": 0.0
        }

    if qr_raw_text:
        tlv_data, qr_status = decode_tlv_qr(qr_raw_text)
        try:
            gross_val = float(tlv_data.get("qr_total_with_vat", 0) or 0)
        except ValueError:
            gross_val = 0.0
        try:
            vat_val = float(tlv_data.get("qr_vat_amount", 0) or 0)
        except ValueError:
            vat_val = 0.0
        vendor = tlv_data.get("qr_vendor_name", "غير محدد")
        vendor_vat = tlv_data.get("qr_tax_number", "غير محدد")
        inv_date = tlv_data.get("qr_date", "")[:10]
    else:
        qr_status = "مفقود أو غير واضح"
        gross_val, vat_val = 0.0, 0.0
        vendor, vendor_vat, inv_date = "يتطلب مراجعة يدوية", "غير محدد", ""

    clean_vat, _, vat_status_msg = validate_saudi_vat(vendor_vat)
    inv_num, client_name, searchable_text = extract_invoice_text_fields(
        file_bytes, filename, expected_company_name, expected_company_vat
    )
    doc_type, sign_multiplier = detect_document_type(searchable_text, filename, inv_num, gross_val)

    gross_val = round(abs(gross_val) * sign_multiplier, 2)
    vat_val = round(abs(vat_val) * sign_multiplier, 2)
    net_val = round(gross_val - vat_val, 2)

    return {
        "اسم الملف": filename.rsplit('/', 1)[-1],
        "نوع المستند": doc_type,
        "رقم الفاتورة": inv_num,
        "تاريخ الفاتورة": inv_date,
        "اسم المورد": vendor,
        "الرقم الضريبي للمورد": clean_vat,
        "حالة الرقم الضريبي": vat_status_msg,
        "اسم العميل": client_name,
        "الإجمالي قبل الضريبة": net_val,
        "ضريبة القيمة المضافة": vat_val,
        "الإجمالي شامل الضريبة": gross_val,
        "حالة الـ QR": qr_status,
        "موقع الـ QR": qr_page_location,
        "إجمالي الـ QR": gross_val,
        "ضريبة الـ QR": vat_val
    }

# ==========================================
# 5. محرك التدقيق المحاسبي وتجميع الموردين
# ==========================================
def perform_accounting_audit(df, vat_rate=0.15, start_date=None, end_date=None, check_period=True):
    """تطبيق جميع الضوابط المحاسبية والضريبية على جدول البيانات"""
    if df.empty:
        return df

    audit_notes = []
    status_list = []
    period_status_list = []
    duplicates = df.duplicated(subset=['اسم المورد', 'تاريخ الفاتورة', 'الإجمالي شامل الضريبة'], keep=False)

    for index, row in df.iterrows():
        notes = []
        net = float(row.get('الإجمالي قبل الضريبة', 0) or 0)
        vat = float(row.get('ضريبة القيمة المضافة', 0) or 0)
        gross = float(row.get('الإجمالي شامل الضريبة', 0) or 0)
        qr_status = str(row.get('حالة الـ QR', ''))
        inv_date = str(row.get('تاريخ الفاتورة', ''))
        vendor_vat = str(row.get('الرقم الضريبي للمورد', ''))

        # 1. فحص الفترة الضريبية
        p_status, p_note = check_invoice_period(inv_date, start_date, end_date, check_period)
        period_status_list.append(p_status)
        if p_note:
            notes.append(p_note)

        # 2. فحص الرقم الضريبي للمورد
        _, is_vat_valid, vat_err_msg = validate_saudi_vat(vendor_vat)
        if abs(vat) > 0 and not is_vat_valid:
            notes.append(f"خطر ضريبي: {vat_err_msg}")

        # 3. فحص صلاحية الـ QR والاتساق الحسابي
        if "صالح" not in qr_status:
            notes.append("الـ QR مفقود أو غير صالح")
        if abs((net + vat) - gross) > 0.05:
            notes.append("خطأ جمع أفقي")

        expected_vat = round(net * vat_rate, 2)
        if abs(net) > 0 and abs(vat - expected_vat) > 0.15:
            notes.append(f"نسبة الضريبة غير مطابقة لـ {int(vat_rate*100)}%")

        # 4. فحص التكرار
        if abs(gross) > 0 and duplicates.iloc[index]:
            notes.append("اشتباه تكرار فاتورة (نفس المورد والتاريخ والمبلغ)")

        if not notes:
            status_list.append("مطابق")
            audit_notes.append("سليمة ومطابقة")
        else:
            status_list.append("يوجد ملاحظات")
            audit_notes.append(" | ".join(notes))

    df['موقف الفترة الضريبية'] = period_status_list
    df['حالة التدقيق'] = status_list
    df['ملاحظات الفحص المحاسبي'] = audit_notes
    return df

def build_vendor_summary(df):
    """بناء جدول ملخص الموردين لمطابقة كشوف الحسابات"""
    if df.empty:
        return pd.DataFrame()

    temp_df = df.copy()
    temp_df['اسم المورد'] = temp_df['اسم المورد'].fillna('غير محدد').replace('', 'غير محدد')
    temp_df['الرقم الضريبي للمورد'] = temp_df['الرقم الضريبي للمورد'].fillna('غير محدد').replace('', 'غير محدد')

    vendor_rows = []
    grouped = temp_df.groupby(['اسم المورد', 'الرقم الضريبي للمورد'], as_index=False)

    for (vendor_name, vendor_vat), group in grouped:
        purchases = group[group['الإجمالي شامل الضريبة'] >= 0]
        returns = group[group['الإجمالي شامل الضريبة'] < 0]
        _, is_valid_vat, _ = validate_saudi_vat(vendor_vat)

        vendor_rows.append({
            "اسم المورد": vendor_name,
            "الرقم الضريبي للمورد": str(vendor_vat),
            "صحة الرقم الضريبي": "صحيح (15 رقماً)" if is_valid_vat else "غير نظامي / مفقود",
            "إجمالي عدد المستندات": len(group),
            "عدد فواتير المشتريات (+)": len(purchases),
            "عدد إشعارات الدائن (-)": len(returns),
            "إجمالي المشتريات (قبل الضريبة)": round(purchases['الإجمالي قبل الضريبة'].sum(), 2),
            "إجمالي المرتجعات (قبل الضريبة)": round(returns['الإجمالي قبل الضريبة'].sum(), 2),
            "صافي المشتريات (قبل الضريبة)": round(group['الإجمالي قبل الضريبة'].sum(), 2),
            "صافي ضريبة القيمة المضافة": round(group['ضريبة القيمة المضافة'].sum(), 2),
            "صافي مستحق المورد (شامل الضريبة)": round(group['الإجمالي شامل الضريبة'].sum(), 2),
            "فواتير بها ملاحظات للمراجعة": len(group[group['حالة التدقيق'] != 'مطابق'])
        })

    summary_df = pd.DataFrame(vendor_rows)
    return summary_df.sort_values(by="صافي مستحق المورد (شامل الضريبة)", ascending=False).reset_index(drop=True)

# ==========================================
# 6. محرك تصدير الإكسيل ذو التبويبات الخمسة (5-Tab Excel Engine)
# ==========================================
def write_invoices_sheet(writer, workbook, df_subset, sheet_name, vat_rate, header_fmt, money_fmt, text_fmt, ok_fmt, err_fmt):
    """كتابة وتنسيق تبويب الفواتير مع معادلات إكسيل الحية وصف الإجمالي"""
    ws = workbook.add_worksheet(sheet_name)
    ws.right_to_left()

    if df_subset.empty:
        ws.write('B2', f'لا توجد سجلات في هذا التبويب ({sheet_name})', header_fmt)
        ws.set_column('B:B', 50)
        return

    headers = list(df_subset.columns) + ['فحص الإجمالي (معادلة إكسيل)', 'فرق الضريبة (معادلة إكسيل)']
    for col_num, col_name in enumerate(headers):
        ws.write(1, col_num, col_name, header_fmt)

    money_columns = ['الإجمالي قبل الضريبة', 'ضريبة القيمة المضافة', 'الإجمالي شامل الضريبة', 'إجمالي الـ QR', 'ضريبة الـ QR']
    net_col_letter = xl_col_to_name(df_subset.columns.get_loc('الإجمالي قبل الضريبة'))
    vat_col_letter = xl_col_to_name(df_subset.columns.get_loc('ضريبة القيمة المضافة'))
    gross_col_letter = xl_col_to_name(df_subset.columns.get_loc('الإجمالي شامل الضريبة'))

    for row_num in range(len(df_subset)):
        excel_row = row_num + 2
        r = excel_row + 1
        for col_num, col_name in enumerate(df_subset.columns):
            val = df_subset.iloc[row_num, col_num]
            if col_name in money_columns:
                ws.write_number(excel_row, col_num, float(val or 0), money_fmt)
            else:
                ws.write_string(excel_row, col_num, str(val if pd.notna(val) else ""), text_fmt)

        f_total = f"=ROUND(({net_col_letter}{r}+{vat_col_letter}{r})-{gross_col_letter}{r}, 2)"
        f_vat = f"=ROUND({vat_col_letter}{r}-({net_col_letter}{r}*{vat_rate}), 2)"
        ws.write_formula(excel_row, len(df_subset.columns), f_total, money_fmt)
        ws.write_formula(excel_row, len(df_subset.columns) + 1, f_vat, money_fmt)

    last_data_row = len(df_subset) + 2
    total_row_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 2, 'num_format': '#,##0.00', 'align': 'center'})
    total_label_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 2, 'align': 'center'})

    ws.write(last_data_row, 0, f"الإجمالي ({len(df_subset)} مستند)", total_label_fmt)
    for col_num in range(1, len(headers)):
        col_name = headers[col_num]
        if col_name in money_columns:
            col_let = xl_col_to_name(col_num)
            ws.write_formula(last_data_row, col_num, f"=SUM({col_let}3:{col_let}{last_data_row})", total_row_fmt)
        else:
            ws.write(last_data_row, col_num, "", total_label_fmt)

    if 'حالة التدقيق' in df_subset.columns:
        status_col_idx = df_subset.columns.get_loc('حالة التدقيق')
        ws.conditional_format(2, status_col_idx, last_data_row - 1, status_col_idx, {
            'type': 'cell', 'criteria': 'equal to', 'value': '"مطابق"', 'format': ok_fmt
        })
        ws.conditional_format(2, status_col_idx, last_data_row - 1, status_col_idx, {
            'type': 'cell', 'criteria': 'not equal to', 'value': '"مطابق"', 'format': err_fmt
        })

    ws.freeze_panes(2, 0)
    ws.autofilter(1, 0, last_data_row - 1, len(headers) - 1)
    ws.set_column(0, len(headers) - 1, 20)
    if 'ملاحظات الفحص المحاسبي' in df_subset.columns:
        notes_idx = df_subset.columns.get_loc('ملاحظات الفحص المحاسبي')
        ws.set_column(notes_idx, notes_idx, 48)

def export_audit_excel(df, vat_rate=0.15):
    """تصدير ملف الإكسيل الشامل المكوّن من 5 تبويبات احترافية"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        header_fmt = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 11,
            'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        money_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'align': 'center'})
        text_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        ok_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center', 'bold': True})
        err_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center', 'bold': True})
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F4E78'})
        total_row_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 2, 'num_format': '#,##0.00', 'align': 'center'})
        total_label_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 2, 'align': 'center'})

        # فصل المقبول عن المرفوض
        if 'موقف الفترة الضريبية' in df.columns:
            accepted_mask = (df['حالة التدقيق'] == 'مطابق') & (df['موقف الفترة الضريبية'].isin(['ضمن الفترة', 'غير مفعل']))
        else:
            accepted_mask = (df['حالة التدقيق'] == 'مطابق')

        accepted_df = df[accepted_mask].reset_index(drop=True)
        rejected_df = df[~accepted_mask].reset_index(drop=True)

        # 1. التبويب الأول: السجل الشامل
        write_invoices_sheet(writer, workbook, df, '1_السجل_الشامل', vat_rate, header_fmt, money_fmt, text_fmt, ok_fmt, err_fmt)

        # 2. التبويب الثاني: الملخص التنفيذي والإقرار الضريبي
        summary_ws = workbook.add_worksheet('2_الملخص_التنفيذي')
        summary_ws.right_to_left()
        summary_ws.write('B2', 'الملخص التنفيذي لتدقيق فواتير المشتريات والإقرار الضريبي', title_fmt)

        purchases_df = df[df['الإجمالي شامل الضريبة'] >= 0]
        credits_df = df[df['الإجمالي شامل الضريبة'] < 0]
        in_period_df = df[df['موقف الفترة الضريبية'] == 'ضمن الفترة'] if 'موقف الفترة الضريبية' in df.columns else df
        prev_period_df = df[df['موقف الفترة الضريبية'] == 'فترة سابقة'] if 'موقف الفترة الضريبية' in df.columns else pd.DataFrame()
        next_period_df = df[df['موقف الفترة الضريبية'].str.contains('فترة لاحقة', na=False)] if 'موقف الفترة الضريبية' in df.columns else pd.DataFrame()

        summary_data = [
            ['إجمالي عدد المستندات المفحوصة', len(df), 'count'],
            ['عدد فواتير المشتريات وإشعارات المدين (+)', len(purchases_df), 'count'],
            ['عدد إشعارات الدائن / مرتجعات المشتريات (-)', len(credits_df), 'count'],
            ['عدد الفواتير المقبولة والجاهزة للإقرار الضريبي (التبويب 4)', len(accepted_df), 'count'],
            ['عدد الفواتير المرفوضة أو المعلقة للمراجعة (التبويب 5)', len(rejected_df), 'count'],
            ['--- تحليل القطع الزمني والفترة الضريبية (Cut-off) ---', 0, 'sep'],
            ['عدد الفواتير الواقعة ضمن الفترة الضريبية', len(in_period_df), 'count'],
            ['عدد فواتير الفترات السابقة (تحتاج تحقق من عدم خصمها سابقاً)', len(prev_period_df), 'count'],
            ['عدد فواتير الفترات اللاحقة (مستبعدة وتُرحّل للإقرار القادم)', len(next_period_df), 'count'],
            ['--- القيم المعتمدة الجاهزة للرفع في الإقرار الضريبي (من التبويب 4) ---', 0, 'sep'],
            ['صافي المشتريات المعتمدة قبل الضريبة (بعد طرح المرتجعات)', accepted_df['الإجمالي قبل الضريبة'].sum(), 'money'],
            ['صافي ضريبة المدخلات المعتمدة القابلة للاسترداد', accepted_df['ضريبة القيمة المضافة'].sum(), 'money'],
            ['الإجمالي المعتمد شامل الضريبة', accepted_df['الإجمالي شامل الضريبة'].sum(), 'money'],
            ['--- القيم المعلقة / المستبعدة لحين التصحيح (من التبويب 5) ---', 0, 'sep'],
            ['إجمالي المشتريات المستبعدة قبل الضريبة', rejected_df['الإجمالي قبل الضريبة'].sum(), 'money'],
            ['إجمالي الضريبة المستبعدة (غير قابلة للخصم حالياً)', rejected_df['ضريبة القيمة المضافة'].sum(), 'money']
        ]

        summary_ws.write('B4', 'البيان المحاسبي والضريبي', header_fmt)
        summary_ws.write('C4', 'القيمة / العدد', header_fmt)

        for i, (label, val, row_type) in enumerate(summary_data, start=4):
            if row_type == 'sep':
                summary_ws.write(i, 1, label, header_fmt)
                summary_ws.write(i, 2, '', header_fmt)
            elif row_type == 'count':
                summary_ws.write(i, 1, label, text_fmt)
                summary_ws.write_number(i, 2, int(val), text_fmt)
            else:
                summary_ws.write(i, 1, label, text_fmt)
                summary_ws.write_number(i, 2, float(val), money_fmt)

        summary_ws.set_column('B:B', 58)
        summary_ws.set_column('C:C', 25)

        # 3. التبويب الثالث: ملخص الموردين ومطابقة الكشوف
        vendor_df = build_vendor_summary(df)
        if not vendor_df.empty:
            v_ws = workbook.add_worksheet('3_ملخص_الموردين')
            v_ws.right_to_left()
            v_headers = list(vendor_df.columns)
            for col_num, col_name in enumerate(v_headers):
                v_ws.write(1, col_num, col_name, header_fmt)

            money_cols = [
                "إجمالي المشتريات (قبل الضريبة)", "إجمالي المرتجعات (قبل الضريبة)",
                "صافي المشتريات (قبل الضريبة)", "صافي ضريبة القيمة المضافة",
                "صافي مستحق المورد (شامل الضريبة)"
            ]
            count_cols = [
                "إجمالي عدد المستندات", "عدد فواتير المشتريات (+)",
                "عدد إشعارات الدائن (-)", "فواتير بها ملاحظات للمراجعة"
            ]

            for row_num in range(len(vendor_df)):
                excel_row = row_num + 2
                for col_num, col_name in enumerate(v_headers):
                    val = vendor_df.iloc[row_num, col_num]
                    if col_name in money_cols:
                        v_ws.write_number(excel_row, col_num, float(val or 0), money_fmt)
                    elif col_name in count_cols:
                        v_ws.write_number(excel_row, col_num, int(val or 0), text_fmt)
                    else:
                        v_ws.write_string(excel_row, col_num, str(val), text_fmt)

            last_v_row = len(vendor_df) + 2
            v_ws.write(last_v_row, 0, "الإجمالي العام لجميع الموردين", total_label_fmt)
            v_ws.write(last_v_row, 1, "-", total_label_fmt)
            v_ws.write(last_v_row, 2, "-", total_label_fmt)

            for col_num in range(3, len(v_headers)):
                col_let = xl_col_to_name(col_num)
                sum_f = f"=SUM({col_let}3:{col_let}{last_v_row})"
                if v_headers[col_num] in money_cols:
                    v_ws.write_formula(last_v_row, col_num, sum_f, total_row_fmt)
                else:
                    v_ws.write_formula(last_v_row, col_num, sum_f, total_label_fmt)

            v_ws.freeze_panes(2, 0)
            v_ws.autofilter(1, 0, last_v_row - 1, len(v_headers) - 1)
            v_ws.set_column(0, 0, 35)
            v_ws.set_column(1, len(v_headers) - 1, 22)

        # 4. التبويب الرابع: الفواتير المقبولة الجاهزة للإقرار الضريبي
        write_invoices_sheet(writer, workbook, accepted_df, '4_الفواتير_المقبولة_للإقرار', vat_rate, header_fmt, money_fmt, text_fmt, ok_fmt, err_fmt)

        # 5. التبويب الخامس: الفواتير المرفوضة والمستبعدة للمراجعة
        write_invoices_sheet(writer, workbook, rejected_df, '5_الفواتير_المرفوضة_للمراجعة', vat_rate, header_fmt, money_fmt, text_fmt, ok_fmt, err_fmt)

    output.seek(0)
    return output.getvalue()

# ==========================================
# 7. واجهة المستخدم (Streamlit UI)
# ==========================================
st.title("📊 منظومة التدقيق المحاسبي والضريبي الشاملة (+3000 فاتورة)")
st.markdown("فحص الدفعات (ZIP Batch) • فلاتر OpenCV للـ QR • تصحيح النص العربي المعكوس • فحص الرقم الضريبي • كشف الإشعارات الدائنة • التبويبات الخمسة في الإكسيل.")

with st.sidebar:
    st.header("🏢 بيانات الشركة والضريبة")
    company_name = st.text_input("اسم شركتك (المشتري)", value="شركة")
    company_vat = st.text_input("الرقم الضريبي لشركتك (15 رقماً)", value="300000000000003")
    vat_rate = st.number_input("نسبة ضريبة القيمة المضافة (%)", value=15.0, step=1.0) / 100.0

    st.divider()
    st.subheader("📅 الفترة الضريبية (Cut-off)")
    enable_period_check = st.checkbox("تفعيل فحص توافق الفترة الضريبية", value=True)
    col_d1, col_d2 = st.columns(2)
    period_start = col_d1.date_input("من تاريخ", value=date(2026, 10, 1))
    period_end = col_d2.date_input("إلى تاريخ", value=date(2026, 12, 31))
    if period_start > period_end:
        st.error("⚠️ تاريخ البداية يجب أن يكون قبل تاريخ النهاية!")

    st.divider()
    st.subheader("⚙️ إعدادات الدفعات والذاكرة")
    batch_size = st.slider("عدد الفواتير في الدفعة الواحدة", min_value=50, max_value=500, value=150, step=50)

    if st.button("🗑️ مسح الذاكرة والبدء من جديد", use_container_width=True):
        st.session_state.processed_files.clear()
        st.session_state.audit_results.clear()
        gc.collect()
        st.rerun()

# قسم التجربة السريعة لإثبات عمل التبويبات الخمسة فوراً بدون رفع ملفات
with st.expander("🧪 تجربة فورية: توليد تقرير إكسيل نموذجي بالتبويبات الخمسة (بدون رفع ملف ZIP)"):
    if st.button("🚀 توليد بيانات تجريبية شاملة وتحميل الإكسيل"):
        st.session_state.audit_results = [
            {
                "اسم الملف": "INV_001.pdf", "نوع المستند": "فاتورة مشتريات", "رقم الفاتورة": "INV-2026-101",
                "تاريخ الفاتورة": "2026-10-15", "اسم المورد": "شركة التوريدات الحديثة",
                "الرقم الضريبي للمورد": "300123456700003", "حالة الرقم الضريبي": "صحيح (15 رقماً)",
                "اسم العميل": f"{company_name} (تطابق نصي مؤكد ✓)", "الإجمالي قبل الضريبة": 10000.00,
                "ضريبة القيمة المضافة": 1500.00, "الإجمالي شامل الضريبة": 11500.00,
                "حالة الـ QR": "صالح (Valid TLV)", "موقع الـ QR": "ص الأولى [مباشر - 150dpi]",
                "إجمالي الـ QR": 11500.00, "ضريبة الـ QR": 1500.00
            },
            {
                "اسم الملف": "CN_002.pdf", "نوع المستند": "إشعار دائن (Credit Note)", "رقم الفاتورة": "CN-2026-009",
                "تاريخ الفاتورة": "2026-11-02", "اسم المورد": "شركة التوريدات الحديثة",
                "الرقم الضريبي للمورد": "300123456700003", "حالة الرقم الضريبي": "صحيح (15 رقماً)",
                "اسم العميل": f"{company_name} (تطابق بالرقم الضريبي ✓)", "الإجمالي قبل الضريبة": -2000.00,
                "ضريبة القيمة المضافة": -300.00, "الإجمالي شامل الضريبة": -2300.00,
                "حالة الـ QR": "صالح (Valid TLV)", "موقع الـ QR": "ص الأخيرة (2) [معالجة تباين (Otsu) - 150dpi]",
                "إجمالي الـ QR": -2300.00, "ضريبة الـ QR": -300.00
            },
            {
                "اسم الملف": "INV_OUT_PERIOD.pdf", "نوع المستند": "فاتورة مشتريات", "رقم الفاتورة": "INV-2027-005",
                "تاريخ الفاتورة": "2027-01-10", "اسم المورد": "مؤسسة التقنية المتطورة",
                "الرقم الضريبي للمورد": "399999999900001", "حالة الرقم الضريبي": "غير نظامي (لا يبدأ أو ينتهي بـ 3)",
                "اسم العميل": "غير واضح في النص", "الإجمالي قبل الضريبة": 5000.00,
                "ضريبة القيمة المضافة": 750.00, "الإجمالي شامل الضريبة": 5750.00,
                "حالة الـ QR": "صالح (Valid TLV)", "موقع الـ QR": "ص الأولى [تكبير ذكي (2x) - 300dpi]",
                "إجمالي الـ QR": 5750.00, "ضريبة الـ QR": 750.00
            }
        ]
        st.rerun()

# قسم رفع ومعالجة ملف الـ ZIP الفعلي
uploaded_zip = st.file_uploader("📦 ارفع ملف الـ ZIP الذي يحتوي على الفواتير (PDF / JPG / PNG)", type=['zip'])

if uploaded_zip is not None:
    with zipfile.ZipFile(uploaded_zip, 'r') as z:
        valid_extensions = ('.pdf', '.png', '.jpg', '.jpeg')
        all_files = [
            f for f in z.namelist()
            if not f.endswith('/') and not f.startswith('__MACOSX') and f.lower().endswith(valid_extensions)
        ]

        total_files = len(all_files)
        remaining_files = [f for f in all_files if f not in st.session_state.processed_files]
        processed_count = total_files - len(remaining_files)

        col1, col2, col3 = st.columns(3)
        col1.metric("إجمالي المستندات في الملف", total_files)
        col2.metric("تم فحصها وتدقيقها", processed_count)
        col3.metric("المتبقي للفحص", len(remaining_files))

        st.progress(processed_count / total_files if total_files > 0 else 0)

        c1, c2 = st.columns(2)
        run_batch = c1.button(f"▶️ تشغيل الدفعة التالية ({min(batch_size, len(remaining_files))} فاتورة)", disabled=(len(remaining_files) == 0), use_container_width=True)
        run_all = c2.button("⏩ معالجة كل المتبقي تلقائياً (مع تفريغ الرامات)", disabled=(len(remaining_files) == 0), use_container_width=True)

        if run_batch or run_all:
            target_files = remaining_files if run_all else remaining_files[:batch_size]
            status_text = st.empty()
            sub_progress = st.progress(0)

            for idx, file_path in enumerate(target_files):
                status_text.text(f"⏳ جاري فحص وتدقيق ({idx+1} من {len(target_files)}): {file_path}")
                file_bytes = z.read(file_path)
                invoice_data = process_single_invoice(file_bytes, file_path, company_name, company_vat)

                st.session_state.audit_results.append(invoice_data)
                st.session_state.processed_files.add(file_path)

                del file_bytes
                if (idx + 1) % 25 == 0:
                    gc.collect()

                sub_progress.progress((idx + 1) / len(target_files))

            gc.collect()
            status_text.success("✅ اكتملت معالجة الدفعة بنجاح وتم تفريغ الذاكرة المؤقتة!")
            st.rerun()

# عرض النتائج وتوليد ملف الإكسيل ذي التبويبات الخمسة
if st.session_state.audit_results:
    st.divider()
    raw_df = pd.DataFrame(st.session_state.audit_results)
    audited_df = perform_accounting_audit(raw_df, vat_rate, period_start, period_end, enable_period_check)

    accepted_count = len(audited_df[(audited_df['حالة التدقيق'] == 'مطابق') & (audited_df['موقف الفترة الضريبية'].isin(['ضمن الفترة', 'غير مفعل']))])
    rejected_count = len(audited_df) - accepted_count

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("إجمالي المستندات المفحوصة", len(audited_df))
    m2.metric("الفواتير المقبولة للإقرار (تبويب 4)", accepted_count)
    m3.metric("الفواتير المرفوضة/المستبعدة (تبويب 5)", rejected_count)
    m4.metric("صافي ضريبة المدخلات المعتمدة", f"{audited_df[(audited_df['حالة التدقيق'] == 'مطابق') & (audited_df['موقف الفترة الضريبية'].isin(['ضمن الفترة', 'غير مفعل']))]['ضريبة القيمة المضافة'].sum():,.2f}")

    st.subheader("📋 معاينة سريعة لآخر 50 مستنداً تم فحصها")
    st.dataframe(audited_df.tail(50), use_container_width=True)

    excel_data = export_audit_excel(audited_df, vat_rate)
    st.download_button(
        label=f"📥 تحميل التقرير المحاسبي الشامل - 5 تبويبات إكسيل ({len(audited_df)} مستند)",
        data=excel_data,
        file_name=f"Audit_Report_{period_start}_to_{period_end}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
