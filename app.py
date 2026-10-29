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
from datetime import datetime
import pandas as pd
import numpy as np
import fitz  # PyMuPDF
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

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
# 3. دوال استخراج البيانات الذكية (فائقة السرعة)
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
    return date_str

def extract_invoice_date(text: str, filename: str) -> str:
    explicit_patterns = [
        r'(?:Invoice\s+Issue\s+Date|Invoice\s+Date|Issue\s+Date|Document\s+Date|\bDATE\b|تاريخ\s+اصدار\s+الفاتورة|تاريخ\s+الفاتورة)[\s\S]{1,50}?(\d{4}[-/]\d{2}[-/]\d{2}|\d{2}[-/\.]\d{2}[-/\.]\d{4})',
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
        r'(HS-\d{4}/\d+-\d+)',
        r'(PSI/\d{4}/\d+)',
        r'(SA-B2B-INV-\d+)',
        r'(SINV\d+)',
        r'(INV/\d{4}/\d{2}/\d+)',
        r'(INV/\d{4}/\d+)',
        r'(SCR\d+)',
        r'(\d{13})',
        r'Invoice\s*(?:No\.?|Number|#)\s*[:\|]?\s*([A-Za-z0-9\-/]+)',
        r'رقم\s+الفاتورة\s*[:\|]?\s*([A-Za-z0-9\-/]+)',
        r'INVOICE\s+(\d{3,15})\b',
        r'Document\s*Number[\s\S]{1,30}?(\d{5,15})\b',
    ]
    for pat in inv_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            cand = m.group(1).strip()
            if cand.lower() not in {"date", "total", "tax", "no", "number"} and cand not in CUSTOMER_VAT_NUMBERS:
                return cand

    fname_patterns = [
        r'(HS-[A-Za-z0-9\-/]+)',
        r'(PSI[_/]\d{4}[_/]\d+)',
        r'(INV[_/]\d{4}[_/]\d+)',
        r'(SA-B2B-INV-\d+)',
        r'(SINV\d+)',
        r'(SCR\d+)',
        r'(\d{6,15})',
    ]
    for pat in fname_patterns:
        m = re.search(pat, filename, re.IGNORECASE)
        if m:
            val = m.group(1)
            return val.replace("_", "/")
    return ""

def extract_amounts(text: str) -> tuple:
    tb, vat, ta = 0.0, 0.0, 0.0

    tb_m = re.search(r'(?:Total\s+Before\s+Tax|Total\s+Excl\.\s+VAT|SUBTOTAL|الإجمالي\s+الفرعي|المبلغ\s+بدون\s+ضريبة)[\s:.\-\|SRSAR]*([\d,]+\.\d{1,2})', text, re.IGNORECASE)
    vat_m = re.search(r'(?:VAT\s*15\.?0?%|Total\s+VAT|ضريبة\s+القيمة\s+المضافة)[\s:.\-\|SRSAR]*([\d,]+\.\d{1,2})', text, re.IGNORECASE)
    ta_m = re.search(r'(?:Total\s+Invoice\s+After\s+VAT|Total\s+Amount\s+Due|Grand\s+Total|TOTAL|الإجمالي\s+الشامل|إجمالي\s+المستحقات|Net\s+Due\s+Invoice)[\s:.\-\|SRSAR]*([\d,]+\.\d{1,2})', text, re.IGNORECASE)

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

def local_smart_extract(text: str, filename: str) -> dict:
    doc_type = detect_document_type(text, filename)
    vendors_db = load_vendors_db()
    
    all_vats = re.findall(r'\b(3\d{14})\b', text)
    vendor_vats = [v for v in all_vats if v not in CUSTOMER_VAT_NUMBERS]
    tax_id = vendor_vats[0] if vendor_vats else ""

    vendor_name = vendors_db.get(tax_id, "")
    if not vendor_name:
        for line in [l.strip() for l in text.split('\n') if len(l.strip()) > 3][:10]:
            if any(k in line for k in ["شركة", "مؤسسة", "مطابع", "Co", "Company"]) and "غلوبيال" not in line:
                vendor_name = line
                break
        if not vendor_name:
            vendor_name = re.sub(r'\.pdf$|\.jpg$|\.png$', '', filename, flags=re.IGNORECASE)

    if tax_id and vendor_name and tax_id not in CUSTOMER_VAT_NUMBERS:
        save_vendor_to_db(tax_id, vendor_name)

    inv_date = extract_invoice_date(text, filename)
    inv_no = extract_invoice_number(text, filename)
    tb, vat, ta = extract_amounts(text)

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

def perform_audit(data: dict) -> dict:
    tb = round(float(data.get("total_before_vat") or 0.0), 2)
    vat = round(float(data.get("vat_value") or 0.0), 2)
    ta = round(float(data.get("total_after_vat") or 0.0), 2)

    data["total_before_vat"] = tb
    data["vat_value"] = vat
    data["total_after_vat"] = ta

    calc_total = round(tb + vat, 2)
    diff = round(ta - calc_total, 2)
    
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
        error_reasons.append("الرقم الضريبي غير مكتمل")
    if not data.get("invoice_date"):
        error_reasons.append("تاريخ الفاتورة مفقود")

    is_clean = len(error_reasons) == 0
    data["is_valid_zatca"] = is_clean
    data["qr_status"] = "✓ فحص نصي سريع"
    data["error_reason"] = "✓ سليم ومطابق" if is_clean else " ، ".join(error_reasons)
    return data

# ==========================================
# 4. المعالجة الفائقة والسرعة القسوى (Turbo Text-Extraction)
# ==========================================
def process_single_file_task(task_info: tuple) -> list:
    task_type, path_or_zip, internal_name, display_name = task_info
    try:
        if task_type == "zip":
            with zipfile.ZipFile(path_or_zip, "r") as zf:
                file_bytes = zf.read(internal_name)
        else:
            with open(path_or_zip, "rb") as fp:
                file_bytes = fp.read()

        if display_name.lower().endswith(".pdf"):
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            text_content = "\n".join([page.get_text("text") for page in doc])
            doc.close()
        else:
            text_content = ""

        local_data = local_smart_extract(text_content, display_name)
        local_data["اسم_الملف"] = display_name
        audited = perform_audit(local_data)
        return [audited]
    except Exception as e:
        return [{
            "invoice_date": "",
            "invoice_number": "",
            "doc_type": "خطأ قراءة",
            "vendor_name": f"خطأ: {str(e)[:30]}",
            "tax_id": "",
            "total_before_vat": 0.0,
            "vat_value": 0.0,
            "total_after_vat": 0.0,
            "qr_status": "❌ تعذر الفحص",
            "audit_status": "❌ ملف تالف",
            "is_valid_zatca": False,
            "error_reason": "تعذر فتح الملف",
            "اسم_الملف": display_name
        }]

# ==========================================
# 5. تصدير الإكسيل
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
    return df_out

def export_to_excel(records: list, output_path: str):
    df = pd.DataFrame(records)
    if df.empty:
        return

    credit_notes_mask = df["doc_type"].astype(str).str.contains("دائن", na=False)
    df_credit_notes_raw = df[credit_notes_mask].copy()

    valid_mask = (df["is_valid_zatca"] == True) & (~credit_notes_mask)
    df_valid_raw = df[valid_mask].copy()

    error_mask = (df["is_valid_zatca"] == False) & (~credit_notes_mask)
    df_error_raw = df[error_mask].copy()

    df_sheet_valid = build_formatted_df(df_valid_raw, include_error_reason=False)
    df_sheet_credit = build_formatted_df(df_credit_notes_raw, include_error_reason=True)
    df_sheet_error = build_formatted_df(df_error_raw, include_error_reason=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df_sheet_valid.to_excel(writer, index=False, sheet_name="الفواتير والإشعارات المدينة")
        df_sheet_credit.to_excel(writer, index=False, sheet_name="الإشعارات الدائنة")
        df_sheet_error.to_excel(writer, index=False, sheet_name="فواتير بها أخطاء")

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            ws.views.sheetView[0].rightToLeft = True
            for col in ws.columns:
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = 22

# ==========================================
# 6. واجهة المستخدم
# ==========================================
st.title("🧾 نظام تحليل المشتريات والفرز الضريبي (ZATCA)")
st.markdown("وضع **السرعة الفائقة** معالج خصخصياً لفك وفحص آلاف الفواتير الضخمة في أجزاء من الثانية بدقة تامة.")

uploaded_files = st.file_uploader(
    "ارفع ملف ZIP أو ملفات الفواتير (PDF / صور)",
    accept_multiple_files=True,
    type=["zip", "pdf", "png", "jpg", "jpeg", "webp"]
)

if uploaded_files:
    if st.button("🚀 بدء التحليل الفائق السرعة", type="primary"):
        temp_dir = tempfile.mkdtemp()
        task_queue = []
        valid_exts = (".pdf", ".png", ".jpg", ".jpeg", ".webp")

        with st.spinner("جاري فك وفهرسة الملفات بسرعة صاروخية..."):
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
        st.info(f"📂 تم تجهيز **{total_files}** مستند. جاري الاستخراج...")

        raw_results = []
        progress_bar = st.progress(0)
        
        for i, t_info in enumerate(task_queue):
            res = process_single_file_task(t_info)
            if res:
                raw_results.extend(res)
            progress_bar.progress((i + 1) / max(total_files, 1))

        st.session_state["last_audit_results"] = raw_results
        st.success("✓ تم الانتهاء من التحليل بنجاح تام!")

if "last_audit_results" in st.session_state and st.session_state["last_audit_results"]:
    raw_results = st.session_state["last_audit_results"]
    
    output_excel = "تحليل_المشتريات_المصنف_ZATCA.xlsx"
    export_to_excel(raw_results, output_excel)

    with open(output_excel, "rb") as f:
        st.download_button(
            label="📥 تحميل تقرير الإكسيل النهائي",
            data=f,
            file_name="تحليل_المشتريات_المصنف.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
