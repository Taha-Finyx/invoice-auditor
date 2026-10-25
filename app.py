import streamlit as st

st.set_page_config(page_title="نظام تدقيق وتحليل الفواتير الضريبية", page_icon="🧾", layout="wide")

# درع حماية شامل يمنع ظهور شاشة Oh no نهائياً
try:
    import os
    import re
    import json
    import base64
    import urllib.request
    import urllib.error
    import pandas as pd
    import numpy as np
    from concurrent.futures import ThreadPoolExecutor, as_completed

    try:
        import fitz
        HAS_FITZ = True
    except Exception:
        HAS_FITZ = False

    try:
        import cv2
        HAS_CV2 = True
    except Exception:
        HAS_CV2 = False

    try:
        from openpyxl.styles import PatternFill, Font, Alignment
        HAS_OPENPYXL = True
    except Exception:
        HAS_OPENPYXL = False

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
        if not HAS_CV2 or img_bgr is None:
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
            if not data["vendor_name"]:
                clean_fname = re.sub(r'\.pdf$|\.jpg$\vert{}\.png$', '', filename, flags=re.IGNORECASE)
                data["vendor_name"] = clean_fname.split('_')[0].strip()
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
                clean_fname = re.sub(r'\.pdf$|\.jpg$\vert{}\.png$', '', filename, flags=re.IGNORECASE)
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
                data["subtotal"] = round(float
