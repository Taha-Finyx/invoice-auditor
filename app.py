def process_single_invoice(file_bytes, filename, expected_company_name="", expected_company_vat=""):
    """معالجة شاملة لفاتورة واحدة مع حماية كاملة ضد ملفات الـ HTTP التالفة والباركود غير الصالح"""
    qr_raw_text = None
    qr_page_location = "غير موجود"

    try:
        # حماية إضافية: التأكد من أن الملف ليس رابط HTTP أو صفحة HTML تالفة
        if file_bytes.startswith(b'HTTP/') or file_bytes.startswith(b'<!DOCTYPE') or file_bytes.startswith(b'<html'):
            return {
                "اسم الملف": filename.rsplit('/', 1)[-1], "نوع المستند": "ملف تالف/رابط",
                "رقم الفاتورة": "خطأ", "تاريخ الفاتورة": "", "اسم المورد": "",
                "الرقم الضريبي للمورد": "", "حالة الرقم الضريبي": "خطأ", "اسم العميل": "",
                "الإجمالي قبل الضريبة": 0.0, "ضريبة القيمة المضافة": 0.0, "الإجمالي شامل الضريبة": 0.0,
                "حالة الـ QR": "ملف HTML أو رابط تالف (ليس PDF)", "موقع الـ QR": "خطأ",
                "إجمالي الـ QR": 0.0, "ضريبة الـ QR": 0.0
            }

        if filename.lower().endswith('.pdf'):
            pdf_info = pdfinfo_from_bytes(file_bytes)
            total_pages = int(pdf_info.get("Pages", 1))
            pages_to_check = [1]
            if total_pages > 1:
                pages_to_check.append(total_pages)

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
            "حالة الـ QR": f"استثناء قراءة ({str(e)[:25]})", "موقع الـ QR": "خطأ",
            "إجمالي الـ QR": 0.0, "ضريبة الـ QR": 0.0
        }

    # استخراج بيانات الـ QR وتحليلها...
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
