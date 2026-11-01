import streamlit as st
import pandas as pd
import io
import base64
from PIL import Image
from pyzbar.pyzbar import decode
import xlsxwriter

st.set_page_config(page_title="نظام التدقيق المحاسبي والضريبي للفواتير", layout="wide")

# ==========================================
# 1. دالة فك تشفير QR Code (معيار TLV Base64)
# ==========================================
def decode_tlv_qr(base64_string):
    """فك تشفير الكيو آر كود الخاص بالفواتير الضريبية (TLV Format)"""
    try:
        raw_data = base64.b64decode(base64_string)
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
        return result, "صالح (Valid TLV)"
    except Exception:
        return {}, "غير متوافق / نص عادي"

# ==========================================
# 2. دالة التدقيق المحاسبي الآلي للبيانات
# ==========================================
def perform_accounting_audit(df, expected_client_name, vat_rate=0.15):
    """إجراء الفحوصات المحاسبية والضريبية على جدول الفواتير"""
    audit_Notes = []
    status_list = []

    # كشف الفواتير المكررة (نفس المورد + نفس رقم الفاتورة)
    duplicates = df.duplicated(subset=['رقم الفاتورة', 'اسم المورد'], keep=False)

    for index, row in df.iterrows():
        notes = []
        net = float(row.get('الإجمالي قبل الضريبة', 0) or 0)
        vat = float(row.get('ضريبة القيمة المضافة', 0) or 0)
        gross = float(row.get('الإجمالي شامل الضريبة', 0) or 0)
        qr_gross = float(row.get('إجمالي الـ QR', 0) or 0)
        qr_vat = float(row.get('ضريبة الـ QR', 0) or 0)
        client = str(row.get('اسم العميل', '')).strip()
        qr_status = str(row.get('حالة الـ QR', ''))

        # 1. فحص المعادلة المحاسبية الأساسية
        if abs((net + vat) - gross) > 0.05:
            notes.append(f"خطأ جمع أفقى: الصافي + الضريبة لا يساوي الإجمالي (الفرق {round((net+vat)-gross, 2)})")

        # 2. فحص نسبة الضريبة (15%)
        expected_vat = round(net * vat_rate, 2)
        if net > 0 and abs(vat - expected_vat) > 0.10:
            notes.append(f"انحراف نسبة الضريبة: الضريبة المتوقعة ({expected_vat}) والمسجلة ({vat})")

        # 3. مطابقة بيانات الفاتورة مع الـ QR Code
        if "صالح" in qr_status:
            if abs(gross - qr_gross) > 0.05 or abs(vat - qr_vat) > 0.05:
                notes.append("عدم تطابق القيم المالية مع بيانات الـ QR المشفرة")
        else:
            notes.append("الـ QR مفقود أو تالف")

        # 4. فحص التكرار
        if duplicates.iloc[index]:
            notes.append("فاتورة مكررة (نفس الرقم والمورد)")

        # 5. فحص اسم العميل
        if expected_client_name and expected_client_name not in client:
            notes.append("اسم العميل غير مطابق للشركة")

        if not notes:
            audit_Notes.append("سليمة ومطابقة")
            status_list.append("مطابق")
        else:
            audit_Notes.append(" | ".join(notes))
            status_list.append("يوجد ملاحظات للمراجعة")

    df['حالة التدقيق'] = status_list
    df['ملاحظات الفحص المحاسبي'] = audit_Notes
    return df

# ==========================================
# 3. محرك تصدير الإكسيل الاحترافي (Excel Engine)
# ==========================================
def export_audit_excel(df, vat_rate=0.15):
    """إنشاء ملف إكسيل تفاعلي يحتوي على ملخص تنفيذي، معادلات حية، وتنسيق شرطي"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # --- إعداد التنسيقات (Styles) ---
        header_fmt = workbook.add_format({
            'bold': True, 'font_name': 'Arial', 'font_size': 11,
            'bg_color': '#1F4E78', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        money_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1, 'align': 'center'})
        text_fmt = workbook.add_format({'border': 1, 'align': 'center'})
        ok_fmt = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', 'border': 1, 'align': 'center', 'bold': True})
        err_fmt = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006', 'border': 1, 'align': 'center', 'bold': True})
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#1F4E78'})

        # ----------------------------------------------------
        # التبويب الأول: سجل الفواتير والتدقيق التفصيلي
        # ----------------------------------------------------
        sheet_name = 'سجل_تدقيق_الفواتير'
        df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1)
        ws = writer.sheets[sheet_name]
        ws.right_to_left() # اتجاه الصفحة من اليمين لليسار

        # كتابة رؤوس الأعمدة بالتنسيق الاحترافي
        headers = list(df.columns) + ['فحص الإجمالي (معادلة إكسيل)', 'فرق الضريبة (معادلة إكسيل)']
        for col_num, col_name in enumerate(headers):
            ws.write(1, col_num, col_name, header_fmt)

        # إدراج البيانات ومعادلات الإكسيل الحية لكل صف
        for row_num in range(len(df)):
            excel_row = row_num + 2  # لأن البيانات تبدأ من الصف رقم 3 في الإكسيل
            for col_num, col_name in enumerate(df.columns):
                val = df.iloc[row_num, col_num]
                if col_name in ['الإجمالي قبل الضريبة', 'ضريبة القيمة المضافة', 'الإجمالي شامل الضريبة', 'إجمالي الـ QR', 'ضريبة الـ QR']:
                    ws.write_number(excel_row, col_num, float(val or 0), money_fmt)
                else:
                    ws.write(excel_row, col_num, str(val), text_fmt)

            # إضافة معادلات إكسيل حية للمراجع (الأعمدة E, F, G تمثل الصافي، الضريبة، الإجمالي)
            # العمود E: الإجمالي قبل الضريبة | العمود F: الضريبة | العمود G: الإجمالي الشامل
            formula_total_diff = f"=ROUND((E{excel_row+1}+F{excel_row+1})-G{excel_row+1}, 2)"
            formula_vat_diff = f"=ROUND(F{excel_row+1}-(E{excel_row+1}*{vat_rate}), 2)"

            ws.write_formula(excel_row, len(df.columns), formula_total_diff, money_fmt)
            ws.write_formula(excel_row, len(df.columns) + 1, formula_vat_diff, money_fmt)

        # تطبيق التنسيق الشرطي (Conditional Formatting) لعمود حالة التدقيق
        status_col_idx = df.columns.get_loc('حالة التدقيق')
        ws.conditional_format(2, status_col_idx, len(df) + 1, status_col_idx, {
            'type': 'cell', 'criteria': 'equal to', 'value': '"مطابق"', 'format': ok_fmt
        })
        ws.conditional_format(2, status_col_idx, len(df) + 1, status_col_idx, {
            'type': 'cell', 'criteria': 'not equal to', 'value': '"مطابق"', 'format': err_fmt
        })

        # تجميد صف الرؤوس وإضافة الفلتر التلقائي وضبط عرض الأعمدة
        ws.freeze_panes(2, 0)
        ws.autofilter(1, 0, len(df) + 1, len(headers) - 1)
        ws.set_column(0, len(headers) - 1, 22)
        ws.set_column(df.columns.get_loc('ملاحظات الفحص المحاسبي'), df.columns.get_loc('ملاحظات الفحص المحاسبي'), 45)

        # ----------------------------------------------------
        # التبويب الثاني: الملخص التنفيذي والإقرار الضريبي
        # ----------------------------------------------------
        summary_ws = workbook.add_worksheet('الملخص_التنفيذي')
        summary_ws.right_to_left()
        summary_ws.write('B2', 'ملخص تدقيق فواتير المشتريات (تقرير المراجعة)', title_fmt)

        total_invoices = len(df)
        valid_invoices = len(df[df['حالة التدقيق'] == 'مطابق'])
        error_invoices = total_invoices - valid_invoices

        summary_data = [
            ['إجمالي عدد الفواتير المفحوصة', total_invoices],
            ['عدد الفواتير المطابقة والسليمة', valid_invoices],
            ['عدد الفواتير التي تتطلب مراجعة (بها ملاحظات)', error_invoices],
            ['إجمالي المشتريات قبل الضريبة', df['الإجمالي قبل الضريبة'].sum()],
            ['إجمالي ضريبة القيمة المضافة (المدخلات)', df['ضريبة القيمة المضافة'].sum()],
            ['الإجمالي العام شامل الضريبة', df['الإجمالي شامل الضريبة'].sum()]
        ]

        summary_ws.write('B4', 'البيان المحاسبي', header_fmt)
        summary_ws.write('C4', 'القيمة / العدد', header_fmt)

        for i, (label, val) in enumerate(summary_data, start=4):
            summary_ws.write(i, 1, label, text_fmt)
            if i >= 7:
                summary_ws.write_number(i, 2, float(val), money_fmt)
            else:
                summary_ws.write_number(i, 2, int(val), text_fmt)

        summary_ws.set_column('B:B', 42)
        summary_ws.set_column('C:C', 25)

    output.seek(0)
    return output.getvalue()

# ==========================================
# 4. واجهة المستخدم على Streamlit Cloud
# ==========================================
st.title("📊 منصة التدقيق الآلي لفواتير المشتريات والـ QR Code")
st.markdown("معالجة الفواتير بكميات كبيرة (+3000 فاتورة) مع التحقق الضريبي وتصدير تقرير المراجعة.")

with st.sidebar:
    st.header("⚙️ إعدادات التدقيق المحاسبي")
    company_name = st.text_input("اسم شركتك (لفحص اسم العميل)", value="شركة")
    vat_percentage = st.number_input("نسبة ضريبة القيمة المضافة (%)", value=15.0, step=1.0) / 100.0
    st.info("نصيحة لـ 3000 فاتورة: يفضل رفع البيانات على دفعات أو استخدام ملف البيانات المجمعة لضمان سرعة الاستجابة.")

# إنشاء بيانات تجريبية لمحاكاة التصدير الفوري والتأكد من عمل الكود
if st.button("🚀 توليد وتحميل تقرير إكسيل نموذجي (إثبات نجاح التصدير)"):
    sample_data = pd.DataFrame([
        {
            "رقم الفاتورة": "INV-2026-001", "تاريخ الفاتورة": "2026-10-15",
            "اسم المورد": "شركة التوريدات الحديثة", "اسم العميل": company_name,
            "الإجمالي قبل الضريبة": 1000.00, "ضريبة القيمة المضافة": 150.00, "الإجمالي شامل الضريبة": 1150.00,
            "حالة الـ QR": "صالح (Valid TLV)", "إجمالي الـ QR": 1150.00, "ضريبة الـ QR": 150.00
        },
        {
            "رقم الفاتورة": "INV-2026-002", "تاريخ الفاتورة": "2026-10-16",
            "اسم المورد": "مؤسسة التقنية المتطورة", "اسم العميل": company_name,
            "الإجمالي قبل الضريبة": 2000.00, "ضريبة القيمة المضافة": 250.00, "الإجمالي شامل الضريبة": 2250.00,
            "حالة الـ QR": "صالح (Valid TLV)", "إجمالي الـ QR": 2300.00, "ضريبة الـ QR": 300.00
        },
        {
            "رقم الفاتورة": "INV-2026-001", "تاريخ الفاتورة": "2026-10-15",
            "اسم المورد": "شركة التوريدات الحديثة", "اسم العميل": "عميل آخر غير مطابق",
            "الإجمالي قبل الضريبة": 1000.00, "ضريبة القيمة المضافة": 150.00, "الإجمالي شامل الضريبة": 1150.00,
            "حالة الـ QR": "مفقود أو تالف", "إجمالي الـ QR": 0.00, "ضريبة الـ QR": 0.00
        }
    ])

    audited_df = perform_accounting_audit(sample_data, company_name, vat_percentage)
    st.dataframe(audited_df, use_container_width=True)

    excel_bytes = export_audit_excel(audited_df, vat_percentage)
    st.download_button(
        label="📥 تحميل تقرير المراجعة المحاسبية (Excel)",
        data=excel_bytes,
        file_name="Purchase_Invoices_Audit_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
