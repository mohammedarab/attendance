import pandas as pd
import streamlit as st
from io import BytesIO

# إعداد واجهة الويب
st.set_page_config(page_title="أداة سجلات الحضور - Excel", layout="centered")

st.title("📂 معالج سجلات الحضور (Excel)")
st.write("ارفع ملف الإكسل لاستخراج البيانات المنسقة فوراً")

# رفع الملف بصيغة xlsx
uploaded_file = st.file_uploader("اختر ملف Excel (.xlsx)", type=['xlsx'])

if uploaded_file is not None:
    try:
        # 1. قراءة ملف الإكسل (تحميل البيانات بالكامل)
        # نقرأ الملف أولاً لفحص السطر الأول
        df_check = pd.read_excel(uploaded_file, nrows=1)
        
        # إذا كان الملف يحتوي على سطر "Transaction" في البداية نقوم بتخطيه
        if 'Transaction' in df_check.columns or 'Transaction' in str(df_check.iloc[0,0]):
            df = pd.read_excel(uploaded_file, skiprows=1)
        else:
            df = pd.read_excel(uploaded_file)

        # 2. تحديد الأعمدة المطلوبة بناءً على ملفاتك المرفقة 
        # نقوم بمطابقة المسميات سواء كانت بالإنجليزية أو العربية
        mapping = {
            'Employee ID': 'الرقم الوظيفي',
            'First Name': 'اسم الموظف',
            'Date': 'التاريخ',
            'Time': 'وقت البصمة',
            'الرقم الوظيفي': 'الرقم الوظيفي',
            'اسم الموظف': 'اسم الموظف',
            'التاريخ': 'التاريخ',
            'وقت البصمة': 'وقت البصمة'
        }
        
        # تصفية الأعمدة الموجودة فقط
        available_cols = [c for c in df.columns if c in mapping]
        result_df = df[available_cols].rename(columns=mapping)
        
        # 3. توحيد تنسيقات التاريخ والوقت كما طلبت سابقاً
        # التاريخ: DD/MM/YYYY
        result_df['التاريخ'] = pd.to_datetime(result_df['التاريخ'], errors='coerce').dt.strftime('%d/%m/%Y')
        
        # الوقت: HH:MM
        result_df['وقت البصمة'] = pd.to_datetime(result_df['وقت البصمة'], errors='coerce').dt.strftime('%H:%M')

        # حذف أي صفوف فارغة ناتجة عن التنسيق
        result_df.dropna(subset=['الرقم الوظيفي', 'التاريخ'], inplace=True)

        st.success("✅ تم استخراج كافة البيانات بنجاح!")
        st.dataframe(result_df) # عرض البيانات للتأكد

        # 4. تجهيز ملف المخرجات بصيغة Excel (.xlsx)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Sheet1')
        
        # زر التحميل
        st.download_button(
            label="تحميل ملف Excel المنسق",
            data=output.getvalue(),
            file_name="Attendance_Report_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"حدث خطأ أثناء قراءة الملف: {e}")
