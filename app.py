import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="أداة سجلات الحضور", layout="centered")

st.title("📂 معالج سجلات الحضور")
st.write("قم برفع ملف الـ CSV لتحويله إلى تنسيق الإكسل المطلوب")

uploaded_file = st.file_uploader("اختر ملف CSV", type=['csv'])

if uploaded_file is not None:
    try:
        # قراءة الملف ومعرفة ما إذا كان يحتاج لتخطي السطر الأول
        content = uploaded_file.getvalue().decode('utf-8').splitlines()
        skip = 1 if 'Transaction' in content[0] else 0
        
        uploaded_file.seek(0)
        df = pd.read_csv(uploaded_file, skiprows=skip)

        # تحديد الأعمدة المطلوبة وتنسيقها
        mapping = {
            'Employee ID': 'الرقم الوظيفي', 'First Name': 'اسم الموظف',
            'Date': 'التاريخ', 'Time': 'وقت البصمة',
            'الرقم الوظيفي': 'الرقم الوظيفي', 'اسم الموظف': 'اسم الموظف',
            'التاريخ': 'التاريخ', 'وقت البصمة': 'وقت البصمة'
        }
        
        cols_to_keep = [c for c in df.columns if c in mapping]
        result_df = df[cols_to_keep].rename(columns=mapping)
        
        # تنسيق التاريخ والوقت وفق القالب [cite: 1]
        result_df['التاريخ'] = pd.to_datetime(result_df['التاريخ'], errors='coerce').dt.strftime('%d/%m/%Y')
        result_df['وقت البصمة'] = pd.to_datetime(result_df['وقت البصمة'], errors='coerce').dt.strftime('%H:%M')

        st.success("✅ تم معالجة البيانات بنجاح!")
        st.dataframe(result_df.head()) # عرض عينة من البيانات

        # تحويل البيانات لملف إكسل في الذاكرة للتحميل
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            result_df.to_excel(writer, index=False)
        
        st.download_button(
            label="تحميل ملف Excel الجاهز",
            data=output.getvalue(),
            file_name="Attendance_Report_Cleaned.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"حدث خطأ: {e}")