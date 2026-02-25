import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Royan Excel Report", page_icon="📊", layout="centered")

st.title("📊 نظام تقارير الإدارة - مجموعة رويان")
st.markdown("---")
st.success("هذا النظام يقوم بتوليد ملف إكسيل **تفاعلي (يحتوي على معادلات حقيقية)** ليكون أداة الإقناع الأقوى في اجتماعاتك.")

# --- 1. تجهيز البيانات (بوضع أصفار في أماكن المجاميع ليتم استبدالها بمعادلات لاحقاً) ---
df_invest_flexo = pd.DataFrame({
    "البند": ["ماكينة طباعة فلكسو CI (8 ألوان)", "ماكينة تركيب البليتات (Mounter)", "مبرد الهواء والكمبروسر", "إجمالي استثمار الفلكسو"],
    "التكلفة (ريال)": [8000000, 150000, 400000, 0] # الصفر سيستبدل بمعادلة
})

df_invest_roto = pd.DataFrame({
    "البند": ["ماكينة طباعة روتوجرافيور (8 ألوان)", "غلاية الزيت الحراري (Thermal Boiler)", "معدات نقل وتخزين السلندرات", "إجمالي استثمار الروتو"],
    "التكلفة (ريال)": [9000000, 1500000, 300000, 0] # الصفر سيستبدل بمعادلة
})

df_opex = pd.DataFrame({
    "بند التكلفة الشهرية": ["الرواتب والأجور", "الإيجار والمصاريف الإدارية", "فاتورة الطاقة (الماكينة + الغلاية)", "إجمالي المصاريف الشهرية"],
    "التكلفة في الفلكسو (ريال)": [150000, 50000, 25000, 0], 
    "التكلفة في الروتو (ريال)": [150000, 60000, 65000, 0]  
})

df_scenario = pd.DataFrame({
    "عناصر تكلفة الطلبية": [
        "تكلفة المواد الخام", 
        "تكلفة التجهيز (بليتات مقابل سلندرات)", 
        "تكلفة هالك التشغيل والتجهيز", 
        "تكلفة المستهلكات (أنيلوكس/رول مطاطي)",
        "إجمالي تكلفة الطلبية",
        "تكلفة الطن الواحد"
    ],
    "تقنية الفلكسو (ريال)": [45000, 3200, 450, 200, 0, 0],
    "تقنية الروتو (ريال)": [45000, 12000, 2250, 150, 0, 0]
})

df_client_mix = pd.DataFrame({
    "الهيكل المطلوب للعميل": ["طبقة واحدة (38 mic label white / 40 mic clear)", "طبقتين (20 opp + 20 met)", "3 طبقات (12 pet + 7 alu + 50 pe)"],
    "النسبة من إجمالي الطلب": ["60%", "30%", "10%"],
    "سعر البيع المستهدف للعميل - فلكسو (ريال/كجم)": [12.0, 13.0, 15.0],
    "سعر البيع المستهدف للعميل - روتو (ريال/كجم)": [13.0, 13.5, 15.0]
})

# --- 2. إنشاء الإكسيل وحقن المعادلات ---
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    workbook = writer.book
    worksheet = workbook.add_worksheet('دراسة الجدوى التفاعلية')
    worksheet.right_to_left() 
    
    # تنسيقات الإكسيل
    header_format = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    money_format = workbook.add_format({'num_format': '#,##0', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    formula_format = workbook.add_format({'num_format': '#,##0', 'bold': True, 'bg_color': '#E2EFDA', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
    title_format = workbook.add_format({'bold': True, 'font_size': 13, 'bg_color': '#D9E1F2', 'align': 'center', 'border': 1})
    input_format = workbook.add_format({'bold': True, 'bg_color': '#FFF2CC', 'border': 1, 'align': 'center', 'font_color': 'red'})
    
    # إضافة التنسيق الذي كان مفقوداً وتسبب في الخطأ
    normal_format = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})

    # --- كتابة الجداول ---
    worksheet.merge_range('A1:B1', '1. استثمار الفلكسو (CAPEX)', title_format)
    df_invest_flexo.to_excel(writer, sheet_name='دراسة الجدوى التفاعلية', startrow=1, startcol=0, index=False)
    
    worksheet.merge_range('D1:E1', '1. استثمار الروتو (CAPEX)', title_format)
    df_invest_roto.to_excel(writer, sheet_name='دراسة الجدوى التفاعلية', startrow=1, startcol=3, index=False)

    worksheet.merge_range('A8:C8', '2. التكاليف التشغيلية الشهرية للمصنع (OPEX)', title_format)
    df_opex.to_excel(writer, sheet_name='دراسة الجدوى التفاعلية', startrow=8, startcol=0, index=False)

    # خلية تفاعلية لتغيير حجم الطلبية
    worksheet.write('A15', '👈 غير حجم الطلبية هنا لاختبار السعر (بالطن):', title_format)
    worksheet.write('B15', 5, input_format) # خلية قابلة للتعديل باللون الأصفر
    worksheet.write('C15', 'الأسفل سيتغير تلقائياً', normal_format) # تم تصحيح الخطأ هنا

    worksheet.merge_range('A16:C16', '3. سيناريوهات التكلفة للطلبية (تتفاعل مع الخلية أعلاه)', title_format)
    df_scenario.to_excel(writer, sheet_name='دراسة الجدوى التفاعلية', startrow=16, startcol=0, index=False)

    worksheet.merge_range('A25:D25', '4. تحليل منتجات العميل', title_format)
    df_client_mix.to_excel(writer, sheet_name='دراسة الجدوى التفاعلية', startrow=25, startcol=0, index=False)

    # --- تطبيق التنسيقات على الأعمدة العادية ---
    for col_num, value in enumerate(df_invest_flexo.columns.values):
        worksheet.write(1, col_num, value, header_format)
    for col_num, value in enumerate(df_invest_roto.columns.values):
        worksheet.write(1, col_num + 3, value, header_format)
    for col_num, value in enumerate(df_opex.columns.values):
        worksheet.write(8, col_num, value, header_format)
    for col_num, value in enumerate(df_scenario.columns.values):
        worksheet.write(16, col_num, value, header_format)
    for col_num, value in enumerate(df_client_mix.columns.values):
        worksheet.write(25, col_num, value, header_format)

    # ----------------------------------------------------
    # حقن المعادلات الرياضية (Formulas) الحقيقية في الإكسيل
    # ----------------------------------------------------
    worksheet.write_formula('B6', '=SUM(B3:B5)', formula_format)
    worksheet.write_formula('E6', '=SUM(E3:E5)', formula_format)
    worksheet.write_formula('B13', '=SUM(B10:B12)', formula_format)
    worksheet.write_formula('C13', '=SUM(C10:C12)', formula_format)
    worksheet.write_formula('B22', '=SUM(B18:B21)', formula_format)
    worksheet.write_formula('C22', '=SUM(C18:C21)', formula_format)
    
    # قسمة إجمالي التكلفة على خلية "حجم الطلبية" (B15)
    worksheet.write_formula('B23', '=B22/B15', formula_format)
    worksheet.write_formula('C23', '=C22/B15', formula_format)

    # --- توسيع الأعمدة لتناسب النص ---
    worksheet.set_column('A:A', 45)
    worksheet.set_column('B:B', 20, money_format)
    worksheet.set_column('D:D', 45)
    worksheet.set_column('E:E', 20, money_format)
    worksheet.set_column('C:C', 20, money_format)

# --- 3. زر التحميل ---
st.download_button(
    label="📥 تحميل ملف الإكسيل التفاعلي",
    data=buffer.getvalue(),
    file_name="Interactive_Feasibility_Report.xlsx",
    mime="application/vnd.ms-excel"
)
