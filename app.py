import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import pandas as pd

# ۱. تنظیمات اولیه پروداکت دیزاین استریم‌لیت
st.set_page_config(
    page_title="پلتفرم جامع مدل‌سازی مالی و ارزیابی لوکیشن گالری طلا",
    page_icon="👑",
    layout="wide"
)

# ۲. تزریق فونت متن‌باز Vazirmatn، سایز ۱۲ و رنگ مشکی خالص
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100..900&display=swap');

    html, body, [data-testid="stAppViewContainer"], .stApp {
        direction: RTL !important;
        text-align: right !important;
        font-family: 'Vazirmatn', sans-serif !important;
        font-size: 12px !important;
        color: #000000 !important;
        background-color: #F8F9FA;
    }

    .stDataFrame, div[data-testid="stDataFrameDataGird"], [role="gridcell"], th, td, input, select, textarea {
        font-family: 'Vazirmatn', sans-serif !important;
        font-size: 12px !important;
        color: #000000 !important;
    }
    
    input[data-testid="stDataFrameDataGirdCellInput"] { color: #000000 !important; }

    button[data-baseweb="tab"] {
        font-family: 'Vazirmatn', sans-serif !important;
        font-size: 12px !important;
        color: #444444 !important;
    }
    
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #000000 !important;
        font-weight: bold !important;
        border-bottom-color: #1B365D !important;
    }

    div[data-testid="stMetricValue"], div[data-testid="stMetricLabel"] {
        color: #000000 !important;
        font-family: 'Vazirmatn', sans-serif !important;
    }
    
    h1, h2, h3, h4, h5, h6, p, span, label {
        font-family: 'Vazirmatn', sans-serif !important;
        color: #000000 !important;
    }
    .stSlider label, .stNumberInput label { color: #000000 !important; }
    </style>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# ۳. مدیریت داده‌های پایه پشت صحنه در Session State (تفکیک عیار و لوکیشن‌ها)
# -----------------------------------------------------------------------------
if 'prices' not in st.session_state:
    st.session_state.prices = {
        "قیمت طلا ۱۸ عیار (گرم)": 4500000,
        "قیمت طلا ۲۴ عیار (گرم)": 6000000,
        "قیمت هر قطعه سکه امامی": 50000000
    }

if 'df_inventory' not in st.session_state:
    st.session_state.df_inventory = pd.DataFrame([
        {"نوع دارایی": "طلای ۱۸ عیار (ویترین/ساخته شده)", "وزن/تعداد اولیه": 1500, "واحد": "گرم"},
        {"نوع دارایی": "طلای ۲۴ عیار (آبشده/شمش)", "وزن/تعداد اولیه": 300, "واحد": "گرم"},
        {"نوع دارایی": "سکه تمام امامی (پشتوانه نقدینگی)", "وزن/تعداد اولیه": 50, "واحد": "عدد"}
    ])

if 'df_locations' not in st.session_state:
    st.session_state.df_locations = pd.DataFrame([
        {"لوکیشن": "لوکیشن الف (عفیف‌آباد شیراز)", "رهن ملک (تومان)": 1500000000, "اجاره ماهانه": 120000000, "ترافیک پاخور روزانه": 12, "نرخ تبدیل مشتری": 0.25},
        {"لوکیشن": "لوکیشن ب (فرشته تهران)", "رهن ملک (تومان)": 4000000000, "اجاره ماهانه": 350000000, "ترافیک پاخور روزانه": 20, "نرخ تبدیل مشتری": 0.30},
        {"لوکیشن": "لوکیشن ج (چهارباغ اصفهان)", "رهن ملک (تومان)": 2000000000, "اجاره ماهانه": 150000000, "ترافیک پاخور روزانه": 15, "نرخ تبدیل مشتری": 0.22}
    ])

if 'df_partners' not in st.session_state:
    st.session_state.df_partners = pd.DataFrame([
        {"شریک": "شریک اول (سرمایه‌گذار مالی)", "درصد سهم از آورده": 75.0, "درصد سهم از سود": 60.0},
        {"شریک": "شریک دوم (مدیر اجرایی/فنی)", "درصد سهم از آورده": 25.0, "درصد سهم از سود": 40.0}
    ])

# -----------------------------------------------------------------------------
# ۴. ساختار ناوبری و تب‌های برنامه
# -----------------------------------------------------------------------------
st.title("👑 سیستم یکپارچه تحلیل مالی و ارزیابی سه لوکیشن گالری طلا")
st.caption("منطبق بر ضوابط محاسبه قیمت، اجرت، سود مصوب ۷٪ اتحادیه و مالیات بر ارزش افزوده بازار ایران")
st.markdown("---")

tab_dash, tab_inventory, tab_locations, tab_partners = st.tabs([
    "📊 داشبورد مقایسه‌ای ۳ لوکیشن", 
    "💎 مدیریت موجودی پشت صحنه (عیار و سکه)", 
    "📍 تنظیمات تخصصی لوکیشن‌ها", 
    "👥 ساختار سهم‌الشرکه و سرمایه"
])

# --- تب ۲: مدیریت موجودی و عیار ---
with tab_inventory:
    st.subheader("تخصیص دارایی اولیه گالری")
    col_p1, col_p2, col_p3 = st.columns(3)
    with col_p1: st.session_state.prices["قیمت طلا ۱۸ عیار (گرم)"] = st.number_input("مبنای طلا ۱۸ (تومان)", value=st.session_state.prices["قیمت طلا ۱۸ عیار (گرم)"])
    with col_p2: st.session_state.prices["قیمت طلا ۲۴ عیار (گرم)"] = st.number_input("مبنای طلا ۲۴ (تومان)", value=st.session_state.prices["قیمت طلا ۲۴ عیار (گرم)"])
    with col_p3: st.session_state.prices["قیمت هر قطعه سکه امامی"] = st.number_input("مبنای سکه امامی (تومان)", value=st.session_state.prices["قیمت هر قطعه سکه امامی"])
    
    st.session_state.df_inventory = st.data_editor(st.session_state.df_inventory, use_container_width=True, key="inv_ed")

# --- تب ۳: تنظیمات تخصصی لوکیشن‌ها ---
with tab_locations:
    st.subheader("مشخصات فیزیکی و تجاری لوکیشن‌های پیشنهادی")
    st.session_state.df_locations = st.data_editor(st.session_state.df_locations, use_container_width=True, key="loc_ed")

# --- تب ۴: ساختار سهم‌الشرکه ---
with tab_partners:
    st.subheader("توزیع آورده مالی و سود عملیاتی بین شرکا")
    st.session_state.df_partners = st.data_editor(st.session_state.df_partners, use_container_width=True, key="part_ed")

# -----------------------------------------------------------------------------
# ۵. موتور محاسبات زنده بازار ایران برای ۳ لوکیشن به صورت همزمان
# -----------------------------------------------------------------------------
# محاسبه ارزش کل انبار طلا و سکه بر اساس نرخ روز بازار ایران
val_18 = st.session_state.df_inventory.iloc[0]["وزن/تعداد اولیه"] * st.session_state.prices["قیمت طلا ۱۸ عیار (گرم)"]
val_24 = st.session_state.df_inventory.iloc[1]["وزن/تعداد اولیه"] * st.session_state.prices["قیمت طلا ۲۴ عیار (گرم)"]
val_coin = st.session_state.df_inventory.iloc[2]["وزن/تعداد اولیه"] * st.session_state.prices["قیمت هر قطعه سکه امامی"]
total_inventory_value = val_18 + val_24 + val_coin

# پارامترهای پیش‌فرض محاسبات فیزیکی بازار طلا
avg_ojarat_percent = 0.18  # ۱۸ درصد اجرت ساخت میانگین طلا ساخته شده
union_profit_percent = 0.07  # ۷ درصد سود مصوب گالری داران
vat_percent = 0.09  # ۹ درصد مالیات بر ارزش افزوده روی اجرت و سود

location_results = []

for idx, row in st.session_state.df_locations.iterrows():
    loc_name = row["لوکیشن"]
    rahn = row["رهن ملک (تومان)"]
    rent_monthly = row["اجاره ماهانه"]
    traffic = row["ترافیک پاخور روزانه"]
    conv_rate = row["نرخ تبدیل مشتری"]
    
    # کل هزینه اولیه راه اندازی این لوکیشن (رهن ملک + دکوراسیون فرضی ۵۰۰ میلیونی + ارزش انبار طلا)
    total_startup_cost = rahn + 500000000 + total_inventory_value
    
    # محاسبه حجم معاملات ماهانه (فروش طلای ۱۸ عیار ساخته شده)
    # ۲۶ روز کاری در ماه، میانگین فاکتور هر مشتری: ۶ گرم
    monthly_sales_grams = traffic * conv_rate * 6 * 26
    
    # قیمت پایه طلا مبدا معاملات
    base_gold_cost = monthly_sales_grams * st.session_state.prices["قیمت طلا ۱۸ عیار (گرم)"]
    
    # محاسبه درآمد خالص گالری مطابق فرمول اتحادیه طلا (سود ۷ درصد گالری)
    total_ojarat_earned = base_gold_cost * avg_ojarat_percent
    gallery_net_profit_monthly = (base_gold_cost + total_ojarat_earned) * union_profit_percent
    
    # مجموع کل دریافتی ماهانه گالری از مشتریان (شامل اصل طلا، اجرت، سود و مالیات)
    vat_collected = (total_ojarat_earned + gallery_net_profit_monthly) * vat_percent
    total_monthly_revenue_from_customers = base_gold_cost + total_ojarat_earned + gallery_net_profit_monthly + vat_collected
    
    # هزینه های جاری ماهانه (اوپکس) = اجاره ملک + ۱۰۰ میلیون هزینه‌های پرسنل و مارکتینگ ثابت
    total_monthly_opex = rent_monthly + 100000000
    
    # سود خالص ماهانه پس از کسر اوپکس گالری
    final_monthly_net_profit = gallery_net_profit_monthly - total_monthly_opex
    final_annual_net_profit = final_monthly_net_profit * 12
    
    # شاخص سرعت گردش دارایی (Gold Turnover Velocity) = میزان وزن طلای فروخته شده سالانه تقسیم بر کل موجودی اولیه انبار طلا
    total_yearly_sales_grams = monthly_sales_grams * 12
    inventory_turnover_velocity = total_yearly_sales_grams / st.session_state.df_inventory.iloc[0]["وزن/تعداد اولیه"]
    
    # نرخ بازگشت سرمایه
    roi = (final_annual_net_profit / total_startup_cost) * 100 if total_startup_cost > 0 else 0
    
    location_results.append({
        "لوکیشن": loc_name,
        "سرمایه اولیه کل (تومان)": total_startup_cost,
        "فروش ماهانه طلا (گرم)": monthly_sales_grams,
        "درآمد ناخالص ماهانه (تومان)": gallery_net_profit_monthly,
        "هزینه جاری ماهانه (تومان)": total_monthly_opex,
        "سود خالص ماهانه (تومان)": final_monthly_net_profit,
        "سرعت گردش دارایی طلا (بار در سال)": inventory_turnover_velocity,
        "نرخ بازگشت سرمایه (ROI)": roi
    })

df_res = pd.DataFrame(location_results)

# --- تب ۱: نمایش داشبورد کلان مقایسه‌ای ---
with tab_dash:
    st.subheader("⚡ نتایج ارزیابی و مقایسه همزمان سه لوکیشن تجاری")
    
    # نمایش کاردهای KPI تفکیک دارایی کلان انبار
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("کل ارزش انبار طلا و سکه شما", f"{total_inventory_value:,.0f} تومان")
    with c2: st.metric("موجودی طلا ۱۸ عیار ویترین", f"{st.session_state.df_inventory.iloc[0]['وزن/تعداد اولیه']:,.0f} گرم")
    with c3: st.metric("پشتوانه سکه امامی گالری", f"{st.session_state.df_inventory.iloc[2]['وزن/تعداد اولیه']:,.0f} عدد")
    
    st.markdown("---")
    
    # ایجاد مقایسه ستونی لوکیشن‌ها جهت سهولت تصمیم‌گیری
    cols_loc = st.columns(3)
    for idx, row in df_res.iterrows():
        with cols_loc[idx]:
            st.markdown(f"### 📍 {row['لوکیشن']}")
            st.metric("سود خالص ماهانه", f"{row['سود خالص ماهانه (تومان)']:,.0f} تومان")
            st.metric("نرخ بازگشت سرمایه (ROI)", f"{row['نرخ بازگشت سرمایه (ROI)']:.1f}%")
            st.metric("سرعت گردش دارایی انبار", f"{row['سرعت گردش دارایی طلا (بار در سال)']:.2f} بار در سال")
            st.metric("وزن فروش ماهانه طلا", f"{row['فروش ماهانه طلا (گرم)']:,.1f} گرم")
            st.markdown("---")
            
            # سهم سود شریک اول و دوم در این لوکیشن
            p1_profit = row['سود خالص ماهانه (تومان)'] * (st.session_state.df_partners.iloc[0]['درصد سهم از سود'] / 100)
            p2_profit = row['سود خالص ماهانه (تومان)'] * (st.session_state.df_partners.iloc[1]['درصد سهم از سود'] / 100)
            st.write(f"🔹 سهم سود ماهانه شریک اول: {p1_profit:,.0f} تومان")
            st.write(f"🔹 سهم سود ماهانه شریک دوم: {p2_profit:,.0f} تومان")

    st.markdown("---")
    st.subheader("📥 دانلود پکیج تجاری و اکسل یکپارچه (Merged Architecture)")
    st.info("با کلیک بر روی دکمه زیر، فایل اکسل جامع طراحی شده که تمامی شیت‌ها را در یک قالب ساختاریافته مرج کرده است، تولید و دانلود خواهد شد.")

    # تابع تولید فایل اکسل کاملاً ادغام شده و تجمیعی
    def generate_unified_excel():
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "مدل ادغام شده مالی گالری طلا"
        ws.views.sheetView[0].showGridLines = True
        
        font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        font_data = Font(name='Calibri', size=11, bold=False, color='000000')
        fill_header = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
        fill_sub = PatternFill(start_color='D4AF37', end_color='D4AF37', fill_type='solid')
        
        # ۱. بخش اطلاعات تخصیص دارایی انبار
        ws.cell(row=1, column=1, value="بخش اول: ساختار تخصیص دارایی و موجودی عیار طلا").font = Font(name='Calibri', size=12, bold=True)
        headers_inv = ["نوع دارایی پشت صحنه", "وزن / تعداد اولیه", "واحد سنجش"]
        for c_idx, h in enumerate(headers_inv, 1):
            cell = ws.cell(row=2, column=c_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            
        for r_idx, row in st.session_state.df_inventory.iterrows():
            ws.cell(row=r_idx+3, column=1, value=row["نوع دارایی"]).font = font_data
            ws.cell(row=r_idx+3, column=2, value=row["وزن/تعداد اولیه"]).font = font_data
            ws.cell(row=r_idx+3, column=3, value=row["واحد"]).font = font_data
            
        # ۲. بخش ادغام شده ارزیابی مقایسه‌ای ۳ لوکیشن بازار ایران
        start_row_loc = 8
        ws.cell(row=start_row_loc, column=1, value="بخش دوم: خروجی ارزیابی و مدل مالی مقایسه‌ای سه لوکیشن").font = Font(name='Calibri', size=12, bold=True)
        
        headers_loc = ["عنوان لوکیشن تجاری", "سرمایه اولیه کل (تومان)", "فروش طلا (گرم/ماه)", "درآمد ناخالص (ماه)", "هزینه جاری (ماه)", "سود خالص ماهانه (تومان)", "سرعت گردش انبار طلا (بار/سال)", "بازدهی سرمایه (ROI)"]
        for c_idx, h in enumerate(headers_loc, 1):
            cell = ws.cell(row=start_row_loc+1, column=c_idx, value=h)
            cell.font = font_header
            cell.fill = fill_sub
            
        for r_idx, row in df_res.iterrows():
            current_r = start_row_loc + 2 + r_idx
            ws.cell(row=current_r, column=1, value=row["لوکیشن"]).font = font_data
            ws.cell(row=current_r, column=2, value=row["سرمایه اولیه کل (تومان)"]).number_format = '#,##0'
            ws.cell(row=current_r, column=3, value=row["فروش ماهانه طلا (گرم)"]).number_format = '#,##0.0'
            ws.cell(row=current_r, column=4, value=row["درآمد ناخالص ماهانه (تومان)"]).number_format = '#,##0'
            ws.cell(row=current_r, column=5, value=row["هزینه جاری ماهانه (تومان)"]).number_format = '#,##0'
            ws.cell(row=current_r, column=6, value=row["سود خالص ماهانه (تومان)"]).number_format = '#,##0'
            ws.cell(row=current_r, column=7, value=row["سرعت گردش دارایی طلا (بار در سال)"]).number_format = '0.00'
            ws.cell(row=current_r, column=8, value=row["نرخ بازگشت سرمایه (ROI)"] / 100).number_format = '0.0%'
            
            for c in range(2, 9):
                ws.cell(row=current_r, column=c).font = font_data

        wb.save(output)
        return output.getvalue()

    excel_data = generate_unified_excel()
    st.download_button(
        label="📥 دانلود فایل اکسل یکپارچه و تجمیعی (Merged Structural Data)",
        data=excel_data,
        file_name="Unified_Gold_Gallery_Model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
