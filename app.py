import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import pandas as pd

# ۱. تنظیمات اولیه پروداکت دیزاین استریم‌لیت
st.set_page_config(
    page_title="پلتفرم هوشمند مدل‌سازی مالی گالری طلا",
    page_icon="👑",
    layout="wide"
)

# ۲. تزریق فونت متن‌باز و استاندارد Vazirmatn همراه با تنظیمات سایز ۱۲ و رنگ مشکی
st.markdown("""
    <style>
    /* وارد کردن فونت متن‌باز و استاندارد Vazirmatn از گوگل فونتس */
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100..900&display=swap');

    /* اعمال گلوبال فونت، سایز ۱۲ و رنگ مشکی برای تمامی بخش‌های اپلیکیشن */
    html, body, [data-testid="stAppViewContainer"], .stApp {
        direction: RTL !important;
        text-align: right !important;
        font-family: 'Vazirmatn', sans-serif !important;
        font-size: 12px !important;
        color: #000000 !important;
        background-color: #F8F9FA;
    }

    /* اعمال تنظیمات فونت، سایز ۱۲ و رنگ مشکی روی جداول تعاملی و ادیتور استریم‌لیت */
    .stDataFrame, div[data-testid="stDataFrameDataGird"], [role="gridcell"], th, td, input, select, textarea {
        font-family: 'Vazirmatn', sans-serif !important;
        font-size: 12px !important;
        color: #000000 !important;
    }
    
    /* تنظیمات رنگ متن داخل سلول‌های در حال ویرایش جدول */
    input[data-testid="stDataFrameDataGirdCellInput"] {
        color: #000000 !important;
    }

    /* تنظیمات تب‌ها (st.tabs) */
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

    /* تنظیمات کامپوننت‌های نمایش شاخص (st.metric) */
    div[data-testid="stMetricValue"], div[data-testid="stMetricLabel"], div[data-testid="stMetricDelta"] {
        color: #000000 !important;
        font-family: 'Vazirmatn', sans-serif !important;
    }
    
    /* تنظیمات عناوین و نوشته‌ها */
    h1, h2, h3, h4, h5, h6, p, span, label {
        font-family: 'Vazirmatn', sans-serif !important;
        color: #000000 !important;
    }
    
    /* اصلاح رنگ لیبل سایدبار و ورودی‌ها به مشکی */
    .stSlider label, .stNumberInput label {
        color: #000000 !important;
    }
    </style>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# ۳. مدیریت وضعیت داده‌ها (Session State) برای ذخیره آنی تغییرات جزئی پشت صحنه
# -----------------------------------------------------------------------------
if 'live_gold_price' not in st.session_state:
    st.session_state.live_gold_price = 4500000

# داده‌های جزئی هزینه‌های راه‌اندازی (CAPEX) در پشت صحنه
if 'df_capex' not in st.session_state:
    st.session_state.df_capex = pd.DataFrame([
        {"کد هزینه": "CAP-01", "عنوان هزینه": "ودیعه ملک تجاری در عفیف‌آباد", "مبلغ (تومان)": 1200000000, "دسته": "ملک"},
        {"کد هزینه": "CAP-04", "عنوان هزینه": "اجرای دکوراسیون چوب و شیشه لوکس", "مبلغ (تومان)": 380000000, "دسته": "دکوراسیون"},
        {"کد هزینه": "CAP-07", "عنوان هزینه": "گاوصندوق فوق سنگین مکانیکی", "مبلغ (تومان)": 120000000, "دسته": "امنیت"},
        {"کد هزینه": "CAP-08", "عنوان هزینه": "سیستم دوربین مداربسته ضدآستگمات", "مبلغ (تومان)": 80000000, "دسته": "امنیت"},
        {"کد هزینه": "CAP-16", "عنوان هزینه": "مراسم و ایونت رسمی افتتاحیه", "مبلغ (تومان)": 60000000, "دسته": "مارکتینگ"}
    ])

# داده‌های جزئی هزینه‌های جاری ماهانه (OPEX)
if 'df_opex' not in st.session_state:
    st.session_state.df_opex = pd.DataFrame([
        {"کد هزینه": "OPX-01", "عنوان هزینه": "اجاره‌بهای ماهانه مغازه عفیف‌آباد", "مبلغ ماهانه": 120000000, "نوع": "ثابت"},
        {"کد هزینه": "OPX-02", "عنوان هزینه": "حقوق پایه مدیر گالری", "مبلغ ماهانه": 22000000, "نوع": "ثابت"},
        {"کد هزینه": "OPX-03", "عنوان هزینه": "حقوق کارشناس فروش و ادمین", "مبلغ ماهانه": 16000000, "نوع": "ثابت"},
        {"کد هزینه": "OPX-08", "عنوان هزینه": "بودجه مستمر دیجیتال مارکتینگ", "مبلغ ماهانه": 25000000, "نوع": "ثابت"},
        {"کد هزینه": "OPX-11", "عنوان هزینه": "هزینه پذیرایی VIP و تشریفات", "مبلغ ماهانه": 5000000, "نوع": "متغیر"}
    ])

# -----------------------------------------------------------------------------
# ۴. سایدبار کنترلرهای فاندامنتال بازار طلا
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ تنظیمات زنده بازار")
st.session_state.live_gold_price = st.sidebar.number_input("قیمت هر گرم طلای ۱۸ عیار (تومان)", min_value=1000000, value=st.session_state.live_gold_price, step=50000)
avg_wage = st.sidebar.slider("میانگین درصد اجرت ساخت", min_value=0.0, max_value=0.50, value=0.20, step=0.01)
gallery_markup = st.sidebar.slider("حاشیه سود ناخالص گالری (مارک‌آپ)", min_value=0.05, max_value=0.40, value=0.24, step=0.01)

partner_1_cash = st.sidebar.number_input("سرمایه شریک اول (سرمایه‌گذار)", value=5500000000, step=100000000)
partner_2_cash = st.sidebar.number_input("سرمایه شریک دوم (اجرایی)", value=1500000000, step=50000000)
total_pool_investment = partner_1_cash + partner_2_cash

# -----------------------------------------------------------------------------
# ۵. رندر بخش اصلی اپلیکیشن با تب‌های تفکیک‌شده و تعاملی
# -----------------------------------------------------------------------------
st.title("👑 پلتفرم هوشمند مدیریت و مدل‌سازی مالی گالری طلا")
st.subheader("لوکیشن استراتژیک مبنا: شیراز - خیابان عفیف‌آباد")
st.markdown("---")

# ایجاد تب‌ها: تب اول خروجی کلان، تب‌های بعدی برای ویرایش خطی جزئیات پشت صحنه
tab_dash, tab_edit_capex, tab_edit_opex = st.tabs([
    "📊 داشبورد کلان و خروجی اکسل", 
    "🛠️ ویرایش جزئیات هزینه‌های راه‌اندازی (CAPEX)", 
    "💸 ویرایش جزئیات هزینه‌های جاری (OPEX)"
])

# --- تب مدیریت تعاملی هزینه‌های سرمایه‌ای (CAPEX) ---
with tab_edit_capex:
    st.subheader("پشت صحنه هزینه‌های سرمایه‌ای (CAPEX)")
    st.caption("مقادیر یا عناوین هزینه‌ها را تغییر دهید یا سطر جدید اضافه کنید. محاسبات مالی کلان فوراً تغییر خواهند کرد.")
    
    edited_capex_df = st.data_editor(
        st.session_state.df_capex,
        num_rows="dynamic",
        column_config={
            "مبلغ (تومان)": st.column_config.NumberColumn("مبلغ (تومان)", min_value=0, format="%d"),
            "دسته": st.column_config.SelectboxColumn("دسته‌بندی", options=["ملک", "دکوراسیون", "امنیت", "تجهیزات", "مارکتینگ"])
        },
        key="capex_editor_v3",
        use_container_width=True
    )
    st.session_state.df_capex = edited_capex_df

# --- تب مدیریت تعاملی هزینه‌های جاری ماهانه (OPEX) ---
with tab_edit_opex:
    st.subheader("پشت صحنه هزینه‌های جاری ماهانه (OPEX)")
    
    edited_opex_df = st.data_editor(
        st.session_state.df_opex,
        num_rows="dynamic",
        column_config={
            "مبلغ ماهانه": st.column_config.NumberColumn("هزینه ماهانه (تومان)", min_value=0, format="%d"),
            "نوع": st.column_config.SelectboxColumn("نوع هزینه", options=["ثابت", "متغیر"])
        },
        key="opex_editor_v3",
        use_container_width=True
    )
    st.session_state.df_opex = edited_opex_df

# -----------------------------------------------------------------------------
# ۶. موتور محاسبات زنده بر اساس جزئیات متصل به ادیتورها
# -----------------------------------------------------------------------------
capex_fixed_total = st.session_state.df_capex["مبلغ (تومان)"].sum()
gold_working_capital = total_pool_investment - capex_fixed_total
cost_per_gram_built = st.session_state.live_gold_price * (1 + avg_wage)
initial_gold_weight_grams = gold_working_capital / cost_per_gram_built if cost_per_gram_built > 0 else 0

daily_traffic = 10
conversion_rate = 0.30
avg_invoice_weight = 5
working_days = 26

monthly_sales_weight_grams = daily_traffic * conversion_rate * avg_invoice_weight * working_days
sell_price_per_gram = cost_per_gram_built * (1 + gallery_markup)
monthly_revenue = monthly_sales_weight_grams * sell_price_per_gram

annual_revenue_y1 = monthly_revenue * 12
annual_cogs_y1 = annual_revenue_y1 / (1 + gallery_markup)
annual_gross_profit_y1 = annual_revenue_y1 - annual_cogs_y1

total_monthly_opex = st.session_state.df_opex["مبلغ ماهانه"].sum()
annual_opex_y1 = total_monthly_opex * 12

ebitda_y1 = annual_gross_profit_y1 - annual_opex_y1
net_profit_y1 = ebitda_y1 * 0.85
roi_y1 = (net_profit_y1 / total_pool_investment) * 100 if total_pool_investment > 0 else 0
payback_period_years = total_pool_investment / net_profit_y1 if net_profit_y1 > 0 else 0

# --- رندر محتوای تب اصلی داشبورد (نمایش نتایج زنده با فونت جدید) ---
with tab_dash:
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("کل نقدینگی پروژه", f"{total_pool_investment:,.0f} تومان")
    with col2:
        st.metric("شارژ اولیه انبار طلا", f"{gold_working_capital:,.0f} تومان")
    with col3:
        st.metric("وزن طلای موجودی اولیه", f"{initial_gold_weight_grams:,.1f} گرم")
    with col4:
        st.metric("پیش‌بینی درآمد سال اول", f"{annual_revenue_y1:,.0f} تومان")

    col5, col6, col7, col8 = st.columns(4)
    with col5:
        st.metric("سود خالص سال اول (پس از مالیات)", f"{net_profit_y1:,.0f} تومان")
    with col6:
        st.metric("نرخ بازگشت سرمایه (ROI)", f"{roi_y1:.1f}%")
    with col7:
        st.metric("دوره بازگشت سرمایه", f"{payback_period_years * 12:.1f} ماه")
    with col8:
        st.metric("قیمت تمام شده هر گرم", f"{cost_per_gram_built:,.0f} تومان")

    st.markdown("---")
    st.write("### 📥 دریافت فایل مدل تجاری هماهنگ شده")
    
    # تابع تولید اکسل پویا بر اساس مقادیر جدید جدول‌ها
    def generate_excel():
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "خلاصه مدل مالی"
        ws.views.sheetView[0].showGridLines = True
        
        # استایل‌ها
        font_title = Font(name='Calibri', size=16, bold=True, color='FFFFFF')
        font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        font_data = Font(name='Calibri', size=11, bold=False, color='000000')
        fill_title = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
        fill_header = PatternFill(start_color='D4AF37', end_color='D4AF37', fill_type='solid')
        
        ws.merge_cells('A1:D1')
        ws['A1'] = "گزارش شبیه‌سازی و مدل‌سازی مالی گالری طلا"
        ws['A1'].font = font_title
        ws['A1'].fill = fill_title
        ws['A1'].alignment = Alignment(horizontal='center')
        
        headers = ["شاخص مالی", "مقدار محاسباتی", "واحد", "توضیحات استراتژیک"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_num)
            cell.value = header
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal='center')
            
        data = [
            ["کل سرمایه گذاری اولیه", total_pool_investment, "تومان", "مجموع آورده شرکا"],
            ["سرمایه در گردش طلا", gold_working_capital, "تومان", "سرمایه مصرفی برای خرید طلا بعد از کسر کاپکس"],
            ["حجم طلای اولیه انبار", initial_gold_weight_grams, "گرم", "موجودی اولیه ویترین"],
            ["هزینه های راه‌اندازی (CAPEX)", capex_fixed_total, "تومان", "مجموع هزینه‌های ثابت سخت‌افزاری"],
            ["هزینه های جاری سالانه (OPEX)", annual_opex_y1, "تومان", "مجموع کل هزینه‌های عملیاتی سال اول"],
            ["سود خالص سال اول", net_profit_y1, "تومان", "پس از کسر هزینه و مالیات فرضی"],
            ["نرخ بازگشت سرمایه (ROI)", roi_y1 / 100, "درصد", "بازدهی خالص سال اول نسبت به سرمایه"],
        ]
        
        for row_idx, row_data in enumerate(data, 4):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.value = value
                cell.font = font_data
                if col_idx == 2:
                    if isinstance(value, float) and value < 1.0:
                        cell.number_format = '0.0%'
                    else:
                        cell.number_format = '#,##0'
                        
        wb.save(output)
        return output.getvalue()

    excel_data = generate_excel()
    st.download_button(
        label="📥 دانلود پکیج کامل مدل مالی (Excel)",
        data=excel_data,
        file_name="Gold_Gallery_Financial_Model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
