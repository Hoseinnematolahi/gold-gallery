import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import plotly.express as px

# ۱. تنظیمات اولیه پروداکت دیزاین استریم‌لیت
st.set_page_config(
    page_title="سیستم یکپارچه شبیه‌سازی مالی و امکان‌سنجی گالری طلا و سکه ایران",
    page_icon="👑",
    layout="wide"
)

# ۲. تزریق استایل‌های بومی با فونت متن‌باز Vazirmatn، سایز ۱۲ و رنگ متن مشکی خالص
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
# ۳. مدیریت وضعیت داده‌ها (Session State) برای بازار خرده‌فروشی و سرمایه‌گذاران
# -----------------------------------------------------------------------------
if 'market_prices' not in st.session_state:
    st.session_state.market_prices = {
        "طلا ۱۸ عیار (هر گرم)": 4500000,
        "طلا ۲۴ عیار / شمش (هر گرم)": 6000000,
        "طلای آبشده (هر گرم ملاک عمل)": 4480000,
        "سکه تمام امامی (تعداد)": 50000000,
        "نیم سکه (تعداد)": 26000000,
        "ربع سکه (تعداد)": 16000000,
        "سکه پارسیان - میانگین (تعداد)": 4000000
    }

if 'df_portfolio' not in st.session_state:
    st.session_state.df_portfolio = pd.DataFrame([
        {"گروه محصول": "طلای ۱۸ عیار ساخته‌شده (ویترینی)", "سهم از کل سرمایه موجودی (%)": 40.0, "میانگین اجرت (%)": 18.0, "سود مصوب گالری (%)": 7.0, "سرعت گردش (بار در سال)": 3.5},
        {"گروه محصول": "طلای ۲۴ عیار / شمش استاندارد", "سهم از کل سرمایه موجودی (%)": 10.0, "میانگین اجرت (%)": 4.0, "سود مصوب گالری (%)": 3.0, "سرعت گردش (بار در سال)": 2.0},
        {"گروه محصول": "طلای آبشده (خرید و فروش/بنکداری)", "سهم از کل سرمایه موجودی (%)": 15.0, "میانگین اجرت (%)": 0.0, "سود مصوب گالری (%)": 1.5, "سرعت گردش (بار در سال)": 8.0},
        {"گروه محصول": "سکه تمام امامی", "سهم از کل سرمایه موجودی (%)": 15.0, "میانگین اجرت (%)": 0.0, "سود مصوب گالری (%)": 2.0, "سرعت گردش (بار در سال)": 6.0},
        {"گروه محصول": "نیم سکه و ربع سکه", "سهم از کل سرمایه موجودی (%)": 12.0, "میانگین اجرت (%)": 0.0, "سود مصوب گالری (%)": 2.5, "سرعت گردش (بار در سال)": 5.0},
        {"گروه محصول": "سکه‌های پارسیان (کادویی)", "سهم از کل سرمایه موجودی (%)": 8.0, "میانگین اجرت (%)": 5.0, "سود مصوب گالری (%)": 4.0, "سرعت گردش (بار در سال)": 7.0}
    ])

if 'df_locations' not in st.session_state:
    st.session_state.df_locations = pd.DataFrame([
        {"موقعیت تجاری": "لوکیشن الف (منطقه لوکس / عفیف‌آباد)", "رهن ملک (تومان)": 2000000000, "اجاره ماهانه (تومان)": 150000000, "ترافیک روزانه مشتری": 15, "نرخ خرید موفق (%)": 25.0},
        {"موقعیت تجاری": "لوکیشن ب (مرکز بازار طلا / بنکداری)", "رهن ملک (تومان)": 4500000000, "اجاره ماهانه (تومان)": 300000000, "ترافیک روزانه مشتری": 35, "نرخ خرید موفق (%)": 20.0},
        {"موقعیت تجاری": "لوکیشن ج (مجتمع تجاری نوپا)", "رهن ملک (تومان)": 1000000000, "اجاره ماهانه (تومان)": 80000000, "ترافیک روزانه مشتری": 10, "نرخ خرید موفق (%)": 30.0}
    ])

# -----------------------------------------------------------------------------
# ۴. ساختار اصلی برنامه و منوی تب‌ها
# -----------------------------------------------------------------------------
st.title("👑 طرح امکان‌سنجی هوشمند و داینامیک راه‌اندازی گالری طلا، آبشده و سکه")
st.markdown("---")

tab_dashboard, tab_market_setup, tab_location_setup, tab_financial_rules = st.tabs([
    "📊 داشبورد مدیریتی و ارزیابی اقتصادی",
    "💰 نرخ روز بازار و سبد دارایی بر مبنای آورده",
    "📍 شبیه‌سازی و مقایسه لوکیشن‌ها",
    "⚖️ قوانین صنف، تورم و نرخ بانکی"
])

# --- تب ۲: نرخ روز بازار و سبد دارایی ---
with tab_market_setup:
    st.subheader("۱. نرخ‌های پایه بازار خرده‌فروشی طلا و سکه ایران")
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.session_state.market_prices["طلا ۱۸ عیار (هر گرم)"] = st.number_input("قیمت گرم طلا ۱۸ (تومان)", value=st.session_state.market_prices["طلا ۱۸ عیار (هر گرم)"])
        st.session_state.market_prices["طلا ۲۴ عیار / شمش (هر گرم)"] = st.number_input("قیمت شمش ۲۴ عیار (تومان)", value=st.session_state.market_prices["طلا ۲۴ عیار / شمش (هر گرم)"])
    with col_m2:
        st.session_state.market_prices["طلای آبشده (هر گرم ملاک عمل)"] = st.number_input("قیمت طلا آبشده (تومان)", value=st.session_state.market_prices["طلای آبشده (هر گرم ملاک عمل)"])
        st.session_state.market_prices["سکه تمام امامی (تعداد)"] = st.number_input("قیمت سکه امامی (تومان)", value=st.session_state.market_prices["سکه تمام امامی (تعداد)"])
    with col_m3:
        st.session_state.market_prices["نیم سکه (تعداد)"] = st.number_input("قیمت نیم سکه (تومان)", value=st.session_state.market_prices["نیم سکه (تعداد)"])
        st.session_state.market_prices["ربع سکه (تعداد)"] = st.number_input("قیمت ربع سکه (تومان)", value=st.session_state.market_prices["ربع سکه (تعداد)"])

    st.markdown("---")
    st.subheader("۲. تنظیمات تخصیص سبد خرید بر اساس آورده مالی موجودی")
    st.session_state.df_portfolio = st.data_editor(st.session_state.df_portfolio, use_container_width=True, key="portfolio_ed")

# --- تب ۳: شبیه‌سازی و مقایسه لوکیشن‌ها ---
with tab_location_setup:
    st.subheader("تنظیمات پارامترهای فیزیکی و تجاری لوکیشن‌ها")
    st.session_state.df_locations = st.data_editor(st.session_state.df_locations, use_container_width=True, key="loc_ed_main")

# --- تب ۴: قوانین صنف، تورم و نرخ بانکی ---
with tab_financial_rules:
    st.subheader("قوانین مالیاتی، تورم انتظاری و نرخ جایگزین")
    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        inflation_rate = st.slider("نرخ تورم سالانه انتظاری (%)", min_value=10, max_value=80, value=45)
    with col_f2:
        bank_interest_rate = st.slider("نرخ سود بدون ریسک بانکی / صکوک (%)", min_value=10, max_value=40, value=28)
    with col_f3:
        vat_rate = st.number_input("نرخ مالیات بر ارزش افزوده روی سود و اجرت (%)", value=9.0)

    st.subheader("👥 آورده و ساختار سهامداران")
    col_s1, col_s2 = st.columns(2)
    with col_s1:
        partner_1_capital = st.number_input("میزان آورده نقدی شریک اول (تومان)", value=10000000000, step=500000000)
    with col_s2:
        partner_2_capital = st.number_input("میزان آورده نقدی شریک دوم (تومان)", value=3000000000, step=100000000)
    
    total_pool_capital = partner_1_capital + partner_2_capital

# -----------------------------------------------------------------------------
# ۵. موتور محاسباتی جامع و تخصصی بازار ایران
# -----------------------------------------------------------------------------
total_initial_capex_fixed = 600000000  # دکوراسیون، گاوصندوق، سیستم‌های امنیتی و دوربین ضدآستگمات
gold_working_capital_pool = total_pool_capital - total_initial_capex_fixed

# تخصیص دارایی و تشکیل سبد
portfolio_results = []
total_sum_shares = st.session_state.df_portfolio["سهم از کل سرمایه موجودی (%)"].sum()

# تصحیح درصدهای سبد در صورتی که ۱۰۰ نباشند
shares_multiplier = 100.0 / total_sum_shares if total_sum_shares > 0 else 1.0

for idx, row in st.session_state.df_portfolio.iterrows():
    p_name = row["گروه محصول"]
    share = row["سهم از کل سرمایه موجودی (%)"] * shares_multiplier
    allocated_cash = gold_working_capital_pool * (share / 100.0)
    
    # تعیین قیمت پایه هر واحد جهت سنجش حجم انبار
    if "۱۸ عیار" in p_name:
        unit_price = st.session_state.market_prices["طلا ۱۸ عیار (هر گرم)"]
    elif "۲۴ عیار" in p_name:
        unit_price = st.session_state.market_prices["طلا ۲۴ عیار / شمش (هر گرم)"]
    elif "آبشده" in p_name:
        unit_price = st.session_state.market_prices["طلای آبشده (هر گرم ملاک عمل)"]
    elif "تمام امامی" in p_name:
        unit_price = st.session_state.market_prices["سکه تمام امامی (تعداد)"]
    elif "نیم سکه" in p_name:
        unit_price = st.session_state.market_prices["نیم سکه (تعداد)"]
    else:
        unit_price = st.session_state.market_prices["سکه پارسیان - میانگین (تعداد)"]
        
    initial_units = allocated_cash / unit_price if unit_price > 0 else 0
    
    portfolio_results.append({
        "گروه محصول": p_name,
        "سرمایه تخصیص یافته (تومان)": allocated_cash,
        "حجم موجودی اولیه (گرم/تعداد)": initial_units,
        "اجرت (%)": row["میانگین اجرت (%)"],
        "سود گالری (%)": row["سود مصوب گالری (%)"],
        "سرعت گردش": row["سرعت گردش (بار در سال)"]
    })

df_portfolio_calculated = pd.DataFrame(portfolio_results)

# محاسبه خروجی و عایدی برای هر ۳ لوکیشن به صورت زنده
location_comparison = []

for idx, loc in st.session_state.df_locations.iterrows():
    loc_name = loc["موقعیت تجاری"]
    rahn_cost = loc["رهن ملک (تومان)"]
    monthly_rent = loc["اجاره ماهانه (تومان)"]
    traffic = loc["ترافیک روزانه مشتری"]
    conversion = loc["نرخ خرید موفق (%)"] / 100.0
    
    # کل هزینه های اولیه این لوکیشن (رهن ملک + سرمایه ثابت + طلا)
    loc_total_startup = rahn_cost + total_initial_capex_fixed + gold_working_capital_pool
    
    # محاسبه فروش ماهانه بر اساس پاخور و سرعت گردش کالا
    # فرض مبنا در بازار خرده فروشی ایران: ۲۶ روز کاری در ماه
    monthly_transactions = traffic * conversion * 26
    
    # توزیع درآمدهای ماهانه با فرمول رسمی صنف ایران
    loc_monthly_gross_profit = 0
    loc_total_monthly_turnover = 0
    
    for p_idx, p_row in df_portfolio_calculated.iterrows():
        p_name = p_row["گروه محصول"]
        velocity = p_row["سرعت گردش"]
        allocated_val = p_row["سرمایه تخصیص یافته (تومان)"]
        
        # درآمد سالانه فرضی ناشی از گردش دارایی
        yearly_product_turnover = allocated_val * velocity
        monthly_product_turnover = yearly_product_turnover / 12.0
        
        # محاسبه سود خالص گالری از این محصول بر اساس فرمول صنف ایران
        # در ایران سود گالری روی (اصل طلا + اجرت) اعمال می‌شود
        ojarat_amount = monthly_product_turnover * (p_row["اجرت"] / 100.0)
        profit_amount = (monthly_product_turnover + ojarat_amount) * (p_row["سود گالری"] / 100.0)
        
        loc_monthly_gross_profit += profit_amount
        loc_total_monthly_turnover += monthly_product_turnover + ojarat_amount + profit_amount
        
    # کل هزینه‌های جاری ماهانه (اجاره بها + ۱۲۰ میلیون هزینه حقوق طلافروش، بیمه، طراح ویترین و امنیت)
    loc_monthly_opex = monthly_rent + 120000000
    
    # سود خالص ماهانه گالری
    loc_monthly_net_profit = loc_monthly_gross_profit - loc_monthly_opex
    loc_annual_net_profit = loc_monthly_net_profit * 12
    
    # نرخ بازگشت سرمایه و دوره بازگشت سرمایه واقعی
    loc_roi = (loc_annual_net_profit / loc_total_startup) * 100 if loc_total_startup > 0 else 0
    loc_payback = loc_total_startup / loc_annual_net_profit if loc_annual_net_profit > 0 else 0
    
    # مقایسه با سود بانکی
    bank_alternative_annual_yield = loc_total_startup * (bank_interest_rate / 100.0)
    alpha_over_bank = loc_annual_net_profit - bank_alternative_annual_yield
    
    location_comparison.append({
        "موقعیت تجاری": loc_name,
        "کل سرمایه اولیه مورد نیاز": loc_total_startup,
        "گردش فروش ماهانه (تومان)": loc_total_monthly_turnover,
        "سود ناخالص ماهانه (صنف)": loc_monthly_gross_profit,
        "هزینه جاری ماهانه": loc_monthly_opex,
        "سود خالص ماهانه گالری": loc_monthly_net_profit,
        "سود خالص سالانه": loc_annual_net_profit,
        "نرخ بازگشت سرمایه (ROI)": loc_roi,
        "دوره بازگشت (سال)": loc_payback,
        "مازاد بازدهی نسبت به بانک": alpha_over_bank
    })

df_loc_res = pd.DataFrame(location_comparison)

# --- تب ۱: نمایش داشبورد مدیریتی و ارزیابی اقتصادی ---
with tab_dashboard:
    st.subheader("⚡ کاردهای شاخص‌های کلیدی عملکرد (KPI) - کلیات طرح")
    c_k1, c_k2, c_k3 = st.columns(3)
    with c_k1:
        st.metric("کل سرمایه آورده شرکا", f"{total_pool_capital:,.0f} تومان")
    with c_k2:
        st.metric("بودجه خالص تامین طلا و سکه", f"{gold_working_capital_pool:,.0f} تومان")
    with c_k3:
        st.metric("هزینه ثابت راه‌اندازی و تجهیزات", f"{total_initial_capex_fixed:,.0f} تومان")
        
    st.markdown("---")
    st.subheader("📊 مقایسه زنده و ستونی پتانسیل تجاری ۳ موقعیت انتخاب شده")
    
    cols_visual = st.columns(3)
    for idx, r in df_loc_res.iterrows():
        with cols_visual[idx]:
            st.markdown(f"#### 📍 {r['موقعیت تجاری']}")
            st.metric("سود خالص ماهانه گالری", f"{r['سود خالص ماهانه گالری']:,.0f} تومان")
            st.metric("نرخ بازگشت سرمایه (ROI)", f"{r['نرخ بازگشت سرمایه (ROI)']:.1f}%")
            
            if r['مازاد بازدهی نسبت به بانک'] > 0:
                st.success(f"📈 مازاد بر سود بانک: {r['مازاد بازدهی نسبت به بانک']:,.0f} تومان")
            else:
                st.error(f"⚠️ توجیه اقتصادی کمتر از بانک: {abs(r['مازاد بازدهی نسبت به بانک']):,.0f} تومان کمتر")
                
            st.write(f"🔹 کل سرمایه‌گذاری این لوکیشن: {r['کل سرمایه اولیه مورد نیاز']:,.0f} تومان")
            st.write(f"🔹 دوره بازگشت سرمایه: {r['دوره بازگشت (سال)'] * 12:.1f} ماه")
            st.markdown("---")

    # افزودن نمودار مقایسه‌ای سود خالص لوکیشن‌ها با Plotly
    st.subheader("📈 نمودار تحلیلی مقایسه سودآوری لوکیشن‌ها")
    fig = px.bar(df_loc_res, x="موقعیت تجاری", y="سود خالص ماهانه گالری", 
                 title="سود خالص ماهانه به تفکیک موقعیت‌های جغرافیایی بازار ایران",
                 labels={"سود خالص ماهانه گالری": "سود خالص (تومان)", "موقعیت تجاری": "لوکیشن"})
    fig.update_layout(font_family="Vazirmatn", textangle=0)
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("---")
    st.subheader("📥 دریافت فایل امکان‌سنجی جامع و مرج‌شده (Merged Excel Output)")
    st.info("تمام شیت‌ها، سبد دارایی طلا و سکه، فرمول‌های صنف و ارزیابی لوکیشن‌ها در یک شیت ساختاریافته ادغام شده‌اند.")

    # تابع تولید اکسل کاملاً مرج شده و ساختاریافته
    def generate_comprehensive_excel():
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "طرح امکان‌سنجی گالری طلا"
        ws.views.sheetView[0].showGridLines = True
        
        font_main_title = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
        font_h = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
        font_d = Font(name='Calibri', size=11, bold=False, color='000000')
        
        fill_title = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
        fill_sub = PatternFill(start_color='D4AF37', end_color='D4AF37', fill_type='solid')
        
        # عنوان بزرگ جدول
        ws.merge_cells('A1:G1')
        ws['A1'] = "گزارش نهایی و طرح امکان‌سنجی یکپارچه گالری طلا، سکه و آبشده ایران"
        ws['A1'].font = font_main_title
        ws['A1'].fill = fill_title
        ws['A1'].alignment = Alignment(horizontal='center')
        
        # جدول اول: سبد دارایی تخصیص یافته
        ws.cell(row=3, column=1, value="۱. سبد خرید و تخصیص دارایی طلا و سکه").font = Font(name='Calibri', size=12, bold=True)
        headers_p = ["گروه محصول طلا و سکه", "سرمایه تخصیص یافته (تومان)", "حجم موجودی اولیه", "اجرت ساخت (%)", "سود گالری (%)", "سرعت گردش (سالانه)"]
        for c_idx, h in enumerate(headers_p, 1):
            cell = ws.cell(row=4, column=c_idx, value=h)
            cell.font = font_h
            cell.fill = fill_sub
            
        for r_idx, row in df_portfolio_calculated.iterrows():
            curr_r = 5 + r_idx
            ws.cell(row=curr_r, column=1, value=row["گروه محصول"]).font = font_d
            ws.cell(row=curr_r, column=2, value=row["سرمایه تخصیص یافته (تومان)"]).number_format = '#,##0'
            ws.cell(row=curr_r, column=3, value=row["حجم موجودی اولیه (گرم/تعداد)"]).number_format = '#,##0.0'
            ws.cell(row=curr_r, column=4, value=row["اجرت (%)"] / 100.0).number_format = '0.0%'
            ws.cell(row=curr_r, column=5, value=row["سود گالری (%)"] / 100.0).number_format = '0.0%'
            ws.cell(row=curr_r, column=6, value=row["سرعت گردش"]).number_format = '0.0'
            
        # جدول دوم: مقایسه لوکیشن‌ها و ارزیابی اقتصادی
        start_r_loc = 14
        ws.cell(row=start_r_loc, column=1, value="۲. تحلیل مقایسه‌ای لوکیشن‌ها و نرخ بازگشت سرمایه").font = Font(name='Calibri', size=12, bold=True)
        headers_l = ["موقعیت تجاری گالری", "کل سرمایه اولیه مورد نیاز", "گردش فروش ماهانه", "سود خالص ماهانه", "نرخ بازگشت (ROI)", "دوره بازگشت سرمایه (سال)", "مازاد بازدهی نسبت به بانک"]
        for c_idx, h in enumerate(headers_l, 1):
            cell = ws.cell(row=start_r_loc+1, column=c_idx, value=h)
            cell.font = font_h
            cell.fill = fill_title
            
        for r_idx, row in df_loc_res.iterrows():
            curr_r = start_r_loc + 2 + r_idx
            ws.cell(row=curr_r, column=1, value=row["موقعیت تجاری"]).font = font_d
            ws.cell(row=curr_r, column=2, value=row["کل سرمایه اولیه مورد نیاز"]).number_format = '#,##0'
            ws.cell(row=curr_r, column=3, value=row["گردش فروش ماهانه (تومان)"]).number_format = '#,##0'
            ws.cell(row=curr_r, column=4, value=row["سود خالص ماهانه گالری"]).number_format = '#,##0'
            ws.cell(row=curr_r, column=5, value=row["نرخ بازگشت سرمایه (ROI)"] / 100.0).number_format = '0.0%'
            ws.cell(row=curr_r, column=6, value=row["دوره بازگشت (سال)"]).number_format = '0.0'
            ws.cell(row=curr_r, column=7, value=row["مازاد بازدهی نسبت به بانک"]).number_format = '#,##0'
            
            for col in range(2, 8):
                ws.cell(row=curr_r, column=col).font = font_d

        wb.save(output)
        return output.getvalue()

    excel_unified_data = generate_comprehensive_excel()
    st.download_button(
        label="📥 دانلود فایل امکان‌سنجی ادغام شده و یکپارچه (Comprehensive Merged Excel)",
        data=excel_unified_data,
        file_name="Comprehensive_Gold_Boutique_Feasibility_Model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
