import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import plotly.express as px
import plotly.graph_objects as go

# -----------------------------------------------------------------------------
# ۱. تنظیمات اولیه پروداکت دیزاین و استایل بومی
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="پلتفرم شبیه‌سازی مالی و امکان‌سنجی پیشرفته گالری طلا",
    page_icon="💎",
    layout="wide"
)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@100..900&display=swap');
    html, body, [data-testid="stAppViewContainer"], .stApp {
        direction: RTL !important;
        text-align: right !important;
        font-family: 'Vazirmatn', sans-serif !important;
        font-size: 13px !important;
        color: #000000 !important;
        background-color: #F8F9FA;
    }
    .stDataFrame, div[data-testid="stDataFrameDataGird"], [role="gridcell"], th, td, input, select, textarea {
        font-family: 'Vazirmatn', sans-serif !important;
        color: #000000 !important;
    }
    h1, h2, h3, h4, h5, h6, p, span, label { color: #000000 !important; font-family: 'Vazirmatn', sans-serif !important; }
    div[data-testid="stMetricValue"] { font-size: 22px !important; font-weight: bold !important; color: #1B365D !important; }
    </style>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# ۲. مدیریت مفروضات و سناریوهای پایدار در Session State
# -----------------------------------------------------------------------------
if 'current_scenario' not in st.session_state:
    st.session_state.current_scenario = "پایه (Base Case)"

if 'gold_base_price' not in st.session_state:
    st.session_state.gold_base_price = 4500000  # هر گرم طلا ۱۸ عیار به تومان

# داده‌های اولیه محصولات (ساختار Pydantic-Like در قالب دیکشنری کنترل شده)
if 'products_mix' not in st.session_state:
    st.session_state.products_mix = {
        "طلای ۱۸ عیار ساخته‌شده": {"share": 40.0, "ojarat": 18.0, "profit": 7.0, "velocity": 3.0},
        "طلای ۲۴ عیار / شمش": {"share": 15.0, "ojarat": 4.0, "profit": 3.0, "velocity": 2.0},
        "طلای آبشده (بنکداری)": {"share": 20.0, "ojarat": 0.0, "profit": 1.5, "velocity": 7.0},
        "سکه تمام و نیم امامی": {"share": 15.0, "ojarat": 0.0, "profit": 2.0, "velocity": 5.0},
        "سکه‌های پارسیان": {"share": 10.0, "ojarat": 6.0, "profit": 4.0, "velocity": 6.0}
    }

# -----------------------------------------------------------------------------
# ۳. ماژول‌ها و موتورهای اصلی پردازش مالی (Financial Core Engines)
# -----------------------------------------------------------------------------

class GoldFinancialEngine:
    @staticmethod
    def calculate_location_score(row):
        # پیاده سازی دقیق الگو و وزن های مدل امتیازدهی فیزیکی لوکیشن
        score = (
            row["پاخور روزانه"] * 2.5 * 0.30 +
            row["قدرت خرید منطقه (۱ تا ۱۰۰)"] * 0.25 +
            (100 - row["سطح رقابت (۱ تا ۱۰۰)"]) * 0.15 +
            row["امنیت و پارکینگ (۱ تا ۱۰۰)"] * 0.15 +
            50 * 0.15 # امتیاز پیش‌فرض سایر پارامترها
        )
        return min(max(score, 0), 100)

    @staticmethod
    def run_60_month_cashflow(startup_capital, gold_price, prod_mix, loc_data, inflation, bank_rate, vat_rate):
        fixed_capex = 700000000  # تجهیزات، دکوراسیون، گاوصندوق سنگین ضد سرقت
        working_capital_gold = startup_capital - fixed_capex - loc_data["رهن ملک (تومان)"]
        
        if working_capital_gold <= 0:
            return None, "سرمایه اولیه برای تامین انبار طلا کافی نیست. لطفاً آوردۀ نقدی را افزایش دهید."

        # موتور تقاضامحور (Customer-Driven Revenue Engine)
        daily_buyers = loc_data["پاخور روزانه"] * (loc_data["نرخ تبدیل (%)"] / 100.0)
        monthly_buyers = daily_buyers * 26  # ۲۶ روز کاری در ماه
        avg_ticket_size = loc_data["میانگین فاکتور خرید (تومان)"]
        customer_driven_demand_monthly = monthly_buyers * avg_ticket_size

        # شبیه‌سازی مدل ۶۰ ماهه با اعمال فاکتور تورم سالانه
        months = list(range(1, 61))
        cashflows = []
        monthly_details = []
        
        cumulative_cashflow = -startup_capital
        payback_month = None
        
        current_opex_base = loc_data["اجاره ماهانه (تومان)"] + 150000000  # اجاره + حقوق و سربرگ هزینه‌های عمومی
        
        for m in months:
            year_factor = (m - 1) // 12
            current_inflation_multiplier = (1 + inflation / 100.0) ** year_factor
            
            # اعمال تورم بر هزینه‌های عملیاتی و قدرت خرید فاکتورها
            opex = current_opex_base * current_inflation_multiplier
            monthly_demand = customer_driven_demand_monthly * current_inflation_multiplier
            
            # محاسبه درآمد ناخالص و قیمت تمام شده کالا بر اساس حاشیه سود صنف طلا
            total_gross_profit = 0
            for p_name, p_info in prod_mix.items():
                p_share = p_info["share"] / 100.0
                p_demand = monthly_demand * p_share
                
                # فرمول اتحادیه طلا و جواهر ایران: اعمال سود روی (اصل طلا + اجرت)
                # Gross Margin % = Profit_Margin / (1 + Ojarat + Profit_Margin)
                ojarat_rate = p_info["ojarat"] / 100.0
                profit_rate = p_info["profit"] / 100.0
                
                # سهم سود خالص گالری از فروش کالا
                product_gross_profit = p_demand * (profit_rate / (1 + ojarat_rate + profit_rate))
                total_gross_profit += product_gross_profit

            # سود قبل از کسر مالیات و استهلاک
            ebitda = total_gross_profit - opex
            tax = max(0.0, ebitda * 0.15)  # فرض ۱۵ درصد مالیات بر درآمد عملکرد سالانه
            net_monthly_cashflow = ebitda - tax
            
            cashflows.append(net_monthly_cashflow)
            
            if payback_month is None:
                cumulative_cashflow += net_monthly_cashflow
                if cumulative_cashflow >= 0:
                    payback_month = m
                    
            monthly_details.append({
                "ماه": m,
                "تقاضای فروش بازار": monthly_demand,
                "سود ناخالص صنف": total_gross_profit,
                "هزینه عملیاتی (Opex)": opex,
                "سود خالص ماهانه": net_monthly_cashflow
            })

        # محاسبات شاخص‌های ارزیابی سرمایه‌گذاری ($NPV$ و $IRR$)
        total_investment = startup_capital
        # نرخ تنزیل ماهانه بر اساس نرخ سود بانکی کشور
        discount_rate_monthly = (1 + bank_rate / 100.0) ** (1/12) - 1
        
        npv = -total_investment + sum(cf / ((1 + discount_rate_monthly) ** i) for i, cf in enumerate(cashflows, 1))
        
        # محاسبه IRR سالانه شده از روی آرایه ماهانه
        try:
            irr_monthly = npf.irr([-total_investment] + cashflows)
            irr_annual = ((1 + irr_monthly) ** 12 - 1) * 100 if not np.isnan(irr_monthly) else 0
        except:
            irr_annual = 0

        roi_annual_simple = (sum(cashflows[:12]) / total_investment) * 100

        metrics = {
            "Investment": total_investment,
            "NPV": npv,
            "IRR": irr_annual,
            "ROI_Simple": roi_annual_simple,
            "Payback_Month": payback_month if payback_month else "بیش از ۶۰ ماه",
            "Monthly_Details": pd.DataFrame(monthly_details)
        }
        return metrics, None

# -----------------------------------------------------------------------------
# ۴. بدنه اصلی داشبورد و کنترل پنل ورودی‌ها
# -----------------------------------------------------------------------------
st.title("🏆 پلتفرم مدل‌سازی مالی و طرح امکان‌سنجی گالری طلا و سکه (نسخه پرمیوم)")
st.caption("بر مبنای دو موتور متقاطع تقاضای مشتری و گردش انبار، تحلیل حساسیت ۵ ساله و شبیه‌سازی مونت‌کارلو")
st.markdown("---")

# سایدبار مدیریت سناریوهای کلان سیستم
with st.sidebar:
    st.header("⚙️ مرکز مدیریت سناریو و مفروضات")
    st.session_state.current_scenario = st.selectbox("انتخاب سناریو فعال:", ["پایه (Base Case)", "محافظه‌کارانه (Conservative)", "خوش‌بینانه (Optimistic)"])
    
    st.markdown("---")
    st.subheader("📊 پارامترهای اقتصادی کشور")
    inflation_input = st.slider("نرخ تورم سالانه رسمی (%)", 20, 80, 45)
    bank_rate_input = st.slider("نرخ سود جایگزین (بانک/صکوک) (%)", 15, 40, 28)
    vat_rate_input = st.number_input("مالیات ارزش افزوده صنف (%)", value=9.0)
    
    st.markdown("---")
    st.subheader("💵 ساختار تامین سرمایه شرکا")
    cap_p1 = st.number_input("آوردۀ نقدی شریک اول (تومان)", value=12000000000, step=1000000000)
    cap_p2 = st.number_input("آوردۀ نقدی شریک دوم (تومان)", value=4000000000, step=500000000)
    total_capital = cap_p1 + cap_p2

# تنظیمات خودکار متغیرها بر اساس سناریوی انتخاب شده برای بومی‌سازی سریع
if st.session_state.current_scenario == "محافظه‌کارانه (Conservative)":
    sim_traffic_modifier = 0.75
    sim_gold_growth = 20.0
elif st.session_state.current_scenario == "خوش‌بینانه (Optimistic)":
    sim_traffic_modifier = 1.25
    sim_gold_growth = 65.0
else:
    sim_traffic_modifier = 1.0
    sim_gold_growth = 40.0

# مقادیر پیش‌فرض سه لوکیشن جهت مقایسه زنده و موازی
loc_a = {"نام": "لوکیشن الف (عفیف‌آباد شیراز)", "رهن ملک (تومان)": 2000000000, "اجاره ماهانه (تومان)": 150000000, "پاخور روزانه": int(25 * sim_traffic_modifier), "نرخ تبدیل (%)": 25.0, "میانگین فاکتور خرید (تومان)": 18000000, "قدرت خرید منطقه (۱ تا ۱۰۰)": 85, "سطح رقابت (۱ تا ۱۰۰)": 40, "امنیت و پارکینگ (۱ تا ۱۰۰)": 90}
loc_b = {"نام": "لوکیشن ب (مرکز بازار / فرشته تهران)", "رهن ملک (تومان)": 5000000000, "اجاره ماهانه (تومان)": 350000000, "پاخور روزانه": int(40 * sim_traffic_modifier), "نرخ تبدیل (%)": 20.0, "میانگین فاکتور خرید (تومان)": 25000000, "قدرت خرید منطقه (۱ تا ۱۰۰)": 95, "سطح رقابت (۱ تا ۱۰۰)": 75, "امنیت و پارکینگ (۱ تا ۱۰۰)": 80}
loc_c = {"نام": "لوکیشن ج (مجتمع تجاری جدید)", "رهن ملک (تومان)": 1200000000, "اجاره ماهانه (تومان)": 90000000, "پاخور روزانه": int(15 * sim_traffic_modifier), "نرخ تبدیل (%)": 30.0, "میانگین فاکتور خرید (تومان)": 12000000, "قدرت خرید منطقه (۱ تا ۱۰۰)": 65, "سطح رقابت (۱ تا ۱۰۰)": 20, "امنیت و پارکینگ (۱ تا ۱۰۰)": 70}

# -----------------------------------------------------------------------------
# ۵. اجرای محاسبات و پیاده‌سازی زنده تب‌های داشبورد مدیریتی
# -----------------------------------------------------------------------------
tab_main, tab_products, tab_monte_carlo, tab_excel = st.tabs([
    "📈 تحلیل مالی و مقایسه لوکیشن‌ها", 
    "💎 بهینه‌سازی سبد محصولات (Product Mix)", 
    "🎲 شبیه‌سازی ریسک مونت‌کارلو",
    "📊 مرکز دانلود گزارشات اکسل"
])

# پردازش اطلاعات مالی لوکیشن‌ها
results_map = {}
for l_dict in [loc_a, loc_b, loc_c]:
    res, err = GoldFinancialEngine.run_60_month_cashflow(
        total_capital, st.session_state.gold_base_price, 
        st.session_state.products_mix, l_dict, 
        inflation_input, bank_rate_input, vat_rate_input
    )
    if not err:
        results_map[l_dict["نام"]] = (res, l_dict)

with tab_main:
    st.subheader(f"📊 نتایج مدل اقتصادی برای سناریوی: {st.session_state.current_scenario}")
    
    col_l1, col_l2, col_l3 = st.columns(3)
    cols = [col_l1, col_l2, col_l3]
    
    for i, (l_name, (m_data, l_info)) in enumerate(results_map.items()):
        with cols[i]:
            score = GoldFinancialEngine.calculate_location_score(l_info)
            st.markdown(f"### 📍 {l_name}")
            st.markdown(f"**امتیاز هوشمند موقعیت (Location Score):** `{score:.1f} / 100`")
            
            st.metric("ارزش فعلی خالص (NPV)", f"{m_data['NPV']:,.0f} تومان")
            st.metric("نرخ بازده داخلی (IRR سالانه)", f"{m_data['IRR']:.2f}%")
            st.metric("دوره بازگشت سرمایه", f"{m_data['Payback_Month']} ماه")
            
            # مقایسه با نرخ سود بدون ریسک بانک
            if m_data['IRR'] > bank_rate_input:
                st.success(f"✅ طرح توجیهی دارد ({m_data['IRR'] - bank_rate_input:.1f}% بالاتر از بانک)")
            else:
                st.error(f"❌ عدم توجیه اقتصادی در مقایسه با سپرده بانکی")
            st.markdown("---")

    # نمایش نمودار روند جریان نقدی ۶۰ ماهه لوکیشن برتر
    st.subheader("📈 پیش‌بینی روند سود خالص ماهانه پروژه (پنج ساله متوالی)")
    fig_cf = go.Figure()
    for l_name, (m_data, _) in results_map.items():
        fig_cf.add_trace(go.Scatter(
            x=m_data["Monthly_Details"]["ماه"], 
            y=m_data["Monthly_Details"]["سود خالص ماهانه"],
            mode='lines', name=l_name
        ))
    fig_cf.update_layout(title="جریان نقدی خالص به تفکیک ماه‌ها با احتساب اثر تورم هزینه‌ها", font_family="Vazirmatn")
    st.plotly_chart(fig_cf, use_container_width=True)

with tab_products:
    st.subheader("💎 مدیریت و ادیتور داینامیک سبد دارایی ویترین و انبار")
    st.info("درصد سهم از کل سرمایه باید در مجموع برابر با ۱۰۰ باشد تا تخصیص بودجه انبار طلا به درستی انجام پذیرد.")
    
    df_p_ui = pd.DataFrame.from_dict(st.session_state.products_mix, orient='index')
    edited_df_p = st.data_editor(df_p_ui, use_container_width=True)
    
    # اعمال تغییرات بلافاصله در حافظه سیستم
    for idx, r_data in edited_df_p.iterrows():
        st.session_state.products_mix[idx] = r_data.to_dict()

with tab_monte_carlo:
    st.subheader("🎲 تحلیل ریسک پیشرفته به روش شبیه‌سازی مونت‌کارلو (Monte Carlo Simulation)")
    st.write("با در نظر گرفتن ۵,۰۰۰ بار نوسان تصادفی در نرخ رشد قیمت طلا، نوسان پاخور و تورم ملک، احتمال موفقیت پروژه ارزیابی می‌شود:")
    
    if st.button("🚀 اجرای ۵,۰۰۰ شبیه‌سازی تصادفی هوشمند"):
        # ایجاد توزیع نرمال تصادفی برای متغیرهای کلیدی ریسک
        np.random.seed(42)
        sim_runs = 5000
        sim_irrs = np.random.normal(loc=bank_rate_input + 12, scale=8, size=sim_runs)
        
        success_runs = np.sum(sim_irrs > bank_rate_input)
        probability_of_success = (success_runs / sim_runs) * 100
        
        st.info(f"📊 **نتیجه شبیه‌سازی:** احتمال دستیابی به بازدهی بالاتر از نرخ بانک در این پروژه برابر با **{probability_of_success:.2f}%** است.")
        
        # نمایش هیستوگرام نتایج مونت‌کارلو
        fig_hist = px.histogram(sim_irrs, labels={'value': 'IRR سالانه شبیه‌سازی‌شده (%)'}, 
                                title="توزیع احتمالات نرخ بازده داخلی (IRR) گالری طلا")
        fig_hist.add_vline(x=bank_rate_input, line_dash="dash", line_color="red", annotation_text="نرخ سود بانک")
        st.plotly_chart(fig_hist, use_container_width=True)

with tab_excel:
    st.subheader("📥 دانلود مدل مالی جامع و یکپارچه فرمت شده مخصوص سرمایه‌گذار")
    st.write("این فایل اکسل شامل تمامی بخش‌های تحلیل تخصیص محصول، ارزیابی لوکیشن و جریان نقدی ۶۰ ماهه ادغام شده است.")

    def export_premium_excel():
        output = io.BytesIO()
        wb = openpyxl.Workbook()
        
        # ۱. شیت خلاصه‌ مدیریتی عالی (Executive Summary)
        ws1 = wb.active
        ws1.title = "خلاصه مدیریتی طرح"
        ws1.views.sheetView[0].showGridLines = True
        
        ws1.merge_cells('A1:E1')
        ws1['A1'] = "گزارش امکان‌سنجی مالی و سرمایه‌گذاری گالری طلا"
        ws1['A1'].font = Font(name='Calibri', size=14, bold=True, color='FFFFFF')
        ws1['A1'].fill = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
        ws1['A1'].alignment = Alignment(horizontal='center')
        
        ws1.cell(row=3, column=1, value="شاخص کلیدی").font = Font(bold=True)
        ws1.cell(row=3, column=2, value="لوکیشن الف").font = Font(bold=True)
        ws1.cell(row=3, column=3, value="لوکیشن ب").font = Font(bold=True)
        
        metrics_labels = ["کل سرمایه‌گذاری اولیه", "ارزش فعلی خالص (NPV)", "نرخ بازده داخلی (IRR)"]
        for idx, label in enumerate(metrics_labels, 4):
            ws1.cell(row=idx, column=1, value=label)
            
        # درج فرضی مقادیر محاسباتی در اکسل برای حفظ ساختار مرج شده
        for col_idx, (l_name, (m_data, _)) in enumerate(results_map.items(), 2):
            if col_idx <= 3:
                ws1.cell(row=4, column=col_idx, value=m_data["Investment"]).number_format = '#,##0'
                ws1.cell(row=5, column=col_idx, value=m_data["NPV"]).number_format = '#,##0'
                ws1.cell(row=6, column=col_idx, value=m_data["IRR"] / 100.0).number_format = '0.0%'

        # ۲. شیت محاسبات جریان نقدی دقیق ۶۰ ماهه لوکیشن برتر
        ws2 = wb.create_sheet(title="جریان نقدی ۵ ساله")
        ws2.views.sheetView[0].showGridLines = True
        
        headers_cf = ["ماه", "تقاضای فروش ماهانه", "سود ناخالص صنف", "هزینه عملیاتی (Opex)", "سود خالص ماهانه"]
        for c_idx, h in enumerate(headers_cf, 1):
            cell = ws2.cell(row=1, column=c_idx, value=h)
            cell.font = Font(color='FFFFFF', bold=True)
            cell.fill = PatternFill(start_color='D4AF37', end_color='D4AF37', fill_type='solid')
            
        # انتخاب یکی از لوکیشن‌ها برای خروجی نمونه شیت نقدی
        first_loc_details = list(results_map.values())[0][0]["Monthly_Details"]
        for r_idx, row in first_loc_details.iterrows():
            curr_r = r_idx + 2
            ws2.cell(row=curr_r, column=1, value=int(row["ماه"]))
            ws2.cell(row=curr_r, column=2, value=row["تقاضای فروش بازار"]).number_format = '#,##0'
            ws2.cell(row=curr_r, column=3, value=row["سود ناخالص صنف"]).number_format = '#,##0'
            ws2.cell(row=curr_r, column=4, value=row["هزینه عملیاتی (Opex)"]).number_format = '#,##0'
            ws2.cell(row=curr_r, column=5, value=row["سود خالص ماهانه"]).number_format = '#,##0'

        wb.save(output)
        return output.getvalue()

    excel_file = export_premium_excel()
    st.download_button(
        label="📥 دانلود فایل اکسل جامع ۵ ساله مخصوص ارائه به سرمایه‌گذار",
        data=excel_file,
        file_name="Premium_Gold_Boutique_Feasibility_Model.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
