import streamlit as st
import pandas as pd
import numpy as np
import numpy_financial as npf
import openpyxl
import io
import plotly.express as px

# -----------------------------------------------------------------------------
# ۱. پیکربندی استایل واکنش‌گرا (Responsive) با فونت وزیرمتن، مشکی و سایز ۱۲
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="سیستم مدلسازی مالی گالری طلا شیراز",
    page_icon=None,
    layout="wide"
)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;500;700;800&display=swap');
    
    html, body, [data-testid="stAppViewContainer"], .stApp, p, span, div, label, input, select, textarea {
        direction: RTL !important;
        text-align: right !important;
        font-family: 'Vazirmatn', sans-serif !important;
        color: #000000 !important;
        font-size: 12px !important;
    }
    
    .stMetricValue { font-size: 16px !important; font-weight: 700 !important; color: #000000 !important; }
    h1 { font-family: 'Vazirmatn', sans-serif !important; font-weight: 800 !important; font-size: 18px !important; color: #000000 !important; }
    h2 { font-family: 'Vazirmatn', sans-serif !important; font-weight: 700 !important; font-size: 15px !important; color: #000000 !important; }
    h3 { font-family: 'Vazirmatn', sans-serif !important; font-weight: 600 !important; font-size: 13px !important; color: #000000 !important; }

    /* تنظیمات واکنش‌گرایی برای صفحه نمایش موبایل */
    @media (max-width: 768px) {
        .main .block-container {
            padding-left: 0.5rem !important;
            padding-right: 0.5rem !important;
            padding-top: 1rem !important;
        }
        [data-testid="stHorizontalBlock"] {
            flex-direction: column !important;
        }
        [data-testid="stColumn"] {
            width: 100% !important;
            flex: 1 1 100% !important;
            margin-bottom: 0.75rem !important;
        }
        .stDataFrame, .stTable {
            font-size: 10px !important;
            overflow-x: auto !important;
        }
    }
    </style>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# ۲. پایدارسازی داده‌ها و مفروضات ورودی
# -----------------------------------------------------------------------------
if 'partner_a_cap' not in st.session_state:
    st.session_state.partner_a_cap = 5000000000.0
if 'partner_b_cap' not in st.session_state:
    st.session_state.partner_b_cap = 2000000000.0
if 'gold_price_per_gram' not in st.session_state:
    st.session_state.gold_price_per_gram = 30000000.0

if 'inventory_mix' not in st.session_state:
    st.session_state.inventory_mix = {
        "انگشتر سبک": {"weight_share": 21.5, "ojarat_pct": 18.0, "profit_pct": 7.0},
        "گردنبند": {"weight_share": 21.5, "ojarat_pct": 15.0, "profit_pct": 7.0},
        "دستبند": {"weight_share": 18.4, "ojarat_pct": 14.0, "profit_pct": 7.0},
        "گوشواره": {"weight_share": 12.3, "ojarat_pct": 16.0, "profit_pct": 7.0},
        "حلقه ازدواج": {"weight_share": 9.2, "ojarat_pct": 12.0, "profit_pct": 5.0},
        "نیم‌ست": {"weight_share": 8.0, "ojarat_pct": 20.0, "profit_pct": 7.0},
        "طلای سرمایه‌ای / شمش": {"weight_share": 9.1, "ojarat_pct": 2.0, "profit_pct": 2.0},
    }

# -----------------------------------------------------------------------------
# ۳. موتور پردازش جامع هزینه‌ها و امکان‌سنجی مالی
# -----------------------------------------------------------------------------
class GoldBoutiqueEngine:
    
    @staticmethod
    def calculate_location_matrix(rent_base, traffic_base, conversion_base, ticket_base):
        locations = {
            "ملاصدرا (مرکز سنتی هدفمند)": {"traffic": traffic_base * 1.2, "quality": 85, "competition": 80, "rent": rent_base * 1.4, "brand": 90},
            "عفیف‌آباد (لوکس و خانوادگی)": {"traffic": traffic_base * 1.0, "quality": 95, "competition": 50, "rent": rent_base * 1.2, "brand": 95},
            "معالی‌آباد (مدرن و جوان)": {"traffic": traffic_base * 1.1, "quality": 80, "competition": 45, "rent": rent_base * 1.0, "brand": 85},
            "بازار زرگرها (حجم سنتی بالا)": {"traffic": traffic_base * 1.5, "quality": 70, "competition": 95, "rent": rent_base * 1.5, "brand": 50}
        }
        scored_locs = []
        for name, p in locations.items():
            score = (p["traffic"] / traffic_base * 20) + (p["quality"] * 0.25) + ((100 - p["competition"]) * 0.15) + (20 * (rent_base / p["rent"])) + (p["brand"] * 0.20)
            scored_locs.append({"موقعیت": name, "امتیاز هوشمند (R²)": min(score, 100), "کشش پاخور (μ)": p["traffic"], "نسبت هزینه اجاره (Δ)": p["rent"]})
        return pd.DataFrame(scored_locs)

    @staticmethod
    def run_5_year_feasibility(total_capital, gold_price, inv_mix, daily_traffic, conversion_rate, avg_ticket_gram, monthly_rent, manager_salary, staff_salary, utilities_cost, marketing_monthly, inflation_rate):
        capex_setup = {
            "ودیعه و رهن تجاری ملک": monthly_rent * 20,
            "طراحی داخلی، دکوراسیون و نورپردازی تخصصی": 450000000,
            "ساخت و نصب ویترین‌های ضدسرقت و شیشه لمینت": 250000000,
            "خرید گاوصندوق‌های سنگین بانکی و خزانه امن": 200000000,
            "سیستم مداربسته، مانیتورینگ و دزدگیر پیشرفته": 180000000,
            "راه اندازی شبکه، نرم‌افزار حسابداری و CRM": 60000000,
            "کمپین برندینگ، افتتاحیه و بازاریابی اولیه": 300000000,
            "هزینه اخذ مجوزهای اتحادیه و اتاق اصناف": 50000000,
            "تجهیزات اداری و رفاهی گالری": 40000000
        }
        total_capex_fixed = sum(capex_setup.values())
        working_capital_gold = total_capital - total_capex_fixed
        
        if working_capital_gold <= 0:
            return None, "سرمایه نقدی اولیه جهت خرید طلای انبار و پوشش CAPEX کافی نیست. لطفاً سرمایه را افزایش دهید."

        initial_gold_grams = working_capital_gold / gold_price
        
        monthly_days = 26
        monthly_sales_grams_demand = daily_traffic * (conversion_rate / 100.0) * avg_ticket_gram * monthly_days
        
        monthly_records = []
        cashflows = []
        cumulative_cashflow = -total_capital
        payback_month = None
        
        base_fixed_opex = monthly_rent + manager_salary + staff_salary + utilities_cost + marketing_monthly
        
        for m in range(1, 61):
            year_idx = (m - 1) // 12
            inflation_modifier = (1 + inflation_rate / 100.0) ** year_idx
            current_opex = base_fixed_opex * inflation_modifier
            current_gold_price = gold_price * (1 + (inflation_rate - 5) / 100.0) ** year_idx
            
            total_revenue = 0
            total_gross_profit = 0
            
            for p_name, p_data in inv_mix.items():
                share_grams = monthly_sales_grams_demand * (p_data["weight_share"] / 100.0)
                gold_base_value = share_grams * current_gold_price
                ojarat_value = gold_base_value * (p_data["ojarat_pct"] / 100.0)
                gallery_profit_value = (gold_base_value + ojarat_value) * (p_data["profit_pct"] / 100.0)
                
                total_revenue += (gold_base_value + ojarat_value + gallery_profit_value)
                total_gross_profit += (ojarat_value * 0.3 + gallery_profit_value)
                
            ebitda = total_gross_profit - current_opex
            tax_allowance = max(0.0, ebitda * 0.15)
            net_profit = ebitda - tax_allowance
            
            cashflows.append(net_profit)
            monthly_records.append({
                "ماه": m,
                "حجم فروش (گرم)": monthly_sales_grams_demand,
                "کل گردش مالی (ریال)": total_revenue,
                "سود ناخالص": total_gross_profit,
                "هزینه جاری (OPEX)": current_opex,
                "سود خالص عملیاتی": net_profit
            })
            
            if payback_month is None:
                cumulative_cashflow += net_profit
                if cumulative_cashflow >= 0:
                    payback_month = m

        discount_rate_monthly = (1 + 0.28) ** (1/12) - 1
        npv_value = -total_capital + sum(cf / ((1 + discount_rate_monthly) ** i) for i, cf in enumerate(cashflows, 1))
        
        try:
            irr_m = npf.irr([-total_capital] + cashflows)
            irr_annual = ((1 + irr_m) ** 12 - 1) * 100 if not np.isnan(irr_m) else 0.0
        except:
            irr_annual = 0.0

        return {
            "capex_table": capex_setup,
            "total_capex": total_capex_fixed,
            "base_opex": base_fixed_opex,
            "gold_inventory_weight": initial_gold_grams,
            "monthly_df": pd.DataFrame(monthly_records),
            "npv": npv_value,
            "irr": irr_annual,
            "payback": payback_month if payback_month else "بیش از ۶۰ ماه"
        }, None

# -----------------------------------------------------------------------------
# ۴. کنترل پنل داشبورد هوشمند
# -----------------------------------------------------------------------------
st.title("سیستم مدلسازی مالی و امکان‌سنجی گالری طلا (شیراز)")
st.markdown("---")

with st.sidebar:
    st.header("مدل مالکیت و ساختار سرمایه")
    st.session_state.partner_a_cap = st.number_input("آورده شریک الف (تومان):", value=5000000000, step=500000000)
    st.session_state.partner_b_cap = st.number_input("آورده شریک ب (تومان):", value=2000000000, step=500000000)
    
    total_pool = st.session_state.partner_a_cap + st.session_state.partner_b_cap
    equity_a = (st.session_state.partner_a_cap / total_pool) * 100
    equity_b = (st.session_state.partner_b_cap / total_pool) * 100
    
    st.markdown(f"نسبت سهام:\n* شریک الف: `{equity_a:.1f}%`\n* شریک ب: `{equity_b:.1f}%`")
    
    st.markdown("---")
    st.header("تنظیمات هزینه‌های جاری ماهانه (OPEX)")
    rent_input = st.number_input("اجاره ماهانه ملک (تومان):", value=45000000, step=5000000)
    manager_salary = st.number_input("حقوق مدیریت (تومان):", value=25000000, step=5000000)
    staff_salary = st.number_input("حقوق پرسنل فروش و حراست (تومان):", value=30000000, step=5000000)
    utilities_cost = st.number_input("هزینه آب، برق، اینترنت و شارژ (تومان):", value=5000000, step=1000000)
    marketing_monthly = st.number_input("هزینه مستمر بازاریابی و تبلیغات (تومان):", value=10000000, step=2000000)
    
    st.markdown("---")
    st.header("پارامترهای کلان بازار")
    st.session_state.gold_price_per_gram = st.number_input("قیمت هر گرم طلا (تومان):", value=30000000, step=500000)
    inflation_input = st.slider("نرخ تورم سالانه (%):", 20, 60, 45)

results, err_msg = GoldBoutiqueEngine.run_5_year_feasibility(
    total_pool, st.session_state.gold_price_per_gram, st.session_state.inventory_mix,
    daily_traffic=10, conversion_rate=30.0, avg_ticket_gram=5.0,
    monthly_rent=rent_input, manager_salary=manager_salary, staff_salary=staff_salary,
    utilities_cost=utilities_cost, marketing_monthly=marketing_monthly, inflation_rate=inflation_input
)

if err_msg:
    st.error(err_msg)
else:
    tab_costs, tab_dashboard, tab_inventory, tab_locations, tab_export = st.tabs([
        "هزینه‌های اولیه ($\Sigma \text{ CAPEX}$) و جاری ($\Sigma \text{ OPEX}$)",
        "داشبورد ارزیابی ($\Sigma \text{ KPI}$)", 
        "ویترین و انبار", 
        "لوکیشن",
        "اکسل"
    ])
    
    with tab_costs:
        st.subheader("هزینه‌های راه‌اندازی اولیه ($\Sigma \text{ CAPEX}$)")
        df_capex = pd.DataFrame(list(results["capex_table"].items()), columns=["شرح هزینه", "مبلغ (تومان)"])
        st.dataframe(df_capex, use_container_width=True)
        st.metric("جمع کل ($\Sigma \text{ CAPEX}$)", f"{results['total_capex']:,.0f} تومان")
        
        st.markdown("---")
        st.subheader("هزینه‌های جاری ماهانه ($\Sigma \text{ OPEX}$)")
        base_opex_dict = {
            "اجاره ماهانه ملک": rent_input,
            "حقوق مدیریت اجرایی": manager_salary,
            "حقوق پرسنل فروش و حراست": staff_salary,
            "هزینه انرژی و شارژ": utilities_cost,
            "تبلیغات و بازاریابی ماهانه": marketing_monthly
        }
        df_opex = pd.DataFrame(list(base_opex_dict.items()), columns=["شرح هزینه", "مبلغ ماهانه (تومان)"])
        st.dataframe(df_opex, use_container_width=True)
        st.metric("جمع کل ($\Sigma \text{ OPEX}$)", f"{results['base_opex']:,.0f} تومان")

    with tab_dashboard:
        st.subheader("شاخص‌های اقتصادی طرح ($\mu, NPV, IRR$)")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("کل سرمایه", f"{total_pool:,.0f} تومان")
        with c2:
            st.metric("ارزش فعلی خالص ($NPV$)", f"{results['npv']:,.0f} تومان")
        with c3:
            st.metric("بازده داخلی ($IRR$)", f"{results['irr']:.1f}%")
        with c4:
            st.metric("دوره بازگشت", f"{results['payback']} ماه")
            
        st.markdown("---")
        fig = px.area(results["monthly_df"], x="ماه", y="سود خالص عملیاتی", title="روند ارزش سود خالص ماهیانه")
        st.plotly_chart(fig, use_container_width=True)

    with tab_inventory:
        st.subheader("توزیع وزن و بودجه در سبد ویترین")
        inv_data = []
        for k, v in st.session_state.inventory_mix.items():
            allocated_weight = results['gold_inventory_weight'] * (v["weight_share"] / 100.0)
            allocated_cash = allocated_weight * st.session_state.gold_price_per_gram
            inv_data.append({
                "دسته محصول": k,
                "سهم وزنی (%)": v["weight_share"],
                "وزن (گرم)": round(allocated_weight, 2),
                "ارزش (تومان)": round(allocated_cash)
            })
        st.dataframe(pd.DataFrame(inv_data), use_container_width=True)

    with tab_locations:
        st.subheader("ماتریس رتبه‌بندی مکان‌یابی شیراز")
        df_locs_res = GoldBoutiqueEngine.calculate_location_matrix(rent_input, 10, 30.0, 5.0)
        st.dataframe(df_locs_res, use_container_width=True)

    with tab_export:
        st.subheader("دانلود نسخه اکسل محاسباتی")
        def generate_dynamic_excel(res_data):
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            ws_dash = wb.active
            ws_dash.title = "Financial Summary"
            ws_dash.cell(row=1, column=1, value="ارزیابی طرح مالی گالری طلا")
            ws_dash.cell(row=3, column=1, value="NPV").value = res_data['npv']
            wb.save(output)
            return output.getvalue()
        
        xl_data = generate_dynamic_excel(results)
        st.download_button(
            label="دانلود فایل اکسل تنظیمات مالی",
            data=xl_data,
            file_name="Gold_Financial_Model_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
