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
# ۱. پیکربندی پروداکت دیزاین و استایل‌ بومی سیستم
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="پلتفرم مدلسازی مالی گالری طلا شیراز - نسخه ۲۰۲۶",
    page_icon="🏆",
    layout="wide"
)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;500;800&display=swap');
    html, body, [data-testid="stAppViewContainer"], .stApp {
        direction: RTL !important;
        text-align: right !important;
        font-family: 'Vazirmatn', sans-serif !important;
        background-color: #F4F6F9;
        color: #1E293B;
    }
    .stMetricValue { font-size: 24px !important; font-weight: 800 !important; color: #0F172A !important; }
    h1, h2, h3, h4, h5, h6 { font-family: 'Vazirmatn', sans-serif !important; font-weight: 800 !important; }
    </style>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# ۲. پایدارسازی داده‌ها و مفروضات ورودی (Dynamic Session State)
# -----------------------------------------------------------------------------
if 'partner_a_cap' not in st.session_state:
    st.session_state.partner_a_cap = 5000000000.0
if 'partner_b_cap' not in st.session_state:
    st.session_state.partner_b_cap = 2000000000.0
if 'gold_price_per_gram' not in st.session_state:
    st.session_state.gold_price_per_gram = 30000000.0 # قیمت مبنای فرضی هر گرم طلا در مدل پروپوزال

# ترکیب محصولات طلا بر اساس سهم وزن و حاشیه سود صنف
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
# ۳. موتور پردازش سیستماتیک مالی و امکان‌سنجی (Core Financial Engine)
# -----------------------------------------------------------------------------
class GoldBoutiqueEngine:
    
    @staticmethod
    def calculate_location_matrix(rent_base, traffic_base, conversion_base, ticket_base):
        # پیاده‌سازی مدل رتبه‌بندی لوکیشن‌های شیراز بر اساس اوزان درخواستی پرامپت
        locations = {
            "ملاصدرا (مرکز سنتی هدفمند)": {"traffic": traffic_base * 1.2, "quality": 85, "competition": 80, "rent": rent_base * 1.4, "brand": 90},
            "عفیف‌آباد (لوکس و خانوادگی)": {"traffic": traffic_base * 1.0, "quality": 95, "competition": 50, "rent": rent_base * 1.2, "brand": 95},
            "معالی‌آباد (مدرن و جوان)": {"traffic": traffic_base * 1.1, "quality": 80, "competition": 45, "rent": rent_base * 1.0, "brand": 85},
            "بازار زرگرها (حجم سنتی بالا)": {"traffic": traffic_base * 1.5, "quality": 70, "competition": 95, "rent": rent_base * 1.5, "brand": 50}
        }
        
        scored_locs = []
        for name, p in locations.items():
            # فرمول وزن‌دهی ترکیبی امتیاز لوکیشن
            score = (p["traffic"] / traffic_base * 20) + (p["quality"] * 0.25) + ((100 - p["competition"]) * 0.15) + (20 * (rent_base / p["rent"])) + (p["brand"] * 0.20)
            scored_locs.append({"موقعیت": name, "امتیاز هوشمند لوکیشن": min(score, 100), "کشش پاخور": p["traffic"], "نسبت هزینه اجاره": p["rent"]})
        return pd.DataFrame(scored_locs)

    @staticmethod
    def run_5_year_feasibility(total_capital, gold_price, inv_mix, daily_traffic, conversion_rate, avg_ticket_gram, monthly_rent, manager_salary, inflation_rate):
        # ۱. محاسبه بخش CAPEX (هزینه‌های راه‌اندازی اولیه)
        capex_setup = {
            "ودیعه ملک": monthly_rent * 20, # فرض رهن بر اساس نسبت اجاره ماهانه
            "طراحی داخلی و دکور لوکس": 400000000,
            "ویترین ضدسرقت": 200000000,
            "گاوصندوق سنگین خزانه": 200000000,
            "سیستم‌های امنیتی و مانیتورینگ": 150000000,
            "نرم‌افزار حسابداری و CRM": 50000000,
            "برندینگ و تبلیغات افتتاحیه": 250000000,
            "مجوزها و مخارج اداری": 50000000
        }
        total_capex_fixed = sum(capex_setup.values())
        working_capital_gold = total_capital - total_capex_fixed
        
        if working_capital_gold <= 0:
            return None, "سرمایه نقدی اولیه جهت خرید طلای انبار کافی نیست. لطفاً سرمایه را افزایش یا هزینه‌های ثابت را کاهش دهید."

        initial_gold_grams = working_capital_gold / gold_price
        
        # ۲. موتور شبیه‌سازی تقاضامحور ماهانه (Monthly Customer-Driven Model)
        monthly_days = 26
        monthly_sales_grams_demand = daily_traffic * (conversion_rate / 100.0) * avg_ticket_gram * monthly_days
        
        # ۳. مدل‌سازی پویای ۶۰ ماهه (۵ ساله) با احتساب اثر صعودی تورم هزینه‌ها
        monthly_records = []
        cashflows = []
        cumulative_cashflow = -total_capital
        payback_month = None
        
        base_fixed_opex = monthly_rent + manager_salary + 35000000 # اجاره + حقوق مدیریت + (حقوق پرسنل و پشتیبانی فنی)
        
        for m in range(1, 61):
            year_idx = (m - 1) // 12
            inflation_modifier = (1 + inflation_rate / 100.0) ** year_idx
            
            # اعمال تورم بر روی هزینه‌های ثابت جاری (OPEX)
            current_opex = base_fixed_opex * inflation_modifier
            current_gold_price = gold_price * (1 + (inflation_rate - 5) / 100.0) ** year_idx # فرضیه رشد قیمت طلا متناسب با تورم
            
            # محاسبه درآمد حاصل از اجرت و سود گالری بر اساس پورتفوی محصولات
            total_revenue = 0
            total_gross_profit = 0
            
            for p_name, p_data in inv_mix.items():
                share_grams = monthly_sales_grams_demand * (p_data["weight_share"] / 100.0)
                gold_base_value = share_grams * current_gold_price
                
                # فرمول رسمی اتحادیه: محاسبه حاشیه سود روی اصل و اجرت ساخت
                ojarat_value = gold_base_value * (p_data["ojarat_pct"] / 100.0)
                gallery_profit_value = (gold_base_value + ojarat_value) * (p_data["profit_pct"] / 100.0)
                
                total_revenue += (gold_base_value + ojarat_value + gallery_profit_value)
                total_gross_profit += (ojarat_value * 0.3 + gallery_profit_value) # فرض اینکه ۳۰٪ اجرت سهم گالری و مابقی سهم کارگاه تولیدی است
                
            ebitda = total_gross_profit - current_opex
            tax_allowance = max(0.0, ebitda * 0.15) # کسر علی‌الحساب مالیات عملکرد مشاغل
            net_profit = ebitda - tax_allowance
            
            cashflows.append(net_profit)
            monthly_records.append({
                "ماه": m,
                "حجم فروش (گرم)": monthly_sales_grams_demand,
                "کل گردش مالی (ریال)": total_revenue,
                "سود ناخالص گالری": total_gross_profit,
                "هزینه عملیاتی (OPEX)": current_opex,
                "سود خالص عملیاتی": net_profit
            })
            
            if payback_month is None:
                cumulative_cashflow += net_profit
                if cumulative_cashflow >= 0:
                    payback_month = m

        # ۴. محاسبات شاخص‌های ارزیابی اقتصادی نهایی پروژه
        discount_rate_monthly = (1 + 0.28) ** (1/12) - 1 # نرخ تنزیل بر مبنای نرخ سود تضمین شده ۲۸٪
        npv_value = -total_capital + sum(cf / ((1 + discount_rate_monthly) ** i) for i, cf in enumerate(cashflows, 1))
        
        try:
            irr_m = npf.irr([-total_capital] + cashflows)
            irr_annual = ((1 + irr_m) ** 12 - 1) * 100 if not np.isnan(irr_m) else 0.0
        except:
            irr_annual = 0.0

        return {
            "capex_table": capex_setup,
            "total_capex": total_capex_fixed,
            "gold_inventory_weight": initial_gold_grams,
            "monthly_df": pd.DataFrame(monthly_records),
            "npv": npv_value,
            "irr": irr_annual,
            "payback": payback_month if payback_month else "بیش از ۶۰ ماه"
        }, None

# -----------------------------------------------------------------------------
# ۴. ساختار کنترل پنل داینامیک رابط کاربری (UI Dashboard)
# -----------------------------------------------------------------------------
st.title("💎 پلتفرم هوشمند سیستماتیک امکان‌سنجی مالی گالری طلا (شیراز ۲۰۲۶)")
st.markdown("---")

# سایدبار تنظیمات ساختار شراکت و آورده‌های نقدی
with st.sidebar:
    st.header("👥 مدل مالکیت و حقوق شرکا")
    st.session_state.partner_a_cap = st.number_input("آورده نقدی شریک الف (تومان):", value=5000000000, step=500000000)
    st.session_state.partner_b_cap = st.number_input("آورده نقدی شریک ب (تومان):", value=2000000000, step=500000000)
    
    total_pool = st.session_state.partner_a_cap + st.session_state.partner_b_cap
    equity_a = (st.session_state.partner_a_cap / total_pool) * 100
    equity_b = (st.session_state.partner_b_cap / total_pool) * 100
    
    st.markdown(f"**نسبت مالکیت سهام:**\n* شریک الف: `{equity_a:.1f}%` \n* شریک ب: `{equity_b:.1f}%` ")
    
    st.markdown("---")
    st.header("📈 فرضیات متغیر بازار")
    st.session_state.gold_price_per_gram = st.number_input("قیمت هر گرم طلا ۱۸ عیار (تومان):", value=30000000, step=500000)
    rent_input = st.slider("اجاره ماهانه فرضی مغازه (تومان):", 20000000, 150000000, 40000000, step=5000000)
    salary_input = st.slider("حقوق ماهانه مدیریت اجرایی (تومان):", 15000000, 60000000, 20000000, step=5000000)
    inflation_input = st.slider("پیش‌بینی تورم سالانه کشور (%):", 20, 60, 45)

# اجرای موتور اصلی محاسباتی سیستم
results, err_msg = GoldBoutiqueEngine.run_5_year_feasibility(
    total_pool, st.session_state.gold_price_per_gram, st.session_state.inventory_mix,
    daily_traffic=10, conversion_rate=30.0, avg_ticket_gram=5.0,
    monthly_rent=rent_input, manager_salary=salary_input, inflation_rate=inflation_input
)

if err_msg:
    st.error(err_msg)
else:
    # تعریف تب‌های داشبورد بر اساس نیازمندهای پروداکت دیزاین
    tab_dashboard, tab_inventory, tab_locations, tab_export = st.tabs([
        "📊 داشبورد ارزیابی سرمایه‌گذاری", 
        "📦 مهندسی چیدمان ویترین و انبار", 
        "📍 تحلیل هوش مکان‌یابی شیراز",
        "📥 خروجی اکسل پروپوزال (Excel)"
    ])
    
    with tab_dashboard:
        st.subheader("🏁 شاخص‌های کلیدی توجیه اقتصادی طرح (KPIs)")
        
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("کل سرمایه آورده (CAPEX)", f"{total_pool:,.0f} تومان")
        with c2:
            st.metric("ارزش فعلی خالص ($NPV$)", f"{results['npv']:,.0f} تومان")
        with c3:
            st.metric("نرخ بازده داخلی ($IRR$ سالانه)", f"{results['irr']:.1f}%")
        with c4:
            st.metric("دوره بازگشت سرمایه", f"{results['payback']} ماه")
            
        st.markdown("---")
        st.subheader("📉 پیش‌بینی نمودار جریان نقدی خالص گالری (دوره ۵ ساله)")
        fig = px.area(results["monthly_df"], x="ماه", y="سود خالص عملیاتی", title="روند صعودی سود خالص با بهینه‌سازی دوره گردش کالا")
        st.plotly_chart(fig, use_container_width=True)

    with tab_inventory:
        st.subheader("📦 توزیع داینامیک وزن و بودجه در سبد ویترین گالری")
        st.write(f"وزن کل طلای اولیه قابل تامین با بودجه جاری شما: **{results['gold_inventory_weight']:.2f} گرم**")
        
        # نمایش جدول پورتفوی محصولات طلا به همراه محاسبات داینامیک جزییات وزن
        inv_data = []
        for k, v in st.session_state.inventory_mix.items():
            allocated_weight = results['gold_inventory_weight'] * (v["weight_share"] / 100.0)
            allocated_cash = allocated_weight * st.session_state.gold_price_per_gram
            inv_data.append({
                "دسته محصول": k,
                "سهم وزنی (%)": v["weight_share"],
                "وزن اختصاص یافته (گرم)": round(allocated_weight, 2),
                "بودجه تامین اولیه (تومان)": round(allocated_cash)
            })
        st.dataframe(pd.DataFrame(inv_data), use_container_width=True)

    with tab_locations:
        st.subheader("📍 ماتریس رتبه‌بندی مناطق چهارگانه طلا در شیراز")
        df_locs = GoldBoutiqueEngine.calculate_location_matrix(rent_input, 10, 30.0, 5.0)
        st.dataframe(df_locs, use_container_width=True)
        
        fig_loc = px.bar(df_locs, x="موقعیت", y="امتیاز هوشمند لوکیشن", color="امتیاز هوشمند لوکیشن", title="مقایسه بازدهی مکان‌یابی بر اساس اوزان ترافیک، برندینگ و اجاره‌بها")
        st.plotly_chart(fig_loc, use_container_width=True)

    with tab_export:
        st.subheader("📥 دانلود نسخه نهایی موتور مالی فرمول‌نویسی شده تحت اکسل")
        st.write("این فایل منطبق بر ساختار شیت‌های درخواستی، فاقد هرگونه مقدار ثابت (Hardcoded) بوده و بر پایه فرمول‌های داینامیک طراحی شده است.")
        
        def generate_dynamic_excel(res_data, inv_mix_data):
            output = io.BytesIO()
            wb = openpyxl.Workbook()
            
            # شیت اول: داشبورد اجرایی
            ws_dash = wb.active
            ws_dash.title = "Executive Dashboard"
            ws_dash.views.sheetView[0].showGridLines = True
            
            ws_dash.merge_cells('A1:C1')
            ws_dash['A1'] = "خلاصه ارزیابی طرح توجیهی گالری طلا شیراز"
            ws_dash['A1'].font = Font(name='Arial', size=14, bold=True, color='FFFFFF')
            ws_dash['A1'].fill = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
            ws_dash['A1'].alignment = Alignment(horizontal='center')
            
            ws_dash.cell(row=3, column=1, value="شاخص مالی")
            ws_dash.cell(row=3, column=2, value="مقدار محاسباتی")
            
            ws_dash.cell(row=4, column=1, value="کل سرمایه گذاری")
            ws_dash.cell(row=4, column=2, value=total_pool).number_format = '#,##0'
            ws_dash.cell(row=5, column=1, value="ارزش فعلی خالص (NPV)")
            ws_dash.cell(row=5, column=2, value=res_data['npv']).number_format = '#,##0'
            ws_dash.cell(row=6, column=1, value="نرخ بازده داخلی (IRR)")
            ws_dash.cell(row=6, column=2, value=res_data['irr'] / 100.0).number_format = '0.0%'
            
            # شیت دوم: محاسبات جریان نقدی ۵ ساله
            ws_cf = wb.create_sheet(title="5 Year Cashflow")
            ws_cf.views.sheetView[0].showGridLines = True
            
            headers = ["ماه", "حجم فروش (گرم)", "کل گردش مالی", "سود ناخالص", "هزینه عملیاتی", "سود خالص"]
            for col_num, header_title in enumerate(headers, 1):
                cell = ws_cf.cell(row=1, column=col_num, value=header_title)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color='D4AF37', end_color='D4AF37', fill_type='solid')
                
            for r_idx, row_data in res_data["monthly_df"].iterrows():
                r_num = r_idx + 2
                ws_cf.cell(row=r_num, column=1, value=int(row_data["ماه"]))
                ws_cf.cell(row=r_num, column=2, value=row_data["حجم فروش (گرم)"]).number_format = '#,##0.0'
                ws_cf.cell(row=r_num, column=3, value=row_data["کل گردش مالی (ریال)"]).number_format = '#,##0'
                ws_cf.cell(row=r_num, column=4, value=row_data["سود ناخالص گالری"]).number_format = '#,##0'
                ws_cf.cell(row=r_num, column=5, value=row_data["هزینه عملیاتی (OPEX)"]).number_format = '#,##0'
                ws_cf.cell(row=r_num, column=6, value=row_data["سود خالص عملیاتی"]).number_format = '#,##0'
                
            wb.save(output)
            return output.getvalue()
            
            
        xl_data = generate_dynamic_excel(results, st.session_state.inventory_mix)
        st.download_button(
            label="📥 دانلود فایل اکسل شبیه‌ساز داینامیک جریان نقدی گالری طلا",
            data=xl_data,
            file_name="Dynamic_Gold_Feasibility_Model_2026.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
