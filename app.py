import streamlit as st
from components.styles import load_css


st.set_page_config(
    page_title="Gold Gallery OS",
    page_icon="🟡",
    layout="wide"
)


load_css()


st.title(
"🟡 Gold Gallery Investment OS"
)


st.markdown("""
سیستم هوشمند امکان‌سنجی راه‌اندازی گالری طلا

شیراز | اهواز | مرودشت | ارسنجان

""")


# Global State

if "model" not in st.session_state:

    st.session_state.model = {

        "city":"شیراز",

        "gold_price":7000000,

        "inventory":750,

        "cash":1200000000,

        "sales":150,

        "rent":150000000,

        "salary":80000000

    }


st.sidebar.success(
"سیستم آماده تحلیل است"
)


st.sidebar.write(
st.session_state.model
)
