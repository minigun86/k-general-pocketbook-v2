import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="K총무 E포켓북",
    page_icon="🏢",
    layout="wide"
)

st.markdown("""
<style>
.stApp {background:#111418;}

[data-testid="stSidebar"]{
    background:linear-gradient(
        180deg,
        #2d2826 0%,
        #231f1d 100%
    );

    border-right:3px solid #ff6b2c;
}

.stButton > button {
    width:100%;
    border-radius:10px;
    border:1px solid rgba(255,107,44,.35);
    background:#35302d;
    color:white;
}

.stButton > button:hover {
    background:#ff6b2c;
    color:white;
    border:1px solid #ff6b2c;
}

h1 {
    border-bottom:3px solid #ff5f1f;
    padding-bottom:10px;
}

h2, h3 {
    color:#ffb38d;
}

[data-testid="stMetric"] {
    background:#1c1f24;
    border:1px solid rgba(255,95,31,.35);
    border-radius:15px;
    padding:18px;
}

input {
    border:1px solid rgba(255,95,31,.35)!important;
}

input:focus {
    border:1px solid #ff5f1f!important;
    box-shadow:0 0 10px rgba(255,95,31,.35)!important;
}

hr {
    border-color:rgba(255,95,31,.25);
}
</style>
""", unsafe_allow_html=True)

excel_file = "data/K총무 포켓북.xlsx"

menu_map = {
    "👥 조직도": "조직도",
    "📋 업무현황": "현황",
    "📂 문서실": "1문서실",
    "📞 VOC": "2VOC",
    "🚗 법인차량": "3차량",
    "🅿 사외주차장": "4사외",
    "🏢 회의실": "5회의실",
    "🗂 OA룸·서식류": "6OA서식",
    "📦 비품": "7비품",
    "🎁 VIP기념품": "8VIP",
    "🖨 OA": "9OA",
    "💧 정수기": "10정수기",
    "🧊 제빙기": "11제빙기",
    "🚰 생수기": "12 생수기",
    "☕ 커피": "13 커피",
    "⭐ 구독": "구독",
    "⭐ 구독2": "구독2",
}

st.sidebar.image("images/etners_logo.png", width=100)
st.sidebar.title("K총무 E포켓북")
st.sidebar.caption("총무 업무 통합 포털")

selected_menu = st.sidebar.radio("메뉴 선택", list(menu_map.keys()))
selected_sheet = menu_map[selected_menu]

st.title("K총무 E포켓북")
st.caption("etners General Affairs Portal")

st.header(selected_menu)

try:
    df = pd.read_excel(excel_file, sheet_name=selected_sheet)

    search = st.text_input("🔍 검색어 입력", placeholder="검색어를 입력하세요")

    if search:
        df = df[
            df.astype(str)
            .apply(lambda row: row.str.contains(search, case=False, na=False).any(), axis=1)
        ]

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("📄 현재 메뉴", selected_menu)

    with col2:
        st.metric("📊 데이터", f"{len(df)}개")

    with col3:
        st.metric("🟧 브랜드", "etners")

    st.divider()

    st.dataframe(df, use_container_width=True)

except Exception as e:
    st.error("엑셀 데이터를 불러오지 못했습니다.")
    st.write(e)