import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="K총무 E포켓북",
    page_icon="🏢",
    layout="wide"
)

st.markdown("""
<style>
/* =========================
   ETNERS PORTAL V2.1
   밝은톤 + 모바일 메뉴 확대
========================= */

.stApp {
    background:#FFF8F2;
}

.block-container {
    padding-top:2rem;
    padding-bottom:2rem;
}

/* 사이드바 전체 */
[data-testid="stSidebar"] {
    background:linear-gradient(180deg, #FFFFFF 0%, #FFF1E8 100%);
    border-right:6px solid #FF6B2C;
}

/* 사이드바 텍스트 */
section[data-testid="stSidebar"] * {
    color:#2E2E2E !important;
}

/* 로고 박스 */
.logo-box {
    background:#FFFFFF;
    padding:18px;
    border-radius:20px;
    border:2px solid #FFD3BF;
    margin-bottom:18px;
    box-shadow:0 4px 14px rgba(255,107,44,0.12);
}

/* 사이드바 제목 */
.sidebar-title {
    font-size:26px;
    font-weight:800;
    color:#222222;
    margin-bottom:0px;
}

.sidebar-subtitle {
    font-size:15px;
    color:#E55A1C;
    font-weight:700;
    margin-bottom:18px;
}

/* 메뉴 선택 글씨 */
[data-testid="stSidebar"] label {
    font-size:24px !important;
    font-weight:800 !important;
}

/* 라디오 메뉴 전체 */
[data-testid="stSidebar"] [role="radiogroup"] {
    gap:10px;
}

/* 라디오 메뉴 한 줄 */
[data-testid="stSidebar"] [role="radiogroup"] label {
    background:#FFFFFF;
    border:2px solid #FFD3BF;
    border-radius:16px;
    padding:14px 16px;
    margin-bottom:8px;
    box-shadow:0 2px 8px rgba(0,0,0,0.04);
}

/* 라디오 메뉴 글씨 */
[data-testid="stSidebar"] [role="radiogroup"] label p,
[data-testid="stSidebar"] [role="radiogroup"] label span,
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
    font-size:22px !important;
    font-weight:800 !important;
}

/* 선택/호버 느낌 */
[data-testid="stSidebar"] [role="radiogroup"] label:hover {
    border-color:#FF6B2C;
    background:#FFF4EC;
}

/* 본문 제목 */
h1 {
    color:#2E2E2E;
    border-bottom:5px solid #FF6B2C;
    padding-bottom:14px;
    margin-bottom:10px;
}

h2, h3 {
    color:#E55A1C;
}

p, span, label, div {
    color:#333333;
}

/* 안내 카드 */
.portal-card {
    background:#FFFFFF;
    border:2px solid #FFD3BF;
    border-left:8px solid #FF6B2C;
    padding:20px;
    border-radius:20px;
    margin-bottom:24px;
    box-shadow:0 4px 14px rgba(255,107,44,0.10);
}

/* 메트릭 카드 */
[data-testid="stMetric"] {
    background:#FFFFFF;
    border:2px solid #FFD3BF;
    border-radius:18px;
    padding:20px;
    box-shadow:0 4px 14px rgba(255,107,44,0.10);
}

/* 검색창 */
input {
    background:#FFFFFF !important;
    color:#333333 !important;
    border:2px solid #FFB38D !important;
    border-radius:14px !important;
    font-size:18px !important;
}

input:focus {
    border:2px solid #FF6B2C !important;
    box-shadow:0 0 10px rgba(255,107,44,.35) !important;
}

/* 데이터프레임 */
[data-testid="stDataFrame"] {
    background:#FFFFFF;
    border:2px solid #FFD3BF;
    border-radius:16px;
}

/* 구분선 */
hr {
    border-color:#FFD3BF;
}

/* 모바일 전용 */
@media (max-width: 768px) {
    [data-testid="stSidebar"] label {
        font-size:28px !important;
    }

    [data-testid="stSidebar"] [role="radiogroup"] label {
        padding:18px 18px;
        border-radius:18px;
        margin-bottom:12px;
    }

    [data-testid="stSidebar"] [role="radiogroup"] label p,
    [data-testid="stSidebar"] [role="radiogroup"] label span {
        font-size:26px !important;
        font-weight:900 !important;
    }

    .sidebar-title {
        font-size:30px;
    }

    .sidebar-subtitle {
        font-size:18px;
    }
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

menu_description = {
    "👥 조직도": "조직 구성, 담당자, 연락처 정보를 확인하는 메뉴입니다.",
    "📋 업무현황": "총무 업무별 현황과 주요 관리 항목을 확인하는 메뉴입니다.",
    "📂 문서실": "문서실 위치, 보관 현황, 관리 기준을 확인하는 메뉴입니다.",
    "📞 VOC": "VOC 접수 현황과 처리 이력을 확인하는 메뉴입니다.",
    "🚗 법인차량": "법인차량 운영 현황, 차량 정보, 점검 내용을 확인하는 메뉴입니다.",
    "🅿 사외주차장": "사외주차장 위치, 배정 현황, 점검 내용을 확인하는 메뉴입니다.",
    "🏢 회의실": "회의실 현황, 점검 항목, 운영 기준을 확인하는 메뉴입니다.",
    "🗂 OA룸·서식류": "OA룸, 봉투, 서식류 보관 및 사용 현황을 확인하는 메뉴입니다.",
    "📦 비품": "비품 보유 현황, 지급 기준, 재고 정보를 확인하는 메뉴입니다.",
    "🎁 VIP기념품": "VIP 기념품 보유 현황과 지급 이력을 확인하는 메뉴입니다.",
    "🖨 OA": "프린터, 복합기, 용지, 토너 관련 현황을 확인하는 메뉴입니다.",
    "💧 정수기": "정수기 설치 위치, 점검 일정, 관리 현황을 확인하는 메뉴입니다.",
    "🧊 제빙기": "제빙기 설치 위치, 점검 현황, 관리 정보를 확인하는 메뉴입니다.",
    "🚰 생수기": "생수기 운영 현황과 관리 정보를 확인하는 메뉴입니다.",
    "☕ 커피": "커피머신, 커피캡슐, 관련 소모품 현황을 확인하는 메뉴입니다.",
    "⭐ 구독": "구독서비스 현황을 확인하는 메뉴입니다.",
    "⭐ 구독2": "구독서비스 세부 현황을 확인하는 메뉴입니다.",
}

with st.sidebar:
    st.markdown('<div class="logo-box">', unsafe_allow_html=True)
    st.image("images/etners_logo.png", width=165)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="sidebar-title">K총무 E포켓북</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-subtitle">ETNERS General Affairs Portal</div>', unsafe_allow_html=True)
    st.divider()

    selected_menu = st.radio("📂 메뉴 선택", list(menu_map.keys()))

selected_sheet = menu_map[selected_menu]

st.title("K총무 E포켓북")
st.caption("ETNERS General Affairs Portal · 총무 업무 통합 플랫폼")

st.markdown(
    f"""
    <div class="portal-card">
        <h3>{selected_menu}</h3>
        <p>{menu_description.get(selected_menu, "총무 업무 정보를 확인하는 메뉴입니다.")}</p>
    </div>
    """,
    unsafe_allow_html=True
)

try:
    # header=None → 엑셀 첫째 행도 데이터로 가져옴
    df = pd.read_excel(
        excel_file,
        sheet_name=selected_sheet,
        header=None
    )

    # 병합셀/빈칸 보정
    df = df.ffill().fillna("")

    search = st.text_input("🔍 검색어 입력", placeholder="검색어를 입력하세요")

    total_count = len(df)

    if search:
        df = df[
            df.astype(str)
            .apply(lambda row: row.str.contains(search, case=False, na=False).any(), axis=1)
        ]

    filtered_count = len(df)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("📄 현재 메뉴", selected_menu)

    with col2:
        st.metric("📊 전체 데이터", f"{total_count}개")

    with col3:
        st.metric("🔎 검색 결과", f"{filtered_count}개")

    st.divider()

    st.subheader("📊 데이터 조회")

    # 세련된 표 디자인
    table_html = df.to_html(
        index=False,
        header=False,
        escape=False,
        classes="pretty-table"
    )

    st.markdown("""
    <style>
    .pretty-table {
        width:100%;
        border-collapse:separate;
        border-spacing:0;
        background:white;
        border:2px solid #FFD3BF;
        border-radius:18px;
        overflow:hidden;
        box-shadow:0 4px 14px rgba(255,107,44,0.10);
        font-size:16px;
    }

    .pretty-table td {
        padding:14px 16px;
        border-bottom:1px solid #FFE2D3;
        border-right:1px solid #FFE2D3;
        color:#333333;
    }

    .pretty-table tr:first-child td {
        background:#FF6B2C;
        color:white;
        font-weight:800;
        font-size:17px;
        text-align:center;
    }

    .pretty-table td:first-child {
        background:#FFF4EC;
        font-weight:700;
        color:#333333;
    }

    .pretty-table tr:hover td {
        background:#FFF8F2;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown(table_html, unsafe_allow_html=True)

except Exception as e:
    st.error("엑셀 데이터를 불러오지 못했습니다.")
    st.write(e)