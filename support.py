import streamlit as st
import pandas as pd
import re
import google.generativeai as genai

# 페이지 기본 설정
st.set_page_config(page_title="지원장학 요청서 자동 분석기", layout="wide")
st.title("📊 지원장학 요청서 자동 분석 및 유목화 웹앱")
st.markdown("여러 학교의 지원장학 요청서(Excel/CSV)를 업로드하면 일시, 담당자, 현안문제, 요청사항을 자동으로 추출하고 AI를 통해 유목화합니다.")

# Gemini API 키 입력
api_key = st.text_input("Gemini API 키를 입력하세요 (유목화 및 부서 요청사항 정리에 필요):", type="password")

# 파일 업로더
uploaded_files = st.file_uploader("장학 요청서 파일(Excel 또는 CSV)을 업로드하세요.", type=['xlsx', 'csv'], accept_multiple_files=True)

if st.button("분석 시작") and uploaded_files:
    if not api_key:
        st.warning("유목화 및 부서 요청사항 정리를 위해 Gemini API 키를 입력해 주세요.")
        st.stop()

    schedule_data = []
    issue_data = []
    request_data = []
    
    all_issues_text = ""
    all_requests_text = ""

    st.success(f"총 {len(uploaded_files)}개의 파일을 분석합니다...")
    
    for file in uploaded_files:
        try:
            if file.name.endswith('.csv'):
                df = pd.read_csv(file, header=None, dtype=str).fillna("")
            else:
                df = pd.read_excel(file, header=None, dtype=str).fillna("")

            school_name = "학교명 미상"
            supervisor_name = "장학사 미상"
            visit_date = "일시 미상"
            issue_content = "내용 없음"
            support_content = "내용 없음"

            # 학교명 추출
            for i in range(min(10, len(df))):
                row_text = "".join(df.iloc[i].values)
                if "학교" in row_text and "장학" not in row_text and "서식" not in row_text:
                    school_name = row_text.strip().replace(",", "")
                    break

            # 학교담당장학사 추출
            for i in range(min(10, len(df))):
                row_text = "".join(df.iloc[i].values)
                if "학교담당장학사" in row_text:
                    match = re.search(r'학교담당장학사\s*[:\s]*([가-힣]+)', row_text)
                    if match:
                        supervisor_name = match.group(1)
                    break

            # 표 내부 데이터 추출 (일시, 현안문제, 요청사항)
            for index, row in df.iterrows():
                row_list = list(row.values)
                for col_idx, cell_value in enumerate(row_list):
                    if "일시" in cell_value:
                        visit_date = row_list[col_idx + 1] if col_idx + 1 < len(row_list) else visit_date
                    elif "현안문제" in cell_value:
                        issue_content = row_list[col_idx + 1] if col_idx + 1 < len(row_list) else issue_content
                    elif "지원 요청 사항" in cell_value:
                        support_content = row_list[col_idx + 1] if col_idx + 1 < len(row_list) else support_content

            visit_date = visit_date.strip()
            issue_content = issue_content.strip()
            support_content = support_content.strip()

            # 데이터 취합
            schedule_data.append(f"{school_name} - {visit_date} - {supervisor_name}")
            
            if issue_content:
                issue_string = f"{school_name} - {issue_content}"
                issue_data.append(issue_string)
                all_issues_text += issue_string + "\n"

            if support_content:
                request_string = f"{school_name} - {support_content}"
                request_data.append(request_string)
                all_requests_text += request_string + "\n"

        except Exception as e:
            st.error(f"'{file.name}' 처리 중 오류 발생: {e}")

    # 결과 출력
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("1. 방문 일정 및 담당자")
        for item in schedule_data:
            st.text(item)

    with col2:
        st.subheader("2. 학교 현안문제")
        for item in issue_data:
            st.info(item)

    with col3:
        st.subheader("3. 교육활동 지원 요청")
        for item in request_data:
            st.success(item)

    st.divider()

    # AI 기반 유목화 및 부서 요청사항 도출
    st.subheader("🤖 AI 기반 유목화 및 부서 요청사항 도출")
    
    with st.spinner("Gemini AI가 내용을 분석하고 유목화 중입니다..."):
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.5-flash')

            prompt = f"""
            당신은 교육지원청의 훌륭한 장학사입니다. 아래 제공된 여러 학교의 '현안문제'와 '교육활동 지원 요청 사항'을 분석하여 다음 두 가지 형식으로 완벽하게 정리해 주세요.

            제공된 데이터:
            [학교 현안문제]
            {all_issues_text}

            [교육활동 지원 요청 사항]
            {all_requests_text}

            ---
            요청 1. 내용을 의미 있는 큰 카테고리로 유목화하세요.
            반드시 다음 형식을 지키세요:
            [유목화 주제]
            - 내용 요약 (학교명)
            - 내용 요약 (학교명)

            요청 2. 위에서 유목화한 내용을 바탕으로, 교육지원청의 관련 부서가 실제로 조치해야 할 '부서 요청 사항'으로 다시 작성하세요.
            반드시 다음 형식을 지키세요:
            [유목화 주제] - [부서에 구체적으로 요청/건의할 내용]
            """

            response = model.generate_content(prompt)
            st.markdown(response.text)
            
        except Exception as e:
            st.error(f"AI 분석 중 오류가 발생했습니다. API 키를 확인해 주세요. (에러: {e})")
