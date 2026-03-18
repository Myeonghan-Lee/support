import streamlit as st
import pandas as pd
import io

# 엑셀 다운로드를 위한 변환 함수
def to_excel(df):
    output = io.BytesIO()
    # openpyxl 엔진을 사용하여 엑셀 작성
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# 페이지 기본 설정
st.set_page_config(page_title="지원장학 요청서 자동 분석기", layout="wide")
st.title("📊 지원장학 요청서 자동 분석 및 유목화 웹앱")
st.markdown("지원장학 요청서를 업로드하면 지정된 셀에서 데이터를 추출하고, 키워드 기반 유목화 후 각각 엑셀로 다운로드할 수 있습니다.")

# 파일 업로더
uploaded_files = st.file_uploader("장학 요청서 파일(Excel 또는 CSV)을 업로드하세요.", type=['xlsx', 'csv'], accept_multiple_files=True)

if st.button("분석 시작") and uploaded_files:
    # 데이터프레임 생성을 위해 리스트(딕셔너리) 형태로 데이터 수집
    schedule_list = []
    issue_list = []
    request_list = []
    categorized_list = []
    
    # 키워드 기반 유목화 사전
    categories = {
        "시설 및 환경 개선": ["방송", "공사", "누수", "노후", "수리", "교체", "안전", "공간", "장비"],
        "예산 및 행정 지원": ["예산", "지원금", "품의", "결제", "계약", "행정", "인력", "채용", "강사"],
        "교육과정 및 학사 운영": ["교과", "학점제", "평가", "성적", "교육과정", "자유학기", "수업", "디지털", "코딩"],
        "생활지도 및 학생 지원": ["폭력", "학폭", "상담", "정서", "위기", "징계", "출결", "다문화"],
        "기타(미분류)": []
    }

    st.success(f"총 {len(uploaded_files)}개의 파일을 분석합니다...")
    
    for file in uploaded_files:
        try:
            # 파일 읽기 (헤더 없이 읽어서 인덱스로 접근)
            if file.name.endswith('.csv'):
                df = pd.read_csv(file, header=None, dtype=str).fillna("")
            else:
                df = pd.read_excel(file, header=None, dtype=str).fillna("")

            # 1. 학교명 추출 (A4셀 -> 인덱스 [3, 0])
            school_name = "학교명 미상"
            if len(df) > 3 and df.iloc[3, 0]:
                school_name = str(df.iloc[3, 0]).strip()

            # 2. 장학사 이름 추출 (A5셀의 마지막 3글자 -> 인덱스 [4, 0])
            supervisor_name = "장학사 미상"
            if len(df) > 4 and df.iloc[4, 0]:
                a5_text = str(df.iloc[4, 0]).strip()
                if len(a5_text) >= 3:
                    supervisor_name = a5_text[-3:]
                else:
                    supervisor_name = a5_text

            # 3. 일시, 현안문제, 요청사항 추출 (표 내부 검색)
            visit_date = "일시 미상"
            issue_content = ""
            support_content = ""

            for index, row in df.iterrows():
                row_list = list(row.values)
                for col_idx, cell_value in enumerate(row_list):
                    if "일시" in str(cell_value):
                        visit_date = row_list[col_idx + 1] if col_idx + 1 < len(row_list) else visit_date
                    elif "현안문제" in str(cell_value):
                        issue_content = row_list[col_idx + 1] if col_idx + 1 < len(row_list) else issue_content
                    elif "지원 요청 사항" in str(cell_value):
                        support_content = row_list[col_idx + 1] if col_idx + 1 < len(row_list) else support_content

            visit_date = visit_date.strip()
            issue_content = issue_content.strip()
            support_content = support_content.strip()

            # 추출 데이터 리스트에 추가
            schedule_list.append({"학교명": school_name, "일시": visit_date, "담당장학사": supervisor_name})
            
            if issue_content:
                issue_list.append({"학교명": school_name, "현안문제": issue_content})
            if support_content:
                request_list.append({"학교명": school_name, "지원요청사항": support_content})

            # --- 유목화 처리 함수 ---
            def classify_content(content, kind, school_nm):
                if not content or content == "내용 없음":
                    return
                
                classified = False
                for category, keywords in categories.items():
                    if category == "기타(미분류)":
                        continue
                    if any(keyword in content for keyword in keywords):
                        categorized_list.append({"유목화 주제": category, "구분": kind, "학교명": school_nm, "내용": content})
                        classified = True
                        break 
                
                if not classified:
                    categorized_list.append({"유목화 주제": "기타(미분류)", "구분": kind, "학교명": school_nm, "내용": content})

            # 현안문제와 지원요청사항 각각 유목화 목록에 추가
            classify_content(issue_content, "현안문제", school_name)
            classify_content(support_content, "지원요청사항", school_name)

        except Exception as e:
            st.error(f"'{file.name}' 처리 중 오류 발생: {e}")

    # 리스트를 Pandas DataFrame으로 변환
    df_schedule = pd.DataFrame(schedule_list)
    df_issue = pd.DataFrame(issue_list)
    df_request = pd.DataFrame(request_list)
    df_categorized = pd.DataFrame(categorized_list)

    # --- 결과 출력 및 엑셀 다운로드 버튼 ---
    st.divider()
    st.subheader("📁 데이터 추출 및 다운로드")
    
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown("**[1] 방문 일정 및 담당자**")
        st.dataframe(df_schedule, hide_index=True)
        if not df_schedule.empty:
            st.download_button("📥 일시 엑셀 다운로드", data=to_excel(df_schedule), file_name="방문일정_정리.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with col2:
        st.markdown("**[2] 학교 현안문제**")
        st.dataframe(df_issue, hide_index=True)
        if not df_issue.empty:
            st.download_button("📥 현안문제 엑셀 다운로드", data=to_excel(df_issue), file_name="현안문제_정리.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with col3:
        st.markdown("**[3] 교육활동 지원 요청**")
        st.dataframe(df_request, hide_index=True)
        if not df_request.empty:
            st.download_button("📥 요청사항 엑셀 다운로드", data=to_excel(df_request), file_name="지원요청_정리.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with col4:
        st.markdown("**[4] 키워드 유목화 결과**")
        st.dataframe(df_categorized, hide_index=True)
        if not df_categorized.empty:
            st.download_button("📥 유목화 엑셀 다운로드", data=to_excel(df_categorized), file_name="유목화_정리.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # --- 부서 요청사항 요약 출력 ---
    st.divider()
    st.subheader("🏢 부서별 조치 요청 사항 요약 (웹 확인용)")
    
    dept_mapping = {
        "시설 및 환경 개선": "교육재정상담과(또는 교육시설과)",
        "예산 및 행정 지원": "행정지원국(행정지원과)",
        "교육과정 및 학사 운영": "교육지원국(중등교육과)",
        "생활지도 및 학생 지원": "학생학부모지원센터(또는 학교통합지원센터)",
        "기타(미분류)": "관련 부서 확인 필요"
    }

    if not df_categorized.empty:
        for category in categories.keys():
            # 해당 카테고리만 필터링
            filtered_df = df_categorized[df_categorized["유목화 주제"] == category]
            if not filtered_df.empty:
                target_dept = dept_mapping[category]
                st.warning(f"**[{category} - {target_dept} 요청 사항]**")
                for _, row in filtered_df.iterrows():
                    st.write(f"- {row['내용']} ({row['학교명']})에 대한 구체적인 지원 방안 검토 요망")
