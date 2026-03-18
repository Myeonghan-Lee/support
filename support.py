import streamlit as st
import pandas as pd
import re

# 페이지 기본 설정
st.set_page_config(page_title="지원장학 요청서 자동 분석기", layout="wide")
st.title("📊 지원장학 요청서 자동 분석 및 유목화 웹앱")
st.markdown("여러 학교의 지원장학 요청서(Excel/CSV)를 업로드하면 내용을 자동으로 추출하고 키워드 기반으로 유목화합니다.")

# 파일 업로더
uploaded_files = st.file_uploader("장학 요청서 파일(Excel 또는 CSV)을 업로드하세요.", type=['xlsx', 'csv'], accept_multiple_files=True)

if st.button("분석 시작") and uploaded_files:
    schedule_data = []
    
    # 카테고리 분류를 위한 딕셔너리 준비
    # 구조: { "유목화 주제": [ "키워드1", "키워드2", ... ] }
    categories = {
        "시설 및 환경 개선": ["방송", "공사", "누수", "노후", "수리", "교체", "안전", "공간", "장비"],
        "예산 및 행정 지원": ["예산", "지원금", "품의", "결제", "계약", "행정", "인력", "채용", "강사"],
        "교육과정 및 학사 운영": ["교과", "학점제", "평가", "성적", "교육과정", "자유학기", "수업", "디지털", "코딩"],
        "생활지도 및 학생 지원": ["폭력", "학폭", "상담", "정서", "위기", "징계", "출결", "다문화"],
        "기타(미분류)": [] # 키워드에 걸리지 않는 항목들
    }

    # 분류된 데이터를 담을 결과 딕셔너리
    categorized_issues = {key: [] for key in categories.keys()}
    categorized_requests = {key: [] for key in categories.keys()}

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

            # 학교명 및 장학사 추출
            for i in range(min(10, len(df))):
                row_text = "".join(df.iloc[i].values)
                if "학교" in row_text and "장학" not in row_text and "서식" not in row_text:
                    school_name = row_text.strip().replace(",", "")
                if "학교담당장학사" in row_text:
                    match = re.search(r'학교담당장학사\s*[:\s]*([가-힣]+)', row_text)
                    if match:
                        supervisor_name = match.group(1)

            # 표 내부 데이터 추출
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

            schedule_data.append(f"{school_name} - {visit_date} - {supervisor_name}")
            
            # --- 키워드 기반 유목화 함수 ---
            def classify_content(content, target_dict, school_nm):
                if not content or content == "내용 없음":
                    return
                
                classified = False
                for category, keywords in categories.items():
                    if category == "기타(미분류)":
                        continue
                    # 내용 중에 키워드가 하나라도 포함되어 있으면 해당 카테고리로 분류
                    if any(keyword in content for keyword in keywords):
                        target_dict[category].append(f"{content} ({school_nm})")
                        classified = True
                        break # 하나의 카테고리에만 넣고 종료
                
                # 어떤 키워드에도 해당하지 않으면 기타로 분류
                if not classified:
                    target_dict["기타(미분류)"].append(f"{content} ({school_nm})")

            # 데이터 분류 실행
            classify_content(issue_content, categorized_issues, school_name)
            classify_content(support_content, categorized_requests, school_name)

        except Exception as e:
            st.error(f"'{file.name}' 처리 중 오류 발생: {e}")

    # --- 1. 기본 추출 결과 출력 ---
    st.subheader("1. 방문 일정 및 담당자")
    for item in schedule_data:
        st.text(item)
    st.divider()

    # --- 2. 조건 4: 유목화 결과 출력 ---
    st.subheader("📑 2. 주제별 유목화 결과 (모든 학교 통합)")
    
    col_issue, col_req = st.columns(2)
    
    with col_issue:
        st.markdown("**[학교 현안문제 유목화]**")
        for category, items in categorized_issues.items():
            if items: # 해당 카테고리에 데이터가 있을 때만 출력
                st.info(f"**[{category}]**")
                for item in items:
                    st.write(f"- {item}")

    with col_req:
        st.markdown("**[교육활동 지원 요청 유목화]**")
        for category, items in categorized_requests.items():
            if items:
                st.success(f"**[{category}]**")
                for item in items:
                    st.write(f"- {item}")
                    
    st.divider()

    # --- 3. 조건 5: 부서 요청 사항으로 재정리 ---
    st.subheader("🏢 3. 관련 부서별 조치 요청 사항 요약")
    st.markdown("위 유목화된 요청 사항들을 바탕으로 관련 부서에 건의할 내용으로 정리했습니다.")
    
    # 카테고리별로 담당할 가상의 부서를 매핑합니다.
    dept_mapping = {
        "시설 및 환경 개선": "교육재정상담과(또는 교육시설과)",
        "예산 및 행정 지원": "행정지원국(행정지원과)",
        "교육과정 및 학사 운영": "교육지원국(중등교육과)",
        "생활지도 및 학생 지원": "학생학부모지원센터(또는 학교통합지원센터)",
        "기타(미분류)": "관련 부서 확인 필요"
    }
    
    for category in categories.keys():
        # 현안문제와 요청사항을 합쳐서 부서에 건의
        combined_items = categorized_issues[category] + categorized_requests[category]
        if combined_items:
            target_dept = dept_mapping[category]
            st.warning(f"**[{category} - {target_dept} 요청 사항]**")
            for item in combined_items:
                st.write(f"- {item}에 대한 구체적인 지원 방안 검토 요망")
