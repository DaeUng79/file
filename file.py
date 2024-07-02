import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="엑셀모아모아", layout="wide")

st.title('📊 부서별 엑셀모아모아') 

department_order = [
    "기획예산담당관", "소통담당관", "감사담당관", "민생경제과", "기업지원과", "세무과", "징수과", "위생과", "산업혁신과",
    "공간혁신과", "문화관광과", "교육체육과", "주민생활지원과", "노인장애인과", "여성청소년과", "아동보육과", "행정과",
    "회계과", "정보통계과", "토지정보과", "종합민원과", "기후환경과", "수질관리과", "자원순환과", "하천과", "산림과",
    "공원과", "시민안전과", "도시계획과", "건설도로과", "도로정비과", "교통과", "건축과", "공동주택과", "원스톱허가과",
    "총무과", "문화복지과", "도시관리과", "허가과", "보건행정과", "건강증진과", "웅상보건지소", "농정과", "동물보호과",
    "농업기술과", "수도과", "정수과", "하수과", "차량등록사업소", "시립박물관", "시립도서관", "물금읍", "동면", "원동면",
    "상북면", "하북면", "중앙동", "양주동", "삼성동", "강서동", "서창동", "소주동", "평산동", "덕계동", "의회사무국"
]

col1, col2 = st.columns([3, 7])

with col1:
    # 여러 파일 업로드
    uploaded_files = st.file_uploader("**부서에서 받은 엑셀 파일을 한번에 업로드하세요**", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    with col2:
    
        skiprows_value = st.number_input("**✂️ 엑셀 데이터 전처리(필요 없는 행 개수 설정)**", min_value=0, step=1, value=0)
        
        # 첫 번째 파일에 대해 skiprows 값을 입력받기
        first_file = uploaded_files[0]
        initial_df = pd.read_excel(first_file)
    
        # skiprows 값을 반영하여 첫 번째 파일 데이터프레임 읽기
        df_list = []
        first_df = pd.read_excel(first_file, skiprows=skiprows_value)
        df_list.append(first_df)
        
        # 나머지 파일도 동일한 skiprows 값을 적용하여 데이터프레임 읽기
        for uploaded_file in uploaded_files[1:]:
            df = pd.read_excel(uploaded_file, skiprows=skiprows_value)
            df_list.append(df)
        
        # 모든 데이터프레임을 하나로 합치기
        combined_df = pd.concat(df_list, ignore_index=True)
        st.dataframe(combined_df,hide_index=True, height=300)

    col3, col4 = st.columns([3, 7])

    with col4:

        # 부서명 컬럼 선택
        부서_column = st.selectbox("**🗃️ 부서별 정렬을 위한 '부서명' 컬럼을 선택하세요**", combined_df.columns)
        
        # 부서 순서를 카테고리형으로 설정 후 정렬
        combined_df[부서_column] = pd.Categorical(combined_df[부서_column], categories=department_order, ordered=True)
        combined_df = combined_df.sort_values(by=부서_column)
        
        st.warning(f"**'{부서_column}' 컬럼으로 정렬되었습니다. 추가 정렬 및 수정은 엑셀 파일을 다운로드 받아 활용하시기 바랍니다.**")
        st.dataframe(combined_df,hide_index=True,height=300)

        # 엑셀 파일로 다운로드 버튼 추가
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='Sheet1')
            writer.book.close()  # Save the workbook

        st.download_button(
            label="**엑셀 파일 다운로드**",
            data=buffer.getvalue(),
            file_name="sorted_departments.xlsx",
            mime="application/vnd.ms-excel"
        )


    with col3:        
        # 부서 목록에서 빠진 부서 찾기
        missing_departments = [dept for dept in department_order if dept not in combined_df[부서_column].values]
        missing_count = len(missing_departments)
        st.header(f"미제출 부서(:red[{missing_count}개소])")
        
        if missing_departments:
            st.write(", ".join(missing_departments))
        else:
            st.write("미제출 부서가 없습니다.")
else:
    with col2: 
        st.markdown("""
                <style>
                .spacer { margin-top: 27px; }
                </style>
                <div class="spacer"></div>
                """, unsafe_allow_html=True)
 
        st.success("**부서에서 취합한 여러 개의 엑셀 파일을 한번에 등록 취합하고, 미제출 부서 확인과 정리된 엑셀 파일을 다운로드 할 수 있습니다.**")
        