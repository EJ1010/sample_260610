# 스트림릿 패키지 로드
import streamlit as st

# 제목 출력
st.title('안녕하세요 제목입니다.')

# 마크다운
st.markdown("""
# 이 글은 마크다운으로 작성됨
            1. 항목
            - **중요한 내용**
            """)

# 사이드바
menu = st.sidebar.selectbox(
    "메뉴 선택",
 ["홈", "회원관리", "설정"]
)

# 탭
## st.tabs(['메뉴1', '메뉴2', '메뉴3'])
tab1, tab2, tab3 = st.tabs(["메뉴1", "메뉴2", "메뉴3"])

with tab1:
    st.write("첫 번째 탭입니다.")
    with st.expander("자세히보기") : ## 마우스를 클릭해야지만 보여지는 페이지
        st.write("숨겨진 내용")

with tab2:
    st.write("두 번째 탭입니다.")
    a = st.selectbox("기능", ["복사", "붙여넣기", "잘라내기"])
    st.write(f"선택한 기능은 {a}입니다.")

with tab3:
    st.write("세 번째 탭입니다.")
    if st.button("버튼누르기"):
        st.write("버튼이 클릭되었습니다.")
        st.success("성공적으로 버튼이 클릭되었습니다.")



### 실습했던 코드 ###

import streamlit as st
from openpyxl import *

Mem_info = load_workbook('회원정보.xlsx')
info_sheet = Mem_info.active

# 전역변수(streamlit)
st.session_state.count = 0
#                변수명

# 제목/헤더
st.title('20Hz : 연애시뮬레이션')
st.header('20Hz : 연애시뮬레이션')
st.subheader('KDT 파이썬미션')

# 마크다운
st.markdown(
    '''
    # 입력해야 할 정보
    ## 1. 이름
    ## 2. 성별
    '''
)

# 입력
name = st.text_input('이름: ')
st.write(f'당신의 이름은 {name}이 맞나요?')

# 버튼
if st.button('네 맞아요!'):
    st.success(f'환영합니다 {name}님!')

# 사이드바
st.sidebar.selectbox("메뉴",['게임시작','저장','끝내기'])

# 탭구성
tab1, tab2 = st.tabs(['메인','회원정보입력'])

with tab1:
    # 첫번째 탭 구성
    st.table(list(info_sheet.values))

with tab2:
    # 두번째 탭 구성
    회원정보 = []
    with st.form('회원정보'):
        c1, c2 = st.columns(2)
        id = c1.text_input('아이디')
        name_ = c2.text_input('성명')

        c3, c4 = st.columns(2)
        date = c3.date_input('날짜')
        gen = c4.radio('성별',['남','여'])
        회원정보 = [id, name_, date, gen]

        if st.form_submit_button('저장'): #클릭시
            info_sheet.append(회원정보) #엑셀에 담기
            Mem_info.save('회원정보.xlsx')
            st.success('저장 완료!')

if st.button('새로고침'):
    st.rerun()