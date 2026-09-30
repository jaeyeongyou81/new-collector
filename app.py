import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
import urllib.parse
from datetime import datetime, timedelta
import re
from email.utils import parsedate_to_datetime

# 웹 앱 페이지 설정 (넓은 화면 사용)
st.set_page_config(page_title="건축사업본부 주간 동향 뉴스 수집기", layout="wide")

# 웹 앱 화면 구성
st.title("🏗️ 건축사업본부 주간 동향 뉴스 수집기 (Pro)")
st.markdown("원하는 키워드와 기간을 설정하여 최신 동향 뉴스를 수집하고 보고서용 엑셀로 활용하세요.")

# 사이드바 설정: 검색 키워드 및 기간 설정
st.sidebar.header("🔍 수집 옵션 설정")

# 1. 기간 필터 설정
period_options = {
    "최근 3일": 3,
    "최근 7일 (1주일)": 7,
    "최근 15일": 15,
    "최근 30일 (1개월)": 30,
    "제한 없음 (전체)": 9999
}
selected_period_name = st.sidebar.selectbox("⏱️ 뉴스 수집 기간 선택", list(period_options.keys()), index=1)
days_limit = period_options[selected_period_name]
cutoff_date = datetime.now() - timedelta(days=days_limit)

# 2. 키워드 입력 설정
default_keywords = "건설현장 AI 활용, 스마트 건설 AI, 삼성물산 건설 AI 신기술, 현대건설 스마트건설, 포스코이앤씨 스마트 건설, 현대산업개발 스마트건설, 롯데건설 스마트건설, 현대엔지니어링 스마트건설, 현대산업개발 스마트건설, 모듈러, 건설현장 점검, 건설현장 노란봉투법, 노조 파업, 아파트 브랜드, 건축법 개정, 주택법 개정, 국토교통부 건설 정책"
user_keywords_input = st.sidebar.text_area("수집할 키워드를 쉼표(,)로 구분하여 입력하세요.", value=default_keywords, height=150)

keywords = [kw.strip() for kw in user_keywords_input.split(",") if kw.strip()]
count_per_keyword = st.sidebar.slider("키워드당 최대 수집 개수", min_value=3, max_value=15, value=5)

# 수집 실행 버튼
if st.button("🚀 최신 뉴스 수집 및 표 생성 시작"):
    if not keywords:
        st.warning("최소 하나 이상의 검색 키워드를 입력해주세요.")
    else:
        news_records = []
        raw_data_for_excel = []
        seen_titles = set()
        
        with st.spinner(f"선택하신 '{selected_period_name}' 기준 내 최신 뉴스를 수집하고 있습니다..."):
            for kw in keywords:
                encoded_kw = urllib.parse.quote(kw)
                url = f"https://news.google.com/rss/search?q={encoded_kw}&hl=ko&gl=KR&ceid=KR:ko"
                
                response = requests.get(url)
                soup = BeautifulSoup(response.content, 'xml')
                items = soup.find_all('item')
                
                count = 0
                for item in items:
                    if count >= count_per_keyword:
                        break
                        
                    title = item.title.text if item.title else "제목 없음"
                    raw_link = item.link.text if item.link else ""
                    pub_date_str = item.pubDate.text if item.pubDate else ""
                    
                    # 발행일 파싱 및 기간 필터링 적용
                    news_date = None
                    if pub_date_str:
                        try:
                            news_date = parsedate_to_datetime(pub_date_str).replace(tzinfo=None)
                        except Exception:
                            news_date = datetime.now()
                    else:
                        news_date = datetime.now()
                        
                    # 설정한 기간보다 오래된 기사는 스킵
                    if news_date < cutoff_date:
                        continue
                        
                    # 제목과 언론사 분리
                    clean_title = title.split(" - ")[0]
                    media = title.split(" - ")[1] if " - " in title else "알 수 없음"
                    
                    # 중복 필터링 정규화
                    normalized_title = re.sub(r'[^가-힣a-zA-Z0-9]', '', clean_title)[:15]
                    if normalized_title in seen_titles:
                        continue
                    seen_titles.add(normalized_title)
                    
                    formatted_date = news_date.strftime("%Y-%m-%d")
                    
                    # 화면 표시용 데이터 (제목과 순수 링크를 분리)
                    news_records.append({
                        "검색 키워드": kw,
                        "뉴스 제목": clean_title,
                        "언론사": media,
                        "발행일": formatted_date,
                        "원문 링크": raw_link
                    })
                    
                    # 엑셀 저장용 (HYPERLINK 수식 적용)
                    safe_link = raw_link.replace('"', '%22')
                    excel_hyperlink = f'=HYPERLINK("{safe_link}", "기사보기")' if safe_link else "링크 없음"
                    
                    raw_data_for_excel.append({
                        "검색 키워드": kw,
                        "뉴스 제목": clean_title,
                        "언론사": media,
                        "원본 링크": excel_hyperlink,
                        "발행일": formatted_date
                    })
                    count += 1
                        
            df_display = pd.DataFrame(news_records)
            df_excel = pd.DataFrame(raw_data_for_excel)
            
            file_name = f"건축사업본부_주간보고_{datetime.now().strftime('%m%d')}.xlsx"
            df_excel.to_excel(file_name, index=False)
            
        if len(df_display) == 0:
            st.warning(f"⚠️ 설정하신 기간({selected_period_name}) 내에 수집된 뉴스가 없습니다. 기간을 넓혀서 다시 시도해보세요.")
        else:
            st.success(f"✨ 총 {len(df_display)}건의 최신 대표 기사 수집 및 중복 정제가 완료되었습니다!")
            
            # 표 디자인 출력 (LinkColumn으로 안전한 링크 이동 구현)
            st.markdown("### 📋 수집된 최신 동향 보고서 미리보기")
            st.markdown("우측의 **[🔗 기사 보기]**를 클릭하시면 원본 뉴스 페이지로 바로 이동합니다.")
            
            st.dataframe(
                df_display,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "검색 키워드": st.column_config.TextColumn("검색 키워드", width="medium"),
                    "뉴스 제목": st.column_config.TextColumn("뉴스 제목", width="large"),
                    "언론사": st.column_config.TextColumn("언론사", width="small"),
                    "발행일": st.column_config.TextColumn("발행일", width="small"),
                    "원문 링크": st.column_config.LinkColumn("원문 링크", display_text="🔗 기사 보기", width="small")
                }
            )
            
            # 엑셀 다운로드 버튼 생성
            st.markdown("---")
            with open(file_name, "rb") as f:
                st.download_button(
                    label="📥 보고서용 엑셀 파일 다운로드 (클릭형 링크 포함)",
                    data=f,
                    file_name=file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )