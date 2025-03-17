import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from time import sleep
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select


wb = Workbook()

# 기본 시트 삭제
std = wb['Sheet']  # 기본 시트 이름이 'Sheet'인 경우
wb.remove(std)

url_1 = "https://www.koreabaseball.com/Record/Team/Pitcher/BasicOld.aspx"

driver = webdriver.Chrome()
driver.implicitly_wait(3)

driver.get(url_1)
sleep(3)

# KBO 기록실 홈페이지에서 'KBO 정규시즌' 선택
select_s = Select(driver.find_element(By.ID, 'cphContents_cphContents_cphContents_ddlSeries_ddlSeries'))
select_s.select_by_visible_text('KBO 정규시즌')
sleep(1)



# 연도별 선택 후 데이터 스크래핑 range(2001, 2026)
for y in range(2001, 2025) :

    # 해당 연도 시트 생성, 열제목 추가
    ws = wb.create_sheet('{}년'.format(y))
    ws.append(['순위', '팀명', 'ERA', 'CG', 'SHO', 'QS', 'BSV', 'TBF', 'NP', 'AVG', '2B', '3B', 'SAC', 'SF', 'IBB', 'WP', 'BK'])

    # 연도 옵션 선택자에서 해당 연도 선택
    select_y = Select(driver.find_element(By.ID, 'cphContents_cphContents_cphContents_ddlSeason_ddlSeason'))
    select_y.select_by_value('{}'.format(y))
    sleep(2)

    # 연도 선택한 페이지의 HTML 코드 가져오기
    record_page = driver.page_source
    soup = BeautifulSoup(record_page, 'html.parser')

    # 순위별 데이터 가져오기
    for tr_tag in soup.select('div.record_result tbody tr')[0:] :
        td_tag = tr_tag.select('td')
        row = [
            td_tag[0].get_text(),
            td_tag[1].get_text(),
            td_tag[2].get_text(),
            td_tag[3].get_text(),
            td_tag[4].get_text(),
            td_tag[5].get_text(),
            td_tag[6].get_text(),
            td_tag[7].get_text(),
            td_tag[8].get_text(),
            td_tag[9].get_text(),
            td_tag[10].get_text(),
            td_tag[11].get_text(),
            td_tag[12].get_text(),
            td_tag[13].get_text(),
            td_tag[14].get_text(),
            td_tag[15].get_text(),
            td_tag[16].get_text()
        ]
        ws.append(row)

    # 전체 합계 데이 가져터기
    tf_tag = soup.select('div.record_result tfoot td')
    row_tfood = [
        tf_tag[0].get_text(),
        tf_tag[0].get_text(),
        tf_tag[1].get_text(),
        tf_tag[2].get_text(),
        tf_tag[3].get_text(),
        tf_tag[4].get_text(),
        tf_tag[5].get_text(),
        tf_tag[6].get_text(),
        tf_tag[7].get_text(),
        tf_tag[8].get_text(),
        tf_tag[9].get_text(),
        tf_tag[10].get_text(),
        tf_tag[11].get_text(),
        tf_tag[12].get_text(),
        tf_tag[13].get_text(),
        tf_tag[14].get_text(),
        tf_tag[15].get_text()
    ]
    ws.append(row_tfood)

    sleep(3)


wb.save('투수기록2.xlsx')

sleep(3)
driver.quit()

