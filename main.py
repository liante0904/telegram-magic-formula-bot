# -*- coding:utf-8 -*- 
import os
import sys
import datetime
from datetime import date, timedelta
import math
from openpyxl import Workbook # 엑셀
from pytz import timezone
# import urlparse
import telegram
import requests
import time
import ssl
import json
import re
import pymysql
import pymysql.cursors
from typing import List
from bs4 import BeautifulSoup
#from urllib.parse import urlparse
import urllib.parse as urlparse
import urllib.request

from requests import get  # to make GET request

# 텔레그램 봇
from telegram import InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Updater, CallbackQueryHandler, CommandHandler, MessageHandler , Filters


# 참고 문서
# https://heodolf.tistory.com/75
# https://heodolf.tistory.com/76
# https://heodolf.tistory.com/77
# https://api.telegram.org/bot1609851580:AAHziXYwvVJqANZhDtg682whClHeaElndZM/getUpdates
# https://minmong.tistory.com/312

# 로직 설명
# 1. Main()-> 각 회사별 함수를 통해 반복 (추후 함수명 일괄 변경 예정)
#   - checkNewArticle -> parse -> downloadFile -> Send 
# 2. 연속키의 경우 현재 .key로 저장
#   - 추후 heroku db로 처리 예정(MySQL)
#   - DB연결이 안되는 경우, Key로 처리할수 있도록 예외처리 반영
# 3. 최초 조회되는 게시판 혹은 Key값이 없는 경우 메세지를 발송하지 않음.
# 4. 테스트와 운영을 구분하여 텔레그램 발송 채널 ID 구분 로직 추가
#   - 어떻게 구분지을지 생각해봐야함
# 5. 메시지 발송 방법 변경 (봇 to 사용자 -> 채널에 발송)

############공용 상수############
# 선택 아이템
SELECT_ITEM = (
    "스크리닝 직접 입력모드",               # 0
    "마법공식",             # 1
    "상상인증권",           # 2
    "하나금융투자",          # 3
    "한양증권",              # 4
    "삼성증권",              # 5
    "교보증권"              # 6
)
# 메시지 발송 ID
CHAT_ID = '-1001431056975' # 운영 채널(증권사 신규 레포트 게시물 알림방)

# BOT_API
BOT_API = "1609851580:AAHziXYwvVJqANZhDtg682whClHeaElndZM"

# DATABASE
CLEARDB_DATABASE_URL = 'mysql://b0464b22432146:290edeca@us-cdbr-east-03.cleardb.com/heroku_31ee6b0421e7ff9?reconnect=true'

# 게시글 갱신 시간
REFRESH_TIME = 600

# pymysql 변수
conn    = ''
cursor  = ''

# 연속키URL
NXT_KEY = ''

# 이모지
EMOJI_FIRE = u'\U0001F525'
EMOJI_PICK = u'\U0001F449'

# JSON API 타입
def MagicFormula_crowling(*args):


    # http://wise.thewm.co.kr/ASP/Screener/Screener1.asp?ud=#tabPaging

    # TARGET_URL = 'http://wise.thewm.co.kr/ASP/Screener/data/Screener_Termtabledata.asp?market=0&industry=G0&size=0&workDT=20210305&termCount=4&currentPage=1&orderKey=P1&orderDirect=D&jsonParam=%5B%7B%22Group%22%3A%22I%22%2C%22SEQ%22%3A%222%22%2C%22MIN_VAL%22%3A%226096%22%2C%22MAX_VAL%22%3A%22200000%22%2C%22Ogb%22%3A%223%22%7D%2C%7B%22Group%22%3A%22P%22%2C%22SEQ%22%3A%221%22%2C%22MIN_VAL%22%3A%2210.00%22%2C%22MAX_VAL%22%3A%22100.00%22%2C%22Ogb%22%3A%221%22%7D%2C%7B%22Group%22%3A%22V%22%2C%22SEQ%22%3A%223%22%2C%22MIN_VAL%22%3A%221.00%22%2C%22MAX_VAL%22%3A%2241.00%22%2C%22Ogb%22%3A%221%22%7D%2C%7B%22Group%22%3A%22S%22%2C%22SEQ%22%3A%221%22%2C%22MIN_VAL%22%3A%22-1635%22%2C%22MAX_VAL%22%3A%22100.00%22%2C%22Ogb%22%3A%223%22%7D%5D'
    # TARGET_URL = 'http://wise.thewm.co.kr/ASP/Screener/data/Screener_Termtabledata.asp?market=0&industry=G0&size=0&workDT=20210305&termCount=3&currentPage=1&orderKey=P1&orderDirect=D&jsonParam=%5B%7B%22Group%22%3A%22I%22%2C%22SEQ%22%3A%222%22%2C%22MIN_VAL%22%3A%22100000.00%22%2C%22MAX_VAL%22%3A%22500000.00%22%2C%22Ogb%22%3A%221%22%7D%2C%7B%22Group%22%3A%22V%22%2C%22SEQ%22%3A%221%22%2C%22MIN_VAL%22%3A%221.00%22%2C%22MAX_VAL%22%3A%2230.00%22%2C%22Ogb%22%3A%221%22%7D%2C%7B%22Group%22%3A%22P%22%2C%22SEQ%22%3A%221%22%2C%22MIN_VAL%22%3A%225.00%22%2C%22MAX_VAL%22%3A%2250.00%22%2C%22Ogb%22%3A%221%22%7D%5D'
    # print('구간')    
    # print(GetCurrentDate\('YYYYMMDD'))


    today = date.today()
    yesterday = date.today() - timedelta(1)

    print(today)
    print(yesterday.strftime('%Y%m%d'))
    # sendDocument()
    # workDT = DATE_SPLIT[0] + DATE_SPLIT[1] + DATE_SPLIT[2]

    TARGET_URL = 'http://wise.thewm.co.kr/ASP/Screener/data/Screener_Termtabledata.asp?market=0&industry=G0&size=0&workDT=' +  yesterday.strftime('%Y%m%d') +'&termCount=3&currentPage=1&orderKey=V1&orderDirect=A&jsonParam=%5B%7B%22Group%22%3A%22V%22%2C%22SEQ%22%3A%221%22%2C%22MIN_VAL%22%3A%22-3.93%22%2C%22MAX_VAL%22%3A%2222.89%22%2C%22Ogb%22%3A%221%22%7D%2C%7B%22Group%22%3A%22P%22%2C%22SEQ%22%3A%221%22%2C%22MIN_VAL%22%3A%225.00%22%2C%22MAX_VAL%22%3A%2240.00%22%2C%22Ogb%22%3A%221%22%7D%2C%7B%22Group%22%3A%22V%22%2C%22SEQ%22%3A%2232%22%2C%22MIN_VAL%22%3A%221.00%22%2C%22MAX_VAL%22%3A%228%22%2C%22Ogb%22%3A%222%22%7D%5D'


    try:
        if args[0] == 0:
            print("0번모드")
            TARGET_URL = str(args[1]).strip()
        elif args[0] == 1:
            print("1번모드")
            TARGET_URL = 'http://wise.thewm.co.kr/ASP/Screener/data/Screener_Termtabledata.asp?market=0&industry=G0&size=0&workDT=' +  yesterday.strftime('%Y%m%d') +'&termCount=3&currentPage=1&orderKey=V1&orderDirect=A&jsonParam=%5B%7B%22Group%22%3A%22V%22%2C%22SEQ%22%3A%221%22%2C%22MIN_VAL%22%3A%22-3.93%22%2C%22MAX_VAL%22%3A%2222.89%22%2C%22Ogb%22%3A%221%22%7D%2C%7B%22Group%22%3A%22P%22%2C%22SEQ%22%3A%221%22%2C%22MIN_VAL%22%3A%225.00%22%2C%22MAX_VAL%22%3A%2240.00%22%2C%22Ogb%22%3A%221%22%7D%2C%7B%22Group%22%3A%22V%22%2C%22SEQ%22%3A%2232%22%2C%22MIN_VAL%22%3A%221.00%22%2C%22MAX_VAL%22%3A%228%22%2C%22Ogb%22%3A%222%22%7D%5D'
    except IndexError:
        TARGET_URL = 'http://wise.thewm.co.kr/ASP/Screener/data/Screener_Termtabledata.asp?market=0&industry=G0&size=0&workDT=' +  yesterday.strftime('%Y%m%d') +'&termCount=3&currentPage=1&orderKey=V1&orderDirect=A&jsonParam=%5B%7B%22Group%22%3A%22V%22%2C%22SEQ%22%3A%221%22%2C%22MIN_VAL%22%3A%22-3.93%22%2C%22MAX_VAL%22%3A%2222.89%22%2C%22Ogb%22%3A%221%22%7D%2C%7B%22Group%22%3A%22P%22%2C%22SEQ%22%3A%221%22%2C%22MIN_VAL%22%3A%225.00%22%2C%22MAX_VAL%22%3A%2240.00%22%2C%22Ogb%22%3A%221%22%7D%2C%7B%22Group%22%3A%22V%22%2C%22SEQ%22%3A%2232%22%2C%22MIN_VAL%22%3A%221.00%22%2C%22MAX_VAL%22%3A%228%22%2C%22Ogb%22%3A%222%22%7D%5D'
    

    request = urllib.request.Request(TARGET_URL)
    #검색 요청 및 처리
    response = urllib.request.urlopen(request)
    rescode = response.getcode()
    if rescode != 200 :return print("네이버 뉴스 접속이 원활하지 않습니다 ")

    CMP_PAGE_CNT = 10
    jres = json.loads(response.read().decode('utf-8'))
    #print(jres)
    TOTAL_CMP_CNT = jres['sAllCnt']
    TOTAL_PAGE_CNT = math.ceil(TOTAL_CMP_CNT / CMP_PAGE_CNT) # 페이지 수 이므로 정수가 아닌 경우 +1
    
    print('한 페이지에 회사 수는', CMP_PAGE_CNT , "건 입니다.")
    print('조건에 부합하는 회사 수는 ',TOTAL_CMP_CNT, "건 입니다.")
    print(TOTAL_PAGE_CNT)
    print("VAL 값은 우측 상단의 값임")
    print("반복코드는 나중에")
    
    file = open('hello.txt', 'w')    # hello.txt 파일을 쓰기 모드(w)로 열기. 파일 객체 반환
    
    
    NAVER_URL= 'https://finance.naver.com/item/main.nhn?code='
    jres = jres['resultList']
    for idx in range(1, TOTAL_PAGE_CNT+1):
        paging = 'currentPage='
        paging += str(idx)
        
        # print(TARGET_URL)
        request = urllib.request.Request(TARGET_URL)
        #검색 요청 및 처리
        response = urllib.request.urlopen(request)
        rescode = response.getcode()
        if rescode != 200 :return print("네이버 뉴스 접속이 원활하지 않습니다 ")
        jres = json.loads(response.read().decode('utf-8'))
        jres = jres['resultList']
        # print(idx)
        
        for r in jres:
            write = ''
            write += NAVER_URL + r['CMP_CD'] + '\t' +'종목명:' + r['CMP_NM_KOR'] + '\n'
            write += fnguide_parse(r['CMP_CD']) + '\n'
            print(write)
            file.write(write)      # 파일에 문자열 저장
            
            
        
        TARGET_URL = TARGET_URL.replace('currentPage='+ str(idx), 'currentPage='+ str(idx+1)  )
        
    file.close()                     # 파일 객체 닫기

    sendDocument()
    sendText('/start 를 눌러 시작해보세요.')

    return True

def NAVERNews_downloadFile(LIST_ARTICLE_URL, LIST_ATTACT_FILE_NAME):
    global ATTACH_FILE_NAME
    ATTACH_FILE_NAME = LIST_ATTACT_FILE_NAME #BeautifulSoup(webpage.content, "html.parser").select_one('#contents > div > div.bbs_a_view > dl.b_bottom > dd > em:nth-child(1)> a').text.strip()
    
    DownloadFile(URL = LIST_ARTICLE_URL, FILE_NAME = ATTACH_FILE_NAME)
    time.sleep(5) # 모바일 알림을 받기 위해 8초 텀을 둠(loop 호출시)
    return True


def sendPhoto(ARTICLE_URL): # 파일의 경우 전역변수로 처리 (downloadFile 함수)
    print('sendPhoto()')

    #생성한 텔레그램 봇 정보 assign (@ebest_noti_bot)
    my_token_key = '1372612160:AAHVyndGDmb1N2yEgvlZ_DmUgShqk2F0d4w'
    bot = telegram.Bot(token = my_token_key)

    bot.sendPhoto(chat_id = GetSendChatId(), photo = ARTICLE_URL)
    time.sleep(8) # 모바일 알림을 받기 위해 8초 텀을 둠(loop 호출시)
    return True

def sendText(sendMessageText): # 가공없이 텍스트를 발송합니다.
    global CHAT_ID

    print('sendText()')

    #생성한 텔레그램 봇 정보 assign (@ebest_noti_bot)
    my_token_key = '1372612160:AAHVyndGDmb1N2yEgvlZ_DmUgShqk2F0d4w'
    #생성한 텔레그램 봇 정보 assign (@ssh_stock_info_noti_bot)
    my_token_key = '1609851580:AAHziXYwvVJqANZhDtg682whClHeaElndZM'
    bot = telegram.Bot(token = my_token_key)

    # if SEC_FIRM_ORDER == 998:
    #     if  ARTICLE_BOARD_ORDER == 0 : 
    #         CHAT_ID = '-1001436418974' # 네이버 실시간 속보 뉴스 채널
    #     else:
    #         CHAT_ID = '-1001150510299' # 네이버 많이본 뉴스 채널
    # elif SEC_FIRM_ORDER == 997:
    #         CHAT_ID = '-1001472616534' # 아이투자
    # else:
    #     CHAT_ID = '-1001431056975' # 운영 채널(증권사 신규 레포트 게시물 알림방)

    bot.sendMessage(chat_id = CHAT_ID(), text = sendMessageText, disable_web_page_preview = True, parse_mode = "Markdown")
    
    time.sleep(8) # 모바일 알림을 받기 위해 8초 텀을 둠(loop 호출시)

def sendDocument(): # 가공없이 텍스트를 발송합니다.
    global CHAT_ID

    print('sendDocument()')

    #생성한 텔레그램 봇 정보 assign (@ebest_noti_bot)
    my_token_key = '1372612160:AAHVyndGDmb1N2yEgvlZ_DmUgShqk2F0d4w'
    #생성한 텔레그램 봇 정보 assign (@ssh_stock_info_noti_bot)
    my_token_key = '1609851580:AAHziXYwvVJqANZhDtg682whClHeaElndZM'
    bot = telegram.Bot(token = my_token_key)

    bot.sendDocument(chat_id = CHAT_ID, document =  open('hello.txt', 'rb'))

    # bot.sendMessage(chat_id = CHAT_ID, text = sendMessageText, disable_web_page_preview = True, parse_mode = "Markdown")
    

# URL에 파일명을 사용할때 한글이 포함된 경우 인코딩처리 로직 추가 
def DownloadFile(URL, FILE_NAME):
    global ATTACH_FILE_NAME
    print("DownloadFile()")

    if SEC_FIRM_ORDER == 6: # 교보증권 예외 로직
        # 로직 사유 : 레포트 첨부파일명에 한글이 포함된 경우 URL처리가 되어 있지 않음
        CONVERT_URL = URL 
        for c in URL: # URL내 한글이 있는 경우 인코딩 처리(URL에 파일명을 이용하여 조합함)
            # 코드셋 기준 파이썬:UTF-8 . 교보증권:EUC-KR
            # 1. 주소에서 한글 문자를 판별
            # 2. 해당 문자를 EUC-KR로 변환후 URL 인코딩
            print("##",c , "##", ord('가') <= ord(c) <= ord('힣') )
            if ord('가') <= ord(c) <= ord('힣'): 
                c_encode = c.encode('euc-kr')
                CONVERT_URL = CONVERT_URL.replace(c, urlparse.quote(c_encode) )
                print(CONVERT_URL)

        if URL != CONVERT_URL: 
            print("기존 URL에 한글이 포함되어 있어 인코딩처리함")
            print("CONVERT_URL", CONVERT_URL)
            URL = CONVERT_URL

    ATTACH_FILE_NAME = re.sub('[\/:*?"<>|]','',FILE_NAME) # 저장할 파일명 : 파일명으로 사용할수 없는 문자 삭제 변환
    print('convert URL:',URL)
    print('convert ATTACH_FILE_NAME:',ATTACH_FILE_NAME)
    with open(ATTACH_FILE_NAME, "wb")as file:  # open in binary mode
        response = get(URL, verify=False)     # get request
        file.write(response.content) # write to file
        
    return True

def GetSendMessageText(INDEX, ARTICLE_BOARD_NAME , ARTICLE_TITLE , ARTICLE_URL):

    # 실제 전송할 메시지 작성
    sendMessageText = ''
    # 발신 게시판 종류
    if INDEX == 1:
        sendMessageText += GetSendMessageTitle(ARTICLE_TITLE) + "\n"
    # 게시글 제목(굵게)
    sendMessageText += "**" + ARTICLE_TITLE + "**" + "\n"
    # 원문 링크
    sendMessageText += EMOJI_PICK  + "[원문링크(클릭)]" + "("+ ARTICLE_URL + ")"
    sendMessageText += "\n" + "\n"

    return sendMessageText

def GetSendMessageTitle(ARTICLE_TITLE):

    SendMessageTitle = ''
    if SEC_FIRM_ORDER == 999:
        msgFirmName = "매매동향"
        ARTICLE_BOARD_NAME = ''
        if  "최종치" in ARTICLE_TITLE:
            print('sedaily의 매매동향 최종치 집계 데이터는 메시지 발송을 하지 않습니다.') # 장마감 최종치는 발송 안함
            return 
    elif SEC_FIRM_ORDER == 998:
        msgFirmName = "네이버 - "
        if  ARTICLE_BOARD_ORDER == 0 :
            ARTICLE_BOARD_NAME = "실시간 뉴스 속보"
        else:
            ARTICLE_BOARD_NAME = "가장 많이 본 뉴스"
    elif SEC_FIRM_ORDER == 997:
        msgFirmName = "아이투자 - "
    else:
        msgFirmName = " "
        # msgFirmName = FIRM_NAME[SEC_FIRM_ORDER] + " - "
        # if SEC_FIRM_ORDER != 6: 
        #     ARTICLE_BOARD_NAME = BOARD_NAME[SEC_FIRM_ORDER][ARTICLE_BOARD_ORDER]

    SendMessageTitle += EMOJI_FIRE + msgFirmName + ARTICLE_BOARD_NAME + EMOJI_FIRE + "\n"
    
    return SendMessageTitle


def GetSendChatId():
    SendMessageChatId = 0
    if SEC_FIRM_ORDER == 998:
        if  ARTICLE_BOARD_ORDER == 0 : 
            SendMessageChatId = '-1001436418974' # 네이버 실시간 속보 뉴스 채널
        else:
            SendMessageChatId = '-1001150510299' # 네이버 많이본 뉴스 채널
    elif SEC_FIRM_ORDER == 997:
            SendMessageChatId = '-1001472616534' # 아이투자
    else:
        SendMessageChatId = '-1001431056975' # 운영 채널(증권사 신규 레포트 게시물 알림방)
    
    return SendMessageChatId





def fnguide_parse(*args):
    global NXT_KEY
    global LIST_ARTICLE_TITLE

    pattern = ''
    CODE = ''
    for pattern in args:
        if len(pattern) > 0 :
            CODE =  pattern

    TARGET_URL = 'http://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?MenuYn=Y&gicode=A'
    TARGET_URL += CODE
    # 005930
    webpage = requests.get(TARGET_URL, verify=False)

    # HTML parse
    soup = BeautifulSoup(webpage.content, "html.parser")
    data_cmp_nm = soup.select_one('#giName').text
    data_cmp_code = soup.select_one('#compBody > div.section.ul_corpinfo > div.corp_group1 > h2').text
    data_stxt1 = soup.select_one('#compBody > div.section.ul_corpinfo > div.corp_group1 > p > span.stxt.stxt1').text
    data_stxt2 = soup.select_one('#strMarketTxt').text
    data_Per = soup.select_one('#corp_group2 > dl:nth-child(1) > dd').text
    data_fwdPer = soup.select_one('#corp_group2 > dl:nth-child(2) > dd').text
    data_dividendYield = soup.select_one('#corp_group2 > dl:nth-child(5) > dd').text
    data_cmp_info = soup.select_one('#bizSummaryContent').text
    #data_ROE = soup.select_one('#svdMainGrid10D > table > tbody > tr:nth-child(7) > td:nth-child(2)')#.text
    r = ''
    r += TARGET_URL + '\n'
    r += '==============================================================' + '\n'
    r += '종목명: ' + data_cmp_nm                                + '\n'
    r += '종목코드: ' + data_cmp_code                            + '\n'
    r += '업종: ' + data_stxt1                                      + '\t' + data_stxt2 + '\n'
    r += 'per(FY0): ' + data_Per                                      + '\n'
    r += '12m fwd per: ' + data_fwdPer                           + '\n'
    r += '시가배당 수익률 '  + data_dividendYield                + '\n'
    r += '기업개요:' + data_cmp_info                 + '\n'
    r += '==============================================================' + '\n'
    #print('ROE', data_ROE)
    
    
    return r

def excel_write_title(*args):
    global write_wb
    global write_ws
    
    pattern = ''
    CODE = ''
    for pattern in args:
        if len(pattern) ==  0 or pattern ==  NoneType :
                # 엑셀파일 쓰기
                write_wb = Workbook()
                # Sheet1에다 입력
                write_ws = write_wb.active
                # 타이틀
                write_ws.cell(1, 1, '링크')
                write_ws.cell(1, 2, '종목명')
                write_ws.cell(1, 3, '종목코드')
                write_ws.cell(1, 4, 'PER')
                write_ws.cell(1, 5, 'fwd-PER')
                write_ws.cell(1, 6, '시가배당')


    # 엑셀파일 쓰기
    # write_wb = Workbook()

    # 이름이 있는 시트를 생성
    # write_ws = write_wb.create_sheet('생성시트')

    # # Sheet1에다 입력
    # write_ws = write_wb.active
    # # 타이틀
    # write_ws['A1'] = '숫자'
    # write_ws['B1'] = '종목명'

    #행 단위로 추가
    write_ws.append([1,2,3])

    #셀 단위로 추가
    write_ws.cell(5, 5, '5행5열')
    write_wb.save("숫자.xlsx")

    # 출처 https://myjamong.tistory.com/51


# 시간 및 날짜는 모두 한국 시간 (timezone('Asia/Seoul')) 으로 합니다.
def GetCurrentDate(*args):
    pattern = ''
    for pattern in args:
        print('pattern 입력값',pattern)
    
    time_now = str(datetime.datetime.now(timezone('Asia/Seoul')))[:19] # 밀리세컨즈 제거

    DATE = time_now[:10].strip()
    DATE_SPLIT = DATE.split("-")

    if pattern == '':
        DATE = time_now[:10].strip()
    elif pattern == 'YY' or pattern == 'yy':
        DATE = DATE_SPLIT[0][2:]
    elif pattern == 'YYYY' or pattern == 'yyyy':
        DATE = DATE_SPLIT[0]
    elif pattern == 'MM' or pattern == 'mm':
        DATE = DATE_SPLIT[1]
    elif pattern == 'DD' or pattern == 'dd':
        DATE = DATE_SPLIT[2]
    elif pattern == 'YYYY/HH/MM' or pattern == 'yyyy/hh/mm':
        DATE = DATE_SPLIT[0] + "/" + DATE_SPLIT[1] + "/" + DATE_SPLIT[2]
    elif pattern == 'YYYY-HH-MM' or pattern == 'yyyy-hh-mm':
        DATE = time_now[:10].strip()
    elif pattern == 'YY-HH-MM' or pattern == 'yy-hh-mm':
        DATE = time_now[2:10].strip()
    elif pattern == 'YYYYMMDD' or pattern == 'yyyymmdd':
        DATE = DATE_SPLIT[0] + DATE_SPLIT[1] + DATE_SPLIT[2]
    else:
        DATE = time_now[:10].strip()

    return DATE

def main():
    global CHAT_ID
    data_selected = 0
    print('########Program Start Run########')


    BOT_TOKEN='672768316:AAHXpYmnMzGp_eH0i-juikUFU6q9y78CBhA'
    BOT_TOKEN= "1609851580:AAHziXYwvVJqANZhDtg682whClHeaElndZM"
    
    bot = telegram.Bot( token=BOT_TOKEN )
    try:
        chat_id = bot.getUpdates()[-1].message.chat.id
        CHAT_ID = bot.getUpdates()[-1].message.chat.id
        print(chat_id)
    except IndexError:
        chat_id = '183792411'
        CHAT_ID = '183792411'
    
    bot.sendMessage(chat_id=CHAT_ID, text='/start 를 눌러 시작해보세요 ')

    updater = Updater( token=BOT_TOKEN, use_context=True )
    # 버튼 UI dispatcher
    dispatcher = updater.dispatcher

    def start(update, context):
        task_buttons =  [
            [ InlineKeyboardButton( '0. 스크리닝 직접 입력모드', callback_data=0 ) ],
            [ InlineKeyboardButton( '1.마법공식 종목받기', callback_data=1 ), InlineKeyboardButton( '2.준비중', callback_data=2 ) ], 
            [ InlineKeyboardButton( '3.준비중', callback_data=3 ) ] 
        ]
        
        reply_markup = InlineKeyboardMarkup( task_buttons )
        
        context.bot.send_message(
            chat_id=update.message.chat_id
            , text='작업을 선택해주세요.'
            , reply_markup=reply_markup
        )
    

    start_handler = CommandHandler('start', start)
    
    dispatcher.add_handler(start_handler)


    def callback_get(update, context):
        print("callback")
        data_selected = int(update.callback_query.data)
        print(data_selected)
        if data_selected == 0:
            # 스크리닝 공식 받기
            context.bot.edit_message_text(text="{}이(가) 선택되었습니다".format(SELECT_ITEM[data_selected]),
                                        chat_id=update.callback_query.message.chat_id,
                                        message_id=update.callback_query.message.message_id)
            bot.sendMessage(chat_id=chat_id, text="가이드 링크 : " + 'https://www.notion.so/shinseunghoon/URL-9b91ddd9b409479ca9a0276d0c5a69be' + '\n' + '\n'+ '스크리닝 링크 : '+ 'http://wise.thewm.co.kr/ASP/Screener/Screener1.asp?ud=#tabPaging')     
            bot.sendMessage(chat_id=chat_id, text="가이드를 참조하여 스크리닝 URL을 입력하세요.")    

        elif data_selected == 1:
            context.bot.edit_message_text(text="{}이(가) 선택되었습니다".format(SELECT_ITEM[data_selected]),
                                        chat_id=update.callback_query.message.chat_id,
                                        message_id=update.callback_query.message.message_id)
            MagicFormula_crowling(1)
        else:
            return 
            print("준비중")
            context.bot.edit_message_text(text="{}이(가) 선택되었습니다".format(SELECT_ITEM[0]),
                                        chat_id=update.callback_query.message.chat_id,
                                        message_id=update.callback_query.message.message_id)
        
    updater.dispatcher.add_handler(CallbackQueryHandler(callback_get))


    def get_screening_url(update, context):
        if data_selected == 0:
            inputURL = update.message.text
            # update.message.reply_text("가이드를 참조하여 스크리닝 URL을 입력하세요.")
            # update.message.reply_text("가이드 링크 : " + 'https://www.notion.so/shinseunghoon/URL-9b91ddd9b409479ca9a0276d0c5a69be')
            # URL 형태가 아닌 경우 다시 입력을 받을 수 있는지 여부 확인
            if 'http://wise.thewm.co.kr/ASP/Screener/data/Screener_Termtabledata.asp' not in inputURL:
                bot.sendMessage(chat_id=chat_id, text="입력 값이 올바르지 않습니다. 다시 입력하세요.")
            bot.sendMessage(chat_id=chat_id, text="입력 받은 조건으로 집계를 시작합니다. ")
            URL = update.message.text
            MagicFormula_crowling(0, URL)

    if data_selected == 0:
        message_handler = MessageHandler(Filters.text, get_screening_url)
        updater.dispatcher.add_handler(message_handler)

    updater.start_polling()
    updater.idle()

if __name__ == "__main__":
	main()
