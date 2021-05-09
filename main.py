# main program
# main.py

import stock_kind as sk
import collect_day.collect_day as cd
import naver_crawling.naver_finance_crawling as nfc
import auto.auto_login as at
import multiple_real_time_finding_listmethod as real_time


def start_collect_stock_kind():     # 종목코드 가져오기
    collect_stock_kind = sk.StockKind()
    return collect_stock_kind.start_collect()


def start_collect_day(codes):           # 일별 데이터 모두 가져오기
    collect_stock_day = cd.DayCollect()
    collect_stock_day.start_get_days_data(codes)


def start_collect_day_update(codes):   # 일별 데이터 갱신 하기
    collect_stock_day = cd.DayCollect()
    collect_stock_day.start_update_days_data(codes)


def start_naver_crawling():         # 네이버 업종별 크롤링
    crawling = nfc.NaverFinanceCrawler()
    crawling.get_upjong()
    while not crawling.upjong_tour():
        crawling.init()
        crawling.upjong_tour()
    print('crawling fin')


def start_get_real_time():      # 실시간 데이터 수집 시작
    real_time.start()


def start_program():
    # start_naver_crawling()          # 네이버 크롤링

    auto_login = at.AutoLogin()     # 자동 로그인
    while not auto_login.login():
        pass

    codes = start_collect_stock_kind()  # 주식 종목 가져오기
    # start_collect_day()               # 모든 종목 일별 데이터 수집(전체)
    start_collect_day_update(codes)     # 모든 종목 일별 데이터 갱신

    # start_get_real_time()             # 실시간 데이터 수집. 네이버 크롤링 선행 필요

start_program()

# start_collect_day()               # 모든 종목 일별 데이터 수집(전체)
# start_collect_day_update(codes)   # 모든 종목 일별 데이터 갱신
