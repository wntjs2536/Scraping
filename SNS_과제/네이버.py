"""네이버 영화 리뷰 수집 스크립트."""

import os
import re
import sys
import time
from typing import Dict, List

import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver


BASE_SAVE_DIR = "./output"
CHROME_DRIVER_PATH = "./chromedriver.exe"


def create_save_directory(run_timestamp: str, search_keyword: str) -> str:
    """출력 파일을 저장할 디렉터리를 생성하고 경로를 반환합니다."""
    folder_name = f"{run_timestamp} {search_keyword}"
    save_directory = os.path.join(BASE_SAVE_DIR, folder_name)
    os.makedirs(save_directory, exist_ok=True)
    return save_directory


def open_text_log(save_directory: str, run_timestamp: str, search_keyword: str):
    """로그 파일 핸들을 엽니다."""
    log_path = os.path.join(save_directory, f"{run_timestamp} {search_keyword}.txt")
    return open(log_path, "a", encoding="utf-8")


def fetch_review_total(page_soup: BeautifulSoup) -> int:
    """현재 페이지에서 전체 리뷰 건수를 추출합니다."""
    review_count_text = (
        page_soup.find("div", class_="score_total")
        .find("strong", class_="total")
        .find("em")
        .get_text()
    )
    digits_only = re.findall("\\d", review_count_text)
    return int("".join(digits_only))


def extract_review(page_soup: BeautifulSoup, item_index: int) -> Dict[str, str]:
    """주어진 위치의 리뷰 정보를 추출합니다."""
    base_selector = (
        "body > div > div > div.score_result > ul > li:nth-child("
        f"{item_index})"
    )
    review_score = page_soup.select(f"{base_selector}> div.star_score > em")[0].text
    review_text = page_soup.select(f"#_filtered_ment_{item_index - 1}")[0].text.strip()
    reviewer = page_soup.select(
        f"{base_selector} > div.score_reple > dl > dt > em:nth-child(1) > a > span"
    )[0].text
    review_date = page_soup.select(
        f"{base_selector} > div.score_reple > dl > dt > em:nth-child(2)"
    )[0].text
    sympathy = page_soup.select(
        f"{base_selector} > div.btn_area > a._sympathyButton > strong"
    )[0].text
    not_sympathy = page_soup.select(
        f"{base_selector} > div.btn_area > a._notSympathyButton > strong"
    )[0].text

    return {
        "별점": review_score,
        "리뷰내용": review_text,
        "작성자": reviewer,
        "작성일자": review_date,
        "공감": sympathy,
        "비공감": not_sympathy,
    }


def log_review(log_handle, review: Dict[str, str]) -> None:
    """리뷰 정보를 콘솔과 로그 파일에 출력합니다."""
    for label, value in review.items():
        print(f"{label}: {value}")
        log_handle.write(f"{label}: {value}\n")
    print("")
    log_handle.write("\n")


def navigate_to_reviews(driver: webdriver.Chrome, search_keyword: str) -> BeautifulSoup:
    """네이버 영화 리뷰 페이지로 이동하고 BeautifulSoup 객체를 반환합니다."""
    driver.get("https://movie.naver.com")
    time.sleep(2)

    search_input = driver.find_element_by_id("ipt_tx_srch")
    search_input.send_keys(search_keyword)
    driver.find_element_by_class_name("btn_srch").click()
    driver.find_element_by_xpath('//*[@id="old_content"]/ul[1]/li[2]/a').click()
    time.sleep(2)

    driver.find_element_by_class_name("result_thumb").click()
    driver.find_element_by_class_name("end_sub_tab").click()
    time.sleep(2)

    driver.switch_to.frame("pointAfterListIframe")
    return BeautifulSoup(driver.page_source, "html.parser")


def save_results(
    save_directory: str,
    run_timestamp: str,
    search_keyword: str,
    reviews: List[Dict[str, str]],
) -> None:
    """리뷰 수집 결과를 CSV와 Excel로 저장합니다."""
    data_frame = pd.DataFrame(reviews)
    csv_path = os.path.join(save_directory, f"{run_timestamp} {search_keyword}.csv")
    excel_path = os.path.join(save_directory, f"{run_timestamp} {search_keyword}.xls")
    data_frame.to_csv(csv_path, encoding="utf-8-sig", index=True)
    data_frame.to_excel(excel_path, index=True)


def main():
    run_timestamp = time.strftime("%y-%m-%d_%H-%M-%S")
    driver = webdriver.Chrome(CHROME_DRIVER_PATH)

    print("--------------------------------------------------------------")
    print("연습문제 8-1 :네이버 영화 리뷰 정보 수집하기")
    print("==============================================================")
    print("\n\n")

    search_keyword = str(input("검색명: "))
    desired_review_count = int(input("리뷰건수: "))

    save_directory = create_save_directory(run_timestamp, search_keyword)
    print(f"결과 저장경로: {save_directory}")

    page_soup = navigate_to_reviews(driver, search_keyword)

    total_review_count = fetch_review_total(page_soup)
    log_handle = open_text_log(save_directory, run_timestamp, search_keyword)
    log_file_path = os.path.join(save_directory, f"{run_timestamp} {search_keyword}.txt")
    print(log_file_path)

    if total_review_count < desired_review_count:
        print("--------------------------------------------------------------")
        print(f"리뷰건수 초과 {total_review_count}건만 수집합니다.")
        print("==============================================================")

        log_handle.write("--------------------------------------------------------------\n")
        log_handle.write(f"리뷰건수 초과 {total_review_count}건만 수집합니다.\n")
        log_handle.write("==============================================================\n")
        desired_review_count = total_review_count

    print("")
    log_handle.write("")

    collected_reviews: List[Dict[str, str]] = []
    page_review_position = 0

    for collected_count in range(desired_review_count):
        page_review_position += 1

        print("--------------------------------------------------------------")
        print(
            f"총 {desired_review_count}건 중 {collected_count + 1} 번째 리뷰 데이터를 수집합니다."
        )
        print("==============================================================")
        print("")

        log_handle.write("--------------------------------------------------------------\n")
        log_handle.write(
            f"총 {desired_review_count}건 중 {collected_count + 1} 번째 리뷰 데이터를 수집합니다.\n"
        )
        log_handle.write("==============================================================\n\n")

        review_details = extract_review(page_soup, page_review_position)
        log_review(log_handle, review_details)
        collected_reviews.append(review_details)

        if page_review_position == 10 and collected_count + 1 < desired_review_count:
            driver.find_element_by_link_text("다음").click()
            time.sleep(2)
            page_soup = BeautifulSoup(driver.page_source, "html.parser")
            page_review_position = 0

    log_handle.close()
    save_results(save_directory, run_timestamp, search_keyword, collected_reviews)

    driver.close()
    print("크롤링 종료")
    os.system("pause")


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(f"실행 중 오류가 발생했습니다: {error}")
        sys.exit(1)
