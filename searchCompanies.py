import os
import time
import base64
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options

def initialize_webdriver():
    """初始化Chrome WebDriver。"""
    try:
        chrome_options = Options()
        # chrome_options.add_argument('--headless')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')

        chrome_driver_path = 'chromedriver.exe'
        service = Service(chrome_driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
        return driver
    except Exception as e:
        print(f"初始化WebDriver时出错: {e}")
        raise

def read_keywords(file_path):
    """读取公司名-关键词表格。"""
    try:
        df = pd.read_excel(file_path, header=None)
        company_names = df.iloc[:, 0].astype(str)
        keywords1 = df.iloc[:, 1].astype(str)
        keywords2 = df.iloc[:, 2].astype(str)
        return company_names, keywords1, keywords2
    except Exception as e:
        print(f"读取关键词文件时出错: {e}")
        raise

def create_folder(result_folder, company):
    """创建结果文件夹。"""
    try:
        company_folder = os.path.join(result_folder, company)
        os.makedirs(company_folder, exist_ok=True)
        return company_folder
    except Exception as e:
        print(f"创建文件夹时出错: {e}")
        raise

def perform_search(driver, company, keyword1, keyword2):
    """在百度上执行搜索并返回搜索URL。"""
    try:
        search_query = f"{company} {keyword1} {keyword2}"
        driver.get("https://www.baidu.com")
        wd = driver.find_element(By.NAME, 'wd')
        wd.send_keys(search_query)
        su = driver.find_element(By.ID, 'su')
        su.click()
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div#content_left"))
        )
        return driver.current_url
    except Exception as e:
        print(f"执行搜索时出错: {e}")
        raise

def save_search_results(driver, company_folder, page_num, search_query):
    """保存搜索结果为PDF。"""
    try:
        pdf_file_path = os.path.join(company_folder, f'{search_query}_page_{page_num}.pdf')

        current_time = time.strftime("%Y-%m-%d %H:%M:%S")
        current_url = driver.current_url
        # 使用Chrome DevTools协议来生成PDF
        result = driver.execute_cdp_cmd("Page.printToPDF", {
            "printBackground": True,
            "scale": 0.35,
            "displayHeaderFooter": True,
            "headerTemplate": f'''
                <div style="width: 100%; text-align: left; font-size: 20px; padding-left: 10px;">
                    {current_time}
                </div>''',
            "footerTemplate": f'''
                <div style="width: 100%; text-align: left; font-size: 20px; padding-right: 20px; position: absolute; bottom: 0; width: 100%;">
                    {current_url}
                </div>''',
            "marginTop": 0.5,
            "marginBottom": 0.5,
            "marginLeft": 0.5,
            "marginRight": 0.5
        })

        pdf_data = base64.b64decode(result['data'])
        with open(pdf_file_path, 'wb') as file:
            file.write(pdf_data)

        print(f"已在'{company_folder}'文件夹中保存了第{page_num}页的搜索结果PDF：{pdf_file_path}")
    except Exception as e:
        print(f"保存第 {page_num} 页搜索结果时出错: {e}")

def go_to_next_page(driver):
    """点击下一页按钮并等待页面加载。"""
    try:
        next_page_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "下一页 >"))
        )
        next_page_button.click()
        time.sleep(1)
    except Exception as e:
        print(f"翻页时出错: {e}")
        raise

def main():
    try:
        # 初始化 WebDriver
        driver = initialize_webdriver()

        # 读取关键词
        keywordFilePath = "./keywords.xlsx"
        company_names, keywords1, keywords2 = read_keywords(keywordFilePath)

        # 创建result文件夹
        result_folder = 'result'
        os.makedirs(result_folder, exist_ok=True)

        # 遍历每一行
        for company, keyword1, keyword2 in zip(company_names, keywords1, keywords2):
            company_folder = create_folder(result_folder, company)
            search_query = f"{company} {keyword1} {keyword2}"
            perform_search(driver, company, keyword1, keyword2)
            time.sleep(1)
            save_search_results(driver, company_folder, 1, search_query)

            for page_num in range(2, 6):
                try:
                    go_to_next_page(driver)
                    save_search_results(driver, company_folder, page_num, search_query)
                except Exception as e:
                    print(f"无法处理第 {page_num} 页: {e}")
                    break

        # 关闭浏览器
        driver.quit()
        print("所有搜索结果已成功导出。")
    except Exception as e:
        print(f"执行主程序时出错: {e}")
        if 'driver' in locals():
            driver.quit()

if __name__ == "__main__":
    main()
