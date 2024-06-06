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
    """初始化Chrome WebDriver."""
    try:
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # 使Chrome浏览器在无头模式下运行。无头模式允许浏览器在后台运行，没有任何可见的窗口
        chrome_options.add_argument('--disable-gpu')  # 无头模式下，某些图形渲染任务可能会出现问题，禁用GPU可以避免这些问题
        chrome_options.add_argument('--no-sandbox')  # 禁用沙箱模式。沙箱模式提供了一层额外的安全保护，但在某些环境下（如Docker容器中），它可能会导致权限问题，禁用沙箱可以解决这些问题
        chrome_options.add_argument('--disable-dev-shm-usage')  # 禁用/dev/shm的使用。/dev/shm是共享内存文件系统，默认情况下，Chrome在这里存储临时文件。在某些系统（如Docker容器）中，/dev/shm的空间可能不足，禁用它可以防止因空间不足导致的错误。

        chrome_driver_path = 'chromedriver.exe'  # chromedriver.exe的路径，此项目中与源文件在同一个目录下，使用相对路径
        service = Service(chrome_driver_path)  # 创建一个 Service 对象，该对象负责管理 ChromeDriver 的启动和停止
        driver = webdriver.Chrome(service=service, options=chrome_options) # 创建一个 Chrome WebDriver 实例，它可以用于控制 Chrome 浏览器的行为
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

def create_folders(result_folder, company, keyword1, keyword2):
    """创建结果文件夹。"""
    try:
        company_folder = os.path.join(result_folder, company)  # 公司名文件夹
        keyword_folder = os.path.join(company_folder, f"{keyword1}_{keyword2}")  # 公司名文件夹下的关键词文件夹
        os.makedirs(keyword_folder, exist_ok=True)
        return keyword_folder
    except Exception as e:
        print(f"创建文件夹时出错: {e}")
        raise

def perform_search(driver, company, keyword1, keyword2):
    """在百度上执行搜索并返回搜索URL。"""
    try:
        search_query = f"{company} {keyword1} {keyword2}"  # 输入的搜索文本是 [公司名] [关键词1] [关键词2]
        driver.get("https://www.baidu.com")
        wd = driver.find_element(By.NAME, 'wd')  # 找到输入文本框
        wd.send_keys(search_query)  # 输入搜索文本
        su = driver.find_element(By.ID, 'su')   # 找到搜索按钮
        su.click()  # 点击搜索按钮
        WebDriverWait(driver, 10).until(    # 最多等待10s
            EC.presence_of_element_located((By.CSS_SELECTOR, "div#content_left"))
        )
        return driver.current_url
    except Exception as e:
        print(f"执行搜索时出错: {e}")
        raise

def save_search_results(driver, keyword_folder, page_num):
    """保存搜索结果为PDF。"""
    try:
        pdf_file_path = os.path.join(keyword_folder, f'search_result_page_{page_num}.pdf')

        # 使用Chrome DevTools协议来生成PDF
        result = driver.execute_cdp_cmd("Page.printToPDF", {
            "printBackground": True
        })

        pdf_data = base64.b64decode(result['data'])
        with open(pdf_file_path, 'wb') as file:
            file.write(pdf_data)

        print(f"已在'{keyword_folder}'文件夹中保存了第{page_num}页的搜索结果PDF：{pdf_file_path}")
    except Exception as e:
        print(f"保存第 {page_num} 页搜索结果时出错: {e}")

def go_to_next_page(driver):
    """点击下一页按钮并等待页面加载。"""
    try:
        next_page_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, "下一页 >"))
        )
        next_page_button.click()
        time.sleep(2)  # 等待页面加载
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
            keyword_folder = create_folders(result_folder, company, keyword1, keyword2)
            perform_search(driver, company, keyword1, keyword2)
            save_search_results(driver, keyword_folder, 1)

            for page_num in range(2, 6):
                try:
                    go_to_next_page(driver)
                    save_search_results(driver, keyword_folder, page_num)
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
