import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import csv
from datetime import datetime

gamename = "丝之歌"  # 要爬取的游戏名称
country = "中国"  # 地区
app_id = "1030300"  # 示例游戏ID
max_reviews = 20000 # 设置要爬取的最大评论数


def setup_driver():
    """设置并返回Edge WebDriver"""
    edge_options = Options()
    # edge_options.add_argument("--headless")
    edge_options.add_argument("--window-size=1920x1080")
    edge_options.add_argument("--disable-gpu")
    edge_options.add_argument("--disable-extensions")

    # 禁用图片加载提高速度
    prefs = {"profile.managed_default_content_settings.images": 2}
    edge_options.add_experimental_option("prefs", prefs)

    # 指定msedgedriver的完整路径
    edge_driver_path = r".\edgedriver_win64\msedgedriver.exe"

    # 创建Edge服务
    service = Service(executable_path=edge_driver_path)

    # 创建Edge浏览器实例
    driver = webdriver.Edge(service=service, options=edge_options)
    return driver


def extract_hours(hours_text):
    """安全提取游戏时长"""
    try:
        # 匹配数字部分（包括小数）
        match = re.search(r'(\d+\.?\d*)', hours_text)
        if match:
            return "{:.1f}".format(float(match.group(1)))
        return "0.0"  # 如果没有匹配到数字，返回0
    except:
        return "0.0"  # 出现任何错误返回0


def format_date(date_text):
    """将Steam日期格式转换为YYYY-MM-DD"""
    # 匹配带年份的格式
    full_date_match = re.search(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日', date_text)
    if full_date_match:
        year, month, day = full_date_match.groups()
        return f"{year}-{month.zfill(2)}-{day.zfill(2)}"

    # 匹配只有月日的格式
    month_day_match = re.search(r'(\d{1,2})\s*月\s*(\d{1,2})\s*日', date_text)
    if month_day_match:
        month, day = month_day_match.groups()
        return f"2025-{month.zfill(2)}-{day.zfill(2)}"  # 手动添加2025年

    return date_text.replace("发布于：", "")


def extract_username(review_element):
    """精确提取用户名"""
    try:
        # 方案1：直接定位最内层的<a>标签
        username = review_element.find_element(By.CSS_SELECTOR,
                                               ".apphub_CardContentAuthorName a:last-child").text.strip()
        if username:  # 如果成功获取到用户名
            return username

        # 方案2：作为备选方案
        username = review_element.find_element(By.CSS_SELECTOR, ".apphub_CardContentAuthorName").text.strip()
        return username if username else "匿名用户"

    except Exception as e:
        print(f"提取用户名时出错: {e}")
        return "匿名用户"


def extract_review_data(review):
    """提取并格式化评论数据以匹配Excel表格"""
    try:
        # 用户名（使用优化后的选择器）
        username = extract_username(review)

        # 评论内容（移除日期行）
        content_element = WebDriverWait(review, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".apphub_CardTextContent"))
        )
        content = "\n".join([line.strip() for line in content_element.text.split("\n")
                             if not line.startswith("发布于：")]).strip()

        # 评分（推荐/不推荐）
        recommended = WebDriverWait(review, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".title"))
        ).text.strip()

        # 日期（格式化YYYY-MM-DD）
        date_text = WebDriverWait(review, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".date_posted"))
        ).text
        date = format_date(date_text)  # 调用日期格式化函数

        # 游戏时长（使用安全提取方法）
        hours_text = WebDriverWait(review, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".hours"))
        ).text
        hours = extract_hours(hours_text)  # 调用安全提取函数

        return {
            "游戏名称": gamename,
            "国家": country,
            "发表时间": date,
            "作者": username,
            "游戏时长": hours,
            "内容": content,
            "是否推荐/好评": recommended,
        }

    except Exception as e:
        print(f"提取评论时出错: {e}")
        return None


def scrape_steam_reviews(app_id, max_reviews=max_reviews, output_file='steam_reviews.csv'):
    """
    爬取Steam游戏评论（实时保存版本）
    """
    driver = setup_driver()
    url = f"https://steamcommunity.com/app/{app_id}/negativereviews/?browsefilter=mostrecent&snr=1_5_100010_&filterLanguage=schinese"

    try:
        print(f"正在访问Steam差评页面: {url}")
        driver.get(url)

        # 等待页面主要元素加载完成
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".apphub_Card"))
        )

        # 打开CSV文件并写入表头（如果不存在）
        fieldnames = [
            "游戏名称", "国家", "发表时间", "作者",
            "游戏时长", "内容", "是否推荐/好评"
        ]
        with open(output_file, mode='w', newline='', encoding='utf-8-sig') as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()

        print("正在加载更多评论...")
        loaded_reviews = 0
        scroll_attempts = 0
        max_scroll_attempts = 30  # 滚动上限
        last_saved_count = 0

        while loaded_reviews < max_reviews and scroll_attempts < max_scroll_attempts:
            reviews = driver.find_elements(By.CSS_SELECTOR, ".apphub_Card")
            current_count = len(reviews)

            if current_count > loaded_reviews:
                loaded_reviews = current_count
                print(f"已加载 {loaded_reviews} 条评论")
                scroll_attempts = 0
            else:
                scroll_attempts += 1

            # 实时提取并写入新评论
            new_reviews = reviews[last_saved_count:current_count]
            if new_reviews:
                with open(output_file, mode='a', newline='', encoding='utf-8-sig') as file:
                    writer = csv.DictWriter(file, fieldnames=fieldnames)
                    for review in new_reviews:
                        data = extract_review_data(review)
                        if data:
                            writer.writerow(data)
                            file.flush()  # ✅ 实时写入磁盘
                    last_saved_count = current_count

            # 滚动加载
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)

        print(f"爬取完成，共保存 {loaded_reviews} 条评论至 {output_file}")

    except TimeoutException:
        print("页面加载超时，请检查网络连接或游戏ID是否正确")
    except Exception as e:
        print(f"发生错误: {e}")
    finally:
        driver.quit()



def save_to_csv(data, filename):
    """将评论数据保存到CSV文件"""
    fieldnames = [
        "游戏名称", "国家", "发表时间", "作者",
        "游戏时长", "内容", "是否推荐/好评"
    ]

    with open(filename, mode='w', newline='', encoding='utf-8-sig') as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)


if __name__ == "__main__":
    output_file = f"steam_reviews_{app_id}_{datetime.now().strftime('%Y%m%d')}.csv"
    scrape_steam_reviews(app_id, max_reviews, output_file)