import requests
import pandas as pd
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random

# 获取当前日期并调整日期范围，这里往前推 7 天
start_date = (datetime.today() - timedelta(days=7)).strftime('%Y-%m-%d')
end_date = datetime.today().strftime('%Y-%m-%d')

# 获取当前运行时间，用于生成文件名
current_time = datetime.now().strftime('%Y-%m-%d %H-%M-%S')
filename = f'定增公告{current_time}.xlsx'

# 巨潮资讯网公告查询接口
url = 'http://www.cninfo.com.cn/new/hisAnnouncement/query'

# 请求头
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3',
    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'
}

# 总页数
total_pages = 10
data_list = []

# 配置 Chrome 浏览器
chrome_options = Options()
chrome_options.add_argument('--headless')  # 无头模式，不显示浏览器窗口
service = Service('D:\Program Files\chromedriver\chromedriver.exe')  # 请替换为你的 ChromeDriver 路径
driver = webdriver.Chrome(service=service, options=chrome_options)

# 打开新的可以选择行业的网页
industry_selection_url = 'http://www.cninfo.com.cn/new/commonUrl/pageOfSearch?url=disclosure/list/search'
driver.get(industry_selection_url)
# 等待页面加载
WebDriverWait(driver, 20).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
)
time.sleep(random.uniform(3, 6))

try:
    # 定位并点击“行业”按钮，使用新的 XPath
    industry_button = WebDriverWait(driver, 1).until(
        EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div[2]/div[1]/div[2]/div/div[2]/form/div[2]/div[4]/div/div/span/button"))
    )
    industry_button.click()
    time.sleep(1)

    # 定位并选中“信息传输、软件和信息技术服务业”标签，使用新的 XPath
    target_label = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[6]/div[1]/label[9]"))
    )
    industry_button.click()
     # 检查标签是否已选中
    if 'is-checked' not in target_label.find_element(By.XPATH, '..').get_attribute('class'):
        try:
            # 尝试直接点击
            target_label.click()
        except:
            # 若直接点击失败，使用 JavaScript 模拟点击
            driver.execute_script("arguments[0].click();", target_label)
        time.sleep(random.uniform(3, 6))
    else:
        print("标签已经被选中，无需再次点击。")
    

except Exception as e:
    print(f"选中标签失败: {e}")
    
'''
for page in range(1, total_pages + 1):
    # 请求参数
    data = {
        'pageNum': page,
        'pageSize': 30,
        'column':'szse',  # 深圳证券交易所，可根据需要调整为 shse（上海证券交易所）等
        'tabName': 'fulltext',
        'plate': '',
       'stock': '',
       'searchkey': '',
       'secid': '',
        'category': '',
        'trade': '',
       'seDate': f'{start_date}~{end_date}',  # 使用调整后的日期范围
       'sortName': '',
       'sortType': '',
        'isHLtitle': 'true'
    }

    try:
        # 发送 POST 请求
        response = requests.post(url, headers=headers, data=data)
        print(f"第 {page} 页公告列表请求状态码: {response.status_code}")
        response.raise_for_status()

        # 解析 JSON 数据
        result = response.json()
        announcements = result['announcements']

        for announcement in announcements:
            secName = announcement['secName']
            secCode = announcement['secCode']
            announcementTitle = announcement['announcementTitle']
            adjunctUrl = 'http://static.cninfo.com.cn/' + announcement['adjunctUrl']
            # 提取公告时间
            announcementTime = datetime.fromtimestamp(announcement['announcementTime'] / 1000).strftime('%Y-%m-%d %H:%M:%S')

            # 筛选包含“特定对象”“定向增发”“定增”的公告
            keywords = ['特定对象', '定向增发', '定增']
            if any(keyword in announcementTitle for keyword in keywords):
                data_list.append([secName, secCode, announcementTitle, adjunctUrl, announcementTime])

    except requests.RequestException as e:
        print(f'第 {page} 页公告列表请求出错: {e}')
    except KeyError as e:
        print(f'第 {page} 页解析 JSON 数据出错: {e}')
    except Exception as e:
        print(f'第 {page} 页发生未知错误: {e}')

driver.quit()

# 创建 DataFrame
df = pd.DataFrame(data_list, columns=['公司名称', '股票代码', '公告标题', '公告链接', '公告时间'])

# 保存到 Excel 文件
df.to_excel(filename, index=False)
print(f'符合条件的公告信息已成功保存到 {filename} 文件中。')
'''
