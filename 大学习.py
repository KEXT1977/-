# -*- coding: utf-8 -*-  
from selenium import webdriver
from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pandas import read_excel
from matplotlib.pyplot import figure,rcParams,table,xticks,yticks,savefig

rcParams['font.sans-serif']=['SimHei'] #用来正常显示中文标签
rcParams['axes.unicode_minus'] = False #用来正常显示负号
from os import path,makedirs,listdir,rename
from time import sleep
from shutil import rmtree


url = "https://tuan.12355.net/bg/index.html"

account = ""

password = ""

path_temp = "temp_for_big_study"


if path.exists(path_temp):
    rmtree(path_temp)    
makedirs(path_temp)
   


options = EdgeOptions()
options.use_chromium = True
options.add_experimental_option("useAutomationExtension", False)
options.add_experimental_option("excludeSwitches", ['enable-automation'])
options.add_experimental_option('excludeSwitches', ['enable-logging']) # 防止出现提示DevTools
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--disable-gpu")
# options.add_argument("--headless")  # 开启无界面模式
# options.add_argument('--user-data-dir='"C:\Users\{}\AppData\Local\Microsoft\Edge\User Data\Default")
prefs = {'profile.default_content_settings.popups': 0, 
 'download.default_directory': path_temp} 
options.add_experimental_option('prefs', prefs) 
driver = Edge(options=options, executable_path=r"D:\msedgedriver.exe") # 相应的浏览器的驱动位置
driver.get(url)


input_text = driver.find_elements_by_class_name('input')
for i,j in zip(input_text,range(2)):
    i.clear()
    input_temp = account
    if j==1:
        input_temp = password

    i.send_keys(input_temp)

driver.find_element_by_css_selector('#login').click()


wait = WebDriverWait(driver,timeout=20,poll_frequency=0.02)
wait.until(lambda x:driver.find_element(By.XPATH,'//div[contains(text(), "青年大学习")]')).click()


WebDriverWait(driver,30,0.02).until(lambda x:len(driver.window_handles)==2)
driver.switch_to.window(driver.window_handles[1])


def refresh_daxuexi():
    try:
        WebDriverWait(driver,10,0.02).until(EC.presence_of_element_located((By.CLASS_NAME,'hamburger'))).click()
    except:
        driver.refresh()
        refresh_daxuexi()
refresh_daxuexi()
sleep(0.2)


try:
    WebDriverWait(driver,30,0.02).until(EC.presence_of_element_located((By.CSS_SELECTOR, '#app > div > div.sidebar-container.el-scrollbar > div.scrollbar-wrapper.el-scrollbar__wrap > div > ul > div:nth-child(2) > a'))).click()
    
except:
    driver.find_element(By.XPATH, '//span[contains(text(), "数据查看")]').click()


WebDriverWait(driver,30,0.02).until(lambda x:driver.find_elements_by_class_name('link-button'))[0].click()


while len(listdir(path_temp))==0:
    try:
        WebDriverWait(driver,10,0.02).until(lambda x:driver.find_element(By.XPATH, '//span[contains(text(), "导出未参学团员名单")]')).click()
        sleep(2)
    except:
        sleep(1)
        driver.refresh()     
sleep(1)
driver.quit()
df2 = read_excel(path.join(path_temp,listdir(path_temp)[0]),header=1)

#undoset = set(df['姓名'])

df2 = df2 [['姓名','期数','学习情况']]
#list(df2.columns)

h = df2.shape[0]
w = df2.shape[1]
figure(figsize=(10,10))
tab = table(cellText=df2.values, 
              colLabels=list(df2.columns),
              loc='center',
              cellLoc='center',
              rowLoc='center')

xticks([]) 
yticks([])
##plt.subplots_adjust(left=0.5, bottom=0)
tab.scale(1,1.5)

savefig("temp_for_big_study\\199青年大学习未做名单.png",dpi=600,bbox_inches='tight')
with open(path.join(path_temp,'199共{}人完成青年大学习，团员40人，非团员0人，参学率{:.2f}%'.format(40-df2.shape[0],(40-df2.shape[0])/40.0 *100)),'w',encoding='utf-8') as f:
    print('Over')

for i in listdir(path_temp):
    if 'xlsx' in i:
        rename(path.join(path_temp,i),path.join(path_temp,'199未参学名单【青年大学习】.xlsx'))
        break
