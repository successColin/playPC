#! /usr/bin/env python
# coding:utf-8

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import sys
import io

# 解决终端输出编码问题
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

driver = webdriver.Firefox()
wait = WebDriverWait(driver, 20)
time.sleep(3)

login_url = 'https://login.taobao.com/member/login.jhtml?style=mini&css_style=taobao'
driver.get(login_url)
time.sleep(5)

try:
    switch_btn = driver.find_element(By.XPATH, "//*[contains(text(),'密码登录') or contains(text(),'账号密码登录')]")
    switch_btn.click()
    time.sleep(2)
except Exception:
    pass

for by, selector in [(By.ID, "fm-login-id"), (By.NAME, "TPL_username")]:
    try:
        inp = wait.until(EC.presence_of_element_located((by, selector)))
        inp.clear()
        inp.send_keys('18257961003')
        print('账号输入成功')
        break
    except Exception:
        continue

for by, selector in [(By.ID, "fm-login-password"), (By.NAME, "TPL_password"), (By.XPATH, "//input[@type='password']")]:
    try:
        pwd = driver.find_element(by, selector)
        pwd.clear()
        pwd.send_keys('colin0828')
        pwd.send_keys(Keys.ENTER)
        print('密码输入，等待登录...')
        break
    except Exception:
        continue

time.sleep(8)

test_url = 'https://shop9g7161130fc31.1688.com/page/contactinfo.htm'
print(f'正在访问: {test_url}')
driver.get(test_url)
# 等待动态内容加载
time.sleep(6)

print('\n=== 页面所有可见文本（前5000字） ===')
try:
    body_text = driver.find_element(By.TAG_NAME, 'body').text
    print(body_text[:5000])
except Exception as e:
    print(f'获取body失败: {e}')

print('\n=== 尝试各种元素选择器 ===')
selectors_to_try = [
    ('class包含tel', By.XPATH, "//*[contains(@class,'tel')]"),
    ('class包含phone', By.XPATH, "//*[contains(@class,'phone')]"),
    ('class包含contact', By.XPATH, "//*[contains(@class,'contact')]"),
    ('class包含address', By.XPATH, "//*[contains(@class,'address')]"),
    ('class包含member', By.XPATH, "//*[contains(@class,'member')]"),
    ('class包含name', By.XPATH, "//*[contains(@class,'name') and not(contains(@class,'company'))]"),
    ('dt标签', By.TAG_NAME, "dt"),
    ('dd标签', By.TAG_NAME, "dd"),
]
for label, by, selector in selectors_to_try:
    try:
        els = driver.find_elements(by, selector)
        if els:
            print(f'\n[{label}] 找到 {len(els)} 个元素:')
            for el in els[:5]:
                text = el.text.strip()
                cls = el.get_attribute('class') or ''
                if text:
                    print(f'  class="{cls}" -> "{text[:80]}"')
    except Exception as e:
        print(f'[{label}] 查询失败: {e}')

driver.quit()
