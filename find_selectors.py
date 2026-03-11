#! /usr/bin/env python
# coding:utf-8
"""
自动探测 1688 页面元素选择器
"""
import sys, io, time, random
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

driver = webdriver.Firefox()
wait = WebDriverWait(driver, 20)
time.sleep(3)

# ── 登录 ──────────────────────────────────────────────
driver.get('https://login.taobao.com/member/login.jhtml?style=mini&css_style=taobao')
time.sleep(5)
try:
    driver.find_element(By.XPATH,
        "//*[contains(text(),'密码登录') or contains(text(),'账号密码登录')]").click()
    time.sleep(2)
except Exception:
    pass
for by, sel in [(By.ID,'fm-login-id'),(By.NAME,'TPL_username')]:
    try:
        el = wait.until(EC.presence_of_element_located((by, sel)))
        el.clear(); el.send_keys('18257961003'); break
    except Exception:
        continue
for by, sel in [(By.ID,'fm-login-password'),(By.NAME,'TPL_password'),(By.XPATH,"//input[@type='password']")]:
    try:
        el = driver.find_element(by, sel)
        el.clear(); el.send_keys('colin0828'); el.send_keys(Keys.ENTER); break
    except Exception:
        continue
print('登录完成，等待跳转...')
time.sleep(8)

# ── 采集搜索页的产品选择器 ─────────────────────────────
search_url = 'https://s.1688.com/company/company_search.htm?keywords=%BB%AF%B9%A4&n=y&spm=a260k.635.1998096057.d1'
driver.get(search_url)
time.sleep(4)

print('\n======== 搜索结果页：产品区域元素 ========')
# 找第一个商家卡片，打印卡片内所有元素
cards = driver.find_elements(By.CSS_SELECTOR, '.company-card, .list-item, .search-result-item, [class*="company"]')
if cards:
    card = cards[0]
    print(f'商家卡片 class: {card.get_attribute("class")}')
    children = card.find_elements(By.XPATH, './/*')
    for ch in children[:40]:
        t = ch.text.strip()
        if t:
            print(f'  <{ch.tag_name} class="{ch.get_attribute("class")}"> {t[:60]}')
else:
    # 直接打印整个body前4000字文本
    print('未找到卡片，打印搜索页body文本:')
    print(driver.find_element(By.TAG_NAME,'body').text[:4000])

# 找第一家公司URL
companies = driver.find_elements(By.CSS_SELECTOR, 'a.company-name')
if not companies:
    print('未找到 a.company-name')
    driver.quit()
    sys.exit(1)

first_company_url = companies[0].get_attribute('href')
first_company_name = companies[0].get_attribute('title') or companies[0].text
print(f'\n第一家公司: {first_company_name}')
print(f'URL: {first_company_url}')

# ── 访问公司主页，探测联系方式入口 ──────────────────────
print('\n======== 公司主页元素 ========')
driver.get(first_company_url)
time.sleep(random.uniform(4,6))

body_text = driver.find_element(By.TAG_NAME,'body').text
# 查找含"联系"的链接
links = driver.find_elements(By.XPATH, "//a[contains(text(),'联系') or contains(text(),'contact')]")
for lk in links[:10]:
    print(f'  链接: "{lk.text}" href={lk.get_attribute("href")}')

# ── 进入联系方式页面 ────────────────────────────────────
contact_url = first_company_url.rstrip('/') + '/page/contactinfo.htm'
print(f'\n直接访问联系页: {contact_url}')
driver.get(contact_url)
time.sleep(random.uniform(4,6))

page_text = driver.find_element(By.TAG_NAME,'body').text
if 'slide to verify' in page_text.lower():
    print('触发验证码！尝试先从主页进入...')
    driver.get(first_company_url)
    time.sleep(5)
    try:
        driver.find_element(By.XPATH,
            "//a[contains(text(),'联系方式') or contains(text(),'联系我们')]").click()
        time.sleep(4)
    except Exception as e:
        print(f'点击联系方式失败: {e}')

print('\n======== 联系信息页 body 全文 ========')
body = driver.find_element(By.TAG_NAME,'body').text
print(body[:5000])

print('\n======== 联系信息页所有有文本的元素 ========')
all_els = driver.find_elements(By.XPATH, '//*[string-length(normalize-space(text()))>0]')
seen = set()
for el in all_els:
    t = el.text.strip()
    cls = el.get_attribute('class') or ''
    tag = el.tag_name
    key = f'{tag}|{cls}|{t[:40]}'
    if key not in seen and len(t) < 100 and tag not in ('script','style'):
        seen.add(key)
        print(f'  <{tag} class="{cls}"> {t[:80]}')

# driver.quit()
print('\n探测完成')
