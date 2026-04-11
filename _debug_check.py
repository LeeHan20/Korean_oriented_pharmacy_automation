from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By

opts = Options()
opts.add_experimental_option("debuggerAddress", "localhost:9222")
d = webdriver.Chrome(options=opts)

for h in d.window_handles:
    d.switch_to.window(h)
    if "ongkihanyak" in d.current_url:
        break

d.switch_to.frame("right")

# 택배엑셀다운 td 주변 HTML 확인
elems = d.find_elements(By.XPATH, "//td[contains(text(),'택배엑셀다운')]")
for e in elems:
    # 부모 tr -> 부모 table 수준 HTML
    try:
        parent_tr = e.find_element(By.XPATH, "..")
        parent_table = parent_tr.find_element(By.XPATH, "..")
        # 해당 테이블 내 select, a 요소
        sels = parent_table.find_elements(By.TAG_NAME, "select")
        anchors = parent_table.find_elements(By.TAG_NAME, "a")
        print("=== 택배엑셀다운 섹션 ===")
        print(f"  select 수: {len(sels)}")
        for s in sels:
            name = s.get_attribute("name")
            opts_list = [o.text for o in s.find_elements(By.TAG_NAME, "option")]
            print(f"    SELECT name={name} options={opts_list}")
        print(f"  a 수: {len(anchors)}")
        for a in anchors:
            print(f"    A text={repr(a.text)} href={a.get_attribute('href')} onclick={a.get_attribute('onclick')}")
    except Exception as ex:
        print("오류:", ex)

d.switch_to.default_content()
