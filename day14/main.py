from helium import *
from selenium.webdriver.common.keys import Keys
from time import sleep


start_chrome("https://www.amazon.cn/")
ipt = TextField()
write("户外帐篷", into=ipt)
press(Keys.ENTER)
sleep(3)

scroll_down(10000)
sleep(1)
for e, f in zip(find_all(S("span.a-size-base-plus")), find_all(S("span.a-price"))):
    print(e.web_element.text, f.web_element.text)
sleep(2)
kill_browser()