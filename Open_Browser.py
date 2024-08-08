from selenium import webdriver
import time

opts = webdriver.ChromeOptions()
opts.add_experimental_option('detach', True)

driver = webdriver.Chrome(options=opts)
driver.get('https://www.primalwear.com/')
print(f'commend executor url: {driver.command_executor._url}')
print(f'session id: {driver.session_id}')

with open('session_info.txt', 'w') as session_info_file:
    session_info_file.write(driver.command_executor._url + '\n')
    session_info_file.write(driver.session_id)    

while True:
    time.sleep(0.1)