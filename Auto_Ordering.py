# web control
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

import pandas as pd # data processing
import time # delays
import os # paths
import numpy as np # nan
import subprocess # to not close window on program exit
import webbrowser # opening res
import pyperclip # clipboard
import json # for reading secrets.json

with open('secrets.json') as secrets_file:
    secrets = json.load(secrets_file)
    
DISCOUNT_CODE = secrets['DISCOUNT_CODE']

BILLING_FIRST_NAME = secrets['BILLING_ADDR']['FIRST_NAME']
BILLING_LAST_NAME = secrets['BILLING_ADDR']['LAST_NAME']
BILLING_COMPANY = secrets['BILLING_ADDR']['COMPANY']
BILLING_ADDRESS = secrets['BILLING_ADDR']['ADDRESS']
BILLING_CITY = secrets['BILLING_ADDR']['CITY']
BILLING_STATE = secrets['BILLING_ADDR']['STATE']
BILLING_ZIP_CODE = secrets['BILLING_ADDR']['ZIP_CODE']

PRODUCT_PAGES = secrets['PRODUCT_PAGES']

RES_PAGE = secrets['RES_PAGE']

# data processing
db = pd.read_excel('Gear0812.xlsx')
db.rename(columns=str.strip, inplace=True) # strip trailing whitespaces
db = db[['Reservation #', 'Guest Name', 'Address', 
         'Country', 'Sex', 'T-Shirt', 'Shorts', 
         'Sport Cut Jersey', 'Womens Racerback Jersey', 
         'Socks', 'Virtuoso', 'Gear Ordered']] # select subset of columns
db.dropna(axis=0, subset=['Guest Name', 'Reservation #'], inplace=True) # drop blank entries
db['Reservation #'] = db['Reservation #'].astype(int) # convert res #s to ints

# conversion from size names in sheet to those used by Primal
SIZE_NAMES = {
    'X-Small': 'XS',
    'Small': 'SM',
    'Medium': 'MD',
    'Large': 'LG',
    'X-Large': 'XL',
    'XX-Large': '2X',
    'XXX-Large': '3X'
}

# reformat sizes
def size_str_cleaner(size_str: str) -> str | float:
    # 'Do not send...' -> NaN
    if size_str[0] == 'D':
        return np.nan
    
    return SIZE_NAMES[size_str.split(' ')[1]]

sized_items = ['T-Shirt', 'Shorts', 'Sport Cut Jersey'] # columns to convert

db[sized_items] = db[sized_items].map(size_str_cleaner, na_action='ignore')

# selenium setup
opts = webdriver.ChromeOptions()

# do not close browser when program stops
opts.add_experimental_option('detach', True)
serv = Service(popen_kw={"creation_flags": subprocess.CREATE_NEW_PROCESS_GROUP})

# retrieve session info from Open_Browser.py's session
driver = None
url = ''
session_id = ''
with open('session_info.txt', 'r') as session_info:
    lines = session_info.readlines()
    url = lines[0].strip()
    session_id = lines[1].strip()

# connect webdriver and close new window that auto-opened
driver = webdriver.Remote(command_executor=url, options=opts)
driver.close()
driver.quit()

driver.session_id = session_id
    
driver.implicitly_wait(5) # will wait 5 secs for an element to appear before crashing

PATH = os.getcwd() # path to script location

# guest t-shirts use 2XL and 3XL instead of 2X and 3X, for some reason
GUEST_SHIRT_SIZES = {
    'SM': 'SM',
    'MD': 'MD',
    'LG': 'LG',
    'XL': 'XL',
    '2X': '2XL',
    '3X': '3XL'
}

FIREFOX_PATH = 'C:\\Program Files\\Mozilla Firefox\\firefox.exe'
webbrowser.register('firefox', None, webbrowser.BackgroundBrowser(FIREFOX_PATH), preferred=True)

print(f'commend executor url: {driver.command_executor._url}')
print(f'session id: {driver.session_id}')
print('MAKE SURE TO LOGIN, AND WAIT FOR AND CLOSE POPUP')

# country names in sheet to country codes in Primal
COUNTRIES = {
    'usa': 'US',
    'canada': 'CA'
}

def checkout(name_rows: list[pd.Series]):
    """Starts checkout, fills name, country, discount code, and address, 
    then waits for user to manually select correct address

    Args:
        name_rows (list[pd.Series]): list of all people added to cart
    """
    
    # navigate to checkout
    driver.get('https://www.primalwear.com/cart')
    checkout_button = driver.find_element(By.NAME, 'checkout')
    checkout_button.click()
    
    # clear address fields
    addr_select = driver.find_element(By.XPATH, '//span[text()="Saved addresses"]') # span with 'Saved addresses' text
    addr_select_id = addr_select.find_element(By.XPATH, '../..').get_attribute('for') # id of dropdown from label surrounding span
    addr_select = Select(driver.find_element(By.ID, addr_select_id)) # dropdown from id
    addr_select.select_by_value('4') # choose different addr to ensure that selecting new clears fields
    addr_select.select_by_value('5') # select 'create new address'
    
    # select country
    country_select = Select(driver.find_element(By.NAME, 'countryCode'))
    country_select.select_by_value(COUNTRIES[name_rows[0]['Country'].lower().strip()])
    
    # name fields
    first_name = driver.find_element(By.NAME, 'firstName')
    last_name = driver.find_element(By.NAME, 'lastName')
    
    # get all first names
    first_names = [row['Guest Name'].split()[0].capitalize() for row in name_rows]
    if len(name_rows) > 1:
        first_names = ', '.join(first_names[:-1]) + ' and ' + first_names[-1] # format as list of names
    else:
        first_names = first_names[0] # just one
    
    # get all last names
    last_names = [row['Guest Name'].split()[-1].capitalize() for row in name_rows]
    last_names = list(dict.fromkeys(last_names)) # remove duplicates
    if len(last_names) > 1:
        last_names = ', '.join(last_names[:-1]) + ' and ' + last_names[-1] # format as list of names
    else:
        last_names = last_names[0] # just one
        
    # input names
    # e.g.: John Smith and Mary Smith -> First: John and Mary, Last: Smith
    # e.g.: John Doe and Mary Smith -> First: John and Mary, Last: Doe and Smith
    first_name.send_keys(first_names)
    last_name.send_keys(last_names)

    # input discount code
    discount = driver.find_element(By.NAME, 'reductions')
    discount.send_keys(DISCOUNT_CODE+'\n')
    
    address_input = driver.find_element(By.ID, 'shipping-address1')
    address_input.send_keys(name_rows[0]['Address'])

def complete_order() -> str:
    """On payment page, enter billing address and complete order

    Returns:
        confirmation_num (str): confirmation number
    """
    diff_addr_radio = driver.find_element(By.ID, 'billing_address_selector-custom')
    diff_addr_radio.click()
    
    addr_select = driver.find_element(By.XPATH, '//span[text()="Saved addresses"]')
    addr_select_id = addr_select.find_element(By.XPATH, '../..').get_attribute('for')

    addr_select = Select(driver.find_element(By.ID, addr_select_id))
    addr_select.select_by_value('4')
    addr_select.select_by_value('5')

    country_select = Select(driver.find_element(By.NAME, 'countryCode'))
    country_select.select_by_value('US')

    first_name = driver.find_element(By.NAME, 'firstName')
    last_name = driver.find_element(By.NAME, 'lastName')
    first_name.send_keys(BILLING_FIRST_NAME)
    last_name.send_keys(BILLING_LAST_NAME)

    company_input = driver.find_element(By.NAME, 'company')
    company_input.send_keys(BILLING_COMPANY)

    address_input = driver.find_element(By.ID, 'billing-address1')
    address_input.send_keys(BILLING_ADDRESS)
    
    city_input = driver.find_element(By.NAME, 'city')
    city_input.send_keys(BILLING_CITY)
    
    state_selection = Select(driver.find_element(By.NAME, 'zone'))
    state_selection.select_by_value(BILLING_STATE)
    
    zip_input = driver.find_element(By.NAME, 'postalCode')
    zip_input.send_keys(BILLING_ZIP_CODE)
    
    submit = driver.find_element(By.XPATH, '//span[text()="Complete order"]/..')
    submit.click()
    
    confirmation = driver.find_element(By.XPATH, '//p[contains(text(), "Confirmation #")]')
    confirmation_num = confirmation.text.split('#')[1].strip()
    
    return confirmation_num

PAGE_LOAD_WAIT = 1.0 # wait after page load before selecting size
ADD_TO_CART_WAIT = 0.5 # wait after selecting size to add to cart
CART_PROCESS_WAIT = 0.2 # wait after added to cart to ensure it has processed

def add_to_cart(row: pd.Series):
    """Adds items to cart for a person

    Args:
        row (pd.Series): sheet entry for person
    """
    if not row.isnull()['T-Shirt']:
        # load page
        driver.get(PRODUCT_PAGES[row['Sex'] + ' Shirt'])
        time.sleep(PAGE_LOAD_WAIT)
        
        # find and select size
        size_button = driver.find_element(By.XPATH, f'//input[@value="{GUEST_SHIRT_SIZES[row['T-Shirt']]}"]')
        size_button.find_element(By.XPATH, '..').click()
        time.sleep(ADD_TO_CART_WAIT)
        
        # add to cart and wait for cart sidebar to appear
        driver.find_element(By.ID, 'addToCart-product').click()
        driver.find_element(By.CLASS_NAME, 'cart-dropdown.is-open')
        time.sleep(CART_PROCESS_WAIT)
        
    if not row.isnull()['Shorts']:
        # load page
        driver.get(PRODUCT_PAGES[row['Sex'] + ' Short'])
        time.sleep(PAGE_LOAD_WAIT)
        
        # find and select size
        size_button = driver.find_element(By.XPATH, f'//input[@value="{row['Shorts']}"]')
        size_button.find_element(By.XPATH, '..').click()
        time.sleep(ADD_TO_CART_WAIT)
        
        # add to cart and wait for cart sidebar to appear
        driver.find_element(By.ID, 'addToCart-product').click()
        driver.find_element(By.CLASS_NAME, 'cart-dropdown.is-open')
        time.sleep(CART_PROCESS_WAIT)
        
    if not row.isnull()['Sport Cut Jersey']:
        # load page
        driver.get(PRODUCT_PAGES[row['Sex'] + ' Prisma'])
        time.sleep(PAGE_LOAD_WAIT)
        
        # find and select size
        size_button = driver.find_element(By.XPATH, f'//input[@value="{row['Sport Cut Jersey']}"]')
        size_button.find_element(By.XPATH, '..').click()
        time.sleep(ADD_TO_CART_WAIT)
        
        # add to cart and wait for cart sidebar to appear
        driver.find_element(By.ID, 'addToCart-product').click()
        driver.find_element(By.CLASS_NAME, 'cart-dropdown.is-open')
        time.sleep(CART_PROCESS_WAIT)
        
    if not row.isnull()['Womens Racerback Jersey']:
        # load page
        driver.get(PRODUCT_PAGES['Racerback'])
        time.sleep(PAGE_LOAD_WAIT)
        
        # find and select size
        size_button = driver.find_element(By.XPATH, f'//input[@value="{SIZE_NAMES[row['Womens Racerback Jersey']]}"]')
        size_button.find_element(By.XPATH, '..').click()
        time.sleep(ADD_TO_CART_WAIT)
        
        # add to cart and wait for cart sidebar to appear
        driver.find_element(By.ID, 'addToCart-product').click()
        driver.find_element(By.CLASS_NAME, 'cart-dropdown.is-open')
        time.sleep(CART_PROCESS_WAIT)
    
    if not row.isnull()['Socks']:
        # load page
        driver.get(PRODUCT_PAGES['Socks'])
        time.sleep(PAGE_LOAD_WAIT)
        
        # determine appropriate size from other items
        size = ( row['T-Shirt'] if not row.isnull()['T-Shirt'] 
                else row['Sport Cut Jersey'] if not row.isnull()['Sport Cut Jersey'] 
                else row['Womens Racerback Jersey'] )
            
        size = 'SM/MD' if size in ['XS', 'SM', 'MD'] else 'LG/XL'
        
        # find and select size
        size_button = driver.find_element(By.XPATH, f'//input[@value="{size}"]')
        size_button.find_element(By.XPATH, '..').click()
        time.sleep(ADD_TO_CART_WAIT)
        
        # add to cart and wait for cart sidebar to appear
        driver.find_element(By.ID, 'addToCart-product').click()
        driver.find_element(By.CLASS_NAME, 'cart-dropdown.is-open')
        time.sleep(CART_PROCESS_WAIT)

prev_rows = [] # entries in cart
prev_index = -1 # used for getting next entry
while True:
    name = input('Enter Name/enter for checkout/"c" to complete order/"n" to go to next name: ')
    
    # checkout
    if name == '' and prev_rows:
        checkout(prev_rows)
        prev_rows = []
        continue
    
    # complete order
    if name == 'c':
        confirmation_num = complete_order()
        pyperclip.copy(confirmation_num)
        print('Copied confirmation #', confirmation_num)
        continue
    
    # go to next name
    if name == 'n':
        name = db.iloc[prev_index+1]['Guest Name']
    
    # catch errors
    try:
        row = db.loc[db['Guest Name'] == name].squeeze(axis=0) # make sure name entry exists
        assert type(row) == pd.Series # assert only one entry for name
    except KeyError:
        print('Name not found...Try again\n')
        continue
    except AssertionError:
        print('Multiple of the same name...Try again\n')
        continue
    
    # update prev_ vars
    prev_rows.append(row)
    prev_index = row.name
    print('Name found...')
    
    # open res page if first person in order
    if len(prev_rows) == 1:
        webbrowser.get('firefox').open(RES_PAGE + str(row['Reservation #']))
        print('Opened page for reservation number: ', row['Reservation #'])
    
    add_to_cart(row)
        
    print('Added all items to cart.\n')