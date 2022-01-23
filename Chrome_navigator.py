import os.path
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from webdriver_manager.chrome import ChromeDriverManager


def verba_connect_login(my_username, my_password):
    # Login to Verba Connect
    driver = webdriver.Chrome(ChromeDriverManager().install())
    # Open the website
    driver.get('https://verbaconnect.com/auth/vst/login')

    # Select the id box
    id_box = driver.find_element_by_id('username')
    # Send id information
    id_box.send_keys(my_username)

    # Find password box
    pass_box = driver.find_element_by_id('password')
    # Send password
    pass_box.send_keys(my_password)

    # Find login button
    login_button_css = 'button[type="submit"]'
    login_button = driver.find_element_by_css_selector(login_button_css)
    # Click login
    login_button.click()
    input('\nLogin to Verba Connect ready? Make the window fullscreen and press Enter to continue.')

    return driver


def automatic_verba_upload(driver, csv_file, Verba_School, Catalog):
    School_Selected = False
    Catalog_Selected = False
    Access_Code_import_Open = False
    File_imported = 'Failed Import'

    # Open school menu
    drop_down_xpath = '/ html / body / div[1] / div / nav / div[2] / div[1] / div[1]'
    drop_down_menu = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, drop_down_xpath)))
    drop_down_menu.click()

    # select school
    try:
        school_menu_xpath = '/ html / body / div[1] / div / nav / div[2] / div[1] / div[2] / div / div[1] / div / select'
        school_menu = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, school_menu_xpath)))
        school_menu_select = Select(school_menu)
        school_menu_select.select_by_visible_text(Verba_School)
        School_Selected = True
        print('School Selected')
        time.sleep(1)

    except:
        print('Could not find school name.')

    # These schools present problems when loading catalog menu, giving time to load.
    bad_schools = ['bnc-walshuniv', 'bnc-technicalclowcountry', 'bnc-uofakron']
    if Verba_School in bad_schools:
        time.sleep(5)

    if School_Selected:
        # Select Catalog
        try:
            # Find and click on list of catalogs
            catalog_menu_xpath = '/ html / body / div[1] / div / nav / div[1] / div[1] / div[1]'
            # Click catalog drop down menu
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, catalog_menu_xpath))).click()
            # Select catalog
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.LINK_TEXT, str(Catalog).upper()))).click()
            Catalog_Selected = True
            print('Catalog Selected')

        except:
            print('Could not find Catalog name.')

    if Catalog_Selected:
        # try open access code import if superadmin is on
        try:
            # Open settings
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, 'Settings'))).click()

            # Open access code import
            access_code_import_xpath = '/ html / body / div[1] / div / div[1] / div / div[1] / div / div / a[9]'
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, access_code_import_xpath))).click()
            Access_Code_import_Open = True
            print('Access Code import OPEN')

        except:
            # Turn on Superadmin
            drop_down_menu.click()

            superadmin_xpath = '/ html / body / div[1] / div / nav / div[2] / div[1] / div[2] / div / a[1]'
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, superadmin_xpath))).click()
            print('Speradmin ON')
            time.sleep(5)

            # Open settings
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.NAME, 'Settings'))).click()

            # Open access code import
            access_code_import_xpath = '/ html / body / div[1] / div / div[1] / div / div[1] / div / div / a[9]'
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, access_code_import_xpath))).click()
            Access_Code_import_Open = True
            print('Access Code import OPEN')

    if Access_Code_import_Open:
        try:
            # Send File
            drop_file_xpath = '/ html / body / div[1] / div / div[1] / div / div[2] / div / div / div[1] / div / div / div[2] / div / input'
            drop_file = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, drop_file_xpath)))

            drop_file.send_keys(str(os.path.abspath(csv_file)))
            upload_button_xpath = '/ html / body / div[1] / div / div[1] / div / div[2] / div / div / div[2] / div / div / div / button'
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, upload_button_xpath))).click()
            print('File UPLOADED')

            finish_import_button_xpath = '/ html / body / div[1] / div / div[1] / div / div[2] / div / div / div[2] / div / div / div[1] / button'
            WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, finish_import_button_xpath))).click()
            File_imported = 'OK'
            print('File IMPORTED')

        except:
            print('File not uploaded.')

    return File_imported


def vital_source_login(my_username='Username', my_password='Password'):
    # Login to manage.vitalsource
    driver = webdriver.Chrome(ChromeDriverManager().install())
    # Open the website
    driver.get('https://manage.vitalsource.com/doorman/vst/login')

    # Select the id box
    id_box = driver.find_element_by_id('username')
    # Send id information
    id_box.send_keys(my_username)

    # Find password box
    pass_box = driver.find_element_by_id('password')
    # Send password
    pass_box.send_keys(my_password)

    # Find login button
    login_button = driver.find_element_by_id('login-form')
    # Click login
    login_button.click()

    return driver


def vital_source_url_search(driver, Billing_ISBN):
    # Find search bar and search for ISBN
    search_bar = driver.find_element_by_id('search_bar')
    search_bar.send_keys(Billing_ISBN)

    # Find book and open
