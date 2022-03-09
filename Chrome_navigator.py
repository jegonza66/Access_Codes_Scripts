import os.path
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
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
    driver.maximize_window()

    return driver


def verba_open_school(driver, Verba_School):
    School_Selected = False
    # select school
    try:
        # Open school menu
        drop_down_xpath = '/ html / body / div[1] / div / nav / div[2] / div[1] / div[1]'
        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, drop_down_xpath)))
        drop_down_menu = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, drop_down_xpath)))
        drop_down_menu.click()

        school_menu_xpath = '/ html / body / div[1] / div / nav / div[2] / div[1] / div[2] / div / div[1] / div / select'
        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, school_menu_xpath)))
        school_menu = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, school_menu_xpath)))
        school_menu_select = Select(school_menu)
        school_menu_select.select_by_visible_text(Verba_School)
        School_Selected = True
        print('\nSchool Selected')
    except:
        # Try Again
        time.sleep(2)
        try:
            # Open school menu
            drop_down_xpath = '/ html / body / div[1] / div / nav / div[2] / div[1] / div[1]'
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, drop_down_xpath)))
            drop_down_menu = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, drop_down_xpath)))
            drop_down_menu.click()

            school_menu_xpath = '/ html / body / div[1] / div / nav / div[2] / div[1] / div[2] / div / div[1] / div / select'
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, school_menu_xpath)))
            school_menu = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, school_menu_xpath)))
            school_menu_select = Select(school_menu)
            school_menu_select.select_by_visible_text(Verba_School)
            School_Selected = True
            print('\nSchool Selected')
        except:
            print('Could not Open School.')

    return School_Selected


def verba_open_catalog(driver, Catalog):

    Catalog_Selected = False
    # Select Catalog
    try:
        # Find and click on list of catalogs
        catalog_menu_xpath = '/ html / body / div[1] / div / nav / div[1] / div[1] / div[1]'
        # Click catalog drop down menu
        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, catalog_menu_xpath)))
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, catalog_menu_xpath))).click()
        # Select catalog
        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.LINK_TEXT, str(Catalog).upper())))
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, str(Catalog).upper()))).click()
        Catalog_Selected = True
        print('Catalog Selected')
    except:
        # Try Again
        try:
            # Find and click on list of catalogs
            catalog_menu_xpath = '/ html / body / div[1] / div / nav / div[1] / div[1] / div[1]'
            # Click catalog drop down menu
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, catalog_menu_xpath)))
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, catalog_menu_xpath))).click()
            # Select catalog
            WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.LINK_TEXT, str(Catalog).upper())))
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.LINK_TEXT, str(Catalog).upper()))).click()
            Catalog_Selected = True
            print('Catalog Selected')
        except:
            print('Could not find Catalog name.')

    return Catalog_Selected


def verba_open_item_menu(driver):
    Items_Menu_Open = False
    # Open Items Menu
    try:
        Items_xpath = '/ html / body / div[1] / div / nav / div[1] / div[2] / a / div'
        Items_tab = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, Items_xpath)))
        # WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, Items_xpath))).click()
        ActionChains(driver).move_to_element(Items_tab).perform()
        # click connect items tab
        connect_tab_xpath = '/ html / body / div[1] / div / nav / div[1] / div[2] / div / div / a[1]'
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, connect_tab_xpath)))
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, connect_tab_xpath))).click()
        # Click search bar to chek if open
        search_bar_xpath = '/ html / body / div[1] / div / div[1] / div[3] / div / div / div[1] / div[1] / input'
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, search_bar_xpath)))
        search_item = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, search_bar_xpath)))
        search_item.click()
        Items_Menu_Open = True
        print('Items Menu Open')
    except:
        time.sleep(1)
        # Try again
        try:
            Items_xpath = '/ html / body / div[1] / div / nav / div[1] / div[2] / a / div'
            Items_tab = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, Items_xpath)))
            # WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, Items_xpath))).click()
            ActionChains(driver).move_to_element(Items_tab).perform()
            # click connect items tab
            connect_tab_xpath = '/ html / body / div[1] / div / nav / div[1] / div[2] / div / div / a[1]'
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, connect_tab_xpath)))
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, connect_tab_xpath))).click()
            # Click search bar to chek if open
            search_bar_xpath = '/ html / body / div[1] / div / div[1] / div[3] / div / div / div[1] / div[1] / input'
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, search_bar_xpath)))
            search_item = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, search_bar_xpath)))
            search_item.click()
            Items_Menu_Open = True
            print('Items Menu Open')
        except:
            print('Could not Open Items Menu')
            driver.refresh()
            time.sleep(3)

    return Items_Menu_Open


def verba_open_item(driver, sku):
    Item_Open = False
    # Search and open Item
    try:
        search_bar_xpath = '/ html / body / div[1] / div / div[1] / div[3] / div / div / div[1] / div[1] / input'
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, search_bar_xpath)))
        search_item = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, search_bar_xpath)))
        search_item.clear()
        search_item.click()
        search_item.send_keys(sku)

        search_button = '/ html / body / div[1] / div / div[1] / div[3] / div / div / div[1] / div[2] / button'
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, search_button)))
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, search_button))).click()

        item_xpath = '/ html / body / div[1] / div / div[1] / div[4] / div / div / div / div / div[1] / div[1] / div[2] / div[1] / h3 / a'
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, item_xpath))).click()
        Item_Open = True
        print('Item Open')
    except:
        time.sleep(2)
        # Try Again
        try:
            search_bar_xpath = '/ html / body / div[1] / div / div[1] / div[3] / div / div / div[1] / div[1] / input'
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, search_bar_xpath)))
            search_item = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, search_bar_xpath)))
            search_item.clear()
            search_item.click()
            search_item.send_keys(sku)

            search_button = '/ html / body / div[1] / div / div[1] / div[3] / div / div / div[1] / div[2] / button'
            WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, search_button)))
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, search_button))).click()

            item_xpath = '/ html / body / div[1] / div / div[1] / div[4] / div / div / div / div / div[1] / div[1] / div[2] / div[1] / h3 / a'
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, item_xpath))).click()
            Item_Open = True
            print('Item Open')
        except:
            print('Could not Open Item')

    return Item_Open


def check_available_codes(driver, Verba_School, Catalog, ISBN, previous_school, previous_catalog):
    Available_Codes = 0

    School_change = Verba_School != previous_school
    Catalog_change = Catalog != previous_catalog

    if School_change:
        School_Selected = verba_open_school(driver=driver, Verba_School=Verba_School)
    else:
        School_Selected = True

    if School_Selected:
        previous_school = Verba_School
        if School_change or Catalog_change:
            Catalog_Selected = verba_open_catalog(driver=driver, Catalog=Catalog)
        else:
            Catalog_Selected = True

        if Catalog_Selected:
            previous_catalog = Catalog
            Items_Menu_Open = verba_open_item_menu(driver=driver)

            if Items_Menu_Open:
                Item_Open = verba_open_item(driver=driver, sku=ISBN)

                if Item_Open:
                    Access_Codes_Open = False
                    # Open Acces Codes
                    try:
                        Access_Codes_xpath = '/ html / body / div[1] / div / div[1] / div / div[3] / div / ul / li[3] / a'
                        WebDriverWait(driver, 20).until(
                            EC.element_to_be_clickable((By.XPATH, Access_Codes_xpath))).click()
                        Access_Codes_Open = True
                        time.sleep(1)
                        print('Access Codes Open')
                    except:
                        print('Could not Open Access Codes')
                        time.sleep(2)
                        # Try Again
                        try:
                            Access_Codes_xpath = '/ html / body / div[1] / div / div[1] / div / div[3] / div / ul / li[3] / a'
                            WebDriverWait(driver, 20).until(
                                EC.element_to_be_clickable((By.XPATH, Access_Codes_xpath))).click()
                            Access_Codes_Open = True
                            time.sleep(1)
                            print('Access Codes Open')
                        except:
                            print('Could not Open Access Codes Again.')

                    if Access_Codes_Open:
                        # Check available Codes
                        try:
                            xpath_codes = '/ html / body / div[1] / div / div[1] / div / div[3] / div / div / div / div / div[1] / div[2] / span / div[1]'
                            Available_Codes = int(driver.find_element_by_xpath(xpath_codes).text)
                            print('{} Available codes'.format(Available_Codes))
                        except:
                            # Check available Codes Again
                            try:
                                xpath_codes = '/ html / body / div[1] / div / div[1] / div / div[3] / div / div / div / div / div[1] / div[2] / span / div[1]'
                                Available_Codes = int(driver.find_element_by_xpath(xpath_codes).text)
                                print('{} Available codes'.format(Available_Codes))
                            except:
                                print('No Avalable Codes.')
                                Available_Codes = 0

    return Available_Codes, previous_school, previous_catalog


def automatic_verba_upload(driver, csv_file, Verba_School, Catalog, previous_school, previous_catalog):

    File_imported = False
    School_change = Verba_School != previous_school
    Catalog_change = Catalog != previous_catalog

    if School_change:
        School_Selected = verba_open_school(driver=driver, Verba_School=Verba_School)
    else:
        School_Selected = True

    if School_Selected:
        previous_school = Verba_School
        if School_change or Catalog_change:
            Catalog_Selected = verba_open_catalog(driver=driver, Catalog=Catalog)
        else:
            Catalog_Selected = True

        if Catalog_Selected:
            previous_catalog = Catalog
            Access_Code_import_Open = False
            File_imported = 'Failed Import'

            # try open access code import if superadmin is on
            try:
                # Open settings
                settings_xpath = '/ html / body / div[1] / div / nav / div[1] / a[7]'
                WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, settings_xpath)))
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, settings_xpath))).click()

                # Open access code import
                access_code_import_xpath = '/ html / body / div[1] / div / div[1] / div / div[1] / div / div / a[9]'
                WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, access_code_import_xpath)))
                WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, access_code_import_xpath))).click()
                Access_Code_import_Open = True
                print('Access Code import OPEN')

            except:
                try:
                    Super_admin = False
                    # Turn on Superadmin
                    drop_down_xpath = '/ html / body / div[1] / div / nav / div[2] / div[1] / div[1]'
                    WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, drop_down_xpath)))
                    drop_down_menu = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, drop_down_xpath)))
                    drop_down_menu.click()

                    superadmin_xpath = '/ html / body / div[1] / div / nav / div[2] / div[1] / div[2] / div / a[1]'
                    WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, superadmin_xpath)))
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, superadmin_xpath))).click()
                    print('Speradmin ON')
                    Super_admin = True
                    time.sleep(3)

                    if Super_admin:
                        # Open settings
                        settings_xpath = '/ html / body / div[1] / div / nav / div[1] / a[7]'
                        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, settings_xpath)))
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, settings_xpath))).click()

                        # Open access code import
                        access_code_import_xpath = '/ html / body / div[1] / div / div[1] / div / div[1] / div / div / a[9]'
                        WebDriverWait(driver, 10).until(
                            EC.visibility_of_element_located((By.XPATH, access_code_import_xpath)))
                        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, access_code_import_xpath))).click()
                        Access_Code_import_Open = True
                        print('Access Code import OPEN')
                except:
                    print('Could not turn on Superadmin mode')

            if Access_Code_import_Open:
                try:
                    # Send File
                    drop_file_xpath = '/ html / body / div[1] / div / div[1] / div / div[2] / div / div / div[1] / div / div / div[2] / div / input'
                    WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, drop_file_xpath)))
                    drop_file = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, drop_file_xpath)))

                    drop_file.send_keys(str(os.path.abspath(csv_file)))
                    upload_button_xpath = '/ html / body / div[1] / div / div[1] / div / div[2] / div / div / div[2] / div / div / div / button'
                    WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, upload_button_xpath)))
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, upload_button_xpath))).click()
                    print('File UPLOADED')

                    finish_import_button_xpath = '/ html / body / div[1] / div / div[1] / div / div[2] / div / div / div[2] / div / div / div[1] / button'
                    WebDriverWait(driver, 10).until(
                        EC.visibility_of_element_located((By.XPATH, finish_import_button_xpath)))
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, finish_import_button_xpath))).click()
                    File_imported = 'OK'
                    print('File IMPORTED')

                except:
                    print('File not uploaded.')

    return File_imported, previous_school, previous_catalog


def vital_source_login(my_username, my_password):
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
