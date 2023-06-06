from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService 
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import undetected_chromedriver as uc
import time
import os
import re
from datetime import datetime
import pandas as pd
import warnings
import sys
import xlsxwriter
from multiprocessing import freeze_support
import calendar 
import shutil
warnings.filterwarnings('ignore')

def initialize_bot(translate):

    # Setting up chrome driver for the bot
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument('--headless')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    # installing the chrome driver
    driver_path = ChromeDriverManager().install()
    chrome_service = ChromeService(driver_path)
    # configuring the driver
    driver = webdriver.Chrome(options=chrome_options, service=chrome_service)
    ver = int(driver.capabilities['chrome']['chromedriverVersion'].split('.')[0])
    driver.quit()
    chrome_options = uc.ChromeOptions()
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--lang=en")
    #chrome_options.add_argument("--incognito")
    chrome_options.add_argument('--headless=new')
    
    # disable location prompts & disable images loading
    if not translate:
        prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 2}  
        chrome_options.page_load_strategy = 'eager'
    else:
        prefs = {"profile.default_content_setting_values.geolocation": 2, "profile.managed_default_content_settings.images": 2, "profile.managed_default_content_settings.notifications": 1, "translate_whitelists": {"zh-TW":"en"},"translate":{"enabled":"true"}, "profile.default_content_setting_values.cookies": 1}
        chrome_options.page_load_strategy = 'normal'

    chrome_options.add_experimental_option("prefs", prefs)
    driver = uc.Chrome(version_main = ver, options=chrome_options) 
    driver.set_window_size(1920, 1080)
    driver.maximize_window()
    driver.set_page_load_timeout(20000)

    return driver

def google_translate(driver, text, source_lang, target_lang):

    driver.get(f"https://translate.google.com/?sl={source_lang}&tl={target_lang}&op=translate")

    # Type the text to be translated into the input box
    input_box = wait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//textarea[@aria-label='Source text']")))
    input_box.clear()
    time.sleep(1)
    input_box.send_keys(text)

    # Wait for the translation to appear and extract it
    translation = wait(driver, 5).until(EC.visibility_of_element_located((By.XPATH, "//span[@class='HwtZe']"))).text
     
    # Return the translated text
    return translation
    
def scrape_articles(driver, driver_tr, output1, page, month, year):

    stamp = datetime.now().strftime("%d_%m_%Y")
    print('-'*75)
    print(f'Scraping The Articles Links from: {page}')
    # getting the full posts list
    links = []
    months = {month: index for index, month in enumerate(calendar.month_abbr) if month}
    full_months = {month: index for index, month in enumerate(calendar.month_name) if month}
    prev_month = month - 1
    if prev_month == 0:
        prev_month = 12

    driver.get(page)
    art_time = ''
    # handling lazy loading
    print('-'*75)
    print("Getting the previous month's articles..." )
    done = False
    for _ in range(50):  
        try:
            height1 = driver.execute_script("return document.body.scrollHeight")
            driver.execute_script(f"window.scrollTo(0, {height1})")
            time.sleep(1)

            # scraping posts urls 
            try:
                posts = wait(driver, 4).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[class='i']")))    
            except:
                break

            for post in posts:
                try:
                    date = wait(post, 1).until(EC.presence_of_element_located((By.TAG_NAME, "i"))).get_attribute('textContent').split('/')[-1].strip() 
                    art_month = int(date.split('-')[1])
                    art_year = int(date.split('-')[0])
                    # for articles from previous year
                    if art_year < year and prev_month != 12:
                        done = True
                        break
                    # for all months except Jan
                    elif art_month < prev_month and prev_month != 12 and art_year == year:
                        done = True
                        break
                    # for Jan
                    elif art_month < prev_month and prev_month == 12 and art_year < year:
                        done = True
                        break
                    elif art_month > prev_month and art_year == year: 
                        continue
                    else:
                        link = wait(post, 1).until(EC.presence_of_element_located((By.TAG_NAME, "a"))).get_attribute('href')
                        if link not in links:
                            links.append(link)
                except:
                    pass

            if done:
                break

            # moving to the next page
            try:
                div = wait(driver, 1).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='plist']")))
                elems = wait(div, 1).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a")))
                for elem in elems:
                    if '下一页' in elem.get_attribute('textContent'):
                        driver.get(elem.get_attribute('href'))
                        time.sleep(1)
                        break                       
            except:
                break

        except Exception as err:
            break


    # scraping posts
    print('-'*75)
    print('Scraping Articles...')
    print('-'*75)

    # reading previously scraped data for duplication checking
    scraped = []
    try:
        df = pd.read_excel(output1)
        scraped = df['unique_id'].values.tolist()
    except:
        pass

    n = len(links)
    data = pd.DataFrame()
    for i, link in enumerate(links):
        try:
            driver.get(link)   
        except:
            print(f'Warning: Failed to load the url: {link}')
            continue

        art_id = ''
        try:
            art_id = link.split('.cn')[-1].replace('.html', '')[1:]
        except:
            pass
        if art_id in scraped: 
            print(f'Article {i+1}\{n} is already scraped, skipping.')
            continue        
        
        if art_id == '': 
            print(f'Warning: Article {i+1}\{n} has unknown ID')
            art_id = 0

        # scrolling across the page for auto translation to be applied
        try:
            total_height = driver.execute_script("return document.body.scrollHeight")
            height = total_height/30
            new_height = 0
            for _ in range(30):
                prev_hight = new_height
                new_height += height             
                driver.execute_script(f"window.scrollTo({prev_hight}, {new_height})")
                time.sleep(0.1)
        except:
            pass

        row = {}
        # English article author and date
        en_author, date = '', ''             
        try:
            info = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='info']")))
            text = info.get_attribute('innerHTML')
            elems = text.replace('/', '').split('<strong>')
            for j, elem in enumerate(elems):
                if '责任编辑' in elem:
                    en_author = elems[j+1]
                elif '时间' in elem:
                    date = elems[j+1]
            date = re.findall('\d+', date)
        except Exception as err:
            pass
            
        # checking if the article date is correct
        try:
            art_month = int(date[1])
            art_year = int(date[0])  
            art_day = int(date[2])       
            if art_month != prev_month: 
                print(f'skipping article with date {art_month}/{art_day}/{art_year}')
                continue
            date = f'{art_day}_{art_month}_{art_year}'
        except:
            continue    

        row['sku'] = art_id
        row['unique_id'] = art_id
        row['articleurl'] = link

        # Chinese article title
        title = ''             
        try:
            title = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='title']"))).get_attribute('textContent').strip()
        except:
            continue               
                
        row['articletitle'] = title            

        # Chinese article description
        des = ''             
        try:
            div = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='content']")))
            elems = wait(div, 4).until(EC.presence_of_all_elements_located((By.TAG_NAME, "p")))
            for elem in elems:
                try:
                    des += elem.get_attribute('textContent').replace('\xa0', '').strip('\n').strip(' ') + '\n'
                except:
                    pass
        except:
            continue               
                
        row['articledescription'] = des.strip('\n')
                     
        # English article title
        for _ in range(3):
            title_en = google_translate(driver_tr, title, "zh-TW", "en")
            if len(title_en) > 10: break           
                        
        asian = re.findall(r'[\u3131-\ucb4c]+',title_en)
        if asian: continue

        row['articletitle in English'] = title_en 
                    
        ## English article description
        for _ in range(3):
            des_en = google_translate(driver_tr, des, "zh-TW", "en")
            if len(des_en) > 10: break           
                   
        asian = re.findall(r'[\u3131-\ucb4c]+',des_en)
        if asian: continue  
                    
        row['articledescription in English'] = des_en
        row['articleauthor'] = en_author
        row['articledatetime'] = date            
            
        # article category
        cat = ''             
        try:
            div = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='placenav']")))
            cat = wait(div, 4).until(EC.presence_of_all_elements_located((By.TAG_NAME, "a")))[-1].get_attribute('textContent').strip()
        except:
            pass 
            
        row['articlecategory'] = cat

        # other columns
        row['domain'] = 'Ladymax'
        row['hype'] = ''   
        row['articletags'] = ''
        row['articleheader'] = ''

        imgs = ''
        try:
            div = wait(driver, 4).until(EC.presence_of_element_located((By.CSS_SELECTOR, "div[class='content']")))
            elems = wait(div, 4).until(EC.presence_of_all_elements_located((By.TAG_NAME, "img")))
            for elem in elems:
                try:
                    imgs += elem.get_attribute('src') + ', '
                except:
                    pass
            imgs = imgs.strip(', ')
        except:
            pass

        if imgs == '': continue
        row['articleimages'] = imgs
        row['articlecomment'] = ''
        row['Extraction Date'] = stamp
        # appending the output to the datafame       
        data = pd.concat([data, pd.DataFrame([row.copy()])], ignore_index=True)
        print(f'Scraping the details of article {i+1}\{n}')
           
    # output to excel
    if data.shape[0] > 0:
        data['articledatetime'] = pd.to_datetime(data['articledatetime'],  errors='coerce', format="%d_%m_%Y")
        data['articledatetime'] = data['articledatetime'].dt.date  
        data['Extraction Date'] = pd.to_datetime(data['Extraction Date'],  errors='coerce', format="%d_%m_%Y")
        data['Extraction Date'] = data['Extraction Date'].dt.date   
        df1 = pd.read_excel(output1)
        if df1.shape[0] > 0:
            df1[['articledatetime', 'Extraction Date']] = df1[['articledatetime', 'Extraction Date']].apply(pd.to_datetime,  errors='coerce', format="%Y-%m-%d")
            df1['articledatetime'] = df1['articledatetime'].dt.date 
            df1['Extraction Date'] = df1['Extraction Date'].dt.date 
        df1 = pd.concat([df1, data], ignore_index=True)
        df1 = df1.drop_duplicates()
        writer = pd.ExcelWriter(output1, date_format='d/m/yyyy')
        df1.to_excel(writer, index=False)
        writer.close()
    else:
        print('-'*75)
        print('No New Articles Found')
        
def get_inputs():
 
    print('-'*75)
    print('Processing The Settings Sheet ...')
    # assuming the inputs to be in the same script directory
    path = os.getcwd()
    if '\\' in path:
        path += '\\Ladymax_settings.xlsx'
    else:
        path += '/Ladymax_settings.xlsx'

    if not os.path.isfile(path):
        print('Error: Missing the settings file "Ladymax_settings.xlsx"')
        input('Press any key to exit')
        sys.exit(1)
    try:
        settings = {}
        urls = []
        df = pd.read_excel(path)
        cols  = df.columns
        for col in cols:
            df[col] = df[col].astype(str)

        inds = df.index
        for ind in inds:
            row = df.iloc[ind]
            link, status = '', ''
            for col in cols:
                if row[col] == 'nan': continue
                elif col == 'Category Link':
                    link = row[col]
                elif col == 'Scrape':
                    status = row[col]
                else:
                    settings[col] = row[col]

            if link != '' and status != '':
                try:
                    status = int(status)
                    urls.append((link, status))
                except:
                    urls.append((link, 0))
    except:
        print('Error: Failed to process the settings sheet')
        input('Press any key to exit')
        sys.exit(1)

    return settings, urls

def initialize_output():

    stamp = datetime.now().strftime("%d_%m_%Y_%H_%M")
    path = os.getcwd() + '\\Scraped_Data\\' + stamp
    if os.path.exists(path):
        shutil.rmtree(path)
    os.makedirs(path)

    file1 = f'Ladymax_{stamp}.xlsx'

    # Windws and Linux slashes
    if os.getcwd().find('/') != -1:
        output1 = path.replace('\\', '/') + "/" + file1
    else:
        output1 = path + "\\" + file1  

    # Create an new Excel file and add a worksheet.
    workbook1 = xlsxwriter.Workbook(output1)
    workbook1.add_worksheet()
    workbook1.close()    

    return output1

def main():

    print('Initializing The Bot ...')
    freeze_support()
    start = time.time()
    output1 = initialize_output()
    settings, urls = get_inputs()
    month = datetime.now().month
    year = datetime.now().year
    try:
        driver = initialize_bot(False)
    except Exception as err:
        print('Failed to initialize the Chrome driver due to the following error:\n')
        print(str(err))
        print('-'*75)
        input('Press any key to exit.')
        sys.exit()

    #initializing the english translator
    try:
        driver_tr = initialize_bot(False)
    except Exception as err:
        print('Failed to load the translation page due to the below error ')
        print(str(err))
        input("press any key to exit")
        sys.exit()

    for url in urls:
        if url[1] == 0: continue
        link = url[0]
        try:
            scrape_articles(driver, driver_tr, output1, link, month, year)
        except Exception as err: 
            print(f'Warning: the below error occurred:\n {err}')
            driver.quit()
            time.sleep(2)
            driver = initialize_bot(False)

    driver.quit()
    driver_tr.quit()
    print('-'*75)
    elapsed_time = round(((time.time() - start)/60), 4)
    input(f'Process is completed in {elapsed_time} mins, Press any key to exit.')
    sys.exit()

if __name__ == '__main__':

    main()
