from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import pytest


class Test_GTScan():

    @pytest.fixture()
    def test_setup(self):
        #initiating chrome driver
        chrome_options=Options()
        chrome_options.add_argument('--start-maximized')
        global driver
        driver=webdriver.Chrome(
            "D:/OneDrive - CACTUS/Python/Sel_python/drivers/chromedriver.exe", chrome_options=chrome_options)
        driver.implicitly_wait(5)

        #initiating file reader
        file=r'D:/OneDrive - CACTUS/Python/Sel_python/GTmetrix/Scan_data.xlsx'
        global df, URLs, Server, Browser, Testsite, Username, Password
        df=pd.read_excel(file,sheet_name='Data_02')
        URLs=df['Urls']
        Server=df['Server']
        Browser=df['Browser']
        Testsite=df['Test server']
        Username=df['Username']
        Password=df['Password']

        #initiating file writer
        global writer
        writer=pd.ExcelWriter(file, engine='openpyxl')
        book=load_workbook(file)
        writer.book=book
        writer.sheets=dict((ws.title,ws)for ws in book.worksheets)

        yield
        driver.close()
        driver.quit()
        print("Test completed")


    def test01_scanpage(self, test_setup):
        #reading from urls column
        col_count=0
        row_count=1

        url_list = len(URLs)
        print("Total urls in the sheet:" + str(url_list))

        # GTmetrix login
        driver.get("https://gtmetrix.com/")
        driver.find_element_by_xpath("//a[@class='js-auth-widget-link'][contains(text(),'Log In')]").click()
        driver.find_element_by_name("email").send_keys("libin.thomas@cactusglobal.com")
        driver.find_element_by_name("password").send_keys("L!b!n20O4")
        driver.find_element_by_xpath("//button[contains(text(),'Log In')]").click()
        wait = WebDriverWait(driver, 300)
        wait.until(EC.presence_of_element_located((By.XPATH, "//h1[@class='page-heading']")))
        driver.implicitly_wait(10)

        while col_count<url_list:
            driver.find_element_by_xpath("//div[@class='header-content clear']//i[@class='sprite-gtmetrix sprite-display-block']").click()

            # Passing urls
            page_url = URLs[col_count]
            driver.find_element_by_name("url").send_keys(page_url)
            print("="*15)
            print("Selected webpage: "+page_url)

            # Server selection
            country = Server[col_count]
            print("Server location: "+country)
            if country == 'India':
                cn_value = '5'
            elif country == 'China':
                cn_value = '7'
            elif country == 'UK':
                cn_value = '2'
            elif country == 'Canada':
                cn_value = '1'
            # print("Value"+cn_value)

            # Browser selection
            browser_option = Browser[col_count]
            print("Browser selected: "+browser_option)
            if browser_option == 'Chrome':
                br_value = '3'
            elif browser_option == 'Firefox':
                br_value = '1'
            # print(br_value)

            #Staging site check
            Staging_server= Testsite[col_count]
            print(Staging_server)

            #fecthing credentials
            User=Username[col_count]
            Pswd=Password[col_count]

            #Conditions
            driver.find_element_by_xpath("//a[@class='btn analyze-form-options-trigger']").click()
            #select country
            select_server=Select(driver.find_element_by_id("af-region"))
            select_server.select_by_value(cn_value)
            #select browser
            select_browser=Select(driver.find_element_by_id("af-browser"))
            select_browser.select_by_value(br_value)
            #entering credentials
            if Staging_server == 'Yes':
                driver.find_element_by_xpath("//a[@id='analyze-form-advanced-options-trigger']").click()
                user_field=driver.find_element_by_xpath("//input[@id='af-username']")
                user_field.click()
                user_field.send_keys(User)
                pswd_field=driver.find_element_by_xpath("//input[@id='af-password']")
                pswd_field.click()
                pswd_field.send_keys(Pswd)

            #Submit
            driver.find_element_by_xpath("//button[contains(text(),'Analyze')]").click()
            wait.until(EC.presence_of_element_located((By.XPATH,"//h1[contains(text(),'Latest Performance Report for:')]")))

            # Saving screenshot
            correction=page_url.replace("/","")
            ss_name=correction.replace(":","")
            # print(correction)
            # print(ss_name)
            driver.implicitly_wait(10)
            driver.save_screenshot('D:/OneDrive - CACTUS/Python/Sel_python/GTmetrix/Screenshots/{}.png'.format(ss_name))
            driver.implicitly_wait(10)

            #Recording page score
            pagescore=driver.find_element_by_xpath(
                "/html/body/div[1]/main/article/div[2]/div[1]/div/div[1]/span/span").text
            print("Pagespeed score of webpage is: "+pagescore)
            df1=pd.DataFrame({'Page score':[pagescore]})
            df1.to_excel(writer, sheet_name='Data_02',index=False, header=None, startcol=3, startrow=row_count)
            writer.save()

            #Recording page grade
            page_grade = driver.find_element_by_xpath(
                "/html/body/div[1]/main/article/div[2]/div[1]/div/div[1]/span/i").get_attribute('class')
            Grade=page_grade.lstrip('sprite-grade-')
            print("Page speed grade is: " + Grade)
            df2 = pd.DataFrame({'Page grade': [Grade]})
            df2.to_excel(writer, sheet_name='Data_02', index=False, header=None, startcol=4, startrow=row_count)
            writer.save()

            #Recording load time
            loaded_time=driver.find_element_by_xpath(
                "/html/body/div[1]/main/article/div[2]/div[2]/div/div[1]/span").text
            print("Fully loaded time is:"+loaded_time)
            df3=pd.DataFrame({'Load time':[loaded_time]})
            df3.to_excel(writer, sheet_name='Data_02',index=False, header=None, startcol=5, startrow=row_count)
            writer.save()

            #Recording page size
            page_size=driver.find_element_by_xpath(
                "/html[1]/body[1]/div[1]/main[1]/article[1]/div[2]/div[2]/div[1]/div[2]/span[1]").text
            print("Total page size is: "+page_size)
            df4=pd.DataFrame({'Page size':[page_size]})
            df4.to_excel(writer, sheet_name='Data_02', index=False, header=None, startcol=6, startrow=row_count)
            writer.save()

            #Additional data
            yslowscore = driver.find_element_by_xpath(
                "/html[1]/body[1]/div[1]/main[1]/article[1]/div[2]/div[1]/div[1]/div[2]/span[1]/span[1]").text
            print("The Yslow score of webpage is: " + yslowscore)

            yslow_grade=driver.find_element_by_xpath(
                "/html/body/div[1]/main/article/div[2]/div[1]/div/div[2]/span/i").get_attribute('class')
            print("Yslow grade is: "+yslow_grade.lstrip('sprite-grade-'))

            #incrementing cell positions
            row_count=row_count+1
            col_count=col_count+1

    def test02_final_report(self, test_setup):

        #GTmetrix login
        driver.get("https://gtmetrix.com/")
        driver.find_element_by_xpath("//a[@class='js-auth-widget-link'][contains(text(),'Log In')]").click()
        driver.find_element_by_name("email").send_keys("libin.thomas@cactusglobal.com")
        driver.find_element_by_name("password").send_keys("L!b!n20O4")
        driver.find_element_by_xpath("//button[contains(text(),'Log In')]").click()
        wait = WebDriverWait(driver, 300)
        wait.until(EC.presence_of_element_located((By.XPATH, "//h1[@class='page-heading']")))

        driver.find_element_by_xpath(
            "//div[@class='header-content clear']//i[@class='sprite-gtmetrix sprite-display-block']").click()

        #Saving screenshot
        time.sleep(2)
        ele=driver.find_element_by_xpath("//a[@class='paginate_button next']")
        actions=ActionChains(driver)
        actions.move_to_element(ele).perform()
        time.sleep(2)
        driver.save_screenshot( 'D:/OneDrive - CACTUS/Python/Sel_python/GTmetrix/Screenshots/Traversed pages.png')

