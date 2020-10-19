from selenium import webdriver
import unittest
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook
import HtmlTestRunner
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

class GTscan(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        #initiating chrome driver
        chrome_options=Options()
        chrome_options.add_argument('--start-maximized')
        cls.driver=webdriver.Chrome(
            "D:/OneDrive - CACTUS/Python/Sel_python/drivers/chromedriver.exe", chrome_options=chrome_options)
        cls.driver.implicitly_wait(5)

        #initiating file reader
        file=r'D:/OneDrive - CACTUS/Python/Sel_python/GTmetrix/Scan_data.xlsx'
        cls.df=pd.read_excel(file,sheet_name='Data_02')
        cls.URLs=cls.df['Urls']
        cls.Server=cls.df['Server']
        cls.Browser=cls.df['Browser']

        #initiating file writer
        cls.writer=pd.ExcelWriter(file, engine='openpyxl')
        book=load_workbook(file)
        cls.writer.book=book
        cls.writer.sheets=dict((ws.title,ws)for ws in book.worksheets)

    def test01_scanpage(self):
        #reading from urls column
        col_count=0
        row_count=1

        url_list = len(self.URLs)
        print("Total urls in the sheet:" + str(url_list))

        # GTmetrix login
        self.driver.get("https://gtmetrix.com/")
        self.driver.find_element_by_xpath("//a[@class='js-auth-widget-link'][contains(text(),'Log In')]").click()
        self.driver.find_element_by_name("email").send_keys("libin.thomas@cactusglobal.com")
        self.driver.find_element_by_name("password").send_keys("L!b!n20O4")
        self.driver.find_element_by_xpath("//button[contains(text(),'Log In')]").click()
        wait = WebDriverWait(self.driver, 300)
        wait.until(EC.presence_of_element_located((By.XPATH, "//h1[@class='page-heading']")))
        self.driver.implicitly_wait(10)

        while col_count<url_list:
            self.driver.find_element_by_xpath("//div[@class='header-content clear']//i[@class='sprite-gtmetrix sprite-display-block']").click()

            # Passing urls
            page_url = self.URLs[col_count]
            self.driver.find_element_by_name("url").send_keys(page_url)

            # Server selection
            country = self.Server[col_count]
            print(country)
            if country == 'India':
                cn_value = '5'
            elif country == 'China':
                cn_value = '7'
            elif country == 'UK':
                cn_value = '2'
            elif country == 'Canada':
                cn_value = '1'
            print(cn_value)

            # Browser selection
            browser_option = self.Browser[col_count]
            print(browser_option)
            if browser_option == 'Chrome':
                br_value = '3'
            elif browser_option == 'Firefox':
                br_value = '1'
            print(br_value)

            #Conditions
            self.driver.find_element_by_xpath("//a[@class='btn analyze-form-options-trigger']").click()
            #select country
            select_server=Select(self.driver.find_element_by_id("af-region"))
            select_server.select_by_value(cn_value)
            #select browser
            select_browser=Select(self.driver.find_element_by_id("af-browser"))
            select_browser.select_by_value(br_value)

            #Submit
            self.driver.find_element_by_xpath("//button[contains(text(),'Analyze')]").click()
            wait.until(EC.presence_of_element_located((By.XPATH,"//h1[contains(text(),'Latest Performance Report for:')]")))

            # Saving screenshot
            correction=page_url.replace("/","")
            ss_name=correction.replace(":","")
            print(correction)
            print(ss_name)
            self.driver.implicitly_wait(10)
            self.driver.save_screenshot('D:/OneDrive - CACTUS/Python/Sel_python/GTmetrix/Screenshots/{}.png'.format(ss_name))
            self.driver.implicitly_wait(10)

            #Recording page score
            pagescore=self.driver.find_element_by_xpath(
                "/html/body/div[1]/main/article/div[2]/div[1]/div/div[1]/span/span").text
            print("Pagespeed score of webpage is: "+pagescore)
            df1=pd.DataFrame({'Page score':[pagescore]})
            df1.to_excel(self.writer, sheet_name='Data_02',index=False, header=None, startcol=3, startrow=row_count)
            self.writer.save()

            #Recording page grade
            page_grade = self.driver.find_element_by_xpath(
                "/html/body/div[1]/main/article/div[2]/div[1]/div/div[1]/span/i").get_attribute('class')
            Grade=page_grade.lstrip('sprite-grade-')
            print("Page speed grade is: " + Grade)
            df2 = pd.DataFrame({'Page grade': [Grade]})
            df2.to_excel(self.writer, sheet_name='Data_02', index=False, header=None, startcol=4, startrow=row_count)
            self.writer.save()

            #Recording load time
            loaded_time=self.driver.find_element_by_xpath(
                "/html/body/div[1]/main/article/div[2]/div[2]/div/div[1]/span").text
            print("Fully loaded time is:"+loaded_time)
            df3=pd.DataFrame({'Load time':[loaded_time]})
            df3.to_excel(self.writer, sheet_name='Data_02',index=False, header=None, startcol=5, startrow=row_count)
            self.writer.save()

            #Recording page size
            page_size=self.driver.find_element_by_xpath(
                "/html[1]/body[1]/div[1]/main[1]/article[1]/div[2]/div[2]/div[1]/div[2]/span[1]").text
            print("Total page size is: "+page_size)
            df4=pd.DataFrame({'Page size':[page_size]})
            df4.to_excel(self.writer, sheet_name='Data_02', index=False, header=None, startcol=6, startrow=row_count)
            self.writer.save()

            #Additional data
            yslowscore = self.driver.find_element_by_xpath(
                "/html[1]/body[1]/div[1]/main[1]/article[1]/div[2]/div[1]/div[1]/div[2]/span[1]/span[1]").text
            print("The Yslow score of webpage is: " + yslowscore)

            yslow_grade=self.driver.find_element_by_xpath(
                "/html/body/div[1]/main/article/div[2]/div[1]/div/div[2]/span/i").get_attribute('class')
            print("Yslow grade is: "+yslow_grade.lstrip('sprite-grade-'))

            #incrementing cell positions
            row_count=row_count+1
            col_count=col_count+1

    def test02_final_report(self):

        #GTmetrix login
        self.driver.get("https://gtmetrix.com/")
        self.driver.implicitly_wait(10)

        self.driver.find_element_by_xpath(
            "//div[@class='header-content clear']//i[@class='sprite-gtmetrix sprite-display-block']").click()

        #Saving screenshot
        time.sleep(2)
        # required_width = self.driver.execute_script('return document.body.parentNode.scrollWidth')
        # required_height = self.driver.execute_script('return document.body.parentNode.scrollHeight')
        # self.driver.set_window_size(required_width, required_height)


        ele=self.driver.find_element_by_xpath("//a[@class='paginate_button next']")
        actions=ActionChains(self.driver)
        actions.move_to_element(ele).perform()
        # self.driver.execute_script("arguments[0].scrollIntoView();", ele)
        time.sleep(2)
        self.driver.save_screenshot( 'D:/OneDrive - CACTUS/Python/Sel_python/GTmetrix/Screenshots/Traversed pages.png')

    @classmethod
    def tearDownClass(cls):
        # time.sleep(10)
        cls.driver.close()
        cls.driver.quit()
        print("Test completed")

if __name__=='__main__':
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(output='D:/OneDrive - CACTUS/Python/Sel_python/GTmetrix'))