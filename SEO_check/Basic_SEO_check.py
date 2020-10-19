from selenium import webdriver
import unittest
import pandas as pd
import sys
sys.path.append("D:/OneDrive - CACTUS/Python/Sel_python")
from SEO_check.Row_market import row_market
import HtmlTestRunner
from openpyxl import load_workbook


class Basic_Seo_check(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        #initating data file and calling Urls
        excel_file=r'D:/OneDrive - CACTUS/Python/Sel_python/SEO_check/SEO_data.xlsx'
        cls.df = pd.read_excel(excel_file, sheet_name='ROW')
        cls.Urls=cls.df['URLs']
        cls.Meta_keys=cls.df['Meta keywords']

        #initiating pandas writer
        cls.writer=pd.ExcelWriter(excel_file,engine='openpyxl')
        cls.book=load_workbook(excel_file)
        cls.writer.book=cls.book
        #writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

        #initiating webdriver
        cls.driver = webdriver.Chrome(executable_path='D:/OneDrive - CACTUS/Python/Sel_python/drivers/chromedriver.exe')
        cls.driver.implicitly_wait(10)
        cls.driver.maximize_window()

    def test_canonical_verification(self):
        #initiating browser
        driver=self.driver
        writer = self.writer

        #initiating canonical tag check
        i=0
        j=(len(self.Urls))
        print("Total Entries in sheet:",j)
        while i<j:
            page_url=self.Urls[i]
            driver.get(page_url)
            print(page_url)
            row_page_check = row_market(driver,writer)
            row_page_check.cano_check()
            i=i+1
            self.assertEqual()

    def test_meta_desc_verify(self):
        # initiating browser
        driver = self.driver
        writer = self.writer

        #initiating meta keywords check
        i = 0
        j = (len(self.Urls))
        k=0
        print("Total Entries in sheet:",j)
        while i < j and k < j:
            page_url = self.Urls[i]
            driver.get(page_url)
            print(page_url)
            keywords=self.Meta_keys[k]
            #print("Data from sheet:"+keywords)
            row_page_check = row_market(driver, writer)
            writer.sheets=dict((ws.title,ws)for ws in self.book.worksheets)
            row_page_check.meta_key_check(keywords)
            #df.to_excel(self.writer, sheet_name='ROW', header='Status', index=False, startcol=3)
            #df2.to_excel(writer, sheet_name='Sheet1', header=None, index=False,
            #             startcol=7, startrow=6)
            k = k + 1
            i = i + 1
        # row_page_check.meta_desc_check()

    @classmethod
    def tearDownClass(cls):
        cls.driver.close()
        cls.driver.quit()
        print("Test execution completed")


if __name__ =='__main__':
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(output='..//Test files//Reports'))



