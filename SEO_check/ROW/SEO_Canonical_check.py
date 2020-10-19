from selenium import webdriver
import unittest
import pandas as pd
import HtmlTestRunner
from openpyxl import load_workbook
import sys
sys.path.append("D:/OneDrive - CACTUS/Python/Sel_python")

class Canonical_check(unittest.TestCase):

    @classmethod
    def setUpClass(cls) -> None:
        #initiating chrome driver
        cls.driver=webdriver.Chrome(executable_path='D:/OneDrive - CACTUS/Python/Sel_python/drivers/chromedriver.exe')
        cls.driver.maximize_window()
        cls.driver.implicitly_wait(5)

        #initaiting file reader
        excel_file= r'D:/OneDrive - CACTUS/Python/Sel_python/SEO_check/ROW/SEO_cano_data.xlsx'
        cls.df=pd.read_excel(excel_file,sheet_name='Canonical')
        cls.Urls=cls.df['URLs']

        #initiating file writer
        cls.writer=pd.ExcelWriter(excel_file, engine='openpyxl')
        cls.book=load_workbook(excel_file)
        cls.writer.book=cls.book
        cls.writer.sheets=dict((ws.title,ws)for ws in cls.book.worksheets)

    def test01_cano_verify(self):
        col_count=0
        list_count=len(self.Urls)
        print("Total entries in the sheet: ",list_count)
        row_count=1
        while col_count<list_count:
            page_url=self.Urls[col_count]
            self.driver.get(page_url)
            print(self.Urls[col_count])
            if self.driver.find_elements_by_xpath("//link[@rel='canonical']"):
                self.cano_tag = self.driver.find_elements_by_xpath("//link[@rel='canonical']")
                for elements in self.cano_tag:
                    cano_tag = elements.get_attribute('href')
                    if self.driver.current_url == cano_tag:
                        output="Canonical tag is correct"
                        print(output)
                        df1=pd.DataFrame({'Status':[output]})
                        df1.to_excel(self.writer, sheet_name='Canonical', header=None, index=False,startrow=row_count,startcol=1)
                        self.writer.save()
                        row_count = row_count + 1
                    else:
                        output = "Error found: "+cano_tag
                        print(output)
                        df2 = pd.DataFrame({'Status': [output]})
                        df2.to_excel(self.writer, sheet_name='Canonical', header=None, index=False, startrow=row_count,startcol=1)
                        self.writer.save()
                        row_count = row_count + 1
            col_count=col_count+1

    @classmethod
    def tearDownClass(cls) -> None:
        cls.driver.close()
        cls.driver.quit()
        print("Test execution completed")

if __name__=='__main__':
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(
        output="D:/OneDrive - CACTUS/Python/Sel_python/SEO_check/ROW"))



