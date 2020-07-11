from selenium import webdriver
import unittest
import pandas as pd
import HtmlTestRunner
from openpyxl import load_workbook
import sys
sys.path.append("D:/OneDrive - CACTUS/Python/Sel_python")

class Meta_keywords_check(unittest.TestCase):

    @classmethod
    def setUpClass(cls) -> None:
        #initiating chrome driver
        cls.driver=webdriver.Chrome(executable_path='D:/OneDrive - CACTUS/Python/Sel_python/drivers/chromedriver.exe')
        cls.driver.maximize_window()
        cls.driver.implicitly_wait(10)

        #initaiting file reader
        excel_file= r'D:/OneDrive - CACTUS/Python/Sel_python/SEO_check/JPN/SEO_metakeys_data.xlsx'
        cls.df=pd.read_excel(excel_file,sheet_name='Meta_keys')
        cls.Urls=cls.df['URLs']
        cls.Meta_keys=cls.df['Meta keywords']

        #initiating file writer
        cls.writer=pd.ExcelWriter(excel_file, engine='openpyxl')
        cls.book=load_workbook(excel_file)
        cls.writer.book=cls.book
        cls.writer.sheets=dict((ws.title,ws)for ws in cls.book.worksheets)

    def test_meta_key_verify(self):
        i = 0
        j = (len(self.Urls))
        k = 0
        row_num=1
        print("Total Entries in sheet:", j)
        while i < j and k < j:
            page_url = self.Urls[i]
            self.driver.get(page_url)
            print(page_url)
            keywords_data = self.Meta_keys[k]
            keywords=self.driver.find_elements_by_xpath("//meta[@name='Keywords']|//meta[@name='keywords']")
            for elements in keywords:
                site_keywords=elements.get_attribute('content')
                if site_keywords==keywords_data:
                    output = "Meta keywords are correct"
                    print(output)
                    df1 = pd.DataFrame({'Status': [output]})
                    df1.to_excel(self.writer, sheet_name='Meta_keys', header=None, index=False, startrow=row_num,
                                 startcol=2)
                    self.writer.save()
                    row_num = row_num + 1
                else:
                    output = "Keyword mismatch:"+site_keywords
                    print(output)
                    df1 = pd.DataFrame({'Status': [output]})
                    df1.to_excel(self.writer, sheet_name='Meta_keys', header=None, index=False, startrow=row_num,
                                 startcol=2)
                    self.writer.save()
                    row_num = row_num + 1

            k = k + 1
            i = i + 1

    @classmethod
    def tearDownClass(cls) -> None:
        cls.driver.close()
        cls.driver.quit()
        print("Test execution completed")

if __name__=='__main__':
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(output="..//SEO_check//Reports"))



