from selenium import webdriver
import unittest
import pandas as pd
import HtmlTestRunner
from openpyxl import load_workbook
import sys
sys.path.append("D:/OneDrive - CACTUS/Python/Sel_python")

class Meta_descrip_check(unittest.TestCase):

    @classmethod
    def setUpClass(cls) -> None:
        #initiating chrome driver
        cls.driver=webdriver.Chrome(executable_path='D:/OneDrive - CACTUS/Python/Sel_python/drivers/chromedriver.exe')
        cls.driver.maximize_window()
        cls.driver.implicitly_wait(10)

        #initaiting file reader
        excel_file= r'D:/OneDrive - CACTUS/Python/Sel_python/SEO_check/KOR/SEO_metadesc_data.xlsx'
        cls.df=pd.read_excel(excel_file,sheet_name='Meta_desc')
        cls.Urls=cls.df['URLs']
        cls.Meta_desc=cls.df['Meta description']

        #initiating file writer
        cls.writer=pd.ExcelWriter(excel_file, engine='openpyxl')
        cls.book=load_workbook(excel_file)
        cls.writer.book=cls.book
        cls.writer.sheets=dict((ws.title,ws)for ws in cls.book.worksheets)

    def test_meta_desc_verify(self):
        i = 0
        j = (len(self.Urls))
        k = 0
        row_num=1
        print("Total Entries in sheet:", j)
        while i < j and k < j:
            page_url = self.Urls[i]
            self.driver.get(page_url)
            print(page_url)
            desc_data = self.Meta_desc[k]
            descrip=self.driver.find_elements_by_xpath("//meta[@name='description']|//meta[@name='Description']")
            for elements in descrip:
                site_descrip=elements.get_attribute('content')
                if site_descrip==desc_data:
                    output = "Meta description is correct"
                    print(output)
                    df1 = pd.DataFrame({'Status': [output]})
                    df1.to_excel(self.writer, sheet_name='Meta_desc', header=None, index=False, startrow=row_num,
                                 startcol=2)
                    self.writer.save()
                    row_num = row_num + 1
                else:
                    output = "Description mismatch:" + site_descrip
                    print(output)
                    df1 = pd.DataFrame({'Status': [output]})
                    df1.to_excel(self.writer, sheet_name='Meta_desc', header=None, index=False, startrow=row_num,
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
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(output="..//SEO_check\/Reports"))



