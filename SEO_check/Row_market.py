import pandas as pd
import unittest

class Data_Mismatch(ValueError):
    pass

class row_market():

    def __init__(self,driver,writer):
        self.driver=driver
        self.writer=writer

        #locating elements
        self.cano_xpath="canonical"
        self.meta_key_xpath="keywords"
        self.meta_desc_xpath="description"

    def cano_check(self):

        self.cano_tag = self.driver.find_elements_by_id(self.cano_xpath)
        for elements in self.cano_tag:
            cano_url = elements.get_attribute('href')
            if self.driver.current_url==cano_url:
                print("Canonical tag is correct")
            else:
                print("Error found: " + cano_url)
                raise Data_Mismatch(cano_url)


    def meta_key_check(self,key_from_sheet):
        self.key_from_sheet=key_from_sheet
        self.meta_key=self.driver.find_elements_by_name(self.meta_key_xpath)
        for elements in self.meta_key:
            self.keywords = elements.get_attribute('content')
            if self.keywords==self.key_from_sheet:
                output="Meta keywords are correct"
                print(output)
                df1=pd.DataFrame({'Status':[output]})
                df1.to_excel(self.writer, sheet_name='ROW', header='Status', index=False)
                #df.to_excel(self.writer, sheet_name='ROW', header='Status', index=False)
            else:
                output="Error found: "+self.key_from_sheet
                print(output)
                df2 = pd.DataFrame({'Status':[output]})
                df2.to_excel(self.writer, sheet_name='ROW', header='Status', index=False)
        self.writer.save()


    def meta_desc_check(self):
        self.meta_desc=self.driver.find_elements_by_name(self.meta_desc_xpath)
        for elements in self.meta_desc:
            description = elements.get_attribute('content')
            print(description)

