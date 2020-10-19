
class data_File():

    def __init__(self, work_file):
        self.work_file=work_file

        #cls.df = openpyxl.load_workbook('D:/OneDrive - CACTUS/Python/Sel_python/SEO_check/SEO_data.xlsx')
    def cell(self,row,column):
        self.row_sheet = self.work_file["ROW"]
        self.url_cell = self.row_sheet.cell(row,column)
        self.url = self.url_cell.value
        #print(self.url)
