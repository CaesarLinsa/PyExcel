from  openpyxl import Workbook
import openpyxl
import os

class Client(object):
    def __init__(self, wb, sheet=None, rows=None, cols=None):
        if not wb.endswith(".xlsx"):
            wb = "%s.xlsx" %wb
        if os.path.exists(wb):
            self.wb = openpyxl.load_workbook(wb)
        else:
            self.wb = Workbook()
        if sheet in self.wb.sheetnames:
            self.sheet = self.wb[sheet]
        else:
            self.sheet = self.wb.create_sheet(sheet, 0)
        self.rows = rows
        self.cols = cols

    def save(self, file):
        if not file.endswith(".xlsx"):
            file = "%s.xlsx" % file
        self.wb.save(file)

    def insert_row(self, stringrow):
        row=stringrow.split(',')
        self.sheet.append(row)

    def get_head(self):
        rows=list(self.sheet.iter_rows())
        if len(rows) == 0:
            return []
        return rows[0]
    
    def get_all(self):
        return list(self.sheet.iter_rows())

    def get_col(self, num):
        datas = list(self.sheet.iter_rows())
        col = []
        for data in datas:
            col.append(data[num-1])
        return col

    def get_sheet_names(self):
        return [ sheetname for sheetname in self.wb.sheetnames if not
        sheetname.startswith("Sheet")]
