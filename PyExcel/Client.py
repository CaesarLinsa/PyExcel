from  openpyxl import Workbook
import openpyxl
import os

class Client(object):
    def __init__(self, wb, sheet=None):
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
        self.rows = list(self.sheet.iter_rows())
        self.head_row = list(self.rows[0]) if self.rows else []

    def save(self, file):
        if not file.endswith(".xlsx"):
            file = "%s.xlsx" % file
        self.wb.save(file)

    def insert_row(self, stringrow):
        row=stringrow.split(',')
        self.sheet.append(row)

    def get_head(self):
        return self.head_row
    
    def get_all(self):
        return list(self.rows)

    def get_col(self, num):
        datas = list(self.rows)
        col = []
        for data in datas:
            col.append(data[num])
        return col

    def get_sheet_names(self):
        return [ sheetname for sheetname in self.wb.sheetnames if not
        sheetname.startswith("Sheet")]

    def get_col_id_by_name(self, col_name):
        for index, colname in enumerate(self.head_row):
            if col_name == colname.value:
                return index
   
    def get_sheet_id_by_name(self, sheet_name):
        for index, sheetname in enumerate(self.wb.sheetnames):
            if sheetname == sheet_name:
                return index

