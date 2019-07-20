from  openpyxl import Workbook
import openpyxl
import os

class Client(object):
    def __init__(self, wb, sheet=None, rows=None, cols=None):
        if os.path.exists(wb):
            self.wb = openpyxl.load_workbook(wb)
        else:
            self.wb = Workbook()
        if sheet in self.wb.sheetnames:
            self.sheet = self.wb[sheet]
        else:
            self.wb.create_sheet(sheet, 0)
        self.rows = rows
        self.cols = cols

    def save(self, file):
        self.wb.save(file)
