import unittest
import mock
from PyExcel.Client import Client
import contextlib2 as contextlib

class TestClient(unittest.TestCase):
    
    client_kargs = {
          'wb': 'caesar'
     }
    def setUp(self):
        with mock.patch("PyExcel.Client.openpyxl"):
            with mock.patch("PyExcel.Client.Workbook"):
                self.client = Client(**self.client_kargs)
    
    def test_save(self):
        file = "caesar.xlsx"
        self.client.save(file)
        self.client.wb.save.assert_called_with(file)
        file = "caesar"
        self.client.save(file)
        self.client.wb.save.assert_called_with("%s.xlsx" % file)
    
    def test_insert_row(self):
        content = "caesar,kafka"
        self.client.insert_row(content)
        self.client.sheet.append.assert_called_with(['caesar', 'kafka'])

    def test_get_head(self):
        rows=tuple()
        with contextlib.ExitStack() as stack:
            stack.enter_context(mock.patch.object(self.client.sheet, 'iter_rows', return_value=rows))
            ret = self.client.get_head()
            self.assertEqual(len(ret), 0)
        rows=((2, 3, 4),(4, 5, 6))
        with contextlib.ExitStack() as stack:
            stack.enter_context(mock.patch.object(self.client.sheet, 'iter_rows', return_value=rows))
            ret = self.client.get_head()
            self.assertEqual(len(ret), 3)
            self.assertEqual(ret, (2, 3, 4))

    def test_get_col(self):
        rows=((2, 3, 4),(4, 5, 6))
        with contextlib.ExitStack() as stack:
            stack.enter_context(mock.patch.object(self.client.sheet, 'iter_rows', return_value=rows))
            ret = self.client.get_col(1)
            self.assertEqual(len(ret), 2)
            self.assertEqual([2,4], ret)
   
    def test_get_all(self):
        rows=((2, 3, 4),(4, 5, 6))
        with contextlib.ExitStack() as stack:
            stack.enter_context(mock.patch.object(self.client.sheet, 'iter_rows', return_value=rows))
            ret = self.client.get_all()
            self.assertEqual([(2,3,4),(4,5,6)], ret)

    def test_get_sheet_names(self):
        self.client.get_sheet_names()
        self.client.wb.get_sheet_names.assert_called
