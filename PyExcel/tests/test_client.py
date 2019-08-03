import unittest
import mock
from PyExcel.Client import Client
import contextlib2 as contextlib

class TestClient(unittest.TestCase):
    
    client_kargs = {
          'wb': 'caesar'
     }
    def setUp(self):
        with mock.patch("PyExcel.Client.openpyxl") as mock_openpyxl:
            with mock.patch("PyExcel.Client.Workbook") as mock_wb:
                self.client = Client(**self.client_kargs)
        
    def test_save(self):
        file='caesar'
        f_save = mock.Mock(side_effect=Client.save)
        f_save(self.client, file)
        f_save.assert_called_with(self.client, file)

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
