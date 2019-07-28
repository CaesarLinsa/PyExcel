from util import args
from Client import  Client
import prettytable


@args('-f', '--file', metavar='<FILE>', help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', help="Excel file sheet name")
def do_create_excel(args):
    """create a excel file"""
    cc = Client(args.file, args.sheetname)
    cc.save(args.file)

@args('-f', '--file', metavar='<FILE>', help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', help="Excel file sheet name")
@args('-r','--row',metavar='<ROW>', help='a row of data')
def do_insert_row(args):
    """insert data into Excel for a row"""
    cc = Client(args.file, args.sheetname)
    cc.insert_row(args.row)
    cc.save(args.file)

@args('-f', '--file', metavar='<FILE>', help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', help="Excel file sheet name")
def do_get_head(args):
    """get the excel first line"""
    cc = Client(args.file, args.sheetname)
    row=cc.get_head()
    row_list = []
    for cell in row:
        row_list.append(cell.value)
    pt = prettytable.PrettyTable(row_list)
    pt.align = 'l'
    print pt

@args('-f', '--file', metavar='<FILE>', help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', help="Excel file sheet name")
@args('-col', '--col', metavar='<COL>', type=int, help="index of column")
def do_get_col(args):
    """ get a column data of Excel"""
    cc = Client(args.file, args.sheetname)
    data = cc.get_col(args.col)
    col_list = []
    for d in data:
        col_list.append(d.value)
    pt = prettytable.PrettyTable()
    pt.align = 'l'
    pt.add_column(col_list[0], col_list[1:])
    print pt

@args('-f', '--file', metavar='<FILE>', help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', help="Excel file sheetname")
@args('-col', '--col', metavar='<COL>', type=int, help="index of column")
def do_get_all(args):
    cc = Client(args.file, args.sheetname)
    data = cc.get_all()
    pt = prettytable.PrettyTable()
    for i, row in enumerate(data):
        row_list = []
        for cell in row:
            row_list.append(cell.value)
        if i == 0:
            pt.field_names = row_list
        else:
            pt.add_row(row_list)
        pt.align = 'l'
    print pt
 
