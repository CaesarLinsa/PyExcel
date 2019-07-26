from util import args
from Client import  Client

@args('-f', '--file', metavar='<FILE>', help="xlsx file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', help="xlsx file sheet name")
def do_create_excel(args):
    """create a excel file"""
    cc = Client(args.file, args.sheetname)
    cc.save(args.file)

@args('-f', '--file', metavar='<FILE>', help="xlsx file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', help="xlsx file sheet name")
@args('-r','--row',metavar='<ROW>', help='a row of data')
def do_insert_row(args):
    """insert data into Excel for a row"""
    cc = Client(args.file, args.sheetname)
    cc.insert_row(args.row)
    cc.save(args.file)

@args('-f', '--file', metavar='<FILE>', help="xlsx file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', help="xlsx file sheet name")
def do_get_head(args):
    """get the excel first line"""
    cc = Client(args.file, args.sheetname)
    row=cc.get_head()
    for cell in row:
        print "%s " %cell.value,

@args('-f', '--file', metavar='<FILE>', help="xlsx file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', help="xlsx file sheet name")
@args('-col', '--col', metavar='<COL>', help="index of column")
def do_get_col(args):
    """ get a column data of Excel"""
    cc = Client(args.file, args.sheetname)
    data = cc.get_col(int(args.col))
    for d in data:
        print d.value
    
