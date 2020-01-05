from util import args
from Client import  Client
import prettytable
import json

def pretty_print_all(data):
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


def raw_print_all(data):
    for i, row in enumerate(data):
        row_list = []
        for cell in row:
            row_list.append(cell.value)
        if i == 0:
            print(" ".join([ i for i in row_list if i ]))
        else:
            print(" ".join([ i for i in row_list if i ]))


def write_csv_data(data, csv_file):
    with open(csv_file, 'a+') as f:
        for i, row in enumerate(data):
            row_list = []
            for cell in row:
              row_list.append(cell.value)
            f.write("{}\n".format(",".join([ cell if cell else '' for cell in
            row_list])))


def pretty_print_row(data):
    pt = prettytable.PrettyTable(data)
    pt.align = 'l'
    print pt


def raw_print_row(data):
    print(" ".join([ i for i in data if i] ))


def pretty_print_col(col_list):
    pt = prettytable.PrettyTable()
    pt.add_column(col_list[0], col_list[1:])
    pt.align = 'l'
    print pt


def raw_print_col(col_list):
    for cell in col_list:
        print(cell)


@args('-f', '--file', metavar='<FILE>', required=True, help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', required=True, help="Excel file sheet name")
def do_create_excel(args):
    """create a excel file"""
    cc = Client(args.file, args.sheetname)
    cc.save(args.file)


@args('-f', '--file', metavar='<FILE>', required=True,help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', required=True,help="Excel file sheet name")
@args('-r','--row',metavar='<ROW>', required=True,help='a row of data')
def do_insert_row(args):
    """insert data into Excel for a row"""
    cc = Client(args.file, args.sheetname)
    cc.insert_row(args.row)
    cc.save(args.file)


@args('-f', '--file', metavar='<FILE>', required=True,help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', required=True,help="Excel file sheet name")
@args('-pretty', '--pretty', action='store_true',default=False, help="pretty table for raw data")
def do_get_head(args):
    """get the excel first line"""
    cc = Client(args.file, args.sheetname)
    row=cc.get_head()
    row_list = []
    for cell in row:
        row_list.append(cell.value)
    if args.pretty:
       pretty_print_row(row_list)
    else:
       raw_print_row(row_list)


@args('-f', '--file', metavar='<FILE>', required=True, help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', required=True, help="Excel file sheet name")
@args('-col', '--col', metavar='<COL>', type=int, help="index of column")
@args('-colname', '--colname', metavar='<COLNAME>',help="name of colum")
@args('-pretty', '--pretty', action='store_true',default=False, help="pretty table for raw data")
def do_get_col(args):
    """ get a column data of Excel"""
    cc = Client(args.file, args.sheetname)
    if args.colname:
        col_id = cc.get_col_id_by_name(args.colname)
        data = cc.get_col(col_id)
    else:
        data = cc.get_col(args.col)
    col_list = []
    for d in data:
        col_list.append(d.value)
    if args.pretty:
       pretty_print_col(col_list)
    else:
       raw_print_col(col_list)


@args('-f', '--file', metavar='<FILE>', required=True,help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', required=True, help="Excel file sheetname")
@args('-pretty', '--pretty', action='store_true',default=False, help="pretty table for raw data")
def do_get_all(args):
    """ get all cell value of a Excel sheet"""
    cc = Client(args.file, args.sheetname)
    data = cc.get_all()
    if args.pretty:
        pretty_print_all(data)
    else:
        raw_print_all(data)


@args('-f', '--file', metavar='<FILE>',required=True, help="Excel file name")
@args('-pretty', '--pretty', action='store_true',default=False, help="pretty table for raw data")
def do_sheet_names(args):
    """ get Excel sheet names"""
    cc = Client(args.file)
    sheets = cc.get_sheet_names()
    if args.pretty:
        pretty_print_row(sheets)
    else:
        raw_print_row(sheets)


@args('-f', '--file', metavar='<FILE>',required=True, help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>',help="the sheetname transfer to csv")
def do_transfer2csv(args):
    """ transfer excel sheet to csv"""
    cc = Client(args.file, args.sheetname)
    csv_raw_data = cc.get_all()
    write_csv_data(csv_raw_data, "%s.csv" %args.sheetname)

@args('-f', '--file', metavar='<FILE>', required=True,help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', required=True, help="Excel file sheetname")
@args('-pretty', '--pretty', action='store_true',default=False, help="pretty table for raw data")
def do_get_all_json(args):
    """ get all cell value of a Excel sheet format json"""
    cc = Client(args.file, args.sheetname)
    data = cc.get_all_json()
    if args.pretty:
        print(json.dumps(data, indent=4, sort_keys=True))
    else:
        print(json.dumps(data))

@args('-f', '--file', metavar='<FILE>', required=True,help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', required=True,help="Excel file sheet name")
@args('-r','--row',metavar='<ROW>', required=True, help='a json of row. for example [{"colname":"cell_value"}]')
def do_insert_json_rows(args):
    """insert data into Excel for a row by json"""
    cc = Client(args.file, args.sheetname)
    row=args.row.replace(" ","").replace("{","{\"").replace("}","\"}").replace(",","\",\"").replace(":","\":\"").replace("}\",\"{","},{")
    data = json.loads(row)
    cc.insert_json_rows(data)
    cc.save(args.file)

@args('-f', '--file', metavar='<FILE>', required=True,help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', required=True,help="Excel file sheet name")
@args('-col','--colname',metavar='<COLNAME>', required=True,help='a colname of sheet')
def do_delete_col(args):
    """delete Excel sheet col by colname"""
    cc = Client(args.file, args.sheetname)
    cc.delete_col(args.colname)
    cc.save(args.file)

@args('-f', '--file', metavar='<FILE>', required=True,help="Excel file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', required=True,help="Excel file sheet name")
@args('-rn','--rownum',metavar='<ROWNUMBER>', required=True,help='the first row number')
@args('-r','--row',metavar='<ROW>', required=True, help='a row of data')
def do_update_row(args):
    """ update role data by row number start with 1 """
    cc = Client(args.file, args.sheetname)
    cc.update_row(args.rownum, args.row)
    cc.save(args.file)

