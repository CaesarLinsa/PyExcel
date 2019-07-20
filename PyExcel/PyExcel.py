from util import args
from Client import  Client
@args('-f', '--file', metavar='<FILE>', help="xlsx file name")
@args('-sn', '--sheetname', metavar='<SHEETNAME>', help="xlsx file sheet name")
def do_create_xlsx(args):
    cc = Client(args.file, args.sheetname)
    cc.save(args.file)
