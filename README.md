# PyExcel

This is a tool for excel operate

### Install

you can use command line  as below: 

```curl
git clone https://github.com/CaesarLinsa/PyExcel.git
cd PyExcel
python setup.py install
```

### pyexcel  help

#### you can use pyexcel, pyexcel hep or pyexcel --help to get the help info, as below:

```
positional arguments:
  <subcommand>
    create-excel  create a excel file
    get-col          get a column data of Excel
    get-head       get the excel first line
    insert-row      insert data into Excel for a row
    help               Display help about this program or one of its subcommands.

optional arguments:
  -v, --version  show program's version number and exit
```
#### you can use pyexcel help subcommand to get the subcommand info, for example:
```
pyexcel  help create-excel

usage: pyexcel create-excel [-f <FILE>] [-sn <SHEETNAME>]

create a excel file

optional arguments:
  -f <FILE>, --file <FILE>
                   Excel file name
  -sn <SHEETNAME>, --sheetname <SHEETNAME>
                   Excel file sheet name
```
