"""
Filename: spreadsheet.py
Author: Cameron Young
E-mail: camyoung@cisco.com
Version: 1.0 

This library is designed to use two other Excel libraries and combines them to
allow one library to edit and save both .xls and .xlsx spreadsheets. Both
libraries are included in the spreadsheet directory for ease of installation.

Look up openpyxl (http://pythonhosted.org/openpyxl/) for how to implement this
library. The spreadsheet.Workbook is identical to openpyxl.Workbook except for
save, which has been modified to allow .xls in addition to .xlsx, and the format
is determined by the extension typed at the end of the filename.


example usage:

import spreadsheet

wb = spreadsheet.load_workbook('example.xls')
sh = wb.get_active_sheet()

print sh.cell(row=0, column=0).value





The MIT License (MIT)

Copyright (c) 2013 Cameron Young

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.

"""

import spreadsheet.xlrd as xlrd         # .xls  spreadsheets
import spreadsheet.xlwt as xlwt         # .xls  spreadsheets
import spreadsheet.openpyxl as openpyxl # .xlsx spreadsheets

def load_workbook(filename): # openpyxl syntax
    wb = Workbook()
    wb.load(filename)
    return wb

def open_workbook(filename): # xlrd syntax
    return load_workbook(filename)

class Workbook(openpyxl.Workbook):
    def load(self, filename):
        """ Loads in an Excel file, using its extension to know the format """
        for sheet in self.get_sheet_names():
            # Remove all current data (mainly for default sheet)
            self.remove_sheet(self.get_sheet_by_name(sheet))
        try:
            if filename.endswith('.xlsx'):
                wb = openpyxl.load_workbook(filename)
                for sheet in wb.get_sheet_names():
                    sh = wb.get_sheet_by_name(sheet)
                    output_sh = self.create_sheet()
                    output_sh.title = sh.title
                    for row in range(len(sh.rows)):
                        for col in range(len(sh.columns)):
                            output_sh.cell(row=row, column=col).value = sh.cell(row=row, column=col).value
            elif filename.endswith('.xls'):
                wb = xlrd.open_workbook(filename)
                for index in range(wb.nsheets):
                    sh = wb.sheet_by_index(index)
                    output_sh = self.create_sheet()
                    output_sh.title = sh.name
                    for row in range(sh.nrows):
                        for col in range(sh.ncols):
                            output_sh.cell(row=row, column=col).value = sh.cell_value(rowx=row, colx=col)
            else:
                raise Exception('Unknown format, please use either .xls or .xlsx')
        except:
            raise Exception('Uh oh... something went wrong with loading the data. The file may be corrupt or have a wrong extension. Check if the file will load in Excel.')
    
    def save(self, filename):
        """ Saves the workbook with the format corresponding to its extension """
        if filename.endswith('.xlsx'):
            # Use the over save for .xlsx
            openpyxl.Workbook.save(self, filename)
        elif filename.endswith('.xls'):
            # Convert to xlwt for .xls and then save
            wb = xlwt.Workbook()
            for sheet in self.get_sheet_names():
                sh = self.get_sheet_by_name(sheet)
                output_sh = wb.add_sheet(sheet)
                for row in range(len(sh.rows)):
                    for col in range(len(sh.columns)):
                        output_sh.write(row, col, sh.cell(row=row, column=col).value)
            wb.save(filename)
        else:
            raise Exception('Please provide an extension (.xlsx, .xls) to determine which format to save as')
