import csv, openpyxl, sys, time, os, re

__author__ = 'Yuuko'

__version__ = '1.0.3'

class csvProcessXlsx:
    """
    
    This is a class for supporting the CSV files merge to a new xlsx file.
   
   -------------------------------------------------------------------------------
    
    Usage 1: python csvConvertor.py [csv files...] [xlsx filename]
             eg: python csvConvertor.py first.csv second.csv third.csv xlsxObject
             attention: Arguments split by space.
    Usage 2: Put the CSV files with the script at same path,
             use shell change dirctory to the path.
             username>_ python csvConvertor.py
             insert the xlsx filename for done.
    
    -------------------------------------------------------------------------------
    ☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆  Enjoy  ☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆☆
    
    """
    
    tableHead = list(['ID', '代号', '名称', '数量', '材料', '单重', '总重', '备注'])
    
    def __init__(self, xlsxname):
        if (not os.path.exists(xlsxname)):
            self.xlsxname = xlsxname
            wb = openpyxl.Workbook()
            sheet = wb.active
            for col in ['B', 'C', 'E', 'H']:
                sheet.column_dimensions[col].width = 20
            wb.save(xlsxname)
        else:
            pass
    
    def csvInsertXlsx(self, csvname, xlsxname, titleInsert):
    
        """
        Get the context of the CSV files, for writting to
        a new xlsx file.
        """
        
        fr = open(csvname, 'r')
        csvr = csv.reader(fr)
        wb = openpyxl.load_workbook(xlsxname)
        sheet = wb.active
        sheet.title = titleInsert
        lix = list(csvr)
        lix.insert(0, self.tableHead)
        colindex = 0
        for row in lix:
            if (not((row[1] == ' ') and (row[2] == ' '))):
                rown = lix.index(row) + 1
                if (self.getSheetRows(xlsxname) != 0):
                    rown += (self.getSheetRows(xlsxname) + 1)
                for item in row:
                    newitem = ''
                    if (item.find('\n')):
                        newitem = item.replace('\n', '')
                    else:
                        pass
                    colindex += 1
                    if (row[0] != 'ID'):
                        if ((colindex == 4)):
                            sheet.cell(row = rown , column = colindex).value = self.funcConvert(item)
                        elif(colindex == 6):
                            sheet.cell(row = rown, column = colindex).value = float(item.replace(' ', '0'))
                        elif (colindex == 7):
                            leftarg = 'D' + str(rown)
                            rightarg = 'F' + str(rown)
                            seventhvalue = '=MMULT(%s,%s)' % (leftarg, rightarg)
                            sheet.cell(row = rown, column = colindex).value = seventhvalue
                        else:
                            sheet.cell(row = rown, column = colindex).value = item
                    else:
                        sheet.cell(row = rown, column = colindex).value = item
                print(row)
                print('Was writed.')
            colindex = 0
        wb.save(xlsxname)
        
    def getSheetRows(self, xlsxname):
        
        """
        Get the table's current rows.
        """
        
        wb = openpyxl.load_workbook(xlsxname)
        sheet = wb.active
        return len(list(sheet.rows))
    
    def mergeCSV(self, csvnames, xlsxname, titleInsert):
    
        """
        Merge the CSV files into the single file.
        """
        
        for csvs in csvnames:
            self.csvInsertXlsx(csvs, xlsxname, titleInsert)
        print('All done.')
        
    def funcConvert(self, stritem):
    
        """
        Convert the strings to calculatable expression,
        for getting the value to the total of the numbers.
        """
    
        firstre = r'\d*([+-xX/])*\d*([+-xX/])*\d*'
        secondre = r'\d*\(\d*\)'
        thirdre = r'\(\d*\)\d*'
        firstcom = re.compile(firstre)
        secondcom = re.compile(secondre)
        firstmatch = firstcom.match(stritem)
        secondmatch = secondcom.match(stritem)
        thirdcom = re.compile(thirdre)
        thirdmatch = thirdcom.match(stritem)
        if (firstmatch):
            return eval(stritem.lower().replace('x', '*'))
        elif (secondmatch):
            newstrtem = stritem.split('(')
            newstrtem.append(newstrtem.pop(1).replace(')', ''))
            tem = 0
            for num in newstrtem:
                tem += int(num)
            return tem
        elif (thirdmatch):
            newstrtem = stritem.split(')')
            newstrtem.append(newstrtem.pop(0).replace('(', ''))
            tem = 0
            for num in newstrtem:
                tem += int(num)
            return tem
        else:
            return int(stritem)
            
if __name__ == '__main__':
    os.system('title CSV Convertor')
    print('    Author: %s' % __author__)
    print('    Version: %s' % __version__)
    print('    Currently running path is   %s' % os.getcwd())
    names = []
    if ((len(sys.argv) > 1) and (sys.argv[1] != '--help')):
        
        for argname in sys.argv[1:]:
            if (argname.endswith('.csv')):
                names.append(argname)
        a = csvProcessXlsx(sys.argv[-1] + '.xlsx')
        start_time = time.time();
        a.mergeCSV(names, sys.argv[-1] + '.xlsx', '明细表')
        end_time = time.time()
        print('Cost time is %.8f' % (end_time - start_time))
    elif ((len(sys.argv) == 2) and (sys.argv[1] == '--help')):
        if (sys.argv[1] == '--help'):
            print(csvProcessXlsx.__doc__)
        else:
            pass
    else:
        for name in os.listdir(os.getcwd()):
            if (name.endswith('.csv')):
                names.append(name)
        xlsxna = input('Please input the xlsx file name: ') + '.xlsx'
        a = csvProcessXlsx(xlsxna)
        start_time = time.time();
        a.mergeCSV(names, xlsxna, '明细表')
        end_time = time.time()
        print('Cost time is %.8f' % (end_time - start_time))
        