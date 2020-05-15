# @Time : 2020/5/10 9:38 

# @Author : like

# @File : test.py 

# @Description: 不同excel中复制到find中同一个表中。后面备注来源
import glob
import os
import sys
import time

import xlwings as xw


# 遍历目录下所有xls文件


# 1，打开获取sheet数

class ExcelSolve():
    def __init__(self,findname,columnname,filepath):
        self.findname=findname
        self.columnname=columnname
        self.filepath=filepath
        self.num = 2
    def findxls(self):

        x = glob.glob(self.filepath + '/*.xls*')
        if len(x) == 0:
            print('未找到文件,检查路径')
        else:
            for i in x:
                print('找到文件:' + i)
        return x

    # 创建find.xlsx
    def createfindbook(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.add()
            wb.save(self.filepath +r'\find.xlsx')
            print('创建find成功')
        except:
            pass
        finally:
            wb.close()
            app.quit()

    def getChar(self,number):
        factor, moder = divmod(number, 26)  # 26 字母个数
        modChar = chr(moder + 64)  # 65 -> 'A'
        if factor != 0:
            modChar = self.getChar(factor - 1) + modChar  # factor - 1 : 商为有效值时起始数为 1 而余数是 0
        return modChar

    # 获得指定字段所在的位置
    def getloc(self,list_value):
        # 在前10行中找表头中的列名
        loc = {}
        #处理空表错误
        try:
            column=len(list_value[1])
        except:
            return loc

        for i in range(10):
            for j in range(column):
                try:
                    if self.columnname in list_value[i][j]:
                        loc = {'row_num': i, 'col_num': j}
                        return loc
                except:
                    pass
        print('未找到列或找到多个，请检查列名')

    # 3.遍历之前的sheet，查找专业列中包含指定字符的行，复制到新sheet中，
    def modcopylist(self,copylist, sheetname):
        for i in copylist:
            i.append(sheetname)
        return copylist

    def startcopy(self,copylist, sheetname, wb2):

        despath = self.filepath+'/find.xlsx'
        try:
            wb2.sheets.add('find')
        except:
            pass
        copylist = self.modcopylist(copylist, sheetname)
        wb2.sheets['find'].range('A' + str(self.num)).value = copylist
        self.num += (len(copylist))
        # 增加一行表头
        # wb2.sheets['find'].range('A' + str(num+1)).value = copylist
        wb2.save(despath)

    def getcopylist(self,list_value, loc):
        count = 0
        copylist = []
        # 加入表头，会导致覆盖。需要在上面+1
        # copylist.append(list_value[loc['row_num']])
        # 加入值
        for i in range(len(list_value)):
            try:
                if self.findname in list_value[i][loc['col_num']]:
                    copylist.append(list_value[i])
                    count += 1
            except:
                pass
        print('找到匹配个数:' + str(len(copylist)))
        return copylist

    def getrange(self,sheet):
        info = sheet.used_range
        nrows = info.last_cell.row
        ncolumns = info.last_cell.column
        col = self.getChar(ncolumns)
        range = 'A2:' + col + str(nrows)

        return range

    def copy(self,wb, wb2, sheetname):
        for sheet in wb.sheets:
            print('开始处理表:' + sheet.name)
            range = self.getrange(sheet)
            list_value = sheet.range(range).value
            loc = self.getloc(list_value)
            copylist = self.getcopylist(list_value, loc)
            print('已找到总匹配个数:' + str(self.num - 2))
            print("开始复制：" + sheet.name)
            t1 = time.time()
            sheetname = sheetname + sheet.name
            self.startcopy(copylist, sheetname, wb2)
            print("复制完毕：" + sheet.name)
            t2 = time.time()
            print('耗时：', (t2 - t1))

    # 保存退出
    def saveandexit(self,wb, x):
        wb.save(x)
        xw.App().quit()

    def open(self,file_path):
        app = xw.App(visible=True, add_book=False)
        wb = app.books(file_path)
        return wb

    def main(self):
        xlslist = self.findxls()
        self.createfindbook()
        despath = self.filepath+'/find.xlsx'
        wb2 = xw.Book(despath)
        count=0
        total=len(xlslist)
        for i in xlslist:
            count+=1
            if i == despath:
                pass
            else:
                print("开始处理：" + i)
                app = xw.App(visible=True, add_book=False)
                wb = app.books.open(i)
                filename = os.path.split(i)[1]
                self.copy(wb, wb2, filename)
                wb.close()
                app.quit()
                #在pyqt输出时会导致异常
                #self.process_bar(count/total)
                print('*' * 100)

    def process_bar(self,percent, width=50):
        use_num = int(percent * width)
        space_num = int(width - use_num)
        percent = percent * 100
        print('[%s%s]%d%%' % (use_num * '#', space_num * ' ', percent), file=sys.stdout, flush=True, end='\r')

if __name__ == '__main__':
    es=ExcelSolve('扶贫办','摘要','C:/Users/Administrator/Desktop/code/python/pyqt5/resource')
    es.main()
