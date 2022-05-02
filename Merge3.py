import tkinter.messagebox
import copy
import traceback
from ParseError import ParseError


class Merge3:
    workbooks = None
    mergeSheets = []

    def init(self, workbooks):
        self.workbooks = workbooks
        # self.check(workbooks)
        self.doMerge(workbooks)
        return self

    def check(self, workbooks):
        length = len(workbooks)
        for i in range(length):
            workbook1 = workbooks[i]
            for j in range(i + 1, length):
                workbook2 = workbooks[j]
                workbook1.compareSize(workbook2)

    # 根据sheet名称获取sheet
    def getSheetByName(self, sheets, name):
        for sheet in sheets:
            if sheet.getName() == name:
                return sheet
        return None

    def doMerge(self, workbooks):
        sheetsList = [list(w.getSheets()) for w in workbooks]

        # 相同sheet名称进行分组
        alikeSheetsList = []
        length = len(sheetsList)
        for i in range(length):
            sheets1 = sheetsList[i]
            for sheet1 in sheets1:
                alikeSheets = [copy.deepcopy(sheet1)]
                alikeSheetsList.append(alikeSheets)

                for j in range(i + 1, length):
                    sheets2 = sheetsList[j]
                    sheet2 = self.getSheetByName(sheets2, sheet1.getName())
                    if sheet2 is not None:
                        alikeSheets.append(sheet2)
                        sheets2.remove(sheet2)

        # 同一分组数字进行合并
        for likeSheets in alikeSheetsList:
            rowNum = max([sheet.getRowNum() for sheet in likeSheets])
            colNum = max([sheet.getColNum() for sheet in likeSheets])

            # 同一组类表格 大小填充成一致
            for sheet in likeSheets:
                dataLists = sheet.getDataLists()
                c = rowNum - len(dataLists)
                dataLists[len(dataLists):]=[[] for i in range(c)]
                for data in dataLists:
                    c = colNum - len(data)
                    data[len(data):] = ['' for i in range(c)]

            # 同组相同位置有一个是数字就进行合并,  相同位置有数字,又有字符 报该位置不应该是字符错误
            meragSheet = likeSheets[0]
            self.mergeSheets.append(meragSheet)
            for r in range(rowNum):
                for c in range(colNum):
                    isNumber = False
                    for sheet in likeSheets:
                        isNumber = self.isNumber(sheet.getDataLists()[r][c])
                        if isNumber: break

                    if isNumber:
                        val = 0
                        for sheet in likeSheets:
                            data = sheet.getDataLists()[r][c]
                            self.checkVal(data, sheet, r, c)
                            val += self.convertNumber(data)
                        meragSheet.getDataLists()[r][c] = val

    # 判断相同位置有数字后, 校验, 相同位置不能既有数字又有字符
    def checkVal(self, data, sheet, r, c):
        if not self.isNumber(data) and type(data) == str and data != '' and data != '——':
            tkinter.messagebox.showerror('合并错误',
                                         ' 文件:  ' + sheet.getFileName() + ' \n' +
                                         ' 表格:  ' + sheet.getName() + ' \n' +
                                         ' 位置:  ' + str(c + 1) + ',' + str(r + 1) + '\n ' +
                                         '不应该为字符!!!')
            raise ParseError("不应该为字符")

    # 判断相同位置有数字后,校验过后, 值处理(字符为0)
    def convertNumber(self, data):
        if not self.isNumber(data):
            return 0
        return float(data)

    # 判断值是否是数字
    def isNumber(self, val):
        try:
            float(val)
            return True
        except Exception as e:
            return False

    # 保存
    def save(self, filePath):
        wb = self.workbooks[0].workbook
        try:
            for mergeSheet in self.mergeSheets:
                sheet = None
                try:
                    sheet = wb.sheets[mergeSheet.getName()]
                except:
                    pass
                if sheet is None:
                    wb.sheets.add()
                    sheet = wb.sheets.active
                sheet.name = mergeSheet.getName()
                sheet.range('A1').value = mergeSheet.getDataLists()
                sheet.used_range.autofit()

            wb.save(filePath)
        except Exception as e:
            traceback.print_exc()
