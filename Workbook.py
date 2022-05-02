import tkinter.messagebox
import xlwings as xw
from Sheet import Sheet
from ParseError import ParseError


class Workbook:
    workbook = None
    sheets = []
    filePath = None

    def init(self, filePath):
        self.filePath = filePath
        self.workbook = self.createWb(filePath)
        self.sheets = self.createSheet(self.workbook)
        self.getFileName()
        return self

    def createSheet(self, workbook):
        sheets = []
        for sheet in workbook.sheets:
            sheets.append(Sheet().init(sheet, self.getFileName()))
        return sheets

    def createWb(self, filePath):
        global wb
        global app
        try:
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False  # 不显示Excel消息框
            app.screen_updating = False  # 关闭屏幕更新,
            wb = app.books.open(filePath)
        except:
            wb.close()
            app.quit()
        return wb

    # 文件名称
    def getFileName(self):
        fileIndex = self.filePath.rfind("/")
        if fileIndex != -1:
            return self.filePath[fileIndex + 1:]
        else:
            return self.filePath

    # 根据sheet名称获取sheet
    def getSheetByName(self, name):
        for sheet in self.sheets:
            if sheet.getName() == name:
                return sheet
        return None

    # 获取所有sheet
    def getSheets(self):
        return self.sheets

    # 关闭
    def close(self):
        return self.workbook.close()

    # 相同名称的sheet 表格大小要相同
    def compareSize(self, workbook):
        sheets = workbook.getSheets()
        for sheet in sheets:
            sheet1 = None
            try:
                sheet1 = self.getSheetByName(sheet.getName())
            except:
                pass
            if sheet1 is not None and (sheet1.getRowNum() != sheet.getRowNum() or sheet1.getColNum() != sheet.getColNum()):
                tkinter.messagebox.showerror('解析错误',
                                             ' 文件:  ' + self.getFileName() + ' ' + str(sheet1.getColNum()) + '×' + str(
                                                 sheet1.getRowNum()) +
                                             '\n 文件:  ' + workbook.getFileName() + ' ' + str(sheet.getColNum()) + '×' + str(
                                                 sheet.getRowNum()) +
                                             '\n 表格:  ' + sheet.name +
                                             '\n 表格大小不相同!!!')
                raise ParseError("sheet大小不相同")
