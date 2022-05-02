class Sheet:
    name = None
    rowNum = None
    colNum = None
    dataLists = []
    fileName = ''

    # 初始化成员变量
    def init(self, sheet, fileName):
        self.name = sheet.name
        self.fileName = fileName
        self.rowNum = sheet.used_range.rows.count
        self.colNum = sheet.used_range.columns.count
        self.dataLists = self.createData(sheet)

        return self

    def getFileName(self):
        return self.fileName

    def createData(self, sheet):
        # dataList = []
        # for row in sheet.used_range.rows:
        #     dataList.append(row.value)
        return sheet.used_range.value

    # sheet名称
    def getName(self):
        return self.name

    # sheet使用行数
    def getRowNum(self):
        return self.rowNum

    # sheet使用列数
    def getColNum(self):
        return self.colNum

    # sheet数据 二维列表
    def getDataLists(self):
        return self.dataLists

    # # 合并两个sheet的数字
    # def mergeSheet(self, sheet):
