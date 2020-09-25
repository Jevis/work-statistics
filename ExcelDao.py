import xlrd
import re
from xlrd import xldate_as_tuple
import datetime


# 0 empty,1 string, 2 number, 3 date, 4 boolean, 5 error
def getRealityTime(timeCell, i, titleIndex):
    data = timeCell[titleIndex:len(timeCell)]
    tt = re.split('[~-]', data)
    if len(tt) == 2:
        queueUpTime = datetime.datetime.now().strftime("%Y/%m/%d ") + tt[0]
        numTime = datetime.datetime.now().strftime("%Y/%m/%d ") + tt[1].lstrip("隔天").lstrip("隔日")
        qt = datetime.datetime.strptime(queueUpTime, '%Y/%m/%d %H:%M')
        rt = datetime.datetime.strptime(numTime, '%Y/%m/%d %H:%M')
        if rt.timestamp() - qt.timestamp() < 0:
            rt = rt + datetime.timedelta(days=1)
        return [tt[0], (rt - qt).seconds]  # 上班时间和上班总时长
    else:
        raise Exception(print("行数：" + str(i + 1) + " 这里的时间上下班的分割有问题" + data))


class ExcelDao():
    def __init__(self, data_path):
        # 定义一个属性接收文件路径
        self.data_path = data_path
        # 使用xlrd模块打开excel表读取数据
        self.data = xlrd.open_workbook(self.data_path)
        # 根据工作表的名称获取工作表中的内容（方式①）
        self.table = self.data.sheet_by_index(0)
        # 获取工作表的有效行数
        self.rowNum = self.table.nrows
        # 获取工作表的有效列数
        self.colNum = self.table.ncols
        # 是不是上午
        self.isDay = True
        # 定义一个空列表
        self.result = []  # 存储所有人的list<list<map>>

        self.time = datetime.datetime.now()

    # 判断这个人是不是实际到的天数和应到的一样
    def toJudgeWorkSumDay(self, result):
        nR = result[0]['nameRow']
        realityDay = int(self.table.cell_value(nR + 1, 5))
        if len(result) > realityDay * 2:
            print("行数 :" + str(nR) + " 这个人的实际天数大于应到天数")
        elif len(result) < realityDay * 2:
            print("行数 :" + str(nR) + " 这个人的实际天数小于应到天数")
        return realityDay * 2 == len(result)

    # 定义一个读取excel表的方法
    def readExcel(self):
        timeData = []  # 存储这个人的应该上下班的时间
        oneData = []  # 存储这个人的实际上下班的时间

        for i in range(self.rowNum):
            cellValue = self.table.cell_value(i, 2)
            cellType = self.table.cell_type(i, 2)
            if 1 == cellType:  # 输入这一行是有数据的 不是上下班标准就是特殊事件
                timeCell = str(cellValue)
                titleIndex = timeCell.find("上班時間")
                if -1 != titleIndex and titleIndex != len(timeCell) - 1:
                    nameRow = i  # 存储用户姓名的行数
                    timeData = getRealityTime(timeCell, i, timeCell.find("間") + 1)  # 这里要算出间隔
                    if len(oneData) > 0 and self.toJudgeWorkSumDay(oneData):
                        self.result.append(oneData.copy())
                        oneData.clear()
                    else:
                        if nameRow > 0:
                            print("行数：" + str(nameRow + 1) + " 上一个人的上下班的次数不对")
                else:
                    print("行数：" + str(i + 1) + " 这一天她使用了特殊的假期：" + timeCell)

            self.toGetTheWorksEveryDay(oneData, nameRow, timeData, i)

            if i == self.rowNum - 1 and len(oneData) > 0:
                if self.toJudgeWorkSumDay(oneData):
                    self.result.append(oneData.copy())
                    oneData.clear()
                else:
                    print("最后一个人的上下班的次数不对")

        return self.result

    def toGetTheWorksEveryDay(self, one, nameRow, timeData, i):
        dateCell = self.table.cell_type(i, 3)
        if dateCell == 3:
            if i == self.rowNum - 1 or 0 != self.table.cell_type(i + 1, 3):
                c_cell = self.table.cell_value(i, 3)
                date = datetime.datetime(*xldate_as_tuple(c_cell, 0))
                if self.isDay:  # 上午
                    self.time = datetime.datetime.strptime(date.strftime("%Y/%m/%d ") + timeData[0], '%Y/%m/%d %H:%M')
                    r = {'type': 1, 'nameRow': nameRow, 'rowNum': i,
                         'rulerTime': date.strftime("%Y/%m/%d ") + timeData[0],
                         'workedTime': date.strftime("%Y/%m/%d %H:%M")}
                    one.append(r)
                    self.isDay = False
                else:
                    r = {'type': 2, 'nameRow': nameRow, 'rowNum': i,
                         'rulerTime': (self.time + datetime.timedelta(seconds=timeData[1])).strftime('%Y/%m/%d %H:%M'),
                         'workedTime': date.strftime("%Y/%m/%d %H:%M")}
                    one.append(r)
                    self.isDay = True
