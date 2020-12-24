import os
import xlrd
import csv
from datetime import datetime

class Param(object):
    """
    参数化测试数据文件解析工具
    """
    def __init__(self, paramConf):
        """
        初始化
        :param paramConf: 数据配置信息（文件路径等）
        """
        self.paramConf = paramConf

    def paramRowsCount(self):
        """
        获取参数行数
        """
        pass

    def paramColsCount(self):
        """
        获取参数列数
        """
        pass

    def paramHeader(self):
        """
        获取参数名列表
        """
        pass

    def paramDataList(self):
        """
        获取全部参数，返回格式为列表
        """
        pass

    def paramDataDict(self):
        """
        获取全部参数，返回格式为字典
        """
        pass

    def paramDataListDict(self):
        """
        获取全部参数，返回格式为列表字典
        """
        pass


class ParamFactory(object):
    """
    参数化测试数据解析工厂
    """
    def chooseParam(self, type, paramConf):
        target = None
        if type == 'xls':
            target = XLS(paramConf)
        elif type == 'csv':
            target = CSV(paramConf)
        else:
            print('暂不支持该文件类型')
        return target


class XLS(Param):
    '''
    XLS数据文件解析工具

    xls文件格式:
    2、第1行是参数名；
    3、第2~N行是参数值；
    '''

    def __init__(self, paramConf):
        '''
        初始化
        :param paramConf: 数据配置信息（文件路径、指定的sheet）
        例如：
        1、{'filepath': path_to_file.xls, 'sheet': 0}
        2、{'filepath': path_to_file.xls, 'sheet': '工作表1'}
        '''
        self.paramConf = paramConf
        self.paramfile = self.paramConf['filepath']
        self.workbook = xlrd.open_workbook(self.paramfile)
        self.paramsheet = self.getParamSheet(self.paramConf['sheet'])

    def getParamSheet(self, sheet):
        '''
        获取指定的sheet
        :param sheet: sheet name or sheet index
        :return: sheet
        '''
        if isinstance(sheet, str):
            return self.workbook.sheet_by_name(sheet)
        return self.workbook.sheet_by_index(sheet)
    
    def paramRowsCount(self):
        '''
        获取参数文件行数
        :return: 参数行数 int
        '''
        return self.paramsheet.nrows

    def paramColsCount(self):
        '''
        获取参数文件列数(参数个数)
        :return: 参数文件列数(参数个数) int
        '''
        return self.paramsheet.ncols

    def getOneCell(self, numberRow, numberCol):
        return self.paramsheet.cell_value(numberRow, numberCol)

    def getOneRow(self, nRow):
        '''
        返回一行数据
        :param nRow: 行数
        :return: 一行数据 []
        '''
        row_values = []
        for iCol in range(self.paramsheet.ncols):
            ctype = self.paramsheet.cell(nRow, iCol).ctype
            cell = self.paramsheet.cell_value(nRow, iCol)
            if ctype == 2 and cell % 1 == 0:  # 整型
                cell = int(cell)
            elif ctype == 3:  # 时间
                date = datetime(*xldate_as_tuple(cell, 0))
                cell = date.strftime('%Y/%d/%m %H:%M:%S')
            elif ctype == 4:  # 布尔值
                cell = True if cell == 1 else False
            row_values.append(cell)           
        return row_values

    def getOneCol(self, nCol):
        '''
        返回一列
        :param nCol: 列数
        :return: 一列数据 []
        '''
        return self.paramsheet.col_values(nCol)

    def paramHeader(self):
        '''
        获取参数名称(Excel 文件中，第一行是给人读取的每一列参数的注释，而所有的 Excel 都是从第二行开始读取的)
        :return: 参数名称[]

        example:
        ['equipmentid', 'exp']
        '''
        return self.getOneRow(0)

    def paramDataList(self):
        '''
        获取全部参数，返回格式为列表
        :return: 全部参数[[]]
        '''
        nCountRows = self.paramRowsCount()
        paramall = []
        iRowStep = 1
        while iRowStep < nCountRows:
            paramall.append(self.getOneRow(iRowStep))
            iRowStep += 1
        return paramall

    def paramDataDict(self):
        '''
        获取全部参数，返回格式为字典
        :return: {{}},其中dict的key值是header的值

        example:
        {
        0: {'equipmentid': 10001.0, 'exp': 'your pick up equipmentid:10001'}, 
        1: {'equipmentid': 10002.0, 'exp': 'your pick up equipmentid:10002'}, 
        2: {'equipmentid': 10003.0, 'exp': 'your pick up equipmentid:10003'}
        }
        '''
        nCountRows = self.paramRowsCount()
        nCountCols = self.paramColsCount()
        paramHeader = self.paramHeader()
        paramAllLineDict = {}
        iRowStep = 1
        iColStep = 0
        while iRowStep < nCountRows:
            paramOneLineList = self.getOneRow(iRowStep)
            # paramOneLineDict = {}
            # while iColStep < nCountCols:
            #     paramOneLineDict[paramHeader[iColStep]] = paramOneLineList[iColStep]
            #     iColStep += 1
            # iColStep = 0
            paramOneLineDict = dict(zip(paramHeader, paramOneLineList))
            paramAllLineDict[iRowStep-1] = paramOneLineDict
            iRowStep += 1
        return paramAllLineDict

    def paramDataListDict(self):
        '''
        获取全部参数，返回格式为列表字典
        :return: {{}},其中dict的key值是header的值

        example:
        [
          {'equipmentid': 10001.0, 'exp': 'your pick up equipmentid:10001'}, 
          {'equipmentid': 10002.0, 'exp': 'your pick up equipmentid:10002'}, 
          {'equipmentid': 10003.0, 'exp': 'your pick up equipmentid:10003'}
        ]        
        '''
        nCountRows = self.paramRowsCount()
        nCountCols = self.paramColsCount()
        paramHeader = self.paramHeader()
        paramDataList = []
        iRowStep = 1
        iColStep = 0
        while iRowStep < nCountRows:
            paramOneLineList = self.getOneRow(iRowStep)
            print(paramOneLineList)
            # paramOneLineDict = {}
            # while iColStep < nCountCols:
            #     paramOneLineDict[paramHeader[iColStep]] = paramOneLineList[iColStep]
            #     iColStep += 1
            # iColStep = 0
            paramOneLineDict = dict(zip(paramHeader, paramOneLineList))
            paramDataList.append(paramOneLineDict)
            iRowStep += 1
        return paramDataList


class CSV(Param):
    """
    CSV数据文件解析工具
    CSV:Comma Separated Values 逗号分隔值（字符分隔值），一种常用的文本格式，用以存储表格数据，包括数字或者字符。
    
    csv文件格式
    1、第1行是参数名；
    2、第2~N行是参数值；
    示例：demo.csv
        username,password
        test1,111111
        test2,222222
        test3,333333
    """
    
    def __init__(self, paramConf):
        self.paramConf = paramConf
        self.paramFile = paramConf['filepath']


    def paramRowsCount(self):
        """
        获取参数行数
        """
        with open(self.paramFile, encoding='utf-8') as fr:
            csvfr = csv.reader(fr)
            rows = [row for row in csvfr]
        return len(rows)

    def paramColsCount(self):
        """
        获取参数列数
        """
        with open(self.paramFile, encoding='utf-8') as fr:
            csvfr = csv.reader(fr)
            for row in csvfr:
                count = len(row)
        return count

    def paramHeader(self):
        """
        获取参数名列表
        """
        with open(self.paramFile, encoding='utf-8') as fr:
            csvfr = csv.reader(fr)
            for row in csvfr:
                header = row
                break
        return header 

    def paramDataList(self):
        """
        获取全部参数，返回格式为列表
        """
        with open(self.paramFile, encoding='utf-8') as fr:
            csvfr = csv.reader(fr)
            rows = [row for row in csvfr]
        return rows[1:]

    def paramDataDict(self):
        """
        获取全部参数，返回格式为字典
        """
        paramAllLineDict = {}
        with open(self.paramFile, encoding='utf-8') as fr:
            csvfr = csv.DictReader(fr)
            for index, value in enumerate(csvfr):
                paramAllLineDict[index] = dict(value)
        return paramAllLineDict

    def paramDataListDict(self):
        """
        获取全部参数，返回格式为列表字典
        """        
        with open(self.paramFile, encoding='utf-8') as fr:
            csvfr = csv.DictReader(fr)
            paramDataList = [dict(item) for item in csvfr]
        return paramDataList


if __name__ == "__main__":
    # 定义存储参数的excel文件路径
    paramfile = './paramiterize/equipments.xls'
    # 参数数据配置信息
    # paramConf = {'filepath': paramfile, 'sheet': 0}
    paramConf = {'filepath': paramfile, 'sheet': '工作表1'}
    # 调用参数类完成参数读取，返回是一个字典，包含全部的excel数据除去excel的第一行表头说明
    xlsObj = ParamFactory().chooseParam('xls', paramConf)
    print(xlsObj.paramRowsCount())
    print(xlsObj.paramColsCount())
    print(xlsObj.paramHeader())
    print(xlsObj.paramDataList())
    print(xlsObj.paramDataDict())
    print(xlsObj.paramDataListDict())

    # 定义存储参数的csv文件路径
    paramfile = './paramiterize/users.csv'
    # 参数数据配置信息
    paramConf = {'filepath': paramfile}
    # 调用参数类完成参数读取，返回是一个字典，包含全部的excel数据除去excel的第一行表头说明
    csvObj = ParamFactory().chooseParam('csv', paramConf)
    print(csvObj.paramRowsCount())
    print(csvObj.paramColsCount())
    print(csvObj.paramHeader())
    print(csvObj.paramDataList())
    dataDict = csvObj.paramDataDict()
    print(dataDict)
    print(dataDict[0])
    print(dataDict[0]['姓名'])
    print(csvObj.paramDataListDict())
