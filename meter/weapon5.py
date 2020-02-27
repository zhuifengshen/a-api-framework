import os
import json
import xlrd
import requests
from websocket import create_connection

'''
演进5：测试数据分离与参数化（参数类）
'''


class Common(object):
    # common的构造函数
    def __init__(self, url_root, api_type):
        '''
        :param api_type:接口类似当前支持http，ws，http就是http协议，ws是Websocket协议  
        :param url_root: 被测系统的跟路由   
        '''
        if api_type == 'ws':
            self.ws = create_connection(url_root)
        elif api_type == 'http':
            self.ws = 'null'
            self.url_root = url_root

    # ws协议的消息发送
    def send(self, params):
        '''
        :param params: websocket接口的参数

        :return: 访问接口的返回值
        '''
        self.ws.send(params)
        res = self.ws.recv()
        return res

    # common类的析构函数，清理没有用的资源

    def __del__(self):
        '''
        :return:
        '''
        if self.ws != 'null':
            self.ws.close()

    def get(self, uri, params=None):
        '''
        封装你自己的get请求，uri是访问路由，params是get请求的参数，如果没有默认为空 
        :param uri: 访问路由 
        :param params: 传递参数，string类型，默认为None 
        :return: 此次访问的response
        '''
        # 拼凑访问地址
        if params is not None:
            url = self.url_root + uri + params
        else:
            url = self.url_root + uri
        # 通过get请求访问对应地址
        res = requests.get(url)
        # 返回request的Response结果，类型为requests的Response类型
        return res

    def post(self, uri, params=None):
        '''
        封装你自己的post方法，uri是访问路由，params是post请求需要传递的参数，如果没有参数这里为空
        :param uri: 访问路由
        :param params: 传递参数，string类型，默认为None
        :return: 此次访问的response
        '''
        # 拼凑访问地址
        url = self.url_root + uri
        if params is not None:
            # 如果有参数，那么通过post方式访问对应的url，并将参数赋值给requests.post默认参数data
            # 返回request的Response结果，类型为requests的Response类型
            res = requests.post(url, data=params)
        else:
            # 如果无参数，访问方式如下
            # 返回request的Response结果，类型为requests的Response类型
            res = requests.post(url)
        return res

    def put(self, uri, params=None):
        '''
        封装你自己的put方法，uri是访问路由，params是put请求需要传递的参数，如果没有参数这里为空
        :param uri: 访问路由
        :param params: 传递参数，string类型，默认为None
        :return: 此次访问的response
        '''
        url = self.url_root+uri
        if params is not None:
            # 如果有参数，那么通过put方式访问对应的url，并将参数赋值给requests.put默认参数data
            # 返回request的Response结果，类型为requests的Response类型
            res = requests.put(url, data=params)
        else:
            # 如果无参数，访问方式如下
            # 返回request的Response结果，类型为requests的Response类型
            res = requests.put(url)
        return res

    def delete(self, uri, params=None):
        '''
        封装你自己的delete方法，uri是访问路由，params是delete请求需要传递的参数，如果没有参数这里为空
        :param uri: 访问路由
        :param params: 传递参数，string类型，默认为None
        :return: 此次访问的response
        '''
        url = self.url_root + uri
        if params is not None:
            # 如果有参数，那么通过put方式访问对应的url，并将参数赋值给requests.put默认参数data
            # 返回request的Response结果，类型为requests的Response类型
            res = requests.delete(url, data=params)
        else:
            # 如果无参数，访问方式如下
            # 返回request的Response结果，类型为requests的Response类型
            res = requests.put(url)
        return res


class Param(object):
    def __init__(self, paramConf='{}'):
        self.paramConf = json.loads(paramConf)

    def paramRowsCount(self):
        pass

    def paramColsCount(self):
        pass

    def paramHeader(self):
        pass

    def paramAllline(self):
        pass

    def paramAlllineDict(self):
        pass


class XLS(Param):
    '''
    xls基本格式(如果要把xls中存储的数字按照文本读出来的话,纯数字前要加上英文单引号:

    第一行是参数的注释,就是每一行参数是什么
    第二行是参数名,参数名和对应模块的po页面的变量名一致
    第3~N行是参数
    最后一列是预期,默认头Exp
    '''

    def __init__(self, paramConf):
        '''
        :param paramConf: xls 文件位置(绝对路径)
        '''
        self.paramConf = paramConf
        self.paramfile = self.paramConf['file']
        self.data = xlrd.open_workbook(self.paramfile)
        self.getParamSheet(self.paramConf['sheet'])

    def getParamSheet(self, nsheets):
        '''
        设定参数所处的sheet
        :param nsheets: 参数在第几个sheet中
        :return:
        '''
        self.paramsheet = self.data.sheets()[nsheets]

    def getOneline(self, nRow):
        '''
        返回一行数据
        :param nRow: 行数
        :return: 一行数据 []
        '''
        return self.paramsheet.row_values(nRow)

    def getOneCol(self, nCol):
        '''
        返回一列
        :param nCol: 列数
        :return: 一列数据 []
        '''
        return self.paramsheet.col_values(nCol)

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

    def paramHeader(self):
        '''
        获取参数名称(Excel 文件中，第一行是给人读取的每一列参数的注释，而所有的 Excel 都是从第二行开始读取的)
        :return: 参数名称[]

        example:
        ['equipmentid', 'exp']
        '''
        return self.getOneline(1)

    def paramAlllineDict(self):
        '''
        获取全部参数
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
        ParamAllListDict = {}
        iRowStep = 2
        iColStep = 0
        ParamHeader = self.paramHeader()
        while iRowStep < nCountRows:
            ParamOneLinelist = self.getOneline(iRowStep)
            ParamOnelineDict = {}
            while iColStep < nCountCols:
                ParamOnelineDict[ParamHeader[iColStep]] = ParamOneLinelist[iColStep]
                iColStep = iColStep+1
            iColStep = 0
            ParamAllListDict[iRowStep-2] = ParamOnelineDict
            iRowStep = iRowStep+1
        return ParamAllListDict

    def paramAllline(self):
        '''
        获取全部参数
        :return: 全部参数[[]]
        '''
        nCountRows = self.paramRowsCount()
        paramall = []
        iRowStep = 2
        while iRowStep < nCountRows:
            paramall.append(self.getOneline(iRowStep))
            iRowStep = iRowStep+1
        return paramall

    def __getParamCell(self, numberRow, numberCol):
        return self.paramsheet.cell_value(numberRow, numberCol)


class ParamFactory(object):
    def chooseParam(self, type, paramConf):
        map_ = {
            'xls': XLS(paramConf)
        }
        return map_[type]


if __name__ == "__main__":
    # uri_login存储战场的选择武器
    uri_selectEq = '/selectEq'
    comm = Common('http://127.0.0.1:12306', api_type='http')
    # 武器编号变量存储武器编号，并且验证返回时是否有参数设计预期结果
    # 获取当前路径绝对值
    curPath = os.path.abspath('.')
    # 定义存储参数的excel文件路径
    searchparamfile = curPath+'/equipmentid_param.xls'
    # 调用参数类完成参数读取，返回是一个字典，包含全部的excel数据除去excel的第一行表头说明
    searchparam_dict = ParamFactory().chooseParam('xls', {'file': searchparamfile, 'sheet': 0}).paramAlllineDict()
    print(searchparam_dict)
    i = 0
    while i < len(searchparam_dict):
        # 读取通过参数类获取的第i行的参数
        payload = 'equipmentid=' + str(int(searchparam_dict[i]['equipmentid']))
        # 读取通过参数类获取的第i行的预期
        exp = searchparam_dict[i]['exp']
        # 进行接口测试
        response_selectEq = comm.post(uri_selectEq, params=payload)
        # 打印返回结果
        print('Response内容：' + response_selectEq.text)
        # 读取下一行excel中的数据
        i = i+1
