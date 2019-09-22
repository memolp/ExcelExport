# -*- coding:utf-8 -*-

"""
将Excel中的配置转成Lua配置
"""

import os
import sys
import xlrd

class DimensionData:
    """ """
    def __init__(self,dimension=1):
        """ """
        self.mDimension = dimension
        self.mValue = {}

    def __setDictValue(self,dictData,keyValueList):
        """ """
        key = keyValueList[0][1]
        dictData[key] = keyValueList
        return dictData

    def __setDimensionValue(self,dictData,keyValueList,dimension=1):
        """ """
        key = keyValueList[0][1]
        if dimension == 1:
            dictData = self.__setDictValue(dictData,keyValueList)
        elif dimension == 21: #二维字典{100:{1:{},2:{}}}
            if not dictData.has_key(key):
                dictData[key] = {}
            temp_dict = self.__setDictValue({},keyValueList[1:])
            dictData[key].update(temp_dict)
        elif dimension == 22: #二维列表{100:{{1},{2}}}
            if not dictData.has_key(key):
                dictData[key] = []
            temp_dict = self.__setDictValue({},keyValueList[1:])
            dictData[key].append(temp_dict)
        elif dimension == 31: # 三维字典{100:{1:{1:{},2:{}},2:{}}}
            if not dictData.has_key(key):
                dictData[key] = {}
            dictData[key].update(self.__setDimensionValue(dictData[key],keyValueList[1:],21))
        elif dimension == 32: #三维列表{100:{1:{{1},{2}}}}
            if not dictData.has_key(key):
                dictData[key] = {}
            dictData[key].update(self.__setDimensionValue(dictData[key],keyValueList[1:],22))
        return dictData

    def setColValues(self,keyValueList):
        """ """
        self.__setDimensionValue(self.mValue,keyValueList,self.mDimension)

    def getValue(self):
        """ """
        return self.mValue

class ExcelSheetParser(object):
    """
    sheet页导出器
    """
    def __init__(self,sheet,export_dir):
        """ sheet 和 导出的路径 """
        self.mSheet = sheet
        self.mExportDir = export_dir

    def GetCellValue(self,cell,ktype):
        """ """
        # 根据字段类型去调整数值 如果为空值 依据字段类型 填上默认值
        if ktype == 'string':
            if cell.ctype == 0:
                return '\"\"'
            else:
                return '\"%s\"' % (cell.value)
        elif ktype == 'int':
            if cell.ctype == 0:
                return -1
            else:
                return int(cell.value)
        elif ktype == 'float':
            if cell.ctype == 0:
                return -1
            else:
                return float(cell.value)
        elif ktype == 'table':
            if cell.ctype == 0:
                return "{}"
            else:
                return cell.value
        else:
            return cell.value

    def GetExportName(self):
        """ 导出的配置文件 """
        try:
            name = self.mSheet.cell_value(0,0)
            if name != "name" :
                return None
            luaConfigName = self.mSheet.cell_value(0,1) 
            if luaConfigName == "":
                return None
            return luaConfigName
        except Exception as e:
            return None
        
    def GetExportDesc(self):
        """ 获取描述 """
        desc = self.mSheet.cell_value(1,0)
        if desc != "desc":
            return ""
        luaConfigDesc = self.mSheet.cell_value(1,1)
        return luaConfigDesc or ""

    def GetDimension(self):
        """ 获取表的维度 """
        dimension = self.mSheet.cell_value(2,0)
        if dimension != "dimension":
            return 1
        luaDimension = self.mSheet.cell_value(2,1)
        if luaDimension == "":
            return 1
        return int(luaDimension)

    def ParserExcel(self,dimension,startRow,endRow,startCol,endCol,keyList):
        """ 
        startCol 开始的列数
        startRow 开始的函数
        """
        excel_data_dict = DimensionData(dimension)
        for row in range(startRow,endRow):
            # 保存数据索引 默认第一列为id
            cell_id = self.mSheet.cell(row, startCol)
            keyValue = self.GetCellValue(cell_id,keyList[startCol][1])
            #assert cell_id.ctype == 2, "found a invalid id in row [%d] !~" % (row)
            row_data_list = []
            for i in range(startCol,endCol):
                cell  = self.mSheet.cell(row,i)
                key,ktype = keyList[i]
                v = self.GetCellValue(cell,ktype)
                #print("kkkk",key,v)
                # 加入列表
                row_data_list.append([key, v])
            excel_data_dict.setColValues(row_data_list)
        return excel_data_dict.getValue()

    def WriteDateToLua(self,luaConfigName,luaConfigDesc,excel_data_dict):
        """ """
        config_path = os.path.join(self.mExportDir,"%s.lua"%luaConfigName)
        config_pf = open(config_path,"w")
        #config_pf.write("-- %s"%luaConfigDesc)
        config_pf.write("%s = {} \n"%luaConfigName)
        config_pf.write("%s.data={ \n"%luaConfigName)
        def writeDict(config_pf,excel_dict,isList = False):
            # 遍历excel数据字典 按格式写入
            for k , v in excel_dict.items():
                if isList :
                    config_pf.write('        {')
                else:
                    config_pf.write('    [%s]={' % k)
                if isinstance(v,list or tuple):
                    for row_data in v:
                        if isinstance(row_data,dict):
                            writeDict(config_pf,row_data,True)
                        elif isinstance(row_data,list or tuple):
                            if isinstance(row_data[1],unicode):
                                row_data[1] = row_data[1].encode("utf-8")
                            if isinstance(row_data[0],unicode or str):
                                config_pf.write('["{0}"]={1},'.format(row_data[0], row_data[1]))
                            else:
                                config_pf.write('[{0}]={1},'.format(row_data[0], row_data[1]))
                elif isinstance(v,dict):
                    writeDict(config_pf,v,False)
                config_pf.write(' },\n')
        writeDict(config_pf,excel_data_dict)
        config_pf.write('}\n')
        config_pf.close()
        
    def Export(self):
        """ """
        #配置名
        luaConfigName = self.GetExportName()
        if not luaConfigName :
            return
        #描述
        luaConfigDesc = self.GetExportDesc()
        #维度
        dimension = self.GetDimension()

        keyList = []
        for i in range(self.mSheet.ncols):
            key = self.mSheet.cell_value(4,i)
            ktype = self.mSheet.cell_value(5,i)
            keyList.append((key,ktype))

        excel_data_dict = self.ParserExcel(dimension,6,self.mSheet.nrows,0,self.mSheet.ncols,keyList)
        #print(excel_data_dict)
        self.WriteDateToLua(luaConfigName,luaConfigDesc,excel_data_dict)

def OpenExcel(path):
    """ """
    try:
        workbook = xlrd.open_workbook(path)
        return workbook
    except Exception as e:
        print(e)
    return None

def CloseExcel(workbook):
    """ """
    pass

# workbook = OpenExcel(u"配置表.xlsx")
# # s = ExcelSheetParser(workbook.sheets()[2],"./")
# # s.Export()
# for sheet in workbook.sheets():
#    s = ExcelSheetParser(sheet,"./")
#    s.Export()
