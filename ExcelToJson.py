import xlrd  # xlrd版本为1.2.0
import json


def exceltojson(excel_path) -> None:
    skinDict = {"elite1": ".png", "elite2": "_2.png", "skin1": "_skin1.png", "skin2": "_skin2.png", "skin3": "_skin3.png"}
    DrData = {}
    DrDictList = ['Name', 'ID', 'DateIn', 'DateNow', 'Assistant', 'Skin', 'LeftSide',
                  'RightSide', 'TopSide', 'FirstLine', 'Gap', 'SizeName', 'SizeID', 'SizeTotal', 'TotalToRight']
    strDrDictList = ['Name', 'DateIn', 'DateNow', 'Assistant', 'Skin']
    intDrDictList = ['ID', 'LeftSide', 'RightSide', 'TopSide', 'FirstLine', 'Gap', 'SizeName', 'SizeID',
                     'SizeTotal', 'TotalToRight']
    if len(strDrDictList) + len(intDrDictList) != len(DrDictList):
        raise Exception('ExcelToJson.py_Line14: len(strDr) plus len(intDr) isnot equal to len(dr)')
    BoxData = []
    BoxDict = ["Name", "Star", "Potential", "Elite", "Level", "Skill1",
               "Skill2", "Skill3", "Skin", "Mod", "ModLevel", 'ModLight']
    strBoxDict = ["Name", "Skin", "Mod", 'ModLight']
    intBoxDict = ["Star", "Potential", "Elite", "Level", "Skill1", "Skill2", "Skill3", "ModLevel"]
    if len(strBoxDict) + len(intBoxDict) != len(BoxDict):
        raise Exception('ExcelToJson.py_Line21: len(strBox) plus len(intBox) isnot equal to len(Box)')

    ExcelBlockLine = 3  # 表前空行数
    excel = xlrd.open_workbook(excel_path, encoding_override="utf-8")
    AllSheets = excel.sheets()  # 获取所有页对象列表
    UseSheet = AllSheets[0]  # 储存第一页对象
    for j in range(len(DrDictList)):
        if DrDictList[j] in strDrDictList:
            DrData[DrDictList[j]] = str(UseSheet.cell_value(j, 1))
        if DrDictList[j] in intDrDictList:
            DrData[DrDictList[j]] = int(UseSheet.cell_value(j, 1))
    with open('./Source/DrData.json', 'w', encoding='utf-8') as f:
        json.dump(DrData, f, ensure_ascii=False, indent=4)
    UseSheet = AllSheets[1]  # 储存第二页对象
    LineMax = UseSheet.nrows  # 数据行数-1
    j = 0
    for NowLine in range(ExcelBlockLine, LineMax):
        BoxData.append({})
        for k in range(len(BoxDict)):
            if BoxDict[k] in strBoxDict:  # if type[data] is str
                BoxData[j][BoxDict[k]] = str(UseSheet.cell_value(NowLine, k))
                if BoxDict[k] == 'Mod' and BoxData[j][BoxDict[k]] == '':   # if it has no mod
                    BoxData[j][BoxDict[k + 1]] = 0     # set ModLevel to 0
                    BoxData[j][BoxDict[k + 2]] = ''  # set ModLight to ''
                    break
            if BoxDict[k] in intBoxDict:    # if type[data] is str
                # if BoxDict[k] == "ModLevel":   # 如果是模组等级列
                #     if BoxData[j]["Mod"] == "":  # 如果没有模组
                #         BoxData[j][BoxDict[k]] = 0
                #         continue   # 跳过本次
                BoxData[j][BoxDict[k]] = int(UseSheet.cell_value(NowLine, k))
        if not (BoxData[j]['Skin'] in skinDict):
            BoxData[j]['Skin'] = "elite2"  # 缺省默认为精二
        j = j + 1
    with open('./Source/BoxData.json', 'w', encoding='utf-8') as f:
        json.dump(BoxData, f, ensure_ascii=False, indent=4)
