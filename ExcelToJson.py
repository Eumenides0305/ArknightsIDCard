import xlrd  # xlrd版本为1.2.0
import json

def exceltojson(excel_path) -> None:
    Skindict = {"精一": ".png", "精二": "_2.png", "skin1": "_skin1.png", "skin2": "_skin2.png", "skin3": "_skin3.png"}
    # excel_path = "./Source/arkE.xlsx"  # 数据表位置

    DrData = {}
    DrDict = ['Name', 'ID', 'DateIn', 'DateNow', 'Assistant', 'Skin', 'LeftSide',
              'RightSide', 'TopSide', 'FirstLine', 'Gap', 'SizeName', 'SizeID', 'SizeTotal', 'TotalToRight']
    strDrDict = ['Name', 'DateIn', 'DateNow', 'Assistant', 'Skin']
    intDrDict = ['ID', 'LeftSide', 'RightSide', 'TopSide', 'FirstLine', 'Gap', 'SizeName', 'SizeID', 'SizeTotal', 'TotalToRight']
    if len(strDrDict) + len(intDrDict) != len(DrDict):
        print('error')
        quit()
    BoxData = []
    BoxDict = ["Name", "Star", "Potential", "Elite", "Level", "Skill1",
               "Skill2", "Skill3", "Mod", "ModLevel", "Skin"]
    strBoxDict = ["Name", "Mod", "Skin"]
    intBoxDict = ["Star", "Potential", "Elite", "Level", "Skill1", "Skill2", "Skill3", "ModLevel"]
    if len(strBoxDict) + len(intBoxDict) != len(BoxDict):
        print('error')
        quit()

    ExcelBlockLine = 3  # 表前空行数
    excel = xlrd.open_workbook(excel_path, encoding_override="utf-8")
    AllSheets = excel.sheets()  # 获取所有页对象列表
    UseSheet = AllSheets[0]  # 储存第一页对象
    for i in range(len(DrDict)):
        if DrDict[i] in strDrDict:
            DrData[DrDict[i]] = str(UseSheet.cell_value(i, 1))
        if DrDict[i] in intDrDict:
            DrData[DrDict[i]] = int(UseSheet.cell_value(i, 1))
    with open('./Source/DrData.json', 'w', encoding='utf-8') as f:
        json.dump(DrData, f, ensure_ascii=False, indent=4)
    UseSheet = AllSheets[1]  # 储存第二页对象
    LineMax = UseSheet.nrows  # 数据行数-1
    i = 0
    for NowLine in range(ExcelBlockLine, LineMax):
        BoxData.append({})
        for j in range(len(BoxDict)):
            if BoxDict[j] in strBoxDict:
                BoxData[i][BoxDict[j]] = str(UseSheet.cell_value(NowLine, j))
            if BoxDict[j] in intBoxDict:
                if BoxDict[j] == "ModLevel":   # 如果是模组等级列
                    if BoxData[i]["Mod"] == "":  # 如果没有模组
                        BoxData[i][BoxDict[j]] = 0
                        continue   # 跳过本次
                BoxData[i][BoxDict[j]] = int(UseSheet.cell_value(NowLine, j))
        if not (BoxData[i]['Skin'] in Skindict):
            BoxData[i]['Skin'] = "精二"  # 缺省默认为精二
        i = i + 1
    with open('./Source/BoxData.json', 'w', encoding='utf-8') as f:
        json.dump(BoxData, f, ensure_ascii=False, indent=4)

