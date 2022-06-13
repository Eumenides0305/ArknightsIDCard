import json
from PIL import Image, ImageFont, ImageDraw  # pillow版本为9.0.0
from ExcelToJson import exceltojson
import ctypes
import re
import Loginit
import os


def win_info(message: str):
    ctypes.windll.user32.MessageBoxW(0, message, "ArknightsIDCard", 0 | 64)


def only_ascii(raw_str: str) -> bool:
    subed_str = re.sub("[A-Za-z0-9]", "", raw_str)  # 删除数字和英文
    if subed_str == '':  # 没有内容
        return True
    else:
        return False


Path_Excel = "./Source/arkE.xlsx"
Skindict = {"精一": ".png", "精二": "_2.png", "skin1": "_skin1.png", "skin2": "_skin2.png", "skin3": "_skin3.png"}
NumSkillThr = 0
UseJson = 0
BKGcheck = 0.01


class BKG:
    def __init__(self, data: dict, num: int) -> None:
        self.data = data
        self.num = num
        self.PathOutDrBKG = "./OutPut/" + self.data['Name'] + "_" + self.data['DateNow'] + ".png"
        self.img, self.ShowList = self.BKGInit()

    def BKGInit(self):
        PathBKG = "./Source/BKG.png"  # 大背景原图文件
        BKGImg = Image.open(PathBKG)
        x, y = BKGImg.size
        for List in range(10, self.num):
            Line = self.num // List + 1
            OutHight = 206 * Line + self.data['Gap'] * (Line - 1) + self.data['TopSide'] + 180 + self.data["FirstLine"]
            OutWide = 180 * List + self.data['Gap'] * (List - 1) + self.data['LeftSide'] + self.data['RightSide']
            if y / x - OutHight / OutWide > BKGcheck:  # 自动适配背景图大小
                OutHight = int(OutWide * y / x)
                BKGImg = BKGImg.resize((OutWide, OutHight), Image.ANTIALIAS)
                return BKGImg, List  # 输出处理后的背景图和干员列数

    def addtitle(self) -> None:
        OnlyEn = self.only_ascii_name()
        EnBig = 15
        if OnlyEn:  # 设置博士名称字体
            FontName = ImageFont.truetype("./Source/TTF/AcuminProBook.ttf", self.data['SizeName'])  # 英文和数字
        else:
            FontName = ImageFont.truetype("./Source/TTF/NotoSansHans-Regular.otf", self.data['SizeName'] - 5)  # 其他
        FontDayAndID = ImageFont.truetype("./Source/TTF/AcuminProBook.ttf", self.data['SizeID'])  # 设置日期和博士id字体
        FontTotal = ImageFont.truetype("./Source/TTF/BenderRegular.otf", self.data['SizeTotal'])  # 总专三数
        PathHead = "./Source/Character/头像_" + self.data['Assistant'] + Skindict[self.data['Skin']]  # 助手头像文件
        PathHeadBlock = "./Source/Dr头像框.png"
        PathTotal = "./Source/Skill/专精_3.png"  # 总专精数
        New = Image.open(PathTotal).convert("RGBA")  # 打开图标
        a = New.split()[3]
        self.img.paste(New, (self.img.size[0] - self.data['TotalToRight'] - New.size[0] - 50, self.data['TopSide']),
                       mask=a)  # 粘贴
        New = Image.open(PathHead).convert("RGBA")  # 获取头像
        a = New.split()[3]
        self.img.paste(New, (self.data['LeftSide'], self.data['TopSide']), mask=a)  # 粘贴头像
        New = Image.open(PathHeadBlock).convert("RGBA")  # 获取头像框
        a = New.split()[3]
        self.img.paste(New, (self.data['LeftSide'] - 3, self.data['TopSide'] - 3), mask=a)  # 粘贴头像框
        draw = ImageDraw.Draw(self.img)  # 准备添加文字
        text = "Dr. " + self.data['Name']
        draw.text((self.data['LeftSide'] + 220, self.data['TopSide'] - 30 + OnlyEn * EnBig), text, font=FontName)  # 写名字
        text = self.data['DateIn'] + "——" + self.data['DateNow'] + "     ID:" + str(self.data['ID'])
        draw.text((self.data['LeftSide'] + 228, self.data['TopSide'] + 137), text, font=FontDayAndID)  # 日期
        global NumSkillThr
        text = str(NumSkillThr)
        draw.text((self.img.size[0] - self.data['TotalToRight'], self.data['TopSide'] + 10), text,
                  font=FontTotal)  # 专三数
        self.img.save(self.PathOutDrBKG)  # 储存
        message = '制作完成，输出路径为：' + f'{self.PathOutDrBKG}'
        win_info(message)

    def only_ascii_name(self) -> bool:
        cut_word = "#"
        IndexOf_Jing = self.data['Name'].find(cut_word)
        Name = self.data['Name'][:IndexOf_Jing]
        return only_ascii(Name)


class GanYuan:
    def __init__(self, data: dict) -> None:
        self.data = data
        self.EditingImg = ''
        self.addskill()
        self.addpot()  # 添加潜能标识
        self.addlevel()  # 添加等级
        self.addmod()  # 添加模组

    def addskill(self) -> None:
        PathSkill = "./Source/Skill/BKG2.png"  # 获取技能背景
        img = Image.open(PathSkill)
        PathOri = "./Source/Character/头像_" + self.data['Name'] + Skindict[self.data['Skin']]  # 获取原始头像
        New = Image.open(PathOri).convert("RGBA")  # 打开原始头像
        a = New.split()[3]
        img.paste(New, (0, 0), mask=a)  # 粘贴
        Skill = [self.data["Skill1"], self.data["Skill2"], self.data["Skill3"]]
        for num in range(3):
            if Skill[num] == 0:
                continue
            if (Skill[num] < 7) or (self.data['Elite'] != 2):
                draw = ImageDraw.Draw(img)
                font = ImageFont.truetype("./Source/TTF/BenderBold.otf", 22)
                text = "Skill: " + str(Skill[1])
                draw.text((180 - 24 * 3 + 5, 206 - 26), text, font=font)
                break
            PathSkill = "./Source/Skill/" + str(Skill[num]) + ".png"  # num从0到2，0代表技能1，类推
            New = Image.open(PathSkill).convert("RGBA")  # 打开技能图标
            a = New.split()[3]
            img.paste(New, (180 - 24 * (3 - num), 206 - 22), mask=a)  # 粘贴
            global NumSkillThr
            if Skill[num] == 10:
                NumSkillThr = NumSkillThr + 1
        self.EditingImg = "./OutPut/头像_" + self.data['Name'] + Skindict[self.data['Skin']]
        img.save(self.EditingImg)  # 储存

    def addpot(self) -> None:  # 添加潜能
        img = Image.open(self.EditingImg)
        PathPot = "./Source/Potential/" + str(self.data['Potential']) + ".png"
        New = Image.open(PathPot).convert("RGBA")  # 打开图标
        a = New.split()[3]
        img.paste(New, (180 - 68, 0), mask=a)  # 粘贴
        img.save(self.EditingImg)  # 储存

    def addlevel(self) -> None:  # 添加等级
        img = Image.open(self.EditingImg)
        if self.data['Elite'] != 2:
            PathElite = "./Source/Elite/" + str(self.data['Elite']) + ".png"
            New = Image.open(PathElite).convert("RGBA")  # 打开图标
            a = New.split()[3]
            img.paste(New, (4, 206 - 22), mask=a)  # 粘贴
            img.save(self.EditingImg)  # 储存
            return
        draw = ImageDraw.Draw(img)
        font = ImageFont.truetype("./Source/TTF/BenderBold.otf", 22)
        text = "Lv." + str(self.data['Level'])
        draw.text((4, 206 - 26), text, font=font)
        img.save(self.EditingImg)  # 储存

    def addmod(self) -> None:  # 包含模组和模组等级
        if not self.data['Mod'] == '':
            img = Image.open(self.EditingImg)
            PathMod = "./Source/Mod/" + self.data['Name'] + ".png"   # 这里后续应该要改成模组名称
            Newmod = Image.open(PathMod).convert("RGBA")  # 这里是在载入模组
            x, y = Newmod.size
            a = Newmod.split()[3]
            img.paste(Newmod, (180 - x, 180 - y), mask=a)  # 粘贴
            PathModLevel = "./Source/Mod/" + str(int(self.data["ModLevel"])) + ".png"
            Newmod = Image.open(PathModLevel).convert("RGBA")  # 这里是在载入模组等级
            Newmod = Newmod.resize((30, 30), Image.ANTIALIAS)   # 重置模组等级图片大小为30*30
            a = Newmod.split()[3]
            img.paste(Newmod, (180 - 30, 180 - y - 10), mask=a)  # 粘贴
            img.save(self.EditingImg)  # 储


'''
这里开始是main
'''


def draw():
    os.makedirs('./OutPut', exist_ok=True)
    if not UseJson:
        exceltojson(Path_Excel)
    with open('./Source/BoxData.json', 'r', encoding='utf-8') as f:
        BoxData = json.load(f)
    with open('./Source/DrData.json', 'r', encoding='utf-8') as f:
        DrData = json.load(f)
    ShowLine = 0
    DrBKG = BKG(DrData, len(BoxData))
    for i in range(len(BoxData)):
        nowganyuan = GanYuan(BoxData[i])
        if (i % DrBKG.ShowList) == 0:  # 计算小头像所在行
            ShowLine = ShowLine + 1
        ShowX = DrBKG.data['LeftSide'] + (i % DrBKG.ShowList) * (180 + DrBKG.data['Gap'])  # 小头像横坐标
        ShowY = DrBKG.data['TopSide'] + ShowLine * (206 + DrBKG.data['Gap']) + DrBKG.data['FirstLine'] - 26  # 小头像纵坐标
        EditingImg = Image.open(nowganyuan.EditingImg).convert("RGBA")  # 打开小头像
        a = EditingImg.split()[3]
        DrBKG.img.paste(EditingImg, (ShowX, ShowY), mask=a)
    DrBKG.addtitle()


try:
    draw()
except Exception as error:
    Path_Log = './draw.log'
    if os.path.exists(Path_Log):
        os.remove(Path_Log)
    message = '运行出错，错误日志路径：'+f'{Path_Log}'
    win_info(message)
    logger = Loginit._get_logger(Path_Log, 'info')  # 输出错误日志
    logger.info(error)
