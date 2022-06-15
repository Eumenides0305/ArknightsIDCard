import json
from PIL import Image, ImageFont, ImageDraw  # pillow版本为9.0.0
from ExcelToJson import exceltojson
import ctypes
import re
import random

NumSkillThr = 0
skinDict = {"精一": ".png", "精二": "_2.png", "skin1": "_skin1.png", "skin2": "_skin2.png", "skin3": "_skin3.png"}


def win_info(message: str):
    ctypes.windll.user32.MessageBoxW(0, message, "ArknightsIDCard", 0 | 64)


def only_ascii(rawStr: str) -> bool:
    subedStr = re.sub("[A-Za-z0-9]", "", rawStr)   # 删除数字和英文
    if subedStr == '':   # 没有内容
        return True
    else:
        return False


def translucent_paste(frontImg: Image, backImg: Image, x: int, y: int) -> Image:    # for paste transparent img
    backImgCrop = backImg.crop((x, y, x+frontImg.size[0], y+frontImg.size[1]))  # crop from backImg
    # print(x + frontImg.size[0])
    # print(y + frontImg.size[1])
    cropForPaste = Image.new("RGBA", backImgCrop.size)  # new img
    cropForPaste = Image.alpha_composite(cropForPaste, backImgCrop)  # paste backImgCrop
    cropForPaste = Image.alpha_composite(cropForPaste, frontImg)  # paste front
    a = cropForPaste.split()[3]
    backImg.paste(cropForPaste, (x, y), mask=a)
    return backImg


class BKG:
    def __init__(self, data: dict, num: int):
        self.data = data
        self.num = num
        self.PathOutDrBKG = "./OutPut/" + self.data['Name'] + "_" + self.data['DateNow'] + ".png"
        self.img, self.ShowList = self.bkg_init()

    def bkg_init(self) -> [Image, int]:
        deviationCheck = 0.01    # larger value means more blank in the bottom of output img
        PathBKG = "./Source/BKG.png"
        BKGImg = Image.open(PathBKG)     # ori BKG
        x, y = BKGImg.size
        for numList in range(10, self.num):
            Line = self.num // numList + 1
            outHeight = 206*Line + self.data['Gap']*(Line-1) + self.data['TopSide'] + 180 + self.data["FirstLine"]
            outWide = 180*numList + self.data['Gap']*(numList-1) + self.data['LeftSide'] + self.data['RightSide']
            if y / x - outHeight / outWide > deviationCheck:  # auto match the size of BKG
                outHeight = int(outWide * y / x)
                BKGImg = BKGImg.resize((outWide, outHeight), Image.ANTIALIAS)
                return BKGImg, numList  # output new BKG and num of list

    def addtitle(self) -> None:
        OnlyEn = self.only_ascii_name()
        EnBig = 15
        if OnlyEn:                # DrName font setting of if only [A-Za-z0-9]
            fontName = ImageFont.truetype("./Source/TTF/AcuminProBook.ttf", self.data['SizeName'])
        else:                       # DrName font setting of other situation
            fontName = ImageFont.truetype("./Source/TTF/NotoSansHans-Regular.otf", self.data['SizeName']-5)
        fontDayAndID = ImageFont.truetype("./Source/TTF/AcuminProBook.ttf", self.data['SizeID'])  # date and ID
        fontTotal = ImageFont.truetype("./Source/TTF/BenderRegular.otf", self.data['SizeTotal'])  # total of skill-elite 3
        pathHead = "./Source/Character/头像_" + self.data['Assistant'] + skinDict[self.data['Skin']]  # path of Assistant
        pathHeadBlock = "./Source/Dr头像框.png"
        pathTotal = "./Source/Skill/专精_3.png"  # 总专精数
        skillElite3Big = Image.open(pathTotal).convert("RGBA")  # 打开图标
        a = skillElite3Big.split()[3]
        self.img.paste(skillElite3Big, (self.img.size[0] - self.data['TotalToRight'] - skillElite3Big.size[0] - 50,
                                        self.data['TopSide']), mask=a)  # 粘贴
        New = Image.open(pathHead).convert("RGBA")  # 获取头像
        a = New.split()[3]
        self.img.paste(New, (self.data['LeftSide'], self.data['TopSide']), mask=a)  # 粘贴头像
        New = Image.open(pathHeadBlock).convert("RGBA")  # 获取头像框
        a = New.split()[3]
        self.img.paste(New, (self.data['LeftSide'] - 3, self.data['TopSide'] - 3), mask=a)  # 粘贴头像框
        titleDraw = ImageDraw.Draw(self.img)  # add name
        text = "Dr. " + self.data['Name']
        titleDraw.text((self.data['LeftSide'] + 220, self.data['TopSide'] - 30 + OnlyEn*EnBig), text, font=fontName)
        text = self.data['DateIn'] + "——" + self.data['DateNow'] + "     ID:" + str(self.data['ID'])
        titleDraw.text((self.data['LeftSide'] + 228, self.data['TopSide'] + 137), text, font=fontDayAndID)  # add date
        global NumSkillThr
        text = str(NumSkillThr)
        titleDraw.text((self.img.size[0] - self.data['TotalToRight'], self.data['TopSide'] + 10), text,
                       font=fontTotal)  # add num of skill-elite3
        self.img.save(self.PathOutDrBKG)  # save
        message = '制作完成，输出路径为：' + f'{self.PathOutDrBKG}'
        win_info(message)

    def only_ascii_name(self) -> bool:
        cutWord = "#"
        indexOf_Jing = self.data['Name'].find(cutWord)
        Name = self.data['Name'][:indexOf_Jing]
        return only_ascii(Name)


class GanYuan:
    def __init__(self, data: dict):
        self.data = data
        self.imgPath = ''
        self.add_skill()
        self.add_pot()  # potential
        self.add_level()  # level
        self.add_mod()  # mod and mod level

    def add_skill(self) -> None:
        pathSkill = "./Source/Skill/BKG2.png"
        editing = Image.open(pathSkill)     # skill bkg
        PathOri = "./Source/Character/头像_" + self.data['Name'] + skinDict[self.data['Skin']]
        headIcon = Image.open(PathOri).convert("RGBA")   # ori head icon
        a = headIcon.split()[3]
        editing.paste(headIcon, (0, 0), mask=a)
        PathLevel = "./Source/Skill/BKG3.png"
        levelBkg = Image.open(PathLevel).convert("RGBA")   # level bkg
        editing = translucent_paste(levelBkg, editing, 0, 0)
        # img2 = Image.open(PathLevel).convert("RGBA")   # level bkg
        # final2 = Image.new("RGBA", editing.size)    # 新建画布
        # final2 = Image.alpha_composite(final2, editing)    # 粘贴原图
        # final2 = Image.alpha_composite(final2, img2)    # 粘贴等级背景
        # editing = final2     # 重命名，插进原代码，假装无事发生过
        Skill = [self.data["Skill1"], self.data["Skill2"], self.data["Skill3"]]
        for num in range(3):      # for skill1,2,3
            if Skill[num] == 0:
                continue
            if (Skill[num] < 7) or (self.data['Elite'] != 2):
                skillDraw = ImageDraw.Draw(editing)
                font = ImageFont.truetype("./Source/TTF/BenderBold.otf", 22)
                text = "Skill: " + str(Skill[1])
                skillDraw.text((180 - 24 * 3 + 5, 206 - 26), text, font=font)
                break
            pathSkill = "./Source/Skill/" + str(Skill[num]) + ".png"  # num from 0 to 2 means skill 1 to 3
            skillEliteIcon = Image.open(pathSkill).convert("RGBA")  # skill-elite icon
            a = skillEliteIcon.split()[3]
            editing.paste(skillEliteIcon, (3 + 24 * num, 206 - 22), mask=a)
            global NumSkillThr             # count total skill-elite 3
            if Skill[num] == 10:
                NumSkillThr = NumSkillThr + 1
        self.imgPath = "./OutPut/头像_" + self.data['Name'] + skinDict[self.data['Skin']]
        editing.save(self.imgPath)  # save

    def add_pot(self) -> None:  # potential
        editing = Image.open(self.imgPath)
        PathPot = "./Source/Potential/" + str(self.data['Potential']) + ".png"
        potIcon = Image.open(PathPot).convert("RGBA")  # open potential icon
        a = potIcon.split()[3]
        editing.paste(potIcon, (180 - 68, 0), mask=a)
        editing.save(self.imgPath)  # save

    def add_level(self) -> None:  # level
        editing = Image.open(self.imgPath)
        if self.data['Elite'] != 2:
            PathElite = "./Source/Elite/" + str(self.data['Elite']) + ".png"
            eliteIcon = Image.open(PathElite).convert("RGBA")  # open elite icon
            a = eliteIcon.split()[3]
            editing.paste(eliteIcon, (4, 206 - 22), mask=a)
            editing.save(self.imgPath)  # save
            return
        levelDraw = ImageDraw.Draw(editing)
        font = ImageFont.truetype("./Source/TTF/bender.regular.otf", 12)  # font for 'Lv.'
        text = "Lv."
        levelDraw.text((4, 206 - 62), text, font=font)        # draw 'Lv.'
        font = ImageFont.truetype("./Source/TTF/bender.regular.otf", 22)  # font for level number
        text = str(self.data['Level'])
        levelDraw.text((4, 206 - 52), text, font=font)          # draw level number
        editing.save(self.imgPath)  # save

    def add_mod(self) -> None:   # cover mod and mod level
        modLightDict = {"蓝色": "blue", "绿色": "green", "灰色": "grey", "红色": "red", "金色": "yellow"}
        modLightList = ["蓝色", "绿色", "灰色", "红色", "金色"]
        stdModSize = 42
        if not self.data['Mod'] == '':    # if Mod exist
            editing = Image.open(self.imgPath)    # open editing file

            PathModBKG = "./Source/Mod/BKG.png"
            modBkg = Image.open(PathModBKG).convert("RGBA")  # open mod bkg
            editing = translucent_paste(modBkg, editing, 0, 0)       # paste mod frame

            try:
                pathModLight = './Source/Mod/'+modLightDict[self.data['ModLight']]+'_shining.png'
            except:
                pathModLight = './Source/Mod/' + modLightDict[random.choice(modLightList)] + '_shining.png'
            modLightImg = Image.open(pathModLight).convert("RGBA")  # open mod bkg
            x, y = modLightImg.size  # 取得分辨率
            z = max(x, y)  # 找出长边
            newLightImgX = int(1.5*stdModSize * x/z)   # new size of mod light
            newLightImgY = int(1.5*stdModSize * y/z)
            modLightImg = modLightImg.resize((newLightImgX, newLightImgY), Image.ANTIALIAS)
            editing = translucent_paste(modLightImg, editing, int(150 - newLightImgX/2), int(176 - newLightImgY/2))
            # paste mod light

            PathMod = "./Source/Mod/模组类型_" + self.data['Mod'] + ".png"
            modIcon = Image.open(PathMod).convert("RGBA")  # 打开模组图标
            x, y = modIcon.size     # 取得分辨率
            z = max(x, y)     # 找出长边
            modIcon = modIcon.resize((int(stdModSize * x/z), int(stdModSize * y/z)), Image.ANTIALIAS)  # 统一模组尺寸
            # a = modIcon.split()[3]
            editing.paste(modIcon, (int(150 - stdModSize*x/z/2), int(176 - stdModSize*y/z/2)), mask=modIcon.split()[3])
            # 以模组背景为基准居中粘贴

            pathModLevel = "./Source/Mod/" + str(int(self.data["ModLevel"])) + ".png"
            modLevelIcon = Image.open(pathModLevel).convert("RGBA")  # 打开模组等级图标
            modLevelIcon = modLevelIcon.resize((20, 20), Image.ANTIALIAS)  # standardize size to 20*20
            # a = modLevelIcon.split()[3]
            editing.paste(modLevelIcon, (180 - 20, 206 - 5 - 60), mask=modLevelIcon.split()[3])  # 粘贴

            editing.save(self.imgPath)  # 储


'''here begin to draw '''


def draw(UseJson: bool, pathExcel: str):
    if not UseJson:
        exceltojson(pathExcel)
    with open('./Source/BoxData.json', 'r', encoding='utf-8') as f:
        BoxData = json.load(f)
        f.close()
    with open('./Source/DrData.json', 'r', encoding='utf-8') as f:
        DrData = json.load(f)
        f.close()
    ShowLine = 0
    DrBKG = BKG(DrData, len(BoxData))
    for i in range(len(BoxData)):
        nowGanYuan = GanYuan(BoxData[i])
        if (i % DrBKG.ShowList) == 0:  # 计算小头像所在行
            ShowLine = ShowLine + 1
        ShowX = DrBKG.data['LeftSide'] + (i % DrBKG.ShowList) * (180+DrBKG.data['Gap'])  # 小头像横坐标
        ShowY = DrBKG.data['TopSide'] + ShowLine * (206+DrBKG.data['Gap']) + DrBKG.data['FirstLine'] - 26  # 小头像纵坐标
        EditingImg = Image.open(nowGanYuan.imgPath).convert("RGBA")  # 打开小头像
        DrBKG.img = translucent_paste(EditingImg, DrBKG.img, ShowX, ShowY)   # paste transparent img
        # im1 = DrBKG.img.crop((ShowX,ShowY,ShowX+180,206+ShowY))             # 从大背景上裁切小头像背景
        # EditingImg = Image.open(nowGanYuan.imgPath).convert("RGBA")  # 打开小头像
        # final3 = Image.new("RGBA", im1.size)  # 新建画布
        # final3 = Image.alpha_composite(final3, im1)  # 粘贴原图
        # final3 = Image.alpha_composite(final3, EditingImg)  # 粘贴模组背景
        # EditingImg = final3                       # 重命名，插进原代码，假装无事发生过
        # a = EditingImg.split()[3]
        # DrBKG.img.paste(EditingImg, (ShowX, ShowY), mask=a)
    DrBKG.addtitle()
