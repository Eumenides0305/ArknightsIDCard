import xlrd  # xlrd版本为1.2.0
from PIL import Image, ImageFont, ImageDraw  # pillow版本为9.0.0

Skindict = {"精一": ".png", "精二": "_2.png", "skin1": "_skin1.png", "skin2": "_skin2.png", "skin3": "_skin3.png"}
excel_path = "./Source/arkE.xlsx"  # 数据表位置


def AddPot(img):  # 添加潜能
    PathPot = "./Source/Potential/" + str(Potential) + ".png"
    New = Image.open(PathPot).convert("RGBA")  # 打开潜能图标
    r, g, b, a = New.split()
    img.paste(New, (180 - 68, 0), mask=a)  # 粘贴
    PathOut = "./OutPut/头像_" + str(Charactor) + Skindict[Skin]
    img.save(PathOut)  # 另存为


def AddSkill(img):  # 添加技能专精
    PathOri = "./Source/Charactor/头像_" + str(Charactor) + Skindict[Skin]  # 获取原始头像
    New = Image.open(PathOri).convert("RGBA")  # 打开原始头像
    r, g, b, a = New.split()
    img.paste(New, (0, 0), mask=a)  # 粘贴
    for num in range(3):
        if Skill[num] == 0:
            continue
        if (Skill[num] < 7) or (Elite != 2):
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype("./Source/TTF/BenderBold.otf", 22)
            text = "Skill: " + str(Skill[1])
            draw.text((180 - 24 * 3 + 5, 206 - 26), text, font=font)
            break
        PathSkill = "./Source/Skill/" + str(Skill[num]) + ".png"  # num从0到2，0代表技能1，类推
        New = Image.open(PathSkill).convert("RGBA")  # 打开技能图标
        r, g, b, a = New.split()
        img.paste(New, (180 - 24 * (3 - num), 206 - 22), mask=a)  # 粘贴
        global NumSkillThr
        if Skill[num] == 10:
            NumSkillThr = NumSkillThr + 1
    PathOut = "./OutPut/头像_" + str(Charactor) + Skindict[Skin]
    img.save(PathOut)  # 储存


def AddLevel(img):  # 添加等级
    if Elite != 2:
        PathMod = "./Source/Elite/" + str(Elite) + ".png"
        New = Image.open(PathMod).convert("RGBA")  # 打开精英化图标
        r, g, b, a = New.split()
        img.paste(New, (4, 206 - 22), mask=a)  # 粘贴
        PathOut = "./OutPut/头像_" + str(Charactor) + Skindict[Skin]
        img.save(PathOut)  # 储存
        return
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype("./Source/TTF/BenderBold.otf", 22)
    text = "Lv." + str(Level)
    draw.text((4, 206 - 26), text, font=font)
    PathOut = "./OutPut/头像_" + str(Charactor) + Skindict[Skin]
    img.save(PathOut)  # 储存


def AddMod(img):  # 添加模组
    PathMod = "./Source/Mod/" + str(Charactor) + ".png"
    Newmod = Image.open(PathMod).convert("RGBA")  # 打开模组图标
    x, y = Newmod.size
    r, g, b, a = Newmod.split()
    img.paste(Newmod, (180 - x, 180 - y), mask=a)  # 粘贴
    PathOut = "./OutPut/头像_" + str(Charactor) + Skindict[Skin]
    img.save(PathOut)  # 储存


def AddToBKG(img):
    PathMod = "./Source/Mod/" + str(Charactor) + ".png"


def BKGInit(Left, Right, Top, Num):    # 输入图片的左边距，右边距，上边距，干员数量
    PathBKG = "./Source/BKG.png"  # 大背景原图文件
    BKGImg = Image.open(PathBKG)
    x, y = BKGImg.size
    for List in range(10, Num):
        Line = Num//List+1
        OutHight = 226*Line+Top+276
        OutWide = 200*(List+1)+Left+Right
        if OutHight/OutWide < y/x:   # 自动适配背景图大小
            OutWide = int(OutHight*x/y)
            BKGImg = BKGImg.resize((OutWide, OutHight), Image.ANTIALIAS)
            return BKGImg, List+1      # 输出处理后的背景图和干员列数




# #### 主函数开始 ################################
# #### 数据预设 ##################################
Skill = [0, 0, 0]
NumElite = 0
NumSkillThr = 0
ShowLine = 0
ExcelBlockLine = 3  # 表前空行数
PathHeadBlock = "./Source/Dr头像框.png"  # 博士头像框文件，使用外框需要修改框粘贴坐标，横纵各减小3

excel = xlrd.open_workbook(excel_path, encoding_override="utf-8")
AllSheets = excel.sheets()  # 获取所有页对象列表
UseSheet = AllSheets[1]  # 储存第二页对象
Name = UseSheet.cell_value(0, 1)  # 读取博士名
DrID = str(int(UseSheet.cell_value(1, 1)))  # 读取id
TimeIn = str(UseSheet.cell_value(2, 1))  # 入职时间
TimeNow = str(UseSheet.cell_value(3, 1))  # 制图时间
Assistant = UseSheet.cell_value(4, 1)  # 助手名称
Skin = UseSheet.cell_value(5, 1)  # 助手皮肤
PathHead = "./Source/Charactor/头像_" + Assistant + Skindict[Skin]  # 助手头像文件
# ShowList = int(UseSheet.cell_value(6, 1))  # 设置每行展示多少个
LeftSide = int(UseSheet.cell_value(7, 1))  # 左侧边距
RightSide = int(UseSheet.cell_value(8, 1))  # 右侧边距
TopSide = int(UseSheet.cell_value(9, 1))  # 上边距
SizeName = int(UseSheet.cell_value(10, 1))  #
SizeID = int(UseSheet.cell_value(11, 1))  #
SizeTotal = int(UseSheet.cell_value(12, 1))  #
PathOutALL = "./OutPut/" + Name + "_" + TimeNow + ".png"
# 输出图片位置，程序执行中会在这个OutPut目录生成大批小头像，运行完后可删可不删，不影响下次执行
# # 中文用：NotoSansHans-Regular.otf    英文用：AcuminProBook.ttf
FontName = ImageFont.truetype("./Source/TTF/NotoSansHans-Regular.otf", SizeName)  # 设置博士名称字体
FontDayAndID = ImageFont.truetype("./Source/TTF/AcuminProBook.ttf", SizeID)  # 设置日期和博士id字体
FontTotal = ImageFont.truetype("./Source/TTF/AcuminProBook.ttf", SizeTotal)  # 总专三数
UseSheet = AllSheets[0]  # 储存第一页对象
LineMax = UseSheet.nrows  # 数据行数-1
ListMax = UseSheet.ncols  # 数据列数-1
# 大背景图初始化
AllImg, ShowList = BKGInit(LeftSide, RightSide, TopSide, LineMax+1-ExcelBlockLine)


# 循环录入数据, 制作小头像, 添加至大背景图
for i in range(1):
    for NowLine in range(ExcelBlockLine, LineMax):
        Charactor = UseSheet.cell_value(NowLine, 0)  # 读取角色名(第0+1列)
        Star = UseSheet.cell_value(NowLine, 1)  # 读取角色星级(第1+1列)
        Potential = int(UseSheet.cell_value(NowLine, 2))  # 读取潜能(第2+1列)
        Elite = int(UseSheet.cell_value(NowLine, 3))  # 读取精英等级(第3+1列)
        Level = int(UseSheet.cell_value(NowLine, 4))  # 读取等级(第4+1列)
        Skill[0] = int(UseSheet.cell_value(NowLine, 5))  # 读取技能一(第5+1列)
        Skill[1] = int(UseSheet.cell_value(NowLine, 6))  # 读取技能二(第6+1列)
        Skill[2] = int(UseSheet.cell_value(NowLine, 7))  # 读取技能三(第7+1列)
        Mod = UseSheet.cell_value(NowLine, 8)  # 读取模组(第8+1列)
        Skin = UseSheet.cell_value(NowLine, 9)  # 读取皮肤(第9+1列)
        if not (Skin in Skindict):
            Skin = "精二"  # 缺省默认为精二头像
        PathSkill = "./Source/Skill/BKG2.png"  # 获取技能背景
        OriImg = Image.open(PathSkill)
        AddSkill(OriImg)  # 添加技能专精
        PathEditing = "./OutPut/头像_" + str(Charactor) + Skindict[Skin]  # 此后使用另存的头像
        EditingImg = Image.open(PathEditing)
        AddPot(EditingImg)  # 添加潜能标识
        AddLevel(EditingImg)  # 添加等级
        if Mod == 1:
            AddMod(EditingImg)  # 添加模组
        # 以上小头像完成，下面添加至大背景图

        # AddToBKG(EditingImg)
        EditingImg = Image.open(PathEditing).convert("RGBA")  # 打开小头像
        r, g, b, a = EditingImg.split()
        if ((NowLine + 1 - ExcelBlockLine) % ShowList) == 1:  # 计算小头像所在行
            ShowLine = ShowLine + 1
        ShowX = LeftSide + (NowLine - ExcelBlockLine) * 200 - (ShowLine - 1) * 200 * ShowList  # 计算小头像横坐标，左边距加已有头像个数减所在行
        ShowY = TopSide + 50 + ShowLine * 226  # 计算小头像纵坐标，上边距加所在行
        AllImg.paste(EditingImg, (ShowX, ShowY), mask=a)  # 每张小头占180*180，左右间隙20，即单个头像占200*200
    AllImg.save(PathOutALL)  # 储存
    print("干员添加完成，准备添加ID")
    AllImg = Image.open(PathOutALL)  # 打开准备添加id
    New = Image.open(PathHead).convert("RGBA")  # 获取头像
    r, g, b, a = New.split()
    AllImg.paste(New, (LeftSide, TopSide), mask=a)  # 粘贴头像
    New = Image.open(PathHeadBlock).convert("RGBA")  # 获取头像框
    r, g, b, a = New.split()
    AllImg.paste(New, (LeftSide - 3, TopSide - 3), mask=a)  # 粘贴头像框
    PathTotal = "./Source/Skill/专精_3.png"  # 总专精数
    New = Image.open(PathTotal).convert("RGBA")  # 打开图标
    r, g, b, a = New.split()
    AllImg.paste(New, (LeftSide + 1750, TopSide), mask=a)  # 粘贴
draw = ImageDraw.Draw(AllImg)  # 准备添加文字
text = "Dr. " + Name
draw.text((LeftSide + 220, TopSide - 15), text, font=FontName)  # 写名字
text = TimeIn + "——" + TimeNow + "     ID:" + DrID
draw.text((LeftSide + 228, TopSide + 137), text, font=FontDayAndID)  # 日期
text = str(NumSkillThr)
draw.text((LeftSide + 1950, TopSide + 30), text, font=FontTotal)  # 专三数
AllImg.save(PathOutALL)  # 储存
print("完成")
