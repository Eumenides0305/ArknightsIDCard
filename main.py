import os
import Loginit
from new_draw import draw, win_info

os.makedirs('./OutPut', exist_ok=True)
pathExcel = "./Source/arkD.xlsx"

# code below is for pycharm
draw(UseJson=False, pathExcel=pathExcel)

# code below is for exe
# pyinstaller -F -w main.py
# try:
#     draw(UseJson=False, pathExcel=pathExcel)
# except Exception as error:
#     Path_Log = './draw.log'
#     if os.path.exists(Path_Log):
#         os.remove(Path_Log)
#     message = '运行出错，错误日志路径：'+f'{Path_Log}'
#     win_info(message)
#     logger = Loginit._get_logger(Path_Log, 'info')  # 输出错误日志
#     logger.info(error)

