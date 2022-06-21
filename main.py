import os
import win
from new_draw import draw

os.makedirs('./OutPut', exist_ok=True)
pathExcel = win.choose_excel(defaultOpenPath='./Source')
if pathExcel == '':
    win.info('ERROR Excel path! Program exit.')
    quit()

# code below is for pycharm
draw(UseJson=False, pathExcel=pathExcel)

# pyinstaller -F -w main.py
# code below is for exe
# try:
#     draw(UseJson=False, pathExcel=pathExcel)
# except Exception as error:
#     Path_Log = './draw.log'
#     if os.path.exists(Path_Log):
#         os.remove(Path_Log)
#     message = '运行出错，错误日志路径：'+f'{Path_Log}'
#     win.info(message)
#     logger = win.get_logger(Path_Log, 'info')  # 输出错误日志
#     logger.info(error)

