# ArknightsIDCard
明日方舟长名片制作器

使用说明：  
依照source里的excel编辑自己的练度，注意干员名不能有错和表有两页  
背景图片可以自行更换，确保文件名为BKG，格式为png即可，分辨率不做要求  
然后运行main.py或者main.exe  
选择自己的excel文件，取消选择三次视为取消制图  
必要资源包括整个source文件夹
  
更新说明：  
2022.5.30更新：添加了背景图大小自适应和每行干员数量自适应  
2022.6.8更新：完全重构了main->new_draw，增加了使用json录入数据的方式（格式见BoxData和DrData  
2022.6.8更新：增加了预览图与参数说明（见图），增加了完成提示  
2022.6.8更新：内置了两种字体，一种用于纯英和数字的博士名称，一种用于含有其他字符的博士名称  
2022.6.8更新：上传了exe文件，供无python环境使用  
2022.6.11更新：添加了模组等级模块（涉及文件source/mod，excel，json和代码）
2022.6.11更新：重新上传了exe  
2022.6.13更新：增加日志，重新设计了考虑模组等级的小头像版式  
2022.6.13更新：自动创建OutPut文件夹  
2022.6.14更新: 1. New layout of operators' card;  
2022.6.14更新: 2. New mod naming with a new format of the .xlsx document;  
2022.6.14更新: 3. Added mod level system;  
2022.6.14更新: 4. Added New fonts;  
2022.6.14更新: 5. Uploaded the resources updated in the latest CN server version patch (Notice: The mod icon of Mystic and Guardian has not yet been obtained, temporarily replace them with the icons of other professions)  

2022.6.15Update: Add mod light sys,standardize code  
2022.6.15Update: 1. Added rarity distinction for operators  
2022.6.15Update: 2. The parameters in excel are unified in English  
2022.6.15Update: 3. Let operator bottom sidebar be translucent  
2022.6.15Update: 4. Created a new xlsx file named arkD so that the two data sets will not overwrite each other  

2022.6.19Update: 1. Added rarity display, optimized level font style, added sub-function in new_draw for text with shadow  

2022.6.22Update: 1. Made level font editable, more detailed error report(to the line in excel)  
2022.6.22Update: 2. New excel format, fixed bug of elite-1 or elite-0   
2022.6.22Update: 3. Added excel selector, Upload new format of output sample

2023.5.8Update: The author is addicted to Star Rail and has done nothing recently.

下一步计划：  
做一点算一点吧第一次写程序边学边做  
1. 模组资源改成模组名称，通过模组名称获取( finished  
2. Complete error report sys, send more error details to .log  
3. Adapt more kinds of BKG image (only .png can be used now.)
4. 双模组显示
5. 特限模组的紫色背景
6. 玩家头像库扩充
7. 时装保有数
8. 雇佣干员进度
9. 模组展示卡
10. 名片皮肤
11. 专九标志


  
