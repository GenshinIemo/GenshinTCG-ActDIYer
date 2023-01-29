import PIL
from PIL import Image, ImageFilter, ImageFont, ImageDraw

import xlrd

import basicdata as bdt
import act

half = bdt.half
shorter = bdt.shorter
colorDict = bdt.colorDict
elements = bdt.elements
SkipRule = []
UsingRule = dict()

#读取Excel数据
excel = xlrd.open_workbook("行动牌.xls")
table = excel.sheet_by_index(1)
LineNum = table.nrows
ColNum = table.ncols

DescriptionTitle = ImageFont.truetype("./GenshinFont.ttf", 18)
DescriptionIntro = ImageFont.truetype("./GenshinFont.ttf", 14)

#读取基本数据
BasicHead = Image.open("./basic_resource/ruleUI/head.png")
BasicBody = Image.open("./basic_resource/ruleUI/body.png")
BasicEnd = Image.open("./basic_resource/ruleUI/end.png")

BasicBody = BasicBody.resize((238,20))
BodyDraw = ImageDraw.Draw(BasicBody)
BodyDraw.line([(234,0),(234,24)],"#313B46")

#开始结束
beg = bdt.settings[0] - 1
end = bdt.settings[1]
if(bdt.settings[1] == '*' or bdt.settings[1] > LineNum):
    end = LineNum
        
#逐字印刷
debugger = ""
def WriteChar(img, txt, X, AddedLines, color = colorDict['HS'], IfLined = False):
    rec = X
    global debugger 
    debugger = debugger + txt
    draw = ImageDraw.Draw(img)
        
    draw.text((X,92+AddedLines*19), txt, font=DescriptionIntro, fill=color)
    if(IfLined):
        if(txt not in bdt.half):
            ul = Image.new("RGB",(14,2),color=color)
            img.paste(ul,(X,88+(AddedLines+1)*19))
        else:
            if(txt in bdt.shorter):
                ul = Image.new("RGB",(7,2),color=color)
                img.paste(ul,(X,88+(AddedLines+1)*19))
            elif(txt == '.'):
                ul = Image.new("RGB",(8,2),color=color)
                img.paste(ul,(X,88+(AddedLines+1)*29))
            else:
                ul = Image.new("RGB",(16,2),color=color)
                img.paste(ul,(X,88+(AddedLines+1)*29))
        
        
    if(txt not in bdt.half):
        X += 14
    else:
        if(txt in bdt.shorter):
            X += 7
        elif(txt == '.'):
            X += 4
        else:
            X += 10
    return X

def DrawIcon(img, IconIndex, X, AddedLines):
    icon = Image.open("./icon/" + IconIndex + ".png")
    icon = icon.resize((14,14))
                    
    r, g, b, a = icon.split()
    img.paste(icon,(X,93+AddedLines*19),mask=a)

def TextDetail(img, txt):
    #换行计数器
    AutoEnter = 0
    AddedLines = 0
    x = 25
    i = 0
    txt = txt + "﹍"
    while(i < len(txt)):
        if(txt[i] == "﹍"):
            break
        
        #强制印刷 - 接下来的字符，将忽略其语法作用被打印在卡面上
        if(txt[i] == '&'):
            
            if(x > 220):
                AddedLines += 1
                AutoEnter = 0
                x = 25
                img.paste(BasicBody, (5,93+19*AddedLines))
            
            x = WriteChar(img, txt[i+1], x, AddedLines)
            AutoEnter += 1
            i += 2
            continue
        
        if(txt[i] == '$'):
            AddedLines += 1
            AutoEnter = 0
            x = 25
            i += 1
            img.paste(BasicBody, (5,93+19*AddedLines))
            continue
        
        #检测到括号，判断内容后写出来
        if(txt[i] == '<'):
            
            EffectIndex = ""
            Iful = False
            
            while(txt[i+1] != '#'):
                i += 1
                EffectIndex += txt[i]
                
            if(EffectIndex == 'U'):
                Iful = True
                i += 1
                EffectIndex = ""
                while(txt[i+1] != '#'):
                    i += 1
                    EffectIndex += txt[i]
                    
            
            i += 1
            while(txt[i+1] != '>'):
                i += 1
                
                if(x > 220):
                    AddedLines += 1
                    AutoEnter = 0
                    x = 25
                    img.paste(BasicBody, (5,93+19*AddedLines))
                
                if(EffectIndex not in elements):
                    EffectIndex = "BS"
                
                #换行
                if(txt[i] == '$'):
                    AddedLines += 1
                    AutoEnter = 0
                    x = 25
                    i += 1
                    img.paste(BasicBody, (5,93+19*AddedLines))
                
                #强制印刷 
                if(txt[i] == '&'):
                    if(x > 220):
                        AddedLines += 1
                        AutoEnter = 0
                        x = 25
                        img.paste(BasicBody, (5,93+19*AddedLines))
            
                    x = WriteChar(img, txt[i+1], x, AddedLines, colorDict[EffectIndex], Iful)
                    AutoEnter += 1
                    i += 1
                    
                    continue
                
                #图标
                if(txt[i] == '['):
                    IconIndex = ""
                    i += 1
                    while(txt[i] != ']'):
                        IconIndex += txt[i]
                        i += 1
                        
                    if(x > 220):
                        AddedLines += 1
                        AutoEnter = 0
                        x = 25
                        img.paste(BasicBody, (5,93+19*AddedLines))
                        
                    #放置图标
                    DrawIcon(img, IconIndex, x, AddedLines)
                    
                    AutoEnter += 1
                    x += 14
                    if(txt[i+1] != '>'):
                        i += 1
                    continue
                    
                x = WriteChar(img, txt[i], x, AddedLines, colorDict[EffectIndex], Iful)
                
                AutoEnter += 1
            i += 1
    
        #检测图标
        if(txt[i] == '['):
            IconIndex = ""
            i += 1
            while(txt[i] != ']'):
                IconIndex += txt[i]
                i += 1
            
            if(x > 220):
                AddedLines += 1
                AutoEnter = 0
                x = 25
                img.paste(BasicBody, (5,93+19*AddedLines))
                
            #放置图标
            DrawIcon(img, IconIndex, x, AddedLines)
            
            
            AutoEnter += 1
            x += 14
            if(txt[i] != '>'):
                i += 1
            
            continue

        if(txt[i] != '>'):
            if(x > 220):
                AddedLines += 1
                AutoEnter = 0
                x = 25
                img.paste(BasicBody, (5,93+19*AddedLines))
                
            x = WriteChar(img, txt[i], x, AddedLines)
            AutoEnter += 1
        i += 1
        
    return AddedLines
        
            

ErrorCount = 0
ErrorInfo = ""
def IfSkip(row):
    global ErrorCount, ErrorInfo
    content = table.row_values(row)
    flag = False
    IconTest = True
    BasicInfo = 'Excel 规则解释 第' + str(row+1) + '行:\n'
    
    #错误一：必要信息没有完全填写
    if(content[0] == ""):
        BasicInfo = BasicInfo + '·必要表格信息没有完全填写\n'
        flag = True
        
    #错误二：高级语法错误
    try:
        txt = content[2]
        i = 0
        while(i < len(txt)): 
            #检查<#>语法
            #检测到括号，判断内容后写出来
            if(txt[i] == '<'):
                EffectIndex = ""
                    
                while(txt[i+1] != '#'):
                    i += 1
                    EffectIndex += txt[i]
                        
                if(EffectIndex == 'U'):
                    i += 1
                    EffectIndex = ""
                    while(txt[i+1] != '#'):
                        i += 1
                        EffectIndex += txt[i]
                    
                i += 1
                while(txt[i+1] != '>'):
                    i += 1
                        
                    if(EffectIndex not in elements):
                        EffectIndex = "BS"
                i += 1
                
            if(txt[i] == '['):
                IconIndex = ""
                i += 1
                while(txt[i] != ']'):
                    IconIndex += txt[i]
                    i += 1
            i += 1
            
    except IndexError:
        BasicInfo = BasicInfo + '·高级语法书写有误\n'
        flag = True
        IconTest = False
    
    #错误三：图标不存在
    if IconTest:
        txt = content[2]
        i = 0
        while(i < len(txt)): 
            if(txt[i] == '[' and (i-1 < 0 or txt[i-1] != '&')):
                IconIndex = ""
                i += 1
                while(txt[i] != ']'):
                    IconIndex += txt[i]
                    i += 1
                    
                #3. 图标文件存在性
                try:
                    f = open("./icon/" + IconIndex + ".png")
                    f.close()
                except FileNotFoundError:
                    BasicInfo = BasicInfo + '·找不到图标 ' + IconIndex + '.png \n'
                    flag = True       
                continue 
            i += 1
    
    if flag:
        ErrorCount += 1
        ErrorInfo = ErrorInfo + BasicInfo + '\n'
    return flag


def execute():
    print("开始生成规则")
    
    global LineNum, beg, end
    
    LineNum = end - beg + 1
    for i in range(beg, end):
        #读取
        content = table.row_values(i)
    
        if IfSkip(i):
            continue
        
        #引入description作为描述图片
        description = Image.new("RGBA",(450,1141))
        description.paste(BasicHead, (5,10))
        description.paste(BasicBody, (5,55))
        description.paste(BasicBody, (5,74))
        description.paste(BasicBody, (5,74+19))
        
        #基本标题
        DesDraw = ImageDraw.Draw(description)
        DesDraw.text((25,63), content[1], font=DescriptionTitle,fill='white')
        
        AddedLines = TextDetail(description, content[2])
        
        description.paste(BasicEnd, (5,92+(AddedLines+1)*19))
        
        description = description.crop((0,0,250,107+(AddedLines+1)*19))
        description.save("./output/规则解释/[" + content[0] + "]" + content[1] + ".png")
        
        #路径、宽、高
        UsingRule[content[0]] = ["./output/规则解释/[" + content[0] + "]" + content[1] + ".png", 250,107+(AddedLines+1)*19]