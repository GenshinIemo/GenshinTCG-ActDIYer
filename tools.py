import PIL
from PIL import Image, ImageFilter, ImageFont, ImageDraw

import xlrd, tkinter
import tkinter.messagebox
import sys

import os

elements = ['BS','HS','B','H','S','L','F','Y','C']

root = tkinter.Tk()
root.withdraw()

try:
    f = open("./basic_resource/cardHide.png")
    f.close()
    f = open("./basic_resource/card.png")
    f.close()
    for i in elements:
        f = open("./basic_resource/elements/" + i + ".png")
        f.close()
    f = open("./行动牌.xls")
    f.close()
    f = open("./GenshinFont.ttf")
    f.close()
except FileNotFoundError:
    tkinter.messagebox.showinfo("七圣DIY快捷生成器","程序必要文件不齐，建议重新下载")
    sys.exit()

RunTest = tkinter.messagebox.askyesno("七圣DIY快捷生成器","程序运行正常\n"+"是否开始生成？")

if(not RunTest):
    tkinter.messagebox.showinfo("七圣DIY快捷生成器","程序已结束")
    sys.exit()

#参考自：https://stackoverflow.com/questions/41556771/is-there-a-way-to-outline-text-with-a-dark-line-in-pil
#传入draw对象，以fillcolor为前景色，shadowcolor是边框色, boldval是粗度
def DrawBorder(draw, x, y, text, font, shadowcolor, fillcolor, boldval):
    #厚边框
    draw.text((x - boldval, y), text, font=font, fill=shadowcolor)
    draw.text((x + boldval, y), text, font=font, fill=shadowcolor)
    draw.text((x, y - boldval), text, font=font, fill=shadowcolor)
    draw.text((x, y + boldval), text, font=font, fill=shadowcolor)
 
    #薄边框
    draw.text((x - boldval, y - boldval), text, font=font, fill=shadowcolor)
    draw.text((x + boldval, y - boldval), text, font=font, fill=shadowcolor)
    draw.text((x - boldval, y + boldval), text, font=font, fill=shadowcolor)
    draw.text((x + boldval, y + boldval), text, font=font, fill=shadowcolor)
 
    #带边框的输出
    draw.text((x, y), text, font=font, fill=fillcolor)

half = "QWERTYUIOPASDFGHJKLZXCVBNMqwertyuiopasdfghjklzxcvbnm1234567890,.?:;()=-+*/（）"
shorter = "iIl1"
colorDict = {
    'BS':'white',
    'HS':'#A9A7A1',
    'B':'#9FFBFC',
    'H':'#ED9595',
    'S':'#86CAFF',
    'L':'#EEA9EF',
    'F':'#80F4D7',
    'Y':'#FFF29F',
    'C':'#84CC35'
}

debugger = ""

def WriteChar(img, txt, X, AddedLines, color = colorDict['HS'], IfLined = False):
    rec = X
    global debugger 
    debugger = debugger + txt
    draw = ImageDraw.Draw(img)
        
    draw.text((X,460+AddedLines*25), txt, font=CardIntro, fill=color)
    if(IfLined):
        if(txt not in half):
            ul = Image.new("RGB",(22,3),color=color)
            img.paste(ul,(X,460+(AddedLines+1)*25))
        else:
            if(txt in shorter):
                ul = Image.new("RGB",(11,3),color=color)
                img.paste(ul,(X,460+(AddedLines+1)*25))
            else:
                ul = Image.new("RGB",(13,3),color=color)
                img.paste(ul,(X,460+(AddedLines+1)*25))
        
        
    if(txt not in half):
        X += 22
    else:
        if(txt in shorter):
            X += 11
        else:
            X += 15
    return X

def TextDetail(img, txt):
    #换行计数器，达到35换一行
    AutoEnter = 0
    AddedLines = 0
    x = 812
    i = 0
        
    while(i < len(txt)):
        if(txt[i] == '$'):
            AddedLines += 1
            AutoEnter = 0
            x = 812
            i += 1
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
                
                if(EffectIndex not in elements):
                    EffectIndex = "BS"
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
                
            #放置图标
            icon = Image.open("./icon/" + IconIndex + ".png")
            icon = icon.resize((22,22))
            
            r, g, b, a = icon.split()
            img.paste(icon,(x,460+AddedLines*25+4),mask=a)
            
            AutoEnter += 1
            x += 24
            i += 1
            continue

        if(txt[i] != '>'):
            x = WriteChar(img, txt[i], x, AddedLines)
            AutoEnter += 1
        i += 1
        
        if(AutoEnter > 35):
            AddedLines += 1
            AutoEnter = 0
            x = 812
            
                
#载入字体
CardTitle = ImageFont.truetype("./GenshinFont.ttf", 48)
CardType = ImageFont.truetype("./GenshinFont.ttf", 20)
CardIntro = ImageFont.truetype("./GenshinFont.ttf", 22)
CardCostTXT = ImageFont.truetype("./GenshinFont.ttf", 72)
    
#读取Excel数据
excel = xlrd.open_workbook("行动牌.xls")
table = excel.sheet_by_index(0)
LineNum = table.nrows
ColNum = table.ncols

ErrorInfo = "上一次运行时错误发生在：\n"
ErrorCount = 0
def IfSkip(row):
    global ErrorCount, ErrorInfo
    content = table.row_values(row)
    flag = False
    IconTest = True
    BasicInfo = 'Excel行 ' + str(row+1) + ':\n'
    
    #错误一：必要信息没有完全填写
    if(content[0] == '' or content[2] == '' or content[3] == '' or content[6] == '' or content[7] == ''):
        BasicInfo = BasicInfo + '·必要表格信息没有完全填写\n'
        flag = True
    
    #错误二：找不到图片
    try:
        f = open("./pictures/" + content[6])
        f.close()
    except FileNotFoundError:
        BasicInfo = BasicInfo + '·找不到卡牌图片\n'
        flag = True
    except PermissionError:
        BasicInfo = BasicInfo + '·图片被打开、被其他应用占用或未填写任何图片\n'
        flag = True
    
    #错误四：高级语法错误
    try:
        txt = content[5]
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
        txt = content[5]
        i = 0
        while(i < len(txt)): 
            if(txt[i] == '['):
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

for i in range(1, LineNum):
    content = table.row_values(i)
    if IfSkip(i):
        continue
    
    #行动牌是否填写隐藏加入按钮
    if(content[7] == '否'):
        HideAdd = False
    else:
        HideAdd = True

    if(HideAdd):
        basis = Image.open("./basic_resource/cardHide.png")
    else:
        basis = Image.open("./basic_resource/card.png")
        
    #打开卡图并缩放为卡牌大小
    cardimg = Image.open("./pictures/" + content[6])
    cardimg = cardimg.resize((400,660))
    
    #引入card作为输出图片，将卡牌模板与卡图拼贴
    card = Image.new("RGB",(2000,1141))
    card.paste(cardimg,(358,244))
    r, g, b, a = basis.split()
    card.paste(basis,(0,0),mask=a)

    #标题、说明、父类
    PictDraw = ImageDraw.Draw(card)
    PictDraw.text((812,358), content[1], font=CardTitle, fill='white')
    #PictDraw.text((812,460), content[5], font=CardIntro, fill='#A9A7A1')
    TextDetail(card, content[5])
    PictDraw.text((1520,362), content[3], font=CardType, fill='white')

    #子类
    SonType = content[4]
    SonX = 1580 - len(SonType)*20
    PictDraw.text((SonX,389), SonType, font=CardType, fill='white')

    #费用
    CostType = content[2][1:]
    if(CostType not in elements):
        CostType = "HS"
    costimg = Image.open("./basic_resource/elements/" + CostType + ".png")
    costimg = costimg.resize((120,120))
    maskcost = costimg
    card.paste(costimg,(333,232),mask=maskcost)
    CostVal = content[2][0]

    #原神字体的1需要单独调整下位置
    if(CostVal == '1'):
        DrawBorder(PictDraw, 372, 248, str(CostVal), CardCostTXT, 'black', 'white', 2)
    else:    
        DrawBorder(PictDraw, 368, 248, str(CostVal), CardCostTXT, 'black', 'white', 2)

    #card.show()
    card.save("./output/" + content[0] + ".png")


tkinter.messagebox.showinfo("七圣DIY快捷生成器",str(LineNum-1-ErrorCount)+"张行动牌生成完毕\n"+str(ErrorCount)+"张行动牌生成错误\n若有卡牌未生成且Excel已保存，请查看错误报告errors.txt")

if ErrorInfo == '上一次运行时错误发生在：\n':
    ErrorInfo = '上一次运行正常，未发生错误，请检查Excel是否保存'

file = open("./errors.txt", 'w')
file.write(ErrorInfo)
file.close()
