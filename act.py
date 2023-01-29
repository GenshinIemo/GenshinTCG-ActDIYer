import PIL
from PIL import Image, ImageFilter, ImageFont, ImageDraw

import xlrd

import basicdata as bdt
import rule, state

#读取Excel数据
excel = xlrd.open_workbook("行动牌.xls")
table = excel.sheet_by_index(0)
LineNum = table.nrows
ColNum = table.ncols

half = bdt.half
shorter = bdt.shorter
colorDict = bdt.colorDict
elements = bdt.elements

#开始结束
beg = bdt.settings[2] - 1
end = bdt.settings[3]
if(bdt.settings[3] == '*' or bdt.settings[3] > LineNum):
    end = LineNum

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


debugger = ""

def WriteChar(img, txt, X, AddedLines, color = colorDict['HS'], IfLined = False):
    rec = X
    global debugger 
    debugger = debugger + txt
    draw = ImageDraw.Draw(img)
        
    draw.text((X,460+AddedLines*29), txt, font=CardIntro, fill=color)
    if(IfLined):
        if(txt not in half):
            ul = Image.new("RGB",(24,3),color=color)
            img.paste(ul,(X,456+(AddedLines+1)*29))
        else:
            if(txt in shorter):
                ul = Image.new("RGB",(12,3),color=color)
                img.paste(ul,(X,456+(AddedLines+1)*29))
            elif(txt == '.'):
                ul = Image.new("RGB",(8,3),color=color)
                img.paste(ul,(X,456+(AddedLines+1)*29))
            else:
                ul = Image.new("RGB",(16,3),color=color)
                img.paste(ul,(X,456+(AddedLines+1)*29))
        
        
    if(txt not in half):
        X += 24
    else:
        if(txt in shorter):
            X += 12
        elif(txt == '.'):
            X += 8
        else:
            X += 16
    return X

def DrawIcon(img, IconIndex, X, AddedLines):
    icon = Image.open("./icon/" + IconIndex + ".png")
    icon = icon.resize((24,24))
                    
    r, g, b, a = icon.split()
    img.paste(icon,(X+1,460+AddedLines*29+2),mask=a)

def TextDetail(img, txt):
    #换行计数器
    AutoEnter = 0
    AddedLines = 0
    x = 812
    i = 0
        
    while(i < len(txt)):
        
        #强制印刷 - 接下来的字符，将忽略其语法作用被打印在卡面上
        if(txt[i] == '&'):
            x = WriteChar(img, txt[i+1], x, AddedLines)
            AutoEnter += 1
            i += 2
            
            if(x > 1530):
                AddedLines += 1
                AutoEnter = 0
                x = 812
            
            continue
        
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
                
                #换行
                if(txt[i] == '$'):
                    AddedLines += 1
                    AutoEnter = 0
                    x = 812
                    i += 1
                
                #强制印刷 
                if(txt[i] == '&'):
                    x = WriteChar(img, txt[i+1], x, AddedLines, colorDict[EffectIndex], Iful)
                    AutoEnter += 1
                    i += 1
                    
                    if(x > 1530):
                        AddedLines += 1
                        AutoEnter = 0
                        x = 812
                    
                    continue
                
                #图标
                if(txt[i] == '['):
                    IconIndex = ""
                    i += 1
                    while(txt[i] != ']'):
                        IconIndex += txt[i]
                        i += 1
                        
                    #放置图标
                    DrawIcon(img, IconIndex, x, AddedLines)
                    
                    if(x > 1530):
                        AddedLines += 1
                        AutoEnter = 0
                        x = 812
                    
                    AutoEnter += 1
                    x += 24
                    if(txt[i+1] != '>'):
                        i += 1
                    continue
                
                x = WriteChar(img, txt[i], x, AddedLines, colorDict[EffectIndex], Iful)
                
                AutoEnter += 1
                
                if(x > 1530):
                    AddedLines += 1
                    AutoEnter = 0
                    x = 812
            i += 1
    
        #检测图标
        if(txt[i] == '['):
            IconIndex = ""
            i += 1
            while(txt[i] != ']'):
                IconIndex += txt[i]
                i += 1
                
            #放置图标
            DrawIcon(img, IconIndex, x, AddedLines)
            
            AutoEnter += 1
            x += 24
            if(txt[i] != '>'):
                i += 1
            
            if(x > 1530):
                AddedLines += 1
                AutoEnter = 0
                x = 812
            
            continue

        if(txt[i] != '>'):
            x = WriteChar(img, txt[i], x, AddedLines)
            AutoEnter += 1
        i += 1
        
        if(x > 1530):
            AddedLines += 1
            AutoEnter = 0
            x = 812
        
            
                
#载入字体
CardTitle = ImageFont.truetype("./GenshinFont.ttf", 48)
CardType = ImageFont.truetype("./GenshinFont.ttf", 20)
CardIntro = ImageFont.truetype("./GenshinFont.ttf", 24)
CardCostTXT = ImageFont.truetype("./GenshinFont.ttf", 72)
    

ErrorInfo = ""
ErrorCount = 0
def IfSkip(row):
    global ErrorCount, ErrorInfo
    content = table.row_values(row)
    flag = False
    IconTest = True
    BasicInfo = 'Excel 行动牌 第' + str(row+1) + '行:\n'
    
    #错误：必要信息没有完全填写
    if(content[0] == '' or content[2] == '' or content[3] == '' or content[6] == '' or content[7] == ''):
        BasicInfo = BasicInfo + '·必要表格信息没有完全填写\n'
        flag = True
    
    #错误：找不到图片
    try:
        f = open("./pictures/" + content[6])
        f.close()
    except FileNotFoundError:
        BasicInfo = BasicInfo + '·找不到卡牌图片\n'
        flag = True
    except PermissionError:
        BasicInfo = BasicInfo + '·图片被打开、被其他应用占用或未填写任何图片\n'
        flag = True
    
    #错误；规则未建立
    #根据井号分割
    if(content[11] != "" and content[11] != "不打印"):
        RuleInput = content[9]
        CutList = []
        pos = 0
        while(pos != -1):
            CutList.append(pos)
            pos = RuleInput.find("#", pos+1)
        CutList.append(len(RuleInput))
        
        for i in range(len(CutList)-1):
                left = CutList[i]
                right = CutList[i+1]
                if(i != 0):
                    left += 1
                CurrentRule = content[9][left:right]

                #创建带规则的图片
                if CurrentRule != "":
                    if CurrentRule not in rule.UsingRule.keys():
                        BasicInfo = BasicInfo + '·规则' + CurrentRule + '未被成功建立\n'
                        flag = True
    #错误；附属物未建立
    #根据井号分割
    if(content[11] != "" and content[11] != "不打印"):
        Input = content[10]
        CutList = []
        pos = 0
        while(pos != -1):
            CutList.append(pos)
            pos = Input.find("#", pos+1)
        CutList.append(len(Input))
        
        for i in range(len(CutList)-1):
                left = CutList[i]
                right = CutList[i+1]
                if(i != 0):
                    left += 1
                Current = content[10][left:right]

                #创建带规则的图片
                if Current != "":
                    if Current not in state.UsingState.keys():
                        BasicInfo = BasicInfo + '·附属物' + Current + '未被成功建立\n'
                        flag = True
    
    #错误：高级语法错误
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
    
    #错误：图标不存在
    if IconTest:
        txt = content[5]
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

def BackgroundDarker(img):
    dark = Image.new("RGBA",(2000,1141),(0,0,0,0))
    draw = ImageDraw.Draw(dark, "RGBA")
    draw.rectangle([0,0,2000,1141], fill = (0,0,0,114))
    img.paste(Image.alpha_composite(img, dark))

def CraftRuleCard(img, PutRule):
    BackgroundDarker(img)
    RuleInfo = rule.UsingRule[PutRule]
    ruleimg = Image.open(RuleInfo[0])
    ruleimg = ruleimg.resize((500,RuleInfo[2]*2))
    
    r, g, b, a = ruleimg.split()
    img.paste(ruleimg,(750,570-RuleInfo[2]),mask=a)
    return img

def CraftStateCard(img, PutState):
    BackgroundDarker(img)
    StateInfo = state.UsingState[PutState]
    stateimg = Image.open(StateInfo[0])
    stateimg = stateimg.resize((int(StateInfo[1]*1.4+4),int(StateInfo[2]*1.4)))
    
    r, g, b, a = stateimg.split()
    img.paste(stateimg,(777,int(570-StateInfo[2]/1.4)),mask=a)
    
    if(StateInfo[3] != '*'):
        cardimg = Image.open(StateInfo[3])
        cardimg = cardimg.resize((242,405))
        
        img.paste(cardimg, (504,426-70))
        r, g, b, a = bdt.PictBase.split()
        img.paste(bdt.PictBase, (500,422-70), mask=a)
    
    return img

def execute():
    print("开始生成卡牌")
    
    global LineNum, beg, end
    
    LineNum = end - beg + 1
    #for i in range(1, 2):
    for i in range(beg, end):
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
        card = Image.new("RGBA",(2000,1141))
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
        
        #可携带数量
        if(not HideAdd):
            if(content[8] == ''):
                content[8] = '2'
            blackblock = Image.new("RGB",(50,30),color='#2B333D')
            card.paste(blackblock,(900,825))
            PictDraw.text((900,824), "0/" + content[8], font=CardIntro, fill='white')

        #费用
        #根据井号分割
        CostCode = content[2]
        CutList = []
        pos = 0
        while(pos != -1):
            CutList.append(pos)
            pos = CostCode.find("#", pos+1)
        CutList.append(len(CostCode))
        
        #逐个处理
        for i in range(len(CutList)-1):
            left = CutList[i]
            right = CutList[i+1]
            if(i != 0):
                left += 1
            
            CostVal = content[2][left]
            CostType = content[2][left+1:right]
            if(CostType not in elements):
                CostType = "HS"
            if(CostType == 'CN'):
                costimg = Image.open("./basic_resource/elements/" + CostType + ".png")
                costimg = costimg.resize((135,135))
                maskcost = costimg
                card.paste(costimg,(326 + 135*(i//5),225+135*i),mask=maskcost)
            else:
                costimg = Image.open("./basic_resource/elements/" + CostType + ".png")
                costbd = Image.open("./basic_resource/elementBase.png")
                costimg = costimg.resize((118,118))
                costbd = costbd.resize((115,130))
                maskcost = costbd
                card.paste(costbd,(335 + 135*(i//5),228+135*(i%5)),mask=maskcost)
                maskcost = costimg
                card.paste(costimg,(335 + 135*(i//5),233+135*(i%5)),mask=maskcost)

            #原神字体的1需要单独调整下位置
            if(CostVal == '1'):
                DrawBorder(PictDraw, 376 + 135*(i//5), 248 + 135*(i%5), str(CostVal), CardCostTXT, 'black', 'white', 2)
            else:    
                DrawBorder(PictDraw, 370 + 135*(i//5), 248 + 135*(i%5), str(CostVal), CardCostTXT, 'black', 'white', 2)
        
        #使用Rule规则和State附属
        #根据井号分割
        if(content[11] != "" and content[11] != "不打印"):
            addsign = 0
            RuleInput = content[9]
            RCutList = []
            Rpos = 0
            while(Rpos != -1):
                RCutList.append(Rpos)
                Rpos = RuleInput.find("#", Rpos+1)
            RCutList.append(len(RuleInput))
            if(RuleInput != ""):
                addsign += len(RCutList)
            
            StateInput = content[10]
            SCutList = []
            Spos = 0
            while(Spos != -1):
                SCutList.append(Spos)
                Spos = StateInput.find("#", Spos+1)
            SCutList.append(len(StateInput))
            if(StateInput != ""):
                if addsign != 0: addsign -= 1
                addsign += len(SCutList)
            
            if(content[11] == "多图拼贴"):
                #准备将图片拼贴
                TogetherBase = Image.new("RGBA", (2000,1141*(addsign)))
                TogetherBase.paste(card, (0,0))
            
            #逐个处理
            putted = 0
            for i in range(len(RCutList)-1):
                left = RCutList[i]
                right = RCutList[i+1]
                if(i != 0):
                    left += 1
                CurrentRule = content[9][left:right]

                #创建带规则的图片
                if CurrentRule != "":
                    CardSave = Image.new("RGBA", (2000,1141))
                    putted += 1
                    CardSave.paste(card, (0,0))
                    RuledCard = CraftRuleCard(CardSave, CurrentRule)
                    
                    if(content[11] == "输出多张图"):
                        CardSave.save("./output/行动牌/[" + content[0] + "-" + CurrentRule + "]" + content[1] + ".png")
                    else:
                        TogetherBase.paste(CardSave, (0,1141*(putted)))
            
            for i in range(len(SCutList)-1):
                left = SCutList[i]
                right = SCutList[i+1]
                if(i != 0):
                    left += 1
                CurrentSt= content[10][left:right]

                #创建带附属的图片
                if CurrentSt != "":
                    CardSave = Image.new("RGBA", (2000,1141))
                    CardSave.paste(card, (0,0))
                    putted += 1
                    StatedCard = CraftStateCard(CardSave, CurrentSt)
                    
                    if(content[11] == "输出多张图"):
                        CardSave.save("./output/行动牌/[" + content[0] + "-" + CurrentSt + "]" + content[1] + ".png")
                    else:
                        TogetherBase.paste(CardSave, (0,1141*(putted)))

        #card.show()
        if(content[11] == "多图拼贴"):
            TogetherBase.save("./output/行动牌/[" + content[0] + "]" + content[1] + ".png")
        else:    
            card.save("./output/行动牌/[" + content[0] + "]" + content[1] + ".png")