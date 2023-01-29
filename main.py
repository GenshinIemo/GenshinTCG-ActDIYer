import PIL
from PIL import Image, ImageFilter, ImageFont, ImageDraw

import xlrd, tkinter
import tkinter.messagebox
import sys

import basicdata as bdt

import act, rule, state
from tkinter import *

root = Tk()
root.withdraw()

try:
    f = open("./basic_resource/cardHide.png")
    f.close()
    f = open("./basic_resource/card.png")
    f.close()
    f = open("./basic_resource/ruleUI/head.png")
    f.close()
    f = open("./basic_resource/ruleUI/body.png")
    f.close()
    f = open("./basic_resource/ruleUI/end.png")
    f.close()
    f = open("./basic_resource/elementBase.png")
    f.close()
    for i in bdt.elements:
        f = open("./basic_resource/elements/" + i + ".png")
        f.close()
    f = open("./行动牌.xls")
    f.close()
    f = open("./GenshinFont.ttf")
    f.close()
except FileNotFoundError:
    tkinter.messagebox.showinfo("七圣DIY快捷生成器","程序必要文件不齐，建议重新下载")
    sys.exit()

if(rule.beg != rule.end and act.beg != act.end and state.beg != state.end):
    if(bdt.settings[0] <= 1 or bdt.settings[0] > rule.LineNum or (bdt.settings[1] != '*' and bdt.settings[1] <= 1) or bdt.settings[2] <= 1 or bdt.settings[2] > act.LineNum or (bdt.settings[3] != '*' and bdt.settings[3] <= 1/
        bdt.settings[4] <= 1 or bdt.settings[4] > rule.LineNum or (bdt.settings[5] != '*' and bdt.settings[5] <= 1) )):
            tkinter.messagebox.showinfo("七圣DIY快捷生成器","设置不合法，请重新设置")
            sys.exit()

setdata = '规则生成区间：[' + str(bdt.settings[0]) + ',' + str(bdt.settings[1]) + ']\n行动牌生成区间：[' +str(bdt.settings[2]) + ',' + str(bdt.settings[3]) + ']\n'
setdata += '附属物生成区间：[' + str(bdt.settings[4]) + ',' + str(bdt.settings[5]) + ']\n'
RunTest = tkinter.messagebox.askyesno("七圣DIY快捷生成器","程序运行正常，参数如下\n"+setdata+"是否开始生成？")

if(not RunTest):
    tkinter.messagebox.showinfo("七圣DIY快捷生成器","程序已结束")
    sys.exit()

if(rule.beg != rule.end):
    rule.execute()
if(state.beg != state.end):
    state.execute()
if(act.beg != act.end):
    act.execute()

ShowedInfo = ""
ShowedInfo += str(rule.LineNum-1-rule.ErrorCount)+"/"+str(rule.LineNum-1)+"条规则建立成功\n"
ShowedInfo += str(state.LineNum-1-state.ErrorCount)+"/"+str(state.LineNum-1)+"个附属物建立成功\n"
ShowedInfo += str(act.LineNum-1-act.ErrorCount)+"/"+str(act.LineNum-1)+"张行动牌生成成功\n"
ShowedInfo += "若有卡牌未生成且Excel已保存，请查看错误报告errors.txt"

tkinter.messagebox.showinfo("七圣DIY快捷生成器",ShowedInfo)

if rule.ErrorInfo == "":
    rule.ErrorInfo = '上一次规则解释运行正常，未发生错误'
if state.ErrorInfo == "":
    state.ErrorInfo = '上一次附属物运行正常，未发生错误'
if act.ErrorInfo == "":
    act.ErrorInfo = '上一次行动牌运行正常，未发生错误'

Errors = "上一次错误发生在：\n---------------规则----------------\n" + rule.ErrorInfo + '\n---------------附属物----------------\n' + state.ErrorInfo
Errors += '\n---------------行动牌----------------\n' + act.ErrorInfo

file = open("./errors.txt", 'w')
file.write(Errors)
file.close()
