import time

import pyautogui
import pyperclip
import xlrd


# 定义鼠标事件

# pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159

def mouseClick(clickTimes, lOrR, img, reTry):
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                break
            print("未找到匹配图片,0.1秒后重试")
            time.sleep(0.1)
    elif reTry == -1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)  # type: ignore
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
            time.sleep(0.1)
    elif reTry > 1:
        i = 1
        while i < reTry + 1:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                print("重复")
                i += 1
            time.sleep(0.1)


# 【数据检查】
# cmdType.value|   单击左键    双击左键  
#              |   单击右键    输入  
#              |   等待        滚轮 
#              |   热键        文本文件路径（键盘输入文件内容)
# -------------+----------------------------------------------------
# ctype        |   空：0
#              |   字符串：1
#              |   数字：2
#              |   日期：3
#              |   布尔：4
#              |   error：5
# --------------+---------------------------------------------------
def dataCheck(sheet1):
    checkCmd = True
    # 行数检查
    if sheet1.nrows < 2:
        print("没数据啊哥")
        checkCmd = False
    # 每行数据检查（i表示行数）
    i = 1
    while i < sheet1.nrows:
        # 【第1列 操作指令类型检查】
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 1 or (
                cmdType.value != '单击左键' and cmdType.value != '双击左键' and cmdType.value != '单击右键'
                and cmdType.value != '输入' and cmdType.value != '等待' and cmdType.value != '滚轮'
                and cmdType.value != '热键' and cmdType.value != '文本文件路径'):
            print('第', i + 1, "行,第1列数据有误,可能输入了错误的或不能识别的操作指令！\n")
            checkCmd = False
        # 【第2列 内容检查】
        cmdValue = sheet1.row(i)[1]
        # 读图点击类型指令，内容必须为字符串类型
        if cmdType.value == '单击左键' or cmdType.value == '双击左键' or cmdType.value == '单击右键':
            if cmdValue.ctype != 1:
                print('第', i + 1, "行,第2列数据有误,应为字符串类型")
                checkCmd = False
        # 输入类型，内容不能为空
        if cmdType.value == '输入':
            if cmdValue.ctype == 0:
                print('第', i + 1, "行,第2列数据有误,输入内容不能为空")
                checkCmd = False
        # 等待类型，内容必须为数字
        if cmdType.value == '等待':
            if cmdValue.ctype != 2:
                print('第', i + 1, "行,第2列数据有误,应为数字类型")
                checkCmd = False
        # 滚轮事件，内容必须为数字
        if cmdType.value == '滚轮':
            if cmdValue.ctype != 2:
                print('第', i + 1, "行,第2列数据有误,应为数字类型")
                checkCmd = False
        i += 1
    return checkCmd


# 执行的主任务
def mainWork(img):
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型
        cmdType = sheet1.row(i)[0]
        # 指令<单击左键>
        if cmdType.value == '单击左键':
            # 取图片名称
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "left", img, reTry)
            print("单击左键", img)
        # 指令<双击左键>
        elif cmdType.value == '双击左键':
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2, "left", img, reTry)
            print("双击左键", img)
        # 指令<右键>
        elif cmdType.value == '右键单击':
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "right", img, reTry)
            print("右键", img)
        # 指令<输入>
        elif cmdType.value == '输入':
            inputValue = sheet1.row(i)[1].value
            pyperclip.copy(inputValue)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("剪贴板输入:", inputValue)
        # 指令<等待>
        elif cmdType.value == '5':
            # 取图片名称
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("等待 ", waitTime, " 秒")
        # 指令<滚轮>
        elif cmdType.value == '滚轮':
            # 取图片名称
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动了 ", int(scroll), " 距离")
        # 指令<热键>
        elif cmdType.value == '热键':
            inputValue = sheet1.row(i)[1].value
            if inputValue == "enter":
                pyautogui.press("enter")
            time.sleep(0.5)
            print("按下：", inputValue, " 热键")
        # 指令<文本文件路径>
        # 输入文本文件路径，键盘输出文件内容
        elif cmdType.value == '文本文件路径':
            inputValue = sheet1.row(i)[1].value
            # 读取文本文件的内容
            filepath = inputValue  # 文件绝对路径
            with open(filepath, 'r', encoding='UTF-8') as file:  # ’r‘只读文件
                copy_input = file.readline()  # 逐行读取,结果是一个list
            keyboard_input = copy_input
            pyautogui.typewrite(keyboard_input, interval=0.025)  # 放在列表里，interval 指输入间隔秒
            print("键盘输入内容：", keyboard_input)
        i += 1


if __name__ == '__main__':
    # 文件名
    file = 'cmdMysql.xls'
    # 打开文件
    wb = xlrd.open_workbook(filename=file)
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print('+********欢迎使用桌面自动程序********+\n'
          '|| 版本 V1.0.0                      ||\n'
          '|| 原作：B站Up@不高兴就喝水         ||\n'
          '|| 改版：GitHub@54Coconi            ||\n'
          '|| 说明：此版支持中文指令           ||\n'
          '||                                  ||\n'
          '+************************************+\n')
    # 数据检查
    checkCmd = dataCheck(sheet1)
    if checkCmd:
        key = input('请选择功能: \n\t输入1只做一次\n'
                    '\t输入2循环n次 \n\t  ')
        if key == '1':
            # 循环拿出每一行指令
            mainWork(sheet1)
        elif key == '2':
            n = input('循环次数n = ')
            while n > 0:  # type: ignore
                mainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")
                n -= 1  # type: ignore
            input()
    else:
        print('按Enter键退出!')
        input()  # 防止程序直接退出，这样需要按下回车才会退出

# print(os.getcwd())
