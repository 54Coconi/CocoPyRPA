import time

import pyautogui
import pyperclip
import xlrd

""" 
=====================================
            定义鼠标点击事件
=====================================
"""


# pyautogui库其他用法 https://blog.csdn.net/qingfengxd1/article/details/108270159

def mouseClick(clickTimes, lOrR, img, reTry):
    """

    :param clickTimes:
    :param lOrR:
    :param img:
    :param reTry:
    """
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
                print("点击第", i, "次")
                i += 1
            time.sleep(0.1)


""" 
=========================================
               【数据检查】
=========================================
"""


# cmdType.value|   1单击左键        2双击左键
#              |   3单击右键        4输入
#              |   5等待            6滚轮
#              |   7按键            8热键组合
#              |   9键盘输入TXT内容（键盘输入文件内容)
#              |   10鼠标相对移动   11鼠标定点移动
# -------------+------------------------------------------------
# ctype        |   空：0
#              |   字符串：1
#              |   数字：2
#              |   日期：3
#              |   布尔：4
#              |   error：5
# --------------+------------------------------------------------


def dataCheck(sheet1):
    """

    :param sheet1:
    :return:
    """
    checkCmd = True
    # 行数检查
    if sheet1.nrows < 2:
        print("表格没有指令数据！")
        checkCmd = False
        return checkCmd
    # 每行数据检查（i表示行数）
    i = 1
    while i < sheet1.nrows:
        # 【第1列 操作指令类型检查】=========================
        cmdType = sheet1.row(i)[0]
        if cmdType.ctype != 1 or (
                cmdType.value != '单击左键' and cmdType.value != '双击左键' and cmdType.value != '单击右键'
                and cmdType.value != '输入' and cmdType.value != '等待' and cmdType.value != '滚轮'
                and cmdType.value != '鼠标相对移动' and cmdType.value != '按键' and cmdType.value != '键盘输入TXT内容'
                and cmdType.value != '鼠标定点移动' and cmdType.value != '热键组合'):
            print('第', i + 1, "行,第1列数据有误,可能输入了错误的或不能识别的操作指令！\n")
            checkCmd = False

        # 【第2列 操作指令内容检查】=================================
        cmdValue = sheet1.row(i)[1]

        # 读图点击类型指令，内容必须为字符串类型
        if (cmdType.value == '单击左键' or
                cmdType.value == '双击左键' or
                cmdType.value == '单击右键'):
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

        # 鼠标相对当前坐标的移动事件，形式必须为'num1=num2'的字符串且x,y取值在屏幕分辨率范围内
        if cmdType.value == '鼠标相对移动':
            list = cmdValue.value.split("=")
            # 将str字符串转为int整型（防止有小数点'.'先转为float浮点型再转int）
            x = int(float(list[0]))
            y = int(float(list[1]))
            width = pyautogui.size().width
            height = pyautogui.size().height
            if abs(x) > width or abs(y) > height:
                print('第', i + 1, "行,第2列数据有误,鼠标移动距离超出范围")
                checkCmd = False

        i += 1
    return checkCmd


"""
====================================
          每个指令执行的任务
====================================
"""


def mainWork(img):
    """

    :param img:
    """
    # 当前行数为 i+1
    i = 1
    while i < sheet1.nrows:
        # 取本行指令的操作类型（第一列）
        cmdType = sheet1.row(i)[0]

        # 1)指令<单击左键>
        if cmdType.value == '单击左键':
            # 取图片名称
            img = sheet1.row(i)[1].value
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "left", img, reTry)
            print("<单击左键>\t======>\t", img)

        # 2)指令<双击左键>
        elif cmdType.value == '双击左键':
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(2, "left", img, reTry)
            print("<双击左键>\t======>\t", img)

        # 3)指令<单击右键>
        elif cmdType.value == '单击右键':
            # 取图片名称
            img = sheet1.row(i)[1].value
            # 取重试次数
            reTry = 1
            if sheet1.row(i)[2].ctype == 2 and sheet1.row(i)[2].value != 0:
                reTry = sheet1.row(i)[2].value
            mouseClick(1, "right", img, reTry)
            print("<单击右键>\t======>\t", img)

        # 4)指令<输入>
        elif cmdType.value == '输入':
            # 取单元格中要输入的内容
            inputValue = sheet1.row(i)[1].value
            # 复制单元格内容
            pyperclip.copy(inputValue)
            # 粘贴内容
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.5)
            print("剪贴板输入\t======>\t", inputValue)

        # 5)指令<等待>
        elif cmdType.value == '等待':
            # 取等待时间
            waitTime = sheet1.row(i)[1].value
            time.sleep(waitTime)
            print("<等待>\t======>\t", waitTime, "秒")

        # 6)指令<滚轮>
        elif cmdType.value == '滚轮':
            # 取单元格中要移动的距离值
            scroll = sheet1.row(i)[1].value
            pyautogui.scroll(int(scroll))
            print("滚轮滑动了\t======>\t", int(scroll), "距离")

        # 7)指令<鼠标相对移动>
        elif cmdType.value == '鼠标相对移动':
            # 取单元格中形如 num1=num2 的字符串
            str = sheet1.row(i)[1].value
            # 分割字符串得到字符串'num1'(x坐标偏移量)和'num2'(y坐标偏移量)，并转为int型
            x = int(float(str.split('=')[0]))  # '='左边的数为横坐标x偏移量（x>0,右移）
            y = int(float(str.split("=")[1]))  # '='右边的数为纵坐标y偏移量（y>0,下移）
            pyautogui.move(x, y)
            print("<鼠标相对移动>\t======>\t",
                  "右移" if x > 0 else "左移" if x < 0 else "不移动", x,
                  "下移" if y > 0 else "上移" if y < 0 else "不移动", y)

        # 8)指令<按键>
        elif cmdType.value == '按键':
            inputValue = sheet1.row(i)[1].value
            if inputValue == "enter":
                pyautogui.press("enter")
            time.sleep(0.5)
            print("<按键>\t======>\t", inputValue)

        # 9)指令<键盘输入TXT内容>
        elif cmdType.value == '键盘输入TXT内容':
            inputValue = sheet1.row(i)[1].value
            # 读取文本文件的内容
            filepath = inputValue  # 文件绝对路径
            with open(filepath, 'r', encoding='UTF-8') as file:  # ‘r’只读文件
                copy_input = file.read()  # 逐行读取,结果是一个list
            keyboard_input = copy_input
            pyautogui.typewrite(keyboard_input, interval=0.025)  # 放在列表里，interval 指输入间隔秒
            print("<键盘输入TXT内容>", keyboard_input)

        # 10)指令<鼠标定点移动>

        # 11)指令<热键组合>

        i += 1


""" 
执行
"""
if __name__ == '__main__':
    # 文件名
    filename = '指令测试.xls'
    # 打开文件
    wb = xlrd.open_workbook(filename)
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print('+********欢迎使用桌面自动程序********+\n'
          '|| 版本 V1.1.0                      ||\n'
          '|| 原作：B站Up@不高兴就喝水         ||\n'
          '|| 改版：GitHub@54Coconi            ||\n'
          '|| 说明：此版新增中文指令           ||\n'
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
            while n > 0:
                mainWork(sheet1)
                time.sleep(0.1)
                print("等待0.1秒")
                n -= 1
                # 弹窗提示❗
            pyautogui.alert('🎉恭喜任务执行完毕🎉\n单击确定退出!')
    else:
        # 弹窗警告❌
        pyautogui.alert('❌数据检查失败❌\n单击确定退出!')

# print(os.getcwd())
