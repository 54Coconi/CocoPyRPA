"""
功能模块,包含各个功能函数
+----------------------------------------------+
|目前包括的功能(指令)，指令实现功能详见说明文档|
+----------------------------------------------+
【about mouse 鼠标类】
1.<单击左键>
2.<双击左键>
3.<单击右键>
4.<滚轮>
5.<鼠标定点移动>
6.<鼠标相对移动>

【about keyboard 按键类】
7.<输入>
8.<按键>
9.<热键组合>
10.<键盘输入TXT内容>
【about control 控制类】
11.<等待>
"""
import time

import pyautogui
import pyperclip

""" 
====================================
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
    pyautogui.move(x,y)
    print("<鼠标相对移动>\t======>\t",
    "右移" if x > 0 else "左移" if x < 0 else "不移动", x,
    "下移" if y > 0 else "上移" if y < 0 else "不移动", y)

# 8)指令<按键>
elif cmdType.value == '按键':
    inputValue = sheet1.row(i)[1].value
    if inputValue == "enter" :
        pyautogui.press("enter")
    time.sleep(0.5)
    print("<按键>\t======>\t", inputValue)

# 9)指令<键盘输入TXT内容>
elif cmdType.value == '键盘输入TXT内容':
    inputValue = sheet1.row(i)[1].value
    # 读取文本文件的内容
    filepath = inputValue  # 文件绝对路径
    with open(filepath, 'r', encoding='UTF-8') as file:  # ‘r’只读文件
            copy_input = file.readline()  # 逐行读取,结果是一个list
    keyboard_input = copy_input
    pyautogui.typewrite(keyboard_input, interval=0.025)  # 放在列表里，interval 指输入间隔秒
    print("<键盘输入TXT内容>",keyboard_input)

==========================================================================
"""


def mouseClick(clickTimes, lOrR, img, reTry):
    """
    定义鼠标点击事件
    :param clickTimes: 点击次数
    :param lOrR: 左键 或 右键
    :param img: 图片路径
    :param reTry: 重复次数
    """
    if reTry == 1:
        while True:
            location = pyautogui.locateCenterOnScreen(img, confidence=0.9)
            if location is not None:
                pyautogui.click(location.x, location.y, clicks=clickTimes, interval=0.2, duration=0.2, button=lOrR)
                break
            print('未找到匹配图片,0.1秒后重试')
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
                print('点击第', i, '次')
                i += 1
            time.sleep(0.1)


class RPA_mouse:
    """
    鼠标指令类
    """

    def __init__(self):
        print('This is RPA_mouse class.')

    # 1.<单击左键> 和 2.<双击左键>
    @staticmethod
    def clickL(sheetName, rowIndex, clickTimes):
        """
        鼠标左键指令功能方法
        :param sheetName: 表名
        :param rowIndex: 行标
        :param clickTimes: 点击次数
        """
        # 取图片名称(路径)
        imgName = sheetName.row(rowIndex)[1].value
        # 设置默认执行次数为1
        reTry = 1
        # 重复次数必须为数字类型且重复次数不为0（若不满足则按默认的执行1次）
        if sheetName.row(rowIndex)[2].ctype == 2 and sheetName.row(rowIndex)[2].value != 0:
            # 取重试次数
            reTry = sheetName.row(rowIndex)[2].value
        mouseClick(clickTimes, 'left', imgName, reTry)
        if clickTimes == 1:
            print('<单击左键>\t\t=============>\t', imgName)
        elif clickTimes == 2:
            print('<双击左键>\t\t=============>\t', imgName)

    # 3.<单击右键>
    @staticmethod
    def clickR(sheetName, rowIndex, clickTimes):
        """
        鼠标右键指令功能方法
        :param sheetName: 表名
        :param rowIndex: 行标
        :param clickTimes: 点击次数
        """
        # 取图片名称
        imgName = sheetName.row(rowIndex)[1].value
        # 设置默认执行次数为1
        reTry = 1
        # 重复次数必须为数字类型且重复次数不为0（若不满足则按默认的执行1次）
        if sheetName.row(rowIndex)[2].ctype == 2 and sheetName.row(rowIndex)[2].value != 0:
            # 取重试次数
            reTry = sheetName.row(rowIndex)[2].value
        mouseClick(clickTimes, 'right', imgName, reTry)
        print('<单击右键>\t\t=============>\t', imgName)

    # 4.<滚轮>
    @staticmethod
    def myScroll(sheetName, rowIndex):
        """
        鼠标滚轮滚动指令方法
        :param sheetName: 表名
        :param rowIndex: 行标
        """
        # 取单元格中要移动的距离值
        scroll = sheetName.row(rowIndex)[1].value
        pyautogui.scroll(int(scroll))
        print('<滚轮>\t\t\t=============>\t',
              '向下滑动了' if int(scroll) < 0 else
              '向上滑动了' if int(scroll) > 0 else
              '滑动了', abs(int(scroll)), '距离')

    # 5.<鼠标定点移动>
    @staticmethod
    def clickMoveTo(sheetName, rowIndex):
        """
        鼠标移动到指定的绝对坐标
        :param sheetName: 表名
        :param rowIndex: 行标
        """
        xy_list = sheetName.row(rowIndex)[1].value.split(',')
        x = int(float(xy_list[0].split('(')[1]))
        y = int(float(xy_list[1].split(')')[0]))
        pyautogui.moveTo(x, y, 0.1)
        print('<鼠标定点移动>\t=============>\t', 'x=', x, 'y=', y)

    # 6.<鼠标相对移动>
    @staticmethod
    def clickRelMove(sheetName, rowIndex):
        """
        鼠标相对当前坐标移动
        :param sheetName: 表名
        :param rowIndex: 行标
        """
        xy_list = sheetName.row(rowIndex)[1].value.split('|')
        x = int(float(xy_list[0].split('(')[1]))
        y = int(float(xy_list[1].split(')')[0]))
        pyautogui.move(x, y, 0.5)
        print("<鼠标相对移动>\t=============>\t",
              "x右移" if x > 0 else "x左移" if x < 0 else "x移动", abs(x),
              '，', "y下移" if y > 0 else "y上移" if y < 0 else "y移动",
              abs(y))


class RPA_keyboard:
    """
    按键指令类
    """

    def __init__(self):
        print('This is RPA_keyboard class.')

    # 7.<输入>
    @staticmethod
    def pasteboardInput(sheetName, rowIndex):
        """
        复制表格单元格里的内容到粘贴板，并粘贴输入内容
        """
        # 取单元格中要输入的内容
        inputVal = sheetName.row(rowIndex)[1].value
        # 复制单元格内容到粘贴板
        pyperclip.copy(inputVal)
        # 粘贴内容
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        print("剪贴板<输入>\t\t=============>\t", inputVal)

    # 8.<按键>
    @staticmethod
    def keystroke(sheetName, rowIndex):
        """
        获取表格单元格内容，并根据内容模拟键盘击键
        执行键盘按键按下，然后松开。
        """
        inputval = sheetName.row(rowIndex)[1].value
        pyautogui.press(inputval)

    # 9.<热键组合>
    @staticmethod
    def hotkeyCombi(sheetName, rowIndex):
        """
        获取表格单元格的内容（形如 'key1+key2+key3+...'）并将其分割得到一个列表
        key_list['key1','key2',...] 并将整个 list 传入 hotkey() 函数中。
        对按顺序传递的参数执行按键按下，然后按相反顺序执行按键释放。
        其效果是调用热键（“ctrl”，“shift”，“c”）将执行 “Ctrl-Shift-C” 热键键盘快捷键。
        """
        inputVal = sheetName.row(rowIndex)[1].value
        key_list = inputVal.split('+')
        pyautogui.hotkey(key_list)

    # 10.<键盘输入TXT内容>
    @staticmethod
    def EnterTxtOnKeyboard(sheetName, rowIndex):
        """
        键盘输入TXT内容(由于是模拟键盘的按键按下和释放，所以不支持中文)
        从表格单元格获取 txt 文件路径，找到文件并打开得到内容，然后对内容中的每个字符执行键盘按键按下，然后释放
        """
        inputVal = sheetName.row(rowIndex)[1].value
        filepath = inputVal
        with open(filepath, 'r', encoding='UTF-8') as file:  # ‘r’只读文件
            message = file.read()
        pyautogui.typewrite(message, interval=0.025)  # interval 指输入间隔秒
        print("<键盘输入TXT内容>\t=============>\t", message)


# @控制类
class RPA_control:
    """
    系统控制类
    """

    def __init__(self):
        print('This is RPA_control class.')

    # 11.<等待>
    @staticmethod
    def waitTime(sheetName, rowIndex):
        """
        从表格单元格获取延时时长
        """
        wait_time = sheetName.row(rowIndex)[1].value
        print('<等待>\t\t\t=============>\t', wait_time, '秒')
        time.sleep(wait_time)
