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


class RPA_mouse:
    """
    鼠标类指令
    """

    def __init__(self):
        print("This is RPA_mouse class.")

    # 1.<单击左键> 和 2.<双击左键>
    @staticmethod
    def clickL(sheetName, rowIndex, clickTimes):
        """
        鼠标左键指令功能方法
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
        mouseClick(clickTimes, "left", imgName, reTry)
        if clickTimes == 1:
            print("<单击左键>\t=============>\t", imgName)
        elif clickTimes == 2:
            print("<双击左键>\t=============>\t", imgName)

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
        mouseClick(clickTimes, "right", imgName, reTry)
        print("<单击右键>\t=============>\t", imgName)

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
        print("滚轮滑动了\t=============>\t", int(scroll), "距离")

    # 5.<鼠标定点移动>
    @staticmethod
    def clickMoveTo(sheetName, rowIndex):
        """
        鼠标移动到指定的绝对坐标
        :param sheetName: 表名
        :param rowIndex: 行标
        """

    # 6.<鼠标相对移动>


class RPA_keyboard:
    """
    按键类指令
    """

    def __init__(self):
        print("This is RPA_keyboard class.")

    # 7.<输入>

    # 8.<按键>

    # 9.<热键组合>

    # 10.<键盘输入TXT内容>


# @控制类
class RPA_control():
    def __init__(self):
        print("This is RPA_control class.")
