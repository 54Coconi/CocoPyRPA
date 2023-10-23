"""
检查模块，用于检查表格填写的指令及其内容
==========================================================
||                   【指令列数据检查】                   || 
==========================================================
# cmdType.value|   1单击左键        
#              |   2双击左键
#              |   3单击右键        
#              |   4输入
#              |   5等待            
#              |   6滚轮
#              |   7按键            
#              |   8热键组合
#              |   9键盘输入TXT内容(键盘输入文件内容)
#              |   10鼠标相对移动   
#              |   11鼠标定点移动
# -------------+------------------------------------------
# ctype        |   空:0
#              |   字符串:1
#              |   数字:2
#              |   日期:3
#              |   布尔:4
#              |   error:5
#--------------+------------------------------------------
"""
import pyautogui
import PyRPA_pkg.utils_mod as um

__name__ = '指令功能检查模块'


def data_check(sheetName):
    """
    检查任务表格第1，2，4列（忽略第一行）的数据类型和内容
    :param sheetName: 表名
    :return: check_cmd-(true/false)
    """
    check_cmd = True
    # 行数检查
    if sheetName.nrows < 2:
        print('表格没有指令数据！')
        check_cmd = False
        return check_cmd
    # 每行数据检查（i表示行数）
    i = 1
    while i < sheetName.nrows:
        # 【第1列 操作指令类型检查】=========================
        cmdType = sheetName.row(i)[0]
        if cmdType.ctype != 1 or (
                cmdType.value != '单击左键' and cmdType.value != '双击左键' and cmdType.value != '单击右键'
                and cmdType.value != '输入' and cmdType.value != '等待' and cmdType.value != '滚轮'
                and cmdType.value != '鼠标相对移动' and cmdType.value != '按键' and cmdType.value != '键盘输入TXT内容'
                and cmdType.value != '鼠标定点移动' and cmdType.value != '热键组合'):
            print('第', i + 1, '行,第1列数据有误,可能输入了错误的或不能识别的操作指令！\n')
            check_cmd = False

        # 【第2列 操作指令内容检查】=================================
        cmdValue = sheetName.row(i)[1]

        # 读图点击类型指令，内容必须为字符串类型
        if (cmdType.value == '单击左键' or
                cmdType.value == '双击左键' or
                cmdType.value == '单击右键'):
            if cmdValue.ctype != 1:
                print('第', i + 1, '行,第2列数据有误,应为字符串类型,实际却为：', cmdValue.value)
                check_cmd = False

        # 输入类型，内容不能为空
        if cmdType.value == '输入':
            if cmdValue.ctype == 0:
                print('第', i + 1, '行,第2列数据有误,输入内容不能为空！')
                check_cmd = False

        # 等待类型，内容必须为数字
        if cmdType.value == '等待':
            if cmdValue.ctype != 2:
                print('第', i + 1, '行,第2列数据有误,应为数字类型,实际却为：', cmdValue.value)
                check_cmd = False

        # 滚轮事件，内容必须为数字
        if cmdType.value == '滚轮':
            if cmdValue.ctype != 2:
                print('第', i + 1, '行,第2列数据有误,应为数字类型,实际却为：', cmdValue.value)
                check_cmd = False

        # 鼠标相对当前坐标的移动事件，形式必须为形如'num1=num2'的字符串且x,y取值在屏幕分辨率范围内
        if cmdType.value == '鼠标相对移动':
            xy_list = cmdValue.value.split('=')
            # 将str字符串转为int整型（防止有小数点'.'先转为float浮点型再转int）
            x = int(float(xy_list[0]))
            y = int(float(xy_list[1]))
            width = pyautogui.size().width
            height = pyautogui.size().height
            if abs(x) > width or abs(y) > height:
                print('第', i + 1, '行,第2列数据有误,鼠标移动距离超出屏幕最大分辨率！')
                check_cmd = False

        # 鼠标移动到绝对坐标事件，形式必须为形如'(x,y)'的字符串且x,y取值在屏幕分辨率范围内
        if cmdType.value == '鼠标定点移动':
            isRestr = um.MyRegexMatch('\(-?\d+,-?\d+\)', cmdValue.value)
            # 检查格式
            if isRestr is False:
                print('第', i + 1, '行,第2列数据有误,输入的坐标格式不对，请检查表格！')
            else:
                xy_list = cmdValue.value.split(',')
                x = int(float(xy_list[0].split('(')[1]))
                y = int(float(xy_list[1].split(')')[0]))
                # print(xy_list)
                # print('x=' + '%d' % x, end=' ')
                # print('y=' + '%d' % y)
                width = pyautogui.size().width
                height = pyautogui.size().height
                if abs(x) > width or abs(y) > height:
                    print('第', i + 1, '行,第2列数据有误,鼠标移动距离超出屏幕最大分辨率！')
                    check_cmd = False

        # 【第4列，操作指令是否执行（只能为空或者数字0）
        isRun_cmd = sheetName.row(i)[3]
        # print(isRun_cmd.value)
        # print(isRun_cmd.ctype)
        if isRun_cmd.ctype != 0:
            if isRun_cmd.value != 0:
                print('第', i + 1, '行,第4列数据有误,应为数字0或空,实际却为：', isRun_cmd.value)
                check_cmd = False

        i += 1
    return check_cmd
