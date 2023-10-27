"""
主执行程序
"""
import datetime
import os
import sys
import traceback

import pyautogui
import xlrd

import PyRPA_pkg.CocoPyRPAlogger as mylog
import PyRPA_pkg.check_mod as cm
import PyRPA_pkg.functions_mod as fm
import ver_desc

filename = os.path.basename(__file__).split('.')[0]
logpath = os.path.abspath('..') + '/logs/'

my_logg = mylog.MyLog(fileNameOfWhoUse=filename, logOutPath=logpath).logger
my_logg.info('<<=======================任务开始=======================>>')
my_logg.info("任务开始执行时间 {}\n".format(datetime.datetime.now()))


# my_logg.debug('看看debug')
# my_logg.error('This is a error' + '\n')


def mainWork(sheetName):
    """
    主执行方法（按行从上至下顺序依次执行）
    :param sheetName: 表名
    """
    i = 1  # 当前行数为 i+1
    while i < sheetName.nrows:
        # 取本行指令的操作类型名称（第一列）
        cmdType = sheetName.row(i)[0]
        # 取本行指令是否需要执行（第四列）
        isRun_cmd = sheetName.row(i)[3]

        # 指令<单击左键>
        if cmdType.value == '单击左键' and isRun_cmd.value != 0:
            fm.RPA_mouse.clickL(sheetName, i, 1)
            my_logg.info(fm.funLogStr)

        # 指令<双击左键>
        if cmdType.value == '双击左键' and isRun_cmd.value != 0:
            fm.RPA_mouse.clickL(sheetName, i, 2)
            my_logg.info(fm.funLogStr)

        # 指令<单击右键>
        if cmdType.value == '单击右键' and isRun_cmd.value != 0:
            fm.RPA_mouse.clickR(sheetName, i, 1)
            my_logg.info(fm.funLogStr)

        # 指令<滚轮>
        if cmdType.value == '滚轮' and isRun_cmd.value != 0:
            fm.RPA_mouse.myScroll(sheetName, i)
            my_logg.info(fm.funLogStr)

        # 指令<鼠标相对移动>
        if cmdType.value == '鼠标相对移动' and isRun_cmd.value != 0:
            fm.RPA_mouse.clickRelMove(sheetName, i)
            my_logg.info(fm.funLogStr)

        # 指令<鼠标定点移动>
        if cmdType.value == '鼠标定点移动' and isRun_cmd.value != 0:
            fm.RPA_mouse.clickMoveTo(sheetName, i)
            my_logg.info(fm.funLogStr)

        # 指令<输入>
        if cmdType.value == '输入' and isRun_cmd.value != 0:
            fm.RPA_keyboard.pasteboardInput(sheetName, i)
            my_logg.info(fm.funLogStr)

        # 指令<按键>
        if cmdType.value == '按键' and isRun_cmd.value != 0:
            fm.RPA_keyboard.keystroke(sheetName, i)
            my_logg.info(fm.funLogStr)

        # 指令<热键组合>
        if cmdType.value == '热键组合' and isRun_cmd.value != 0:
            fm.RPA_keyboard.hotkeyCombi(sheetName, i)
            my_logg.info(fm.funLogStr)

        # 指令<键盘输入TXT内容>
        if cmdType.value == '键盘输入TXT内容' and isRun_cmd.value != 0:
            fm.RPA_keyboard.EnterTxtOnKeyboard(sheetName, i)
            my_logg.info(fm.funLogStr)

        # 指令<等待>
        if cmdType.value == '等待' and isRun_cmd.value != 0:
            fm.RPA_control.waitTime(sheetName, i)
            my_logg.info(fm.funLogStr)

        i += 1


if __name__ == '__main__':
    # 文件名
    # filename1 = 'D:/桌面自动化python程序/PyRPA桌面自动化程序/execute/test/example/PyRPA_v2.0.0指令测试.xls'

    filename = pyautogui.prompt(text='请输入[任务表格]的绝对路径', title='CocoPyRPA--表格路径',
                                default='形如 D:/aa/bb.xls')

    if filename is None:
        my_logg.info('<<**********************任务被取消**********************>>\n')
        sys.exit(0)  # 程序终止

    try:
        w = xlrd.open_workbook(filename)
    except OSError as e:
        my_logg.error('输入的路径错误或不存在\n' + str(traceback.format_exc()) + '\n')
        pyautogui.alert(text='\n\n路径错误！', title='CocoPyRPA--警告', button='退出')
    # 打开文件
    wb = xlrd.open_workbook(filename)
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print(ver_desc.__doc__)
    # 数据检查
    checkCmd = cm.data_check(sheet1)
    if checkCmd:
        pyautogui.alert(text='\n\n数据检查成功', title='CocoPyRPA--提示', button='继续')
        key = pyautogui.confirm(text='\n\n请选择功能:\n输入1只做一次,输入2循环n次', title='CocoPyRPA--功能选择',
                                buttons=['1', '2'])
        if key == '1':
            # 循环拿出每一行指令
            mainWork(sheet1)
            my_logg.info('<<======================任务执行成功======================>>\n')
            # 弹窗提示
            pyautogui.alert('🎉恭喜任务执行完毕🎉\n单击确定退出!')
        elif key == '2':
            n = pyautogui.prompt(text='请输入循环次数', title='CocoPyRPA--循环次数', default='10')
            n = int(n)
            while n > 0:
                mainWork(sheet1)
                if n != 1:
                    my_logg.info("等待0.1秒后再次执行")
                fm.time.sleep(0.1)
                n -= 1
            my_logg.info('<<======================任务执行成功======================>>\n')
            # 弹窗提示
            pyautogui.alert(text='🎉恭喜任务执行完毕🎉', title='CocoPyRPA--提示', button='退出')
    else:
        # 弹窗警告❌
        pyautogui.alert(text='\n\n❌数据检查失败❌', title='CocoPyRPA--退出', button='退出')
