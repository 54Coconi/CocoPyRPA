"""
主执行程序
"""
import sys

import pyautogui
import xlrd

import PyRPA_pkg.check_mod as cm
import PyRPA_pkg.functions_mod as fm
import ver_desc


# tk = Tk()
# tk.title("任务执行目录")
# tk.geometry("330x180")  # 设置对话框的宽度和高度
# res = simpledialog.askstring('任务执行目录', prompt='请输入任务目录的绝对路径', initialvalue='任务目录绝对路径')
# if res is not None:
#     print()
#     print(res)
# tk.mainloop()

def mainWork(sheetName):
    """
    主执行方法
    :param sheetName: 表名
    """
    # 当前行数为 i+1
    i = 1
    while i < sheetName.nrows:
        # 取本行指令的操作类型（第一列）
        cmdType = sheetName.row(i)[0]

        # 指令<单击左键>
        if cmdType.value == '单击左键':
            fm.RPA_mouse.clickL(sheetName, i, 1)

        # 指令<双击左键>
        if cmdType.value == '双击左键':
            fm.RPA_mouse.clickL(sheetName, i, 2)

        # 指令<单击右键>
        if cmdType.value == '单击右键':
            fm.RPA_mouse.clickR(sheetName, i, 1)

        # 指令<滚轮>
        if cmdType.value == '滚轮':
            fm.RPA_mouse.myScroll(sheetName, i)

        # 指令<鼠标相对移动>
        if cmdType.value == '鼠标相对移动':
            fm.RPA_mouse.clickRelMove(sheetName, i)

        # 指令<鼠标定点移动>
        if cmdType.value == '鼠标定点移动':
            fm.RPA_mouse.clickMoveTo(sheetName, i)

        # 指令<输入>
        if cmdType.value == '输入':
            fm.RPA_keyboard.pasteboardInput(sheetName, i)

        # 指令<按键>
        if cmdType.value == '按键':
            fm.RPA_keyboard.keystroke(sheetName, i)

        # 指令<热键组合>
        if cmdType.value == '热键组合':
            fm.RPA_keyboard.hotkeyCombi(sheetName, i)
        # 指令<键盘输入TXT内容>
        if cmdType.value == '键盘输入TXT内容':
            fm.RPA_keyboard.EnterTxtOnKeyboard(sheetName, i)
        # 指令<等待>
        if cmdType.value == '等待':
            fm.RPA_control.waitTime(sheetName, i)
        i += 1


if __name__ == '__main__':
    # 文件名
    filename1 = 'D:\\桌面自动化python程序\\PyRPA桌面自动化程序\\execute\\test\\PyRPA_v2.0.0指令测试.xls'
    filename = pyautogui.prompt(text='请输入表格的绝对路径', title='表格路径',
                                default='形如 D:\\' + '\\aa\\' + '\\bb.xls')
    if filename is None:
        print('<<=============任务被取消=============>>')
        sys.exit(0)  # 程序终止
    # 打开文件
    wb = xlrd.open_workbook(filename)
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print(ver_desc.__doc__)
    # 数据检查
    checkCmd = cm.data_check(sheet1)
    if checkCmd:
        pyautogui.alert(text='\n\n数据检查成功', title='提示', button='继续')
        key = pyautogui.confirm(text='\n\n请选择功能:\n输入1只做一次,输入2循环n次', title='功能选择',
                                buttons=['1', '2'])
        if key == '1':
            # 循环拿出每一行指令
            mainWork(sheet1)
            print('<<=============任务执行成功=============>>')
            # 弹窗提示
            pyautogui.alert('🎉恭喜任务执行完毕🎉\n单击确定退出!')
        elif key == '2':
            n = pyautogui.prompt(text='请输入循环次数', title='循环次数', default='10')
            n = int(n)
            while n > 0:
                mainWork(sheet1)
                fm.time.sleep(0.1)
                print("等待0.1秒")
                n -= 1
            print('<<=============任务执行成功=============>>')
            # 弹窗提示
            pyautogui.alert('🎉恭喜任务执行完毕🎉\n单击确定退出!')
        else:
            pyautogui.alert('输入错误❗❗❗')

    else:
        # 弹窗警告❌
        pyautogui.alert('❌数据检查失败❌\n单击确定退出!')
