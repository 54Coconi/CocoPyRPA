import tkinter.simpledialog as simpledialog

import pyautogui
import xlrd

import PyRPA_pkg.check_mod as cm
import PyRPA_pkg.functions_mod as fm
import ver_desc

# tk = Tk()
# tk.title("任务执行目录")
# tk.geometry("330x180")  # 设置对话框的宽度和高度

res = simpledialog.askstring('任务执行目录', prompt='请输入任务目录的绝对路径', initialvalue='任务目录绝对路径')
if res is not None:
    print()
    print(res)

# tk.mainloop()
""" 
This modules handles dialog boxes.
It contains the following public symbols:
SimpleDialog -- A simple but flexible modal dialog box
Dialog -- a base class for dialogs
askinteger -- get an integer from the user
askfloat -- get a float from the user
askstring -- get a string from the user 
"""


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

        # 指令<鼠标定点移动>
        i += 1


if __name__ == '__main__':
    # 文件名
    filename = 'D:\\桌面自动化python程序\\PyRPA桌面自动化程序\\execute\\test\\PyRPA_v2.0.0指令测试.xls'
    # 打开文件
    wb = xlrd.open_workbook(filename)
    # 通过索引获取表格sheet页
    sheet1 = wb.sheet_by_index(0)
    print(ver_desc.__doc__)
    # 数据检查
    checkCmd = cm.data_check(sheet1)
    if checkCmd:
        key = input('请选择功能: \n\t输入1只做一次\n'
                    '\t输入2循环n次 \n\t  ')
        if key == '1':
            # 循环拿出每一行指令
            mainWork(sheet1)
            # 弹窗提示
            pyautogui.alert('🎉恭喜任务执行完毕🎉\n单击确定退出!')
        elif key == '2':
            n = int(input('循环次数n = '))
            while n > 0:
                mainWork(sheet1)
                fm.time.sleep(0.1)
                print("等待0.1秒")
                n -= 1
            # 弹窗提示
            pyautogui.alert('🎉恭喜任务执行完毕🎉\n单击确定退出!')
        else:
            pyautogui.alert('输入错误❗❗❗')

    else:
        # 弹窗警告❌
        pyautogui.alert('❌数据检查失败❌\n单击确定退出!')
