"""
工具模块
"""
import re


def MyRegexMatch(pattern, restr):
    """
    正则表达式匹配
    :param pattern: 正则表达式匹配规则
    :param restr: 需要匹配的字符串
    :return: isRestr -- true/false
    """
    isRestr = True
    result = re.match(pattern, restr)
    if result is None:
        isRestr = False
    return isRestr

# [a-zA-Z]:\\+.+\.[a-z]+$  文件绝对路径格式匹配
# D:\桌面自动化python程序\PyRPA桌面自动化程序\execute\test\example\PyRPA_v2.x指令测试.xls


# def pathResolution(path, char):
#     raw(path)
#     re.sub(pattern='\\', repl='/', string=path)
#     pathList = path.split('char')
#     for i in range(len(pathList)):
#         print(pathList[i])
