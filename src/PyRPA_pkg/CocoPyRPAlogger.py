"""
CocoPyRPA 的自定义日志模块
"""
import logging
import sys
from time import strftime

# 输出日志路径
# PATH = os.path.abspath('.') + '/logs/'
# 设置日志格式，和时间格式
FMT = '%(asctime)s %(filename)s [line:%(lineno)d]  %(levelname)s:  %(message)s'
DATEFMT = '%Y-%m-%d %H:%M:%S'


class MyLog(object):
    """
    自定义的日志类
    """

    def __init__(self, fileNameOfWhoUse, logOutPath):
        """
        初始化
        :param fileNameOfWhoUse: 调用Mylog类的程序文件名（不含拓展名）
        :param logOutPath: 日志输出目录路径
        """
        self.logger = logging.getLogger()
        # 初始化自定义日志格式
        self.formatter = logging.Formatter(fmt=FMT, datefmt=DATEFMT)
        # 初始化自定义日志文件名
        self.log_filename = '{0}{1}[{2}].log'.format(logOutPath, fileNameOfWhoUse, strftime('%Y-%m-%d'))

        self.logger.addHandler(self.get_file_handler(self.log_filename))
        self.logger.addHandler(self.get_console_handler())
        # 设置日志的默认级别
        self.logger.setLevel(logging.DEBUG)

    def get_file_handler(self, filename):
        """
        输出到文件handler的函数定义
        :param filename:
        :return:
        """
        filehandler = logging.FileHandler(filename, encoding="utf-8")
        filehandler.setFormatter(self.formatter)
        return filehandler

    def get_console_handler(self):
        """
        输出到控制台handler的函数定义
        :return:
        """
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(self.formatter)
        return console_handler
