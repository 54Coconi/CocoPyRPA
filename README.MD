# CocoPyRPA--基于 PyAutoGUI + OpenCV的桌面自动化python程序

## 1. 项目背景

> **声明：原项目由Bilibili UP 主 [@不高兴就喝水](https://space.bilibili.com/412704776 "个人空间")
完成，本人在其基础上稍作了修改，在此感谢 [@不高兴就喝水](https://space.bilibili.com/412704776 "个人空间") 给予的灵感**
>
> 原视频连接：[传送门](https://www.bilibili.com/video/BV1T34y1o73U/ "视频链接")

* 平时工作、学习需要完成许多具有规律性、重复性且又繁杂的任务，这令我们缺少了更多时间去（摸鱼^o^）学习新知识、新技能，故此项目由此诞生
* 原项目的代码冗余和注释都不够好，且所有功能集于一个python文件，其拓展能力以及后期维护或新增功能都比较麻烦，为解决此问题，我在最新的V2.0.0版本中采用模块化开发，对其进行了部分优化
* 本人初学python（主学的是java后端开发），做此项目也算是一次python练手了(>_<)

## 2. 安装说明：

### 2.1 项目结构

```text

../
  |--history/
      |--模板_v1.0.0.py
      |--模板_v1.0.0指令测试.xls
      |--模板_v1.1.0.py
      |--模板_v1.1.0指令测试.xls
  |--logs/    
  |--src/   
      |--PyRPA_pkg/
          |--__pycache__/
          |--__init__.py
          |--check_mod.py
          |--CocoPyRPAlogger.py
          |--functions_mod.py
          |--utils_mod.py
      |--main.py              --> 主执行文件
      |ver_desc.py
  |--test/
      |--example/
          |--PyRPA_v2.0.0指令测试.xls
          |--test1.png
  |--.gitignore
  |--README.md

```

### 2.2 安装之前的准备

> 1. 安装python3.4以上版本，并配置好系统环境变量（目前有装3.9遇到坑的，我个人用的3.10.8），不会配的请自行百度
>
> 2. 安装依赖包
>
>    方法：在cmd中（win+R 输入cmd 回车）输入
>
>    pip install pyperclip 回车
>
>    pip install xlrd 回车
>
>    pip install pyautogui 回车
>
>    pip install opencv-python -i https://pypi.tuna.tsinghua.edu.cn/simple 回车
>
>    pip install pillow 回车
>
>    这几步如果哪步没成功，请自行百度 如 pip install opencv-python失败
>
>    （我个人强烈建议使用 PyCharm 自带的 python 库安装（[PyCharm官网](https://www.jetbrains.com/pycharm/download/?section=windows)），如下图所示）
>
>    ![PyCharm.png](https://s2.loli.net/2023/10/27/GafiOCLH4X2rtR7.png)
>
>    3.

### 2.3 安装

直接安装右侧 **release** 发行版或者 **clone** 本仓库到本地。

## 3. 使用示例：

> 1. 使用方法：把每一步要操作的图标、区域截图保存至一个文件夹中，**绝对路径不能含有中文**字符（例如：*D:/任务图片/xxx.png* 是不合格的），**图片必须为png格式**（注意如果同屏有多个相同图标，程序默认找到最左上的一个，因此怎么截图，截多大的区域很重要，比如输入框只截中间空白部分肯定是不行的，宗旨就是“唯一”）。
>
> 2. 在与任务图片相同目录文件夹下建一个以` .xls`为后缀名的 Excel 表格（如果不是此格式会出问题），在表格里的第一行是命令的类型、内容等等，详见项目文件 `test/example/PyRPA_v2.x指令测试.xls`。表格里是所有你想要执行的功能指令
>
> 3. 如果你之前的准备都完成了，则双击` main.py`文件即可启动程序。程序启动后会弹窗提示（如下图1所示）输入任务表格的**绝对路径**(最好不带中文字符)，点击` OK`后程序就会去电脑中找到你的 excel 表格。如果表格不存在或者输入的路径错误就会弹窗提示（如下图2所示）
>
>    ​																			图1
>
> ![PyRPA-start.png](https://s2.loli.net/2023/10/31/CzHYN9W21jMJVUR.png)
>
> ​																				图2
>
> ![路径警告](https://s2.loli.net/2023/10/31/NREGAOTvt5UYisu.png)
>
> 4. 在程序读取到表格里的内容时，如果表格内填写的指令内容都正确，则会弹出数据检查成功的弹窗，如果数据检查失败也会弹窗提示，并将错误写到 `logs/` 文件夹下的 ` check_mod[xx.xx.xx].log` 日志里，可以打开日志文件查看具体是表格的什么地方数据有误。
>
> 
>
> PyAutoGUI库的一些实例参考：
>
> 1. [Python自动操作GUI神器PyAutoGUI - 知乎](https://zhuanlan.zhihu.com/p/302592540)
> 2. [pyautogui 官方 github 源码](https://github.com/asweigart/pyautogui)
> 3. [pyautogui 官方 github 中文文档](https://github.com/asweigart/pyautogui/blob/master/docs/simplified-chinese.ipynb)
> 4. [pyautogui 官网说明文档（英文）](https://pyautogui.readthedocs.io/en/latest/index.html)

## 4. 配置选项：

暂无

## 5. 文档链接：

暂无

## 7. 版本历史：

> v1.0.0 **支持中文指令**
>
> 新增指令 `<键盘输入TXT内容>` 、`<热键组合>`、`<输入>`。优化代码和注释
>
> 
>
> v1.1.0 **加入gui弹窗提示功能**
>
> 目前只使用了 pyautogui 自带的gui，优化代码和注释
>
> 
>
> v2.0.0 **使用模块化开发重构了项目**
>
> 新增了部分功能（更多`<热键组合>` 、增加`<鼠标定点移动>` ）
>
> 
>
> v2.1.0 **加入日志模块**
>
> 新增将报错信息和执行过程信息写入日志文件，同时支持在控制台查看运行信息
>
> 
>
> 
>
> 
>
> 
