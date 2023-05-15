# encoding=gbk
from PyQt5 import QtCore, QtWidgets
import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QLabel, QComboBox, QMessageBox, QLineEdit

from xmindparser import xmind_to_dict
# 时间处理模块datetime
import datetime
# 使用pandas模块进行导出csv
import pandas as pd

'''
    获取对应匹配的节点
    @:param topic 数据源
    @:param string 需要匹配的主节点名称
    @:return 匹配主节点下所有数据
'''


def get_node(topic, string):
    # 判断是否为数组，数据则进行循环
    if isinstance(topic, list):
        for v in topic:
            node = get_node(v, string)
            if node != None:
                return node
    else:
        # 判断标题匹配，匹配则返回
        if topic.get('title') == string:
            return topic
        # 获取子节点进行递归
        if (topic.get('topics') == None):
            return None
        return get_node(topic.get('topics'), string)


'''
    将节点转化为字典,倒置字典避免key重复
    @:param node 节点
    @:param dic 字典
    @:param key 字典key
'''


def get_dic(node, dic, key=''):
    # 判断是否为数组，数据则进行循环
    if isinstance(node, list):
        for v in node:
            get_dic(v, dic, key)
        return
    # 判断是否存在子节点
    if (node.get('topics') == None):
        key = key + node.get('title')
        dic[key] = node
        return
    else:
        if key != '':
            key += '-'
        key = key + node.get('title')
        if node.get('topics') != None:
            get_dic(node.get('topics'), dic, key)


'''
    解析字典，创建表格
    @:param file_name 文件名
    @:param dic 字典
    @:param encoding 字符编码
'''


def generate_excel(self, path, file_name, dic, encoding):
    # 获取标题和优先级
    title_list = []
    priority_list = []
    empty_list = []
    type_list = []
    for key in dic:
        title = dic[key].get('title')
        makers = dic[key].get('makers')
        maker = 4
        if makers != None:
            for v in makers:
                if v == 'priority-1':
                    maker = 1
                if v == 'priority-2':
                    maker = 2
                if v == 'priority-3':
                    maker = 3
        rstrip = key.rstrip(title)
        title_list.append(f'【{rstrip}】{title}')
        priority_list.append(maker)
        empty_list.append('')
        type_list.append('功能测试')
    # 获取当前时间
    str_time = datetime.datetime.now().strftime('%Y年%m月%d日%H时%M分%S秒')
    # 导出csv
    df = pd.DataFrame({
        '用例编号': empty_list,
        '所属产品': empty_list,
        '所属模块': empty_list,
        '相关故事': empty_list,
        '用例标题': title_list,
        '前置条件': empty_list,
        '步骤': empty_list,
        '预期': empty_list,
        '实际情况': empty_list,
        '关键词': empty_list,
        '优先级': priority_list,
        '用例类型': type_list
    })

    df.to_csv(f'{path}/{file_name}{str_time}.csv', encoding=encoding, index=False)
    QMessageBox.about(None, "转换", "转换完成！")


'''
    解析字典，创建表格
    @:param file_name 文件名
    @:param dic 字典
'''


def convert_handle(self, file_name, path, string, encoding):
    try:
        # 读取xmind数据
        ReadXmind = xmind_to_dict(file_name)
    except:
        QMessageBox.warning(None, "请重新进行设置", "文件解析异常！")
        return
    # 获取画布下节点树
    topic = ReadXmind[0]['topic']
    # 获取主节点对应的树形结构
    node = get_node(topic, string)
    if node == None:
        QMessageBox.warning(None, "请重新进行设置", "找不到主节点！")
        return
    # 获取主节点下面的子节点
    if node["topics"] == None:
        QMessageBox.warning(None, "请重新进行设置", "主节点下没有子节点！")
        return
    node = node["topics"]
    # 定义字典
    dic = {}
    # 循环遍历所有的子节点转换为字典
    get_dic(node, dic)
    # 字典解析转换为excel
    generate_excel(self, path, string, dic, encoding)


# ui界面设置
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        # 主窗口参数设置
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowIcon(QIcon("favicon.ico"))
        MainWindow.resize(848, 721)
        # 关闭最大化
        MainWindow.setWindowFlags(Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        # 设置xmind按键参数
        self.xmind_button = QtWidgets.QPushButton(self.centralwidget)
        self.xmind_button.setGeometry(QtCore.QRect(57, 150, 175, 28))
        self.xmind_button.setObjectName("file")
        self.xmind_button.setStyleSheet(" margin: 0px; padding: 0px;")
        # 设置excel按键参数
        self.excel_button = QtWidgets.QPushButton(self.centralwidget)
        self.excel_button.setGeometry(QtCore.QRect(57, 450, 175, 28))
        self.excel_button.setObjectName("file")
        self.excel_button.setStyleSheet(" margin: 0px; padding: 0px;")

        # 设置xmind显示窗口参数
        self.xmind_text = QtWidgets.QPushButton(self.centralwidget)
        self.xmind_text.setGeometry(QtCore.QRect(300, 150, 480, 28))
        self.xmind_text.setObjectName("file")
        self.xmind_text.setStyleSheet(" margin: 0px; padding: 0px;")

        # 设置excel显示窗口参数
        self.excel_text = QtWidgets.QPushButton(self.centralwidget)
        self.excel_text.setGeometry(QtCore.QRect(300, 450, 480, 28))
        self.excel_text.setObjectName("file")
        self.excel_text.setStyleSheet(" margin: 0px; padding: 0px;")
        # 设置下拉列表对象
        self.q_label = QLabel("请选中导出的Excel字符编码：", self.centralwidget)
        self.q_label.setGeometry(QtCore.QRect(57, 350, 480, 28))
        self.encoding = "UTF-8"
        self.box = QComboBox(self.centralwidget)
        self.box.setGeometry(QtCore.QRect(300, 350, 480, 28))
        self.box.addItems(["UTF-8", "GBK"])
        # 设置转换按键参数
        self.convert = QtWidgets.QPushButton(self.centralwidget)
        self.convert.setGeometry(QtCore.QRect(400, 550, 100, 28))
        self.convert.setStyleSheet(" margin: 0px; padding: 0px;")
        # 设置主节点名称输入参数
        self.node_label = QLabel("请输入Xmind用例主节点名称：", self.centralwidget)
        self.node_label.setGeometry(QtCore.QRect(57, 250, 480, 28))
        self.nodeName = QLineEdit(self.centralwidget)
        self.nodeName.setPlaceholderText('请输入主节点名称')
        self.nodeName.setGeometry(QtCore.QRect(300, 250, 480, 28))
        self.nodeName.setStyleSheet(" margin: 0px; padding: 0px;")
        # 主窗口及菜单栏标题栏设置
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 848, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)
        ################button按钮点击事件回调函数################
        # xmind
        self.xmind_button.clicked.connect(self.xmind_msg)
        # excel
        self.excel_button.clicked.connect(self.excel_msg)
        # 下拉事件回调
        self.box.currentIndexChanged.connect(self.selectChange)
        # 转换按钮
        self.convert.clicked.connect(self.convert_msg)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        # 标题
        MainWindow.setWindowTitle(_translate("MainWindow", "xmind用例转换excel用例"))
        # xmind按钮
        self.xmind_button.setText(_translate("MainWindow", "选中导入Xmind用例路径"))
        # xmind文本
        self.xmind_text.setText(_translate("MainWindow", ""))
        # excel按钮
        self.excel_button.setText(_translate("MainWindow", "选择导出Excel用例路径"))
        # excel文本
        self.excel_text.setText(_translate("MainWindow", ""))
        # 转换文本
        self.convert.setText(_translate("MainWindow", "转换"))

    #########选择xmind用例#########
    def xmind_msg(self, Filepath):
        fileName, fileType = QtWidgets.QFileDialog.getOpenFileName(None, "选取文件", "./", "Xmind Files (*.xmind)")
        self.xmind_text.setText(fileName)

    #########选择excel用例路径#########
    def excel_msg(self, Filepath):
        directory = QtWidgets.QFileDialog.getExistingDirectory(None, "选取文件", "./")
        self.excel_text.setText(directory)

    #########选择字符编码#########
    def selectChange(self, i):
        self.encoding = self.box.currentText()

    #########转换按钮事件#########
    def convert_msg(self):
        if not self.xmind_text.text():
            QMessageBox.warning(None, "参数异常", "请选择xmind用例路径！")
            return
        if not self.excel_text.text():
            QMessageBox.warning(None, "参数异常", "请选择excel导出路径！")
            return
        if not self.nodeName.text():
            QMessageBox.warning(None, "参数异常", "请输入主节点名称！")
            return
        if not self.box.currentText():
            QMessageBox.warning(None, "参数异常", "请选择excel导出字符编码！")
            return
        convert_handle(self, self.xmind_text.text(), self.excel_text.text(), self.nodeName.text(), self.box.currentText())


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(mainWindow)
    mainWindow.show()
    sys.exit(app.exec_())
