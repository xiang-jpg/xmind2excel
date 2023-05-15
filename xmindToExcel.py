# encoding=gbk
from PyQt5 import QtCore, QtWidgets
import sys

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QLabel, QComboBox, QMessageBox, QLineEdit

from xmindparser import xmind_to_dict
# ʱ�䴦��ģ��datetime
import datetime
# ʹ��pandasģ����е���csv
import pandas as pd

'''
    ��ȡ��Ӧƥ��Ľڵ�
    @:param topic ����Դ
    @:param string ��Ҫƥ������ڵ�����
    @:return ƥ�����ڵ�����������
'''


def get_node(topic, string):
    # �ж��Ƿ�Ϊ���飬���������ѭ��
    if isinstance(topic, list):
        for v in topic:
            node = get_node(v, string)
            if node != None:
                return node
    else:
        # �жϱ���ƥ�䣬ƥ���򷵻�
        if topic.get('title') == string:
            return topic
        # ��ȡ�ӽڵ���еݹ�
        if (topic.get('topics') == None):
            return None
        return get_node(topic.get('topics'), string)


'''
    ���ڵ�ת��Ϊ�ֵ�,�����ֵ����key�ظ�
    @:param node �ڵ�
    @:param dic �ֵ�
    @:param key �ֵ�key
'''


def get_dic(node, dic, key=''):
    # �ж��Ƿ�Ϊ���飬���������ѭ��
    if isinstance(node, list):
        for v in node:
            get_dic(v, dic, key)
        return
    # �ж��Ƿ�����ӽڵ�
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
    �����ֵ䣬�������
    @:param file_name �ļ���
    @:param dic �ֵ�
    @:param encoding �ַ�����
'''


def generate_excel(self, path, file_name, dic, encoding):
    # ��ȡ��������ȼ�
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
        title_list.append(f'��{rstrip}��{title}')
        priority_list.append(maker)
        empty_list.append('')
        type_list.append('���ܲ���')
    # ��ȡ��ǰʱ��
    str_time = datetime.datetime.now().strftime('%Y��%m��%d��%Hʱ%M��%S��')
    # ����csv
    df = pd.DataFrame({
        '�������': empty_list,
        '������Ʒ': empty_list,
        '����ģ��': empty_list,
        '��ع���': empty_list,
        '��������': title_list,
        'ǰ������': empty_list,
        '����': empty_list,
        'Ԥ��': empty_list,
        'ʵ�����': empty_list,
        '�ؼ���': empty_list,
        '���ȼ�': priority_list,
        '��������': type_list
    })

    df.to_csv(f'{path}/{file_name}{str_time}.csv', encoding=encoding, index=False)
    QMessageBox.about(None, "ת��", "ת����ɣ�")


'''
    �����ֵ䣬�������
    @:param file_name �ļ���
    @:param dic �ֵ�
'''


def convert_handle(self, file_name, path, string, encoding):
    try:
        # ��ȡxmind����
        ReadXmind = xmind_to_dict(file_name)
    except:
        QMessageBox.warning(None, "�����½�������", "�ļ������쳣��")
        return
    # ��ȡ�����½ڵ���
    topic = ReadXmind[0]['topic']
    # ��ȡ���ڵ��Ӧ�����νṹ
    node = get_node(topic, string)
    if node == None:
        QMessageBox.warning(None, "�����½�������", "�Ҳ������ڵ㣡")
        return
    # ��ȡ���ڵ�������ӽڵ�
    if node["topics"] == None:
        QMessageBox.warning(None, "�����½�������", "���ڵ���û���ӽڵ㣡")
        return
    node = node["topics"]
    # �����ֵ�
    dic = {}
    # ѭ���������е��ӽڵ�ת��Ϊ�ֵ�
    get_dic(node, dic)
    # �ֵ����ת��Ϊexcel
    generate_excel(self, path, string, dic, encoding)


# ui��������
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        # �����ڲ�������
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowIcon(QIcon("favicon.ico"))
        MainWindow.resize(848, 721)
        # �ر����
        MainWindow.setWindowFlags(Qt.WindowMinimizeButtonHint | Qt.WindowCloseButtonHint)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        # ����xmind��������
        self.xmind_button = QtWidgets.QPushButton(self.centralwidget)
        self.xmind_button.setGeometry(QtCore.QRect(57, 150, 175, 28))
        self.xmind_button.setObjectName("file")
        self.xmind_button.setStyleSheet(" margin: 0px; padding: 0px;")
        # ����excel��������
        self.excel_button = QtWidgets.QPushButton(self.centralwidget)
        self.excel_button.setGeometry(QtCore.QRect(57, 450, 175, 28))
        self.excel_button.setObjectName("file")
        self.excel_button.setStyleSheet(" margin: 0px; padding: 0px;")

        # ����xmind��ʾ���ڲ���
        self.xmind_text = QtWidgets.QPushButton(self.centralwidget)
        self.xmind_text.setGeometry(QtCore.QRect(300, 150, 480, 28))
        self.xmind_text.setObjectName("file")
        self.xmind_text.setStyleSheet(" margin: 0px; padding: 0px;")

        # ����excel��ʾ���ڲ���
        self.excel_text = QtWidgets.QPushButton(self.centralwidget)
        self.excel_text.setGeometry(QtCore.QRect(300, 450, 480, 28))
        self.excel_text.setObjectName("file")
        self.excel_text.setStyleSheet(" margin: 0px; padding: 0px;")
        # ���������б����
        self.q_label = QLabel("��ѡ�е�����Excel�ַ����룺", self.centralwidget)
        self.q_label.setGeometry(QtCore.QRect(57, 350, 480, 28))
        self.encoding = "UTF-8"
        self.box = QComboBox(self.centralwidget)
        self.box.setGeometry(QtCore.QRect(300, 350, 480, 28))
        self.box.addItems(["UTF-8", "GBK"])
        # ����ת����������
        self.convert = QtWidgets.QPushButton(self.centralwidget)
        self.convert.setGeometry(QtCore.QRect(400, 550, 100, 28))
        self.convert.setStyleSheet(" margin: 0px; padding: 0px;")
        # �������ڵ������������
        self.node_label = QLabel("������Xmind�������ڵ����ƣ�", self.centralwidget)
        self.node_label.setGeometry(QtCore.QRect(57, 250, 480, 28))
        self.nodeName = QLineEdit(self.centralwidget)
        self.nodeName.setPlaceholderText('���������ڵ�����')
        self.nodeName.setGeometry(QtCore.QRect(300, 250, 480, 28))
        self.nodeName.setStyleSheet(" margin: 0px; padding: 0px;")
        # �����ڼ��˵�������������
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
        ################button��ť����¼��ص�����################
        # xmind
        self.xmind_button.clicked.connect(self.xmind_msg)
        # excel
        self.excel_button.clicked.connect(self.excel_msg)
        # �����¼��ص�
        self.box.currentIndexChanged.connect(self.selectChange)
        # ת����ť
        self.convert.clicked.connect(self.convert_msg)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        # ����
        MainWindow.setWindowTitle(_translate("MainWindow", "xmind����ת��excel����"))
        # xmind��ť
        self.xmind_button.setText(_translate("MainWindow", "ѡ�е���Xmind����·��"))
        # xmind�ı�
        self.xmind_text.setText(_translate("MainWindow", ""))
        # excel��ť
        self.excel_button.setText(_translate("MainWindow", "ѡ�񵼳�Excel����·��"))
        # excel�ı�
        self.excel_text.setText(_translate("MainWindow", ""))
        # ת���ı�
        self.convert.setText(_translate("MainWindow", "ת��"))

    #########ѡ��xmind����#########
    def xmind_msg(self, Filepath):
        fileName, fileType = QtWidgets.QFileDialog.getOpenFileName(None, "ѡȡ�ļ�", "./", "Xmind Files (*.xmind)")
        self.xmind_text.setText(fileName)

    #########ѡ��excel����·��#########
    def excel_msg(self, Filepath):
        directory = QtWidgets.QFileDialog.getExistingDirectory(None, "ѡȡ�ļ�", "./")
        self.excel_text.setText(directory)

    #########ѡ���ַ�����#########
    def selectChange(self, i):
        self.encoding = self.box.currentText()

    #########ת����ť�¼�#########
    def convert_msg(self):
        if not self.xmind_text.text():
            QMessageBox.warning(None, "�����쳣", "��ѡ��xmind����·����")
            return
        if not self.excel_text.text():
            QMessageBox.warning(None, "�����쳣", "��ѡ��excel����·����")
            return
        if not self.nodeName.text():
            QMessageBox.warning(None, "�����쳣", "���������ڵ����ƣ�")
            return
        if not self.box.currentText():
            QMessageBox.warning(None, "�����쳣", "��ѡ��excel�����ַ����룡")
            return
        convert_handle(self, self.xmind_text.text(), self.excel_text.text(), self.nodeName.text(), self.box.currentText())


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(mainWindow)
    mainWindow.show()
    sys.exit(app.exec_())
