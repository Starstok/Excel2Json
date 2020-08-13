# This Python file uses the following encoding: utf-8
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QMessageBox, QFileDialog
from mainUi import Ui_Form
from exc2Json import list2json
from json2Exc import write2xlsx
from excFormula import conFormula

class excTools(QWidget, Ui_Form):
    def __init__(self):
        QWidget.__init__(self)
        self.ui = Ui_Form()
        self.ui.setupUi(self)

        # 绑定触发信号
        self.ui.pushButton.clicked.connect(self.getFlirName)
        self.ui.pushButton_2.clicked.connect(self.getPathName)
        self.ui.pushButton_3.clicked.connect(self.startConJson)
        self.ui.pushButton_4.clicked.connect(self.startConFormula)
        self.ui.pushButton_5.clicked.connect(self.cleanConFormula)
        self.ui.pushButton_6.clicked.connect(self.startConExc)

    def getFlirName(self):
        path, fileType = QFileDialog.getOpenFileName(self, "选取文件", "", "file(*.xlsx *.xls *.json)")
        if not path:
            return
        # if fileType.find('*.xlsx') or fileType.find('*.json'):
        # 获取到文件信息放入lineEdit
        self.ui.lineEdit.setText(path)
        return

    def getPathName(self):
        selectPath = QFileDialog.getExistingDirectory(self,  "选取文件夹")
        if not selectPath:
            return
        else:
            # 获取到路径信息放入lineEdit_2
            self.ui.lineEdit_2.setText(selectPath)
            return

    def startConJson(self):
        # 读取lineEdit内容
        inPutFile = self.ui.lineEdit.text()
        outPutFile = self.ui.lineEdit_2.text()
        sysType = self.ui.comboBox.currentText() #获取当前下拉列表框中的文本信息
        
        if not inPutFile or not outPutFile:
            QMessageBox.critical(self, "严重的", "请选取文件，和保存路径")
            return
        else:
            msgStatus = list2json(inPutFile, outPutFile+"/"+sysType+"_new.json", sysType)
            if msgStatus == 1:
                QMessageBox.critical(self, "严重的", "请选择对应的转换表格 ！！")
            else:
                QMessageBox.information(self, "提示", "转换完成")
            return

    def startConFormula(self):
        # 读取lineEdit内容
        inPutFile = self.ui.lineEdit.text()
        baseData = self.ui.lineEdit_3.text()
        dataType = self.ui.comboBox_2.currentText() #获取当前下拉列表框中的文本信息

        if not inPutFile:
            QMessageBox.critical(self, "严重的", "请选取文件")
            return
        else:
            formulaText = conFormula(inPutFile, dataType, baseData)
            if formulaText == 1:
                QMessageBox.critical(self, "严重的", "请选择对应的转换表格 ！！")
            else:
                self.ui.textEdit.setText(formulaText)
            return

    def cleanConFormula(self):
        self.ui.textEdit.setText('')

    def startConExc(self):
        # 读取lineEdit内容
        inPutFile = self.ui.lineEdit.text()
        outPutFile = self.ui.lineEdit_2.text()
        sysType = self.ui.comboBox.currentText() #获取当前下拉列表框中的文本信息

        if not inPutFile or not outPutFile:
            QMessageBox.critical(self, "严重的", "请选取文件，和保存路径")
            return
        else:
            msgStatus = write2xlsx(inPutFile, outPutFile+"/"+sysType+"_new.xlsx", sysType)
            if msgStatus == 1:
                QMessageBox.critical(self, "严重的", "请选择对应的格式文件 ！！")
            else:
                QMessageBox.information(self, "提示", "转换完成")
            return


if __name__ == "__main__":
    app = QApplication([])
    window = excTools()
    window.show()
    sys.exit(app.exec_())
