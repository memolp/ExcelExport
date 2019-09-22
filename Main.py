#-*- coding:utf-8 -*-

import os,sys

from PyQt4.QtCore import *
from PyQt4.QtGui import *

import ConfigParser

class QMainWindow(QWidget):
    """ """
    def __init__(self):
        """ """
        super(QMainWindow,self).__init__()
        self.setWindowTitle("配置导出工具")
        self.resize(600,600)

        vLayout = QVBoxLayout()

        #第一行 选择配置表的路径
        hLayout = QHBoxLayout()
        hLayout.addWidget(QLabel("选择配置表路径"))
        self.mConfigPath = QLineEdit("")
        self.mConfigPath.setReadOnly(True)
        hLayout.addWidget(self.mConfigPath)
        btnSelect = QPushButton("浏览")
        btnSelect.clicked.connect(self.onSelectConfigPath)
        hLayout.addWidget(btnSelect)

        #第二行 选择导出的路径
        h2Layout = QHBoxLayout()
        h2Layout.addWidget(QLabel("选择导出的路径"))
        self.mExportPath = QLineEdit("")
        self.mExportPath.setReadOnly(True)
        h2Layout.addWidget(self.mExportPath)
        btnEPath = QPushButton("选择")
        btnEPath.clicked.connect(self.onSelectExportPath)
        h2Layout.addWidget(btnEPath)

        #展示EXCEL的表名列表
        h3Layout = QHBoxLayout()
        self.mExcelTablesList = QListWidget()
        h3Layout.addWidget(self.mExcelTablesList)

        #导出
        h4Layout = QHBoxLayout()
        h4Layout.addStretch(1)
        btnExportSelected = QPushButton("选择导出")
        btnExportSelected.clicked.connect(self.onExportSelect)
        h4Layout.addWidget(btnExportSelected)
        btnExportAll = QPushButton("全部导出")
        btnExportAll.clicked.connect(self.onExportAll)
        h4Layout.addWidget(btnExportAll)

        vLayout.addLayout(hLayout)
        vLayout.addLayout(h2Layout)
        vLayout.addLayout(h3Layout)
        vLayout.addLayout(h4Layout)
        self.setLayout(vLayout)

    def onSelectConfigPath(self,event):
        """ """
        fname = QFileDialog.getOpenFileName(self,"选择Excel",".","Excel(*.xlsx);;Excel(*.xls);;All(*.*)")
        if not fname or  fname[0] == '': #沒有选择文件
            return 
        excel_file = fname[0]
        if not excel_file.endswith(".xls") and not excel_file.endswith("xlsx"): 
            QMessageBox.critical(self,"出错提示","请选择Excel配置表")
            return
        if not os.path.exists(excel_file):
            QMessageBox.critical(self,"出错提示","配置表不存在!")
            return
        self.mConfigPath.setText(excel_file)
        #清除旧的数据
        self.mExcelTablesList.clear()
        workbook = ConfigParser.OpenExcel(self.mConfigPath.text())
        for sheetName in workbook.sheet_names():
            item = QListWidgetItem()
            cb   = QCheckBox(sheetName)
            self.mExcelTablesList.addItem(item)
            self.mExcelTablesList.setItemWidget(item,cb)

    def onSelectExportPath(self,event):
        """ """
        fname = QFileDialog.getExistingDirectory(self,"选择导出目录",".")
        if fname == '': #没有选择文件夹
            return
        if not os.path.exists(fname):
            QMessageBox.critical(self,"出错提示","选择的目录不存在")
            return
        self.mExportPath.setText(fname)

    def _checkInputValid(self):
        """ """
        excel_file = self.mConfigPath.text()
        export_dir = self.mExportPath.text()
        if excel_file == "" or export_dir == "":
            QMessageBox.critical(self,"出错提示","导出的配置表或导出目录设置为空!")
            return False
        if not os.path.exists(excel_file):
            QMessageBox.critical(self,"出错提示","配置表不存在!")
            return False
        if not os.path.exists(export_dir):
            QMessageBox.critical(self,"出错提示","导出的路径不存在!")
            return False
        return True

    def onExportAll(self,event):
        """ """
        if not self._checkInputValid():
            return
        workbook = ConfigParser.OpenExcel(self.mConfigPath.text())
        for sheet in workbook.sheets():
            s = ConfigParser.ExcelSheetParser(sheet,self.mExportPath.text())
            s.Export()
        QMessageBox.information(self,"提示","导出完成")

    def onExportSelect(self,event):
        """ """
        if not self._checkInputValid():
            return
        #找到需要导出的表
        export_sheet = []
        for i in range(self.mExcelTablesList.count()):
            cb = self.mExcelTablesList.itemWidget(self.mExcelTablesList.item(i))
            if cb.isChecked():
                export_sheet.append(cb.text())

        #执行导出
        workbook = ConfigParser.OpenExcel(self.mConfigPath.text())
        for sheet in workbook.sheets():
            if sheet.name in export_sheet:
                s = ConfigParser.ExcelSheetParser(sheet,self.mExportPath.text())
                s.Export()
        QMessageBox.information(self,"提示","导出完成")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = QMainWindow()
    win.show()
    sys.exit(app.exec_())