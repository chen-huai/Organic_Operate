# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Organic_Operate_Ui.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(445, 795)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/ch.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.tabWidget.setFont(font)
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.gridLayout_4 = QtWidgets.QGridLayout(self.tab)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.groupBox_2 = QtWidgets.QGroupBox(self.tab)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.groupBox_2.setFont(font)
        self.groupBox_2.setObjectName("groupBox_2")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.groupBox_2)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.textBrowser = QtWidgets.QTextBrowser(self.groupBox_2)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.textBrowser.setFont(font)
        self.textBrowser.setObjectName("textBrowser")
        self.gridLayout_2.addWidget(self.textBrowser, 0, 0, 1, 1)
        self.gridLayout_4.addWidget(self.groupBox_2, 3, 0, 1, 2)
        self.groupBox = QtWidgets.QGroupBox(self.tab)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.groupBox.setFont(font)
        self.groupBox.setObjectName("groupBox")
        self.gridLayout = QtWidgets.QGridLayout(self.groupBox)
        self.gridLayout.setObjectName("gridLayout")
        self.pushButton = QtWidgets.QPushButton(self.groupBox)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton.sizePolicy().hasHeightForWidth())
        self.pushButton.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.gridLayout.addWidget(self.pushButton, 0, 0, 1, 1)
        self.pushButton_4 = QtWidgets.QPushButton(self.groupBox)
        self.pushButton_4.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_4.sizePolicy().hasHeightForWidth())
        self.pushButton_4.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setObjectName("pushButton_4")
        self.gridLayout.addWidget(self.pushButton_4, 1, 0, 1, 1)
        self.lineEdit = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit.sizePolicy().hasHeightForWidth())
        self.lineEdit.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.lineEdit.setFont(font)
        self.lineEdit.setAlignment(QtCore.Qt.AlignCenter)
        self.lineEdit.setObjectName("lineEdit")
        self.gridLayout.addWidget(self.lineEdit, 2, 0, 1, 1)
        self.gridLayout_4.addWidget(self.groupBox, 1, 0, 1, 1)
        self.groupBox_3 = QtWidgets.QGroupBox(self.tab)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.groupBox_3.setFont(font)
        self.groupBox_3.setObjectName("groupBox_3")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.groupBox_3)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.spinBox = QtWidgets.QSpinBox(self.groupBox_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.spinBox.sizePolicy().hasHeightForWidth())
        self.spinBox.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.spinBox.setFont(font)
        self.spinBox.setMinimum(1)
        self.spinBox.setMaximum(999999999)
        self.spinBox.setObjectName("spinBox")
        self.gridLayout_3.addWidget(self.spinBox, 2, 0, 1, 1)
        self.pushButton_3 = QtWidgets.QPushButton(self.groupBox_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_3.sizePolicy().hasHeightForWidth())
        self.pushButton_3.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.pushButton_3.setFont(font)
        self.pushButton_3.setObjectName("pushButton_3")
        self.gridLayout_3.addWidget(self.pushButton_3, 4, 0, 1, 1)
        self.spinBox_2 = QtWidgets.QSpinBox(self.groupBox_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.spinBox_2.sizePolicy().hasHeightForWidth())
        self.spinBox_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.spinBox_2.setFont(font)
        self.spinBox_2.setMaximum(999999999)
        self.spinBox_2.setObjectName("spinBox_2")
        self.gridLayout_3.addWidget(self.spinBox_2, 2, 1, 1, 1)
        self.pushButton_2 = QtWidgets.QPushButton(self.groupBox_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_2.sizePolicy().hasHeightForWidth())
        self.pushButton_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setObjectName("pushButton_2")
        self.gridLayout_3.addWidget(self.pushButton_2, 4, 1, 1, 1)
        self.pushButton_8 = QtWidgets.QPushButton(self.groupBox_3)
        self.pushButton_8.setEnabled(False)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_8.sizePolicy().hasHeightForWidth())
        self.pushButton_8.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.pushButton_8.setFont(font)
        self.pushButton_8.setStatusTip("")
        self.pushButton_8.setObjectName("pushButton_8")
        self.gridLayout_3.addWidget(self.pushButton_8, 0, 0, 1, 1)
        self.pushButton_5 = QtWidgets.QPushButton(self.groupBox_3)
        self.pushButton_5.setEnabled(False)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_5.sizePolicy().hasHeightForWidth())
        self.pushButton_5.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.pushButton_5.setFont(font)
        self.pushButton_5.setObjectName("pushButton_5")
        self.gridLayout_3.addWidget(self.pushButton_5, 0, 1, 1, 1)
        self.pushButton_6 = QtWidgets.QPushButton(self.groupBox_3)
        self.pushButton_6.setEnabled(False)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_6.sizePolicy().hasHeightForWidth())
        self.pushButton_6.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.pushButton_6.setFont(font)
        self.pushButton_6.setObjectName("pushButton_6")
        self.gridLayout_3.addWidget(self.pushButton_6, 3, 0, 1, 1)
        self.doubleSpinBox = QtWidgets.QDoubleSpinBox(self.groupBox_3)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.doubleSpinBox.sizePolicy().hasHeightForWidth())
        self.doubleSpinBox.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.doubleSpinBox.setFont(font)
        self.doubleSpinBox.setDecimals(3)
        self.doubleSpinBox.setSingleStep(0.001)
        self.doubleSpinBox.setProperty("value", 0.01)
        self.doubleSpinBox.setObjectName("doubleSpinBox")
        self.gridLayout_3.addWidget(self.doubleSpinBox, 3, 1, 1, 1)
        self.gridLayout_4.addWidget(self.groupBox_3, 1, 1, 1, 1)
        self.groupBox_4 = QtWidgets.QGroupBox(self.tab)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox_4.sizePolicy().hasHeightForWidth())
        self.groupBox_4.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.groupBox_4.setFont(font)
        self.groupBox_4.setObjectName("groupBox_4")
        self.gridLayout_6 = QtWidgets.QGridLayout(self.groupBox_4)
        self.gridLayout_6.setObjectName("gridLayout_6")
        self.checkBox = QtWidgets.QCheckBox(self.groupBox_4)
        self.checkBox.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBox.sizePolicy().hasHeightForWidth())
        self.checkBox.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.checkBox.setFont(font)
        self.checkBox.setChecked(False)
        self.checkBox.setObjectName("checkBox")
        self.gridLayout_6.addWidget(self.checkBox, 2, 0, 1, 1)
        self.checkBox_2 = QtWidgets.QCheckBox(self.groupBox_4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBox_2.sizePolicy().hasHeightForWidth())
        self.checkBox_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.checkBox_2.setFont(font)
        self.checkBox_2.setChecked(False)
        self.checkBox_2.setObjectName("checkBox_2")
        self.gridLayout_6.addWidget(self.checkBox_2, 4, 0, 1, 1)
        self.pushButton_58 = QtWidgets.QPushButton(self.groupBox_4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_58.sizePolicy().hasHeightForWidth())
        self.pushButton_58.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.pushButton_58.setFont(font)
        self.pushButton_58.setObjectName("pushButton_58")
        self.gridLayout_6.addWidget(self.pushButton_58, 4, 4, 1, 1)
        self.pushButton_57 = QtWidgets.QPushButton(self.groupBox_4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_57.sizePolicy().hasHeightForWidth())
        self.pushButton_57.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.pushButton_57.setFont(font)
        self.pushButton_57.setObjectName("pushButton_57")
        self.gridLayout_6.addWidget(self.pushButton_57, 2, 4, 1, 1)
        self.pushButton_59 = QtWidgets.QPushButton(self.groupBox_4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.pushButton_59.sizePolicy().hasHeightForWidth())
        self.pushButton_59.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.pushButton_59.setFont(font)
        self.pushButton_59.setObjectName("pushButton_59")
        self.gridLayout_6.addWidget(self.pushButton_59, 0, 4, 1, 1)
        self.checkBox_3 = QtWidgets.QCheckBox(self.groupBox_4)
        self.checkBox_3.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.checkBox_3.sizePolicy().hasHeightForWidth())
        self.checkBox_3.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.checkBox_3.setFont(font)
        self.checkBox_3.setChecked(False)
        self.checkBox_3.setObjectName("checkBox_3")
        self.gridLayout_6.addWidget(self.checkBox_3, 0, 0, 1, 1)
        self.spinBox_7 = QtWidgets.QSpinBox(self.groupBox_4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.spinBox_7.sizePolicy().hasHeightForWidth())
        self.spinBox_7.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.spinBox_7.setFont(font)
        self.spinBox_7.setProperty("value", 20)
        self.spinBox_7.setObjectName("spinBox_7")
        self.gridLayout_6.addWidget(self.spinBox_7, 4, 2, 1, 1)
        self.lineEdit_7 = QtWidgets.QLineEdit(self.groupBox_4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_7.sizePolicy().hasHeightForWidth())
        self.lineEdit_7.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.lineEdit_7.setFont(font)
        self.lineEdit_7.setObjectName("lineEdit_7")
        self.gridLayout_6.addWidget(self.lineEdit_7, 4, 3, 1, 1)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.groupBox_4)
        self.lineEdit_2.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_2.sizePolicy().hasHeightForWidth())
        self.lineEdit_2.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.lineEdit_2.setFont(font)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.gridLayout_6.addWidget(self.lineEdit_2, 2, 2, 1, 2)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.groupBox_4)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.lineEdit_3.sizePolicy().hasHeightForWidth())
        self.lineEdit_3.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(9)
        self.lineEdit_3.setFont(font)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.gridLayout_6.addWidget(self.lineEdit_3, 0, 2, 1, 2)
        self.gridLayout_4.addWidget(self.groupBox_4, 0, 0, 1, 2)
        self.tabWidget.addTab(self.tab, "")
        self.verticalLayout.addWidget(self.tabWidget)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 445, 23))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionExport = QtWidgets.QAction(MainWindow)
        self.actionExport.setObjectName("actionExport")
        self.actionImport = QtWidgets.QAction(MainWindow)
        self.actionImport.setObjectName("actionImport")
        self.actionExit = QtWidgets.QAction(MainWindow)
        self.actionExit.setObjectName("actionExit")
        self.actionHelp = QtWidgets.QAction(MainWindow)
        self.actionHelp.setObjectName("actionHelp")
        self.actionAuthor = QtWidgets.QAction(MainWindow)
        self.actionAuthor.setObjectName("actionAuthor")
        self.menuFile.addAction(self.actionExport)
        self.menuFile.addAction(self.actionImport)
        self.menuFile.addSeparator()
        self.menuFile.addAction(self.actionExit)
        self.menuHelp.addAction(self.actionHelp)
        self.menuHelp.addSeparator()
        self.menuHelp.addAction(self.actionAuthor)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Organic Operate"))
        MainWindow.setStatusTip(_translate("MainWindow", "作者：Frank Chen"))
        self.groupBox_2.setTitle(_translate("MainWindow", "信息提示"))
        self.textBrowser.setStatusTip(_translate("MainWindow", "消息提示"))
        self.groupBox.setTitle(_translate("MainWindow", "Starlims操作"))
        self.pushButton.setStatusTip(_translate("MainWindow", "选择Batch文件，并获取Sample ID"))
        self.pushButton.setText(_translate("MainWindow", "选择Starlims Batch"))
        self.pushButton_4.setStatusTip(_translate("MainWindow", "清除内容"))
        self.pushButton_4.setText(_translate("MainWindow", "清零"))
        self.lineEdit.setStatusTip(_translate("MainWindow", "内容为Sample ID时:可以填写；内容为Stop时:无法填写"))
        self.lineEdit.setText(_translate("MainWindow", "Sample ID"))
        self.groupBox_3.setTitle(_translate("MainWindow", "自动填写"))
        self.spinBox.setStatusTip(_translate("MainWindow", "起始编号"))
        self.pushButton_3.setStatusTip(_translate("MainWindow", "停止填写"))
        self.pushButton_3.setText(_translate("MainWindow", "停止填写"))
        self.spinBox_2.setStatusTip(_translate("MainWindow", "结尾编号，0默认表示该单号最后一个"))
        self.pushButton_2.setStatusTip(_translate("MainWindow", "点击开始后，3秒时间选择起始位置，并从起始位置开始向下填写单号"))
        self.pushButton_2.setText(_translate("MainWindow", "开始填写"))
        self.pushButton_8.setText(_translate("MainWindow", "起始编号"))
        self.pushButton_5.setText(_translate("MainWindow", "结尾编号"))
        self.pushButton_6.setText(_translate("MainWindow", "填写速度"))
        self.doubleSpinBox.setStatusTip(_translate("MainWindow", "自动填写速度"))
        self.groupBox_4.setTitle(_translate("MainWindow", "Tlims操作"))
        self.checkBox.setText(_translate("MainWindow", "重复"))
        self.checkBox_2.setText(_translate("MainWindow", "QC"))
        self.pushButton_58.setText(_translate("MainWindow", "查看数据"))
        self.pushButton_57.setText(_translate("MainWindow", "导出数据"))
        self.pushButton_59.setText(_translate("MainWindow", "选择 Tlims Batch"))
        self.checkBox_3.setText(_translate("MainWindow", "质控"))
        self.lineEdit_7.setText(_translate("MainWindow", "CQC"))
        self.lineEdit_2.setText(_translate("MainWindow", ";A;B"))
        self.lineEdit_3.setText(_translate("MainWindow", "BLK;BLK-S;S-S"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "Auto"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.menuHelp.setTitle(_translate("MainWindow", "Help"))
        self.actionExport.setText(_translate("MainWindow", "Export"))
        self.actionImport.setText(_translate("MainWindow", "Import"))
        self.actionExit.setText(_translate("MainWindow", "Exit"))
        self.actionHelp.setText(_translate("MainWindow", "Help"))
        self.actionAuthor.setText(_translate("MainWindow", "Author"))
