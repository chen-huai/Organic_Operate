# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Organic_Operate_Ui.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

# from PyQt5 import QtCore, QtGui, QtWidgets
# from PyQt5.QtWidgets import *
# from PyQt5.QtCore import *
from PyQt5.QtWidgets import QMessageBox, QFileDialog, QMainWindow, QApplication, QItemDelegate, QTableWidgetItem, QInputDialog
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QLineEdit
from Organic_Operate_Ui import *
from Tlims_Data_Operate import *
from Table_Ui import *
from Get_Data import *
import chicon  # 引用图标

class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)


        self.tabWidget.setCurrentIndex(0)
        self.pushButton.clicked.connect(self.getBatch)
        self.pushButton_3.clicked.connect(self.stopMessage)
        self.pushButton_2.clicked.connect(self.autoWrite)
        self.actionExport.triggered.connect(self.exportConfig)
        self.actionImport.triggered.connect(self.importConfig)
        self.actionExit.triggered.connect(MyMainWindow.close)
        self.actionAuthor.triggered.connect(self.showAuthorMessage)
        self.actionHelp.triggered.connect(self.textBrowser.clear)
        self.pushButton_4.clicked.connect(self.clearContent)
        self.pushButton_59.clicked.connect(self.getTlimsBatchsUrl)
        self.pushButton_57.clicked.connect(self.exportTlimsBatch)
        self.pushButton_58.clicked.connect(self.viewData)
        # QtCore.QMetaObject.connectSlotsByName(MyMainWindow)


    def getConfig(self):
        # 初始化，获取或生成配置文件
        global configFileUrl
        global desktopUrl
        global now
        global last_time
        global today
        # getBatch里的
        global labNumber
        global selectBatchFile
        now = int(time.strftime('%Y'))
        last_time = now - 1
        today = time.strftime('%Y%m%d')
        desktopUrl = os.path.join(os.path.expanduser("~"), 'Desktop')
        configFileUrl = '%s/config' % desktopUrl
        configFile = os.path.exists('%s/config_organic.txt' % configFileUrl)
        # print(desktopUrl,configFileUrl,configFile)
        if not configFile:  # 判断是否存在文件夹如果不存在则创建为文件夹
            reply = QMessageBox.question(self, '信息', '确认是否要创建配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                if not os.path.exists(configFileUrl):
                    os.makedirs(configFileUrl)
                MyMainWindow.createConfigContent(self)
                MyMainWindow.getConfigContent(self)
                self.textBrowser.append("创建并导入配置成功")
            else:
                exit()
        else:
            MyMainWindow.getConfigContent(self)

    # def getConfigContent(self):
    #     # 获取配置文件内容
    #     f1 = open('%s/config_organic.txt' % configFileUrl, "r", encoding="utf-8")
    #     global configContent
    #     configContent = {}
    #     i = 0
    #     for line in f1:
    #         if line != '\n':
    #             lineContent = line.split('||||||')
    #             # print(lineContent)
    #             configContent['%s' % lineContent[0]] = lineContent[1].split('\n')[0]
    #         i += 1
    #     # print(configContent)
    #     self.textBrowser.append("获取配置文件成功")
    #
    # def createConfigContent(self):
    #     # 生成默认配置文件
    #     configContentName = ['选择Organic_Batch的输入路径和结果输出路径', 'Organic_Batch_Import_URL']
    #     configContent = ['默认，可更改为自己需要的', 'Z:\\Organic Batch']
    #     f1 = open('%s/config_organic.txt' % configFileUrl, "w", encoding="utf-8")
    #     i = 0
    #     for i in range(len(configContentName)):
    #         f1.write(configContentName[i] + '||||||' + configContent[i] + '\n')
    #         i += 1
    #     self.textBrowser.append("配置文件创建成功")
    #     QMessageBox.information(self, "提示信息",
    #                             "默认配置文件已经创建好，\n如需修改请在用户桌面查找config文件夹中config_organic.txt，\n将相应的文件内容替换成用户需求即可，修改后记得重新导入配置文件。\n切记：中间‘||||||’六根，不能多也不能少！！！",
    #                             QMessageBox.Yes)

    def getConfigContent(self):
        global configContent
        try:
            csvFile = pd.read_csv('%s/config_organic.csv' % configFileUrl, names=['A', 'B', 'C'])
        except:
            reply = QMessageBox.question(self, '信息', 'config文件配置缺少一些参数，是否重新创建并获取新的config文件',
                                         QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                MyMainWindow.createConfigContent(self)
                MyMainWindow.getConfigContent(self)
        else:
            configContent = {}
            content = list(csvFile['A'])
            rul = list(csvFile['B'])
            use = list(csvFile['C'])
            for i in range(len(content)):
                configContent['%s' % content[i]] = rul[i]
            a = len(configContent)
            if (int(configContent['config_num']) != len(configContent)) or (len(configContent) != 11):
                reply = QMessageBox.question(self, '信息', 'config文件配置缺少一些参数，是否重新创建并获取新的config文件',
                                             QMessageBox.Yes | QMessageBox.No,
                                             QMessageBox.Yes)
                if reply == QMessageBox.Yes:
                    MyMainWindow.createConfigContent(self)
                    MyMainWindow.getConfigContent(self)
            try:
                self.textBrowser.append("配置获取成功")
            except AttributeError:
                QMessageBox.information(self, "提示信息", "已获取配置文件内容", QMessageBox.Yes)
            else:
                pass

    def createConfigContent(self):
        months = "JanFebMarAprMayJunJulAugSepOctNovDec"
        n = time.strftime('%m')
        pos = (int(n) - 1) * 3
        monthAbbrev = months[pos:pos + 3]

        configContent = [
            ['config_num', '11', 'config文件条目数量,不能更改数值'],  # getConfigContent()中需要更改配置文件数量
            ['选择ICP_Batch的输入路径和输出路径', '默认，可更改为自己需要的', '以下ICP组Batch相关'],
            ['Organic_Batch_Import_URL', 'Z:\\Organic Batch', 'Batch路径'],
            ['TLims_Repetition_Check', 0, 'TLims是否根据batch重复样品编号,1选中，0未选中'],
            ['TLims_Repetition_Text', "A;B;C", 'TLims根据内容重复样品编号'],
            ['TLims_QC_Check', 0, 'TLims是否添加QC,1选中，0未选中'],
            ['TLims_QC_Msg', "CQC", 'TLims是否添加QC内容'],
            ['TLims_Batch_Import_URL', "Z:\\Inorganic_batch\\Tlims Batch", 'TLims-Batch导入路径'],
            ['TLims_Batch_Export_URL', '%s/config' % desktopUrl, 'TLims-Batch导出路径'],
            ['TLims_Quality_Control_Check', 0, '质控样品,1选中，0未选中'],
            ['TLims_Quality_Control_Sample', "BLK;BLK-S;S-S", '质控样品，自行填写，并用;间隔'],
        ]
        config = np.array(configContent)
        df = pd.DataFrame(config)
        df.to_csv('%s/config_organic.csv' % configFileUrl, index=0, header=0, encoding='utf_8_sig')
        self.textBrowser.append("配置文件创建成功")
        QMessageBox.information(self, "提示信息",
                                "默认配置文件已经创建好，\n如需修改请在用户桌面查找config文件夹中config_organic.csv，\n将相应的文件内容替换成用户需求即可，修改后记得重新导入配置文件。",
                                QMessageBox.Yes)

    def exportConfig(self):
        # 重新导出默认配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要创建默认配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.createConfigContent(self)
        else:
            QMessageBox.information(self, "提示信息", "没有创建默认配置文件，保留原有的配置文件", QMessageBox.Yes)

    def importConfig(self):
        # 重新导入配置文件
        reply = QMessageBox.question(self, '信息', '确认是否要导入配置文件', QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
        if reply == QMessageBox.Yes:
            MyMainWindow.getConfigContent(self)
            MyMainWindow.getDefaultInformation(self)
        else:
            QMessageBox.information(self, "提示信息", "没有重新导入配置文件，将按照原有的配置文件操作", QMessageBox.Yes)

    def showAuthorMessage(self):
        # 关于作者
        QMessageBox.about(self, "关于",
                          "人生苦短，码上行乐。\n\n\n        ----Frank Chen")

    def getDefaultInformation(self):
        # 默认登录TLims界面信息
        try:
            # data处理
            self.checkBox.setChecked(int(configContent['TLims_Repetition_Check']))
            self.checkBox_2.setChecked(int(configContent['TLims_QC_Check']))
            self.checkBox_3.setChecked(int(configContent['TLims_Quality_Control_Check']))
            self.lineEdit_2.setText(configContent['TLims_Repetition_Text'])
            self.lineEdit_7.setText(configContent['TLims_QC_Msg'])
            self.lineEdit_3.setText(configContent['TLims_Quality_Control_Sample'])
        except Exception as msg:
            self.textBrowser.append("错误信息：%s" % msg)
            self.textBrowser.append('----------------------------------')
            app.processEvents()
            reply = QMessageBox.question(self, '信息', '错误信息：%s。\n是否要重新创建配置文件' % msg,
                                         QMessageBox.Yes | QMessageBox.No,
                                         QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                MyMainWindow.createConfigContent(self)
                self.textBrowser.append("创建并导入配置成功")
                self.textBrowser.append('----------------------------------')
                app.processEvents()


    def getBatch(self):
        self.textBrowser.clear()
        self.lineEdit.clear()
        self.lineEdit.setText('Sample ID')
        global labNumber
        global selectBatchFile
        labNumber = []
        selectBatchFile = QFileDialog.getOpenFileNames(self, '选择Batch文件','%s' % configContent['Organic_Batch_Import_URL'],'Excel files(*.xls*;*.docx;*.csv)')
        if selectBatchFile[0] != []:
            self.textBrowser.append('正在获取Sample ID')
            self.textBrowser.append("Sample ID抓取完成后，\n才可以开始下一步骤！！！")
            app.processEvents()
            excel = win32com.gencache.EnsureDispatch('Excel.Application')
            excel.Visible = 0
            excel.Application.DisplayAlerts = True
            n = 1
            for file in selectBatchFile[0]:
                fileName = os.path.split(file)[1]
                fileType = os.path.split(file)[1].split('.')[-1]
                self.textBrowser.append('%s:%s' % (n, fileName))
                app.processEvents()
                if 'xls' in fileType:
                    wb = excel.Workbooks.Open(os.path.join(os.getcwd(), r'%s' % file))
                    ws = wb.Worksheets('sheet1')
                    x = 2
                    while ws.Cells(x, 1).Value is not None:
                        labNumber.append(ws.Cells(x, 1).Value)
                        x += 1

                elif 'doc' in fileType:
                    doc = Document(r"%s" % file.replace('/', '\\'))
                    for table in doc.tables:
                        for row in table.rows:
                            i = 1
                            for cell in row.cells:
                                print(i,cell.text)
                                if i == 2:
                                    if '/' in cell.text:
                                        labNumber.append(cell.text)
                                        i += 1
                                    else:
                                        i += 1
                                else:
                                    i += 1
                elif 'csv' in fileType:
                    batchFile = file.replace('/', '\\')
                    csvFile = pd.read_csv(batchFile)
                    # 去除质控
                    # sampleNum = list(csvFile[' Sample No.'])
                    # leaveNum = []
                    # for each in sampleNum:
                    #     if '/' in each:
                    #         leaveNum.append(each)
                    # csvFile = csvFile.loc[(csvFile[' Sample No.'].isin(leaveNum))]
                    labNumber += list(csvFile[' Sample No.'])
                    app.processEvents()
                n += 1
            if 'xls' in fileType:
                excel.Quit()
            self.textBrowser.append('Sample ID抓取完成\n开始填写前，将输入法换成英文输入法！！！！')
        else:
            self.textBrowser.setText("请重新选择Batch文件")

    def getTlimsBatchsUrl(self):
        # 获取Tlims-Batch文件
        batchFiles = QFileDialog.getOpenFileNames(self, '选择ICP-Batch文件',
                                                  '%s' % configContent['TLims_Batch_Import_URL'],
                                                  'CSV files(*.csv)')
        self.filesUrls = batchFiles[0]
        if self.filesUrls != []:
            self.textBrowser.append('选中文件:')
            self.textBrowser.append('\n'.join(self.filesUrls))
            self.textBrowser.append('----------------------------------')
            self.timsData = MyMainWindow.exportTlimsBatch(self)

        else:
            self.textBrowser.append('无选中文件')
            self.textBrowser.append('----------------------------------')
        app.processEvents()
        return self.filesUrls

    def exportTlimsBatch(self):
        try:
            global labNumber
            if self.filesUrls != []:
                name = 'organic_tlims_batch_data'
                csv_file_oj = Tlims_Data()
                if self.checkBox_3.isChecked():
                    quality_control_sample = self.lineEdit_3.text().split(';')
                else:
                    quality_control_sample = []
                star_num = 1
                batchs_data = csv_file_oj.get_tlims_batchs_data(self.filesUrls, int(star_num), quality_control_sample)
                batch_data = batchs_data[['Sample Id', 'ID']]
                qc_num = self.spinBox_7.text()
                qc_msg = self.lineEdit_7.text()
                duplicate_check = self.checkBox.isChecked()
                qc_check = self.checkBox_2.isChecked()
                # 定义要插入的行
                col_name_len = 1
                if duplicate_check:
                    new_row = pd.DataFrame(
                        [{'Sample Id': qc_msg, 'ID': 'CC', 'Variable': 'C', 'Value': 1, 'F Sample Id': qc_msg}])
                    num = 1
                    col_name = []
                    duplicate_com_list = self.lineEdit_2.text().split(';')
                    col_name_len = int(len(duplicate_com_list))
                    for col in duplicate_com_list:
                        batch_data["A%s" % num] = None
                        batch_data["A%s" % num] = col
                        col_name.append("A%s" % num)
                        num += 1
                    duplicate_data = csv_file_oj.duplicate_data(batch_data, col_name)
                else:
                    new_row = pd.DataFrame(
                        [{'Sample Id': qc_msg, 'ID': 'CC'}])
                    duplicate_data = batch_data
                # 是否添加QC
                if qc_check:
                    export_data = csv_file_oj.add_qc_data(duplicate_data, new_row, col_name_len, qc_num)
                else:
                    export_data = duplicate_data
                try:
                    export_data['Sample Id'] = export_data['F Sample Id']
                    export_data.drop(['Variable', 'Value', 'F Sample Id'], axis=1, inplace=True)
                except:
                    pass
                file_name = '%s/%s-%s.csv' % (configContent['TLims_Batch_Export_URL'], name.capitalize(), today)
                export_data.to_csv(file_name, index=False, header=None, mode='a')
                self.textBrowser.append('保存文件：%s' % file_name)
                self.textBrowser.append('----------------------------------')
                labNumber = export_data['Sample Id']
                return export_data
            else:
                self.textBrowser.append('无选中文件')
                self.textBrowser.append('----------------------------------')
            app.processEvents()
        except Exception as errorMsg:
            self.textBrowser.append('错误信息：%s' % errorMsg)
            self.textBrowser.append('----------------------------------')
            app.processEvents()

    def clearContent(self):
        # 清除填写内容
        self.lineEdit.clear()
        self.lineEdit.setText("Sample ID")
        self.textBrowser.append("可以开始Sample ID填写")

    def stopMessage(self):
        # 自动填写-停止
        stopMessage = 'Stop'
        self.lineEdit.setText(stopMessage)
        self.textBrowser.append("已停止，请清零后重新开始!!!")

    def autoWrite(self):
        # 自动填写 - 开始自动填写
        if self.lineEdit.text() == '' or self.lineEdit.text() == 'stop' or self.lineEdit.text() == 'Stop':
            QMessageBox.information(self, "提示信息", "自动填写中无内容或内容为‘stop’，请清零并填写内容",
                                    QMessageBox.Yes)
        else:
            time.sleep(3)
            starNum = int(self.spinBox.text())
            endNum = int(self.spinBox_2.text())
            velocityNum = float(self.doubleSpinBox.text())
            if self.lineEdit.text() == 'Sample ID' or self.lineEdit.text() == 'sample ID':
                self.textBrowser.append("正在填写样品单号")
                app.processEvents()
                if endNum == 0 or endNum > len(labNumber):
                    m = len(labNumber) - starNum + 1
                else:
                    m = endNum - starNum + 1
                n = starNum - 1
                for i in range(m):
                    app.processEvents()
                    time.sleep(0.1)
                    if self.lineEdit.text() == 'Sample ID' or self.lineEdit.text() == 'sample ID':
                        pyautogui.typewrite('%s' % labNumber[n], velocityNum)
                        pyautogui.typewrite(['Down'])
                        app.processEvents()
                        n += 1
                    elif self.lineEdit.text() == 'stop' or self.lineEdit.text() == 'Stop':
                        break
                self.textBrowser.append('Sample ID填写完成')

    def viewData(self):
        try:
            df = self.timsData
        except:
            myWin.getTlimsBatchsUrl()
        else:
            myTable.createTable(df)
            myTable.showMaximized()



class MyTableWindow(QMainWindow, Ui_TableWindow):
    def __init__(self, parent=None):
        super(MyTableWindow, self).__init__(parent)
        self.setupUi(self)
        # self.pushButton.clicked.connect(self.saveTable)
        self.pushButton_2.clicked.connect(self.createTable)

    # self.QtWidgets.QDialogButtonBox.Save

    def createTable(self, df):
        self.df = df
        self.df = self.df.astype(str)
        self.df_rows = self.df.shape[0]
        self.df_cols = self.df.shape[1]
        self.tableWidget.setRowCount(self.df_rows)
        self.tableWidget.setColumnCount(self.df_cols)

        ##设置水平表头
        self.tableWidget.setHorizontalHeaderLabels(self.df.keys().astype(str))

        # self.tabletWidget.
        for i in range(self.df_rows):
            for j in range(self.df_cols):
                self.tableWidget.setItem(i, j, QTableWidgetItem(self.df.iloc[i, j]))
        # 第1列不允许编辑
        self.tableWidget.setItemDelegateForColumn(0, EmptyDelegate(self))
        # 行颜色
        self.tableWidget.setAlternatingRowColors(True)
        # 显示所有内容
        self.tableWidget.resizeColumnsToContents()
        # 平均分配
        self.tableWidget.horizontalHeader().setSectionResizeMode(True)

    @pyqtSlot()
    def print_my_df(self):
        print(self.df)

    # @pyqtSlot()
    # def saveTable(self):
    #     col = self.tableWidget.columnCount()
    #     row = self.tableWidget.rowCount()
    #     # for currentQTableWidgetItem in self.tableWidget.selectedItems():
    #     # 	print((currentQTableWidgetItem.row(), currentQTableWidgetItem.column()))
    #     data = []
    #     for i in range(col):
    #         data.append(i)
    #         data[i] = []
    #         for j in range(row):
    #             itemData = self.tableWidget.item(j, i).text()
    #             data[i].append(itemData)
    #     configFile = pd.DataFrame({'a': data[0], 'b': data[1], 'c': data[2]})
    #     configFile.to_csv('%s/config_inorganic.csv' % configFileUrl, encoding="utf_8_sig", index=0, header=0)
    #     reply = QMessageBox.question(self, '信息', '配置文件已修改成功，是否重新获取新的config文件内容',
    #                                  QMessageBox.Yes | QMessageBox.No,
    #                                  QMessageBox.Yes)
    #     if reply == QMessageBox.Yes:
    #         MyMainWindow.getConfigContent(self)


# table不可编辑
class EmptyDelegate(QItemDelegate):
    def __init__(self, parent):
        super(EmptyDelegate, self).__init__(parent)

    def createEditor(self, QWidget, QStyleOptionViewItem, QModelIndex):
        return None

if __name__ == "__main__":
    import sys
    import os
    import time
    import pyautogui
    import chicon #引用图标
    import win32com.client as win32com
    import pandas as pd
    from docx import Document
    from win32com.client import Dispatch
    QtCore.QCoreApplication.setAttribute(QtCore.Qt.AA_EnableHighDpiScaling)
    app = QtWidgets.QApplication(sys.argv)  # 创建一个QApplication，也就是你要开发的软件app
    myWin = MyMainWindow()
    myTable = MyTableWindow()
    myWin.show()
    myWin.getConfig()
    sys.exit(app.exec_())  # 使用exit()或者点击关闭按钮退出QApplication
