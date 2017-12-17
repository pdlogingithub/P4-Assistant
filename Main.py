
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from P4 import P4

import os
import math
import subprocess
import xlwt
import xlrd

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("P4 assistant")
        self.resize(600, 800)

        self.p4 = P4()
        self.p4.exception_level = 0

        root = QWidget()
        self.setCentralWidget(root)
        rootLayouot = QGridLayout()
        root.setLayout(rootLayouot)

        connectionGBox = QGroupBox("Connection")
        connectionGBoxLayout = QGridLayout()
        connectionGBox.setLayout(connectionGBoxLayout)
        rootLayouot.addWidget(connectionGBox)

        serverLable = QLabel("Server:")
        connectionGBoxLayout.addWidget(serverLable, 0, 0)
        serverLine = QLineEdit()
        connectionGBoxLayout.addWidget(serverLine, 0, 1, Qt.AlignLeft)
        serverConnectLable = QLabel("Not connected")
        connectionGBoxLayout.addWidget(serverConnectLable, 0, 2,1,4, Qt.AlignLeft)

        userLable = QLabel("User:")
        connectionGBoxLayout.addWidget(userLable, 1, 0)
        userLine = QLineEdit()
        connectionGBoxLayout.addWidget(userLine, 1, 1, Qt.AlignLeft)
        userLoginLable = QLabel("Not logged in")
        connectionGBoxLayout.addWidget(userLoginLable, 1, 2, 1, 4, Qt.AlignLeft)
        userButton = QPushButton("Browse")
        def userButtonCallBack():
            if self.p4.connected():
                users = self.p4.run("users")
                for user in users:
                    print(user["User"])
        userButton.clicked.connect(userButtonCallBack)
        # connectionGBoxLayout.addWidget(userButton, 1, 2)

        passwordLable = QLabel("Password:")
        connectionGBoxLayout.addWidget(passwordLable, 2, 0)
        passwordLine = QLineEdit()
        passwordLine.setEchoMode(QLineEdit.Password)
        connectionGBoxLayout.addWidget(passwordLine, 2, 1, Qt.AlignLeft)

        clientLable = QLabel("Work space:")
        connectionGBoxLayout.addWidget(clientLable, 3, 0)
        clientLine = QLineEdit()
        connectionGBoxLayout.addWidget(clientLine, 3, 1, Qt.AlignLeft)

        pathLable = QLabel("File path")
        connectionGBoxLayout.addWidget(pathLable, 4, 0)
        pathLine = QLineEdit()
        connectionGBoxLayout.addWidget(pathLine, 4,1, 1, 4)
        pathAddButton = QPushButton("Add path")
        def pathAddButtonCallback():
            openFileName = QFileDialog.getExistingDirectory()
            if openFileName:
                pathLine.setText(pathLine.text() + openFileName + ";")
        pathAddButton.clicked.connect(pathAddButtonCallback)
        connectionGBoxLayout.addWidget(pathAddButton, 4, 5)
        pathClearButton = QPushButton("Clear path")
        pathClearButton.clicked.connect(lambda : pathLine.clear())
        connectionGBoxLayout.addWidget(pathClearButton, 4,6)

        settings = open("connection.txt", 'r+') if os.path.exists("connection.txt") else open("connection.txt", 'w+')
        for line in settings.read().split("\n"):
            map = line.split("=")
            if len(map) == 2 and map[0] == "Server":
                serverLine.setText(map[1])
            if len(map) == 2 and map[0] == "user":
                userLine.setText(map[1])
            if len(map) == 2 and map[0] == "Work space":
                clientLine.setText(map[1])
            if len(map) == 2 and map[0] == "File path":
                pathLine.setText(map[1])
        settings.close()

        runHBox = QHBoxLayout()
        runHBox.setAlignment(Qt.AlignCenter)
        runWidget = QWidget()
        runWidget.setLayout(runHBox)
        connectionGBoxLayout.addWidget(runWidget, 5, 0, 1, 7)

        saveButton = QPushButton("Save connection")
        def saveButtonCallBack():
            settings = open("connection.txt", 'w+')
            settings.write("Server=" + serverLine.text() + "\n")
            settings.write("user=" + userLine.text() + "\n")
            settings.write("Work space=" + clientLine.text() + "\n")
            settings.write("File path=" + pathLine.text() + "\n")
            settings.close()
        saveButton.clicked.connect(saveButtonCallBack)
        runHBox.addWidget(saveButton)

        runButton = QPushButton("Run")
        def runButtonCallBack():
            self.runDialog = QMessageBox()
            ret = self.runDialog.question(self,"Confirm","Program will start connecting to p4 soon,"
            "make sure the information you input are correct,"
            "if the path specific contains too many files,it may takes a while.",self.runDialog.Ok | self.runDialog.Cancel)
            if not ret == self.runDialog.Ok:
                return

            if self.p4.connected():
                self.p4.disconnect()
                print("disconnecting p4")
            if not self.p4.connected():
                self.p4.port = serverLine.text()
                self.p4.user = userLine.text()
                self.p4.password = passwordLine.text()
                self.p4.client = clientLine.text()
                self.p4.connect()
                print("connecting p4")
            if self.p4.connected():
                print("Connected to", self.p4.port)
                serverConnectLable.setText("Connected to "+ str(self.p4.port))
            else:
                print("Not connected")
                serverConnectLable.setText("Not connected")
            if userLine.text():
                self.p4.run_login()
                print("loggin p4")
            loginStat = self.p4.run("login", "-s")
            if loginStat:
                userLoginLable.setText("User " + loginStat[0]["User"] + " logged in,ticket expires in " + str(int(int(loginStat[0]["TicketExpiration"])/3600)) + " hours")
                print("p4 logged in")
            else:
                userLoginLable.setText("Not logged in")
                print("p4 not logged in")

            if len(self.p4.errors) < 1:
                self.model.removeRows(0, self.model.rowCount())
                self.baseModel.removeRows(0, self.baseModel.rowCount())
            filePaths = pathLine.text()

            for filePath in filePaths.split(";"):
                if filePath:
                    if not filePath.endswith("/") and not filePath.endswith("\\"):
                        filePath += "/..."
                    else:
                        filePath += "..."

                    fstatMap = {}
                    if clientLine.text():
                        fstatFiles = self.p4.run("fstat",filePath)
                        for fstatFile in fstatFiles:
                            if "depotFile" in fstatFile.keys() and "clientFile" in fstatFile.keys():
                                fstatMap[fstatFile["depotFile"]] = [fstatFile["clientFile"],"synced" if "haveRev" in fstatFile.keys() and fstatFile["haveRev"] == fstatFile["headRev"] else "unsynced"]

                    depotFiles = self.p4.run_filelog("-L",filePath)
                    for depotFile in depotFiles:
                        fileSize = int(depotFile.revisions[0].fileSize)/1024 if depotFile.revisions[0].fileSize else 0
                        fileSize = math.ceil(fileSize)
                        fileSizeStr = str(fileSize)
                        fileSizeStr = "0" * (9 - len(fileSizeStr)) + fileSizeStr
                        fileSizeStr = fileSizeStr[0:3] + "," + fileSizeStr[3:6] + "," + fileSizeStr[6:9] + "KB"

                        self.baseModel.insertRow(0)
                        self.baseModel.setData(self.baseModel.index(0, 0), depotFile.depotFile)
                        self.baseModel.setData(self.baseModel.index(0, 1), fstatMap[depotFile.depotFile][1] if depotFile.depotFile in fstatMap.keys() else "no workspace info")   #"unsynced" if depotFile.depotFile in haveDepotFilesMap.keys() and haveDepotFilesMap[depotFile.depotFile] < depotFile.revisions[0].change else "synced")
                        self.baseModel.setData(self.baseModel.index(0, 2), depotFile.revisions[0].user)
                        self.baseModel.setData(self.baseModel.index(0, 3), str(depotFile.revisions[0].change))
                        self.baseModel.setData(self.baseModel.index(0, 4), depotFile.revisions[0].action)
                        self.baseModel.setData(self.baseModel.index(0, 5), str(depotFile.revisions[0].time))
                        self.baseModel.setData(self.baseModel.index(0, 6), fileSizeStr)
                        self.baseModel.setData(self.baseModel.index(0, 7), str(depotFile.revisions[0].desc))
                        self.baseModel.setData(self.baseModel.index(0, 8), fstatMap[depotFile.depotFile][0] if depotFile.depotFile in fstatMap.keys() else "no workspace info")
            if len(self.p4.errors) > 0:
                self.errorDialog = QErrorMessage()
                self.errorDialog.setWindowTitle("ERROR")
                self.errorDialog.showMessage(str(self.p4.errors))
                return
            else:
                filterButtonCallBack()
        runButton.clicked.connect(runButtonCallBack)
        runHBox.addWidget(runButton)

        filterGBox = QGroupBox("Filter")
        filterGBoxLayout = QGridLayout()
        filterGBox.setLayout(filterGBoxLayout)
        rootLayouot.addWidget(filterGBox)

        filterFileNameLable = QLabel("File name:")
        filterGBoxLayout.addWidget(filterFileNameLable, 0, 0)
        filterFileNameLine = QLineEdit()
        filterFileNameLine.returnPressed.connect(lambda : filterButtonCallBack())
        filterGBoxLayout.addWidget(filterFileNameLine, 0, 1)

        filterSubmitterLable = QLabel("Submitter name:")
        filterGBoxLayout.addWidget(filterSubmitterLable,1,0)
        filterSubmitterLine = QLineEdit()
        filterSubmitterLine.returnPressed.connect(lambda: filterButtonCallBack())
        filterGBoxLayout.addWidget(filterSubmitterLine,1,1)

        filterClLable = QLabel("Changelist")
        filterGBoxLayout.addWidget(filterClLable, 2, 0)
        filterClLine = QLineEdit()
        filterClLine.returnPressed.connect(lambda: filterButtonCallBack())
        filterGBoxLayout.addWidget(filterClLine,2, 1)

        filterDateLable = QLabel("Date")
        filterGBoxLayout.addWidget(filterDateLable, 3, 0)

        fromDateLable = QLabel("From date")
        fromDateEdit = QDateEdit()
        fromDateEdit.setDate(QDate(2012, 12, 1))
        fromDateEdit.setCalendarPopup(True)
        toDateLable = QLabel("To date")
        toDateEdit = QDateEdit()
        toDateEdit.setDate(QDate.currentDate())
        toDateEdit.setCalendarPopup(True)
        dateHBox = QHBoxLayout()
        dateHBox.setAlignment(Qt.AlignLeft)
        dateHBox.addWidget(fromDateLable)
        dateHBox.addWidget(fromDateEdit)
        dateHBox.addWidget(toDateLable)
        dateHBox.addWidget(toDateEdit)
        dateWidget = QWidget()
        dateWidget.setLayout(dateHBox)
        filterGBoxLayout.addWidget(dateWidget,3,1,1,2)

        filterDescLable = QLabel("Description:")
        filterGBoxLayout.addWidget(filterDescLable, 4, 0)
        filterDescLine = QLineEdit()
        filterDescLine.returnPressed.connect(lambda: filterButtonCallBack())
        filterGBoxLayout.addWidget(filterDescLine, 4, 1)

        filterCheckHBox = QHBoxLayout()
        filterCheckHBox.setAlignment(Qt.AlignLeft)
        filterCheckWidget = QWidget()
        filterCheckWidget.setLayout(filterCheckHBox)
        filterGBoxLayout.addWidget(filterCheckWidget, 5, 0, 1, 3)

        filterBaseNameLable = QLabel("Show only file name:")
        filterCheckHBox.addWidget(filterBaseNameLable)
        filterBaseNameCheck = QCheckBox()
        filterBaseNameCheck.setChecked(True)
        filterCheckHBox.addWidget(filterBaseNameCheck)

        filterDepotPathLable = QLabel("Show depot path")
        filterCheckHBox.addWidget(filterDepotPathLable)
        filterDepotPathCheck = QCheckBox()
        if not clientLine.text():
            filterDepotPathCheck.setChecked(True)
        filterCheckHBox.addWidget(filterDepotPathCheck)

        settings = open("filter.txt", 'r+') if os.path.exists("filter.txt") else open("filter.txt", 'w+')
        for line in settings.read().split("\n"):
            map = line.split("=")
            if not len(map) == 2:
                continue
            if map[0] == "File name":
                filterFileNameLine.setText(map[1])
            elif map[0] == "Submitter name":
                filterSubmitterLine.setText(map[1])
            elif map[0] == "Changelist":
                filterClLine.setText(map[1])
            elif map[0] == "From date":
                if len(map[1].split(";")) > 2:
                    fromDateEdit.setDate(QDate(int(map[1].split(";")[0]), int(map[1].split(";")[1]), int(map[1].split(";")[2])))
            elif map[0] == "To date":
                if len(map[1].split(";")) > 2:
                    toDateEdit.setDate(QDate(int(map[1].split(";")[0]), int(map[1].split(";")[1]), int(map[1].split(";")[2])))
            elif map[0] == "Description":
                filterDescLine.setText(map[1])
            elif map[0] == "Only show filename":
                if map[1] == "True":
                    filterBaseNameCheck.setChecked(True)
                else:
                    filterBaseNameCheck.setChecked(False)
            elif map[0] == "Show depot path":
                if map[1] == "True":
                    filterDepotPathCheck.setChecked(True)
                else:
                    filterDepotPathCheck.setChecked(False)
        settings.close()

        filterHBox = QHBoxLayout()
        filterHBox.setAlignment(Qt.AlignCenter)
        filterButtonWidget = QWidget()
        filterButtonWidget.setLayout(filterHBox)
        filterGBoxLayout.addWidget(filterButtonWidget, 6, 0, 1, 3)

        saveFilterButton = QPushButton("Save filter")
        def saveFilterButtonCallBack():
            settings = open("filter.txt", 'w+')
            settings.write("File name=" + filterFileNameLine.text() + "\n")
            settings.write("Submitter name=" + filterSubmitterLine.text() + "\n")
            settings.write("Changelist=" + filterClLine.text() + "\n")
            fromDateStr = fromDateEdit.sectionText(QDateTimeEdit.YearSection) + ";" + fromDateEdit.sectionText(QDateTimeEdit.MonthSection) + ";" + fromDateEdit.sectionText(QDateTimeEdit.DaySection)
            settings.write("From date=" + fromDateStr + "\n")
            if not toDateEdit.date() == QDate.currentDate():
                toDateStr = toDateEdit.sectionText(QDateTimeEdit.YearSection) + ";" + toDateEdit.sectionText(QDateTimeEdit.MonthSection) + ";" + toDateEdit.sectionText(QDateTimeEdit.DaySection)
                settings.write("To date=" + toDateStr + "\n")
            settings.write("Description=" + filterDescLine.text() + "\n")
            if filterBaseNameCheck.isChecked():
                settings.write("Only show filename=" + "True" + "\n")
            else:
                settings.write("Only show filename=" + "False" + "\n")
            if filterDepotPathCheck.isChecked():
                settings.write("Show depot path=" + "True" + "\n")
            else:
                settings.write("Show depot path=" + "False" + "\n")
            settings.close()
        saveFilterButton.clicked.connect(saveFilterButtonCallBack)
        filterHBox.addWidget(saveFilterButton)

        filterButton = QPushButton("Filter")
        def FilteredByKeyWords(str,keyWords):
            for keyWord in keyWords.split(";"):
                if not keyWord:
                    continue
                if keyWord.startswith("-"):
                    if keyWord[1:] in str:
                        return True
                elif keyWord not in str:
                    return True
            return False
        def filterButtonCallBack():
            self.model.removeRows(0, self.model.rowCount())
            for index in range(self.baseModel.rowCount()):
                basename = ""
                if filterDepotPathCheck.isChecked():
                    basename = self.baseModel.data(self.baseModel.index(index, 0))
                else:
                    basename = self.baseModel.data(self.baseModel.index(index, 8))
                if filterBaseNameCheck.isChecked():
                    basename = os.path.basename(basename)

                if filterFileNameLine.text() and FilteredByKeyWords(basename, filterFileNameLine.text()):
                    continue

                submitterName = self.baseModel.data(self.baseModel.index(index, 2))
                if filterSubmitterLine.text() and FilteredByKeyWords(submitterName, filterSubmitterLine.text()):
                    continue

                changeList = self.baseModel.data(self.baseModel.index(index, 3))
                if filterClLine.text() and FilteredByKeyWords(changeList, filterClLine.text()):
                    continue

                desc = self.baseModel.data(self.baseModel.index(index, 7))
                if filterDescLine.text() and FilteredByKeyWords(desc, filterDescLine.text()):
                    continue

                fromMonth = fromDateEdit.sectionText(QDateTimeEdit.MonthSection)
                fromMonth = "0" + fromMonth if len(fromMonth) < 2 else fromMonth
                fromDay = fromDateEdit.sectionText(QDateTimeEdit.DaySection)
                fromDay = "0" + fromDay if len(fromDay) < 2 else fromDay
                if self.baseModel.data(self.baseModel.index(index, 5)).split(" ")[0].replace("-","") < fromDateEdit.sectionText(QDateTimeEdit.YearSection) + fromMonth + fromDay:
                    continue
                toMonth = toDateEdit.sectionText(QDateTimeEdit.MonthSection)
                toMonth = "0" + toMonth if len(toMonth) < 2 else toMonth
                toDay = toDateEdit.sectionText(QDateTimeEdit.DaySection)
                toDay = "0" + toDay if len(toDay) < 2 else toDay
                if self.baseModel.data(self.baseModel.index(index, 5)).split(" ")[0].replace("-","") > toDateEdit.sectionText(QDateTimeEdit.YearSection) + toMonth + toDay:
                    continue

                self.model.insertRow(0)
                self.model.setData(self.model.index(0, 0), basename)
                self.model.setData(self.model.index(0, 1), self.baseModel.data(self.baseModel.index(index, 1)))
                self.model.setData(self.model.index(0, 2), submitterName)
                self.model.setData(self.model.index(0, 3), changeList)
                self.model.setData(self.model.index(0, 4), self.baseModel.data(self.baseModel.index(index, 4)))
                self.model.setData(self.model.index(0, 5), self.baseModel.data(self.baseModel.index(index, 5)))
                self.model.setData(self.model.index(0, 6), self.baseModel.data(self.baseModel.index(index, 6)))
                self.model.setData(self.model.index(0, 7), self.baseModel.data(self.baseModel.index(index, 7)))
            fileViewGBox.setTitle("File view" + " " + str(self.model.rowCount()) + "/" + str(self.baseModel.rowCount()))
        filterButton.clicked.connect(filterButtonCallBack)
        filterHBox.addWidget(filterButton)

        fileViewGBox = QGroupBox("File view")
        fileViewGBoxLayout = QGridLayout()
        fileViewGBox.setLayout(fileViewGBoxLayout)
        rootLayouot.addWidget(fileViewGBox)

        view = QTreeView()
        view.setAlternatingRowColors(True)
        view.setSortingEnabled(True)
        view.setSelectionMode(QAbstractItemView.ExtendedSelection)
        def viewRightClickMenuCallback(position):
            menu = QMenu()
            action = QAction()
            action.setText("Open in windows explorer")
            def actionCallback():
                if len(view.selectionModel().selectedIndexes())<1:
                    return
                index = view.selectionModel().selectedIndexes()[0]
                data = self.model.data(index)
                if os.path.exists(data):
                    subprocess.Popen("explorer /select," + data)
                    return
                if not data.startswith("//"):
                    self.rightClicktDialog = QMessageBox()
                    self.rightClicktDialog.question(self, "Warning","Uncheck \"Show only file name\" and then filter to enable this function.",self.rightClicktDialog.Ok)
                    return
                fstat = self.p4.run("fstat",data)
                if not (len(fstat) > 0 and "clientFile" in fstat[0].keys()):
                    self.rightClicktDialog = QMessageBox()
                    self.rightClicktDialog.question(self, "Warning","Can not get client path,\ndid you run connection with correct work space?",self.rightClicktDialog.Close)
                    return
                subprocess.Popen("explorer /select," + fstat[0]["clientFile"])
            action.triggered.connect(actionCallback)
            menu.addAction(action)

            actionRemove = QAction()
            actionRemove.setText("Remove selection")
            def actionRemoveCallback():
                rows = set(index.row() for index in view.selectionModel().selectedIndexes())
                for row in rows:
                    self.model.removeRow(row)
                fileViewGBox.setTitle("File view" + " " + str(self.model.rowCount()) + "/" + str(self.baseModel.rowCount()))
            actionRemove.triggered.connect(actionRemoveCallback)
            menu.addAction(actionRemove)

            menu.exec_(view.viewport().mapToGlobal(position))
        view.setContextMenuPolicy(Qt.CustomContextMenu)
        view.customContextMenuRequested.connect(viewRightClickMenuCallback)
        fileViewGBoxLayout.addWidget(view,0,0,1,2)

        self.baseModel = QStandardItemModel(0, 9)
        self.baseModel.setHeaderData(0, Qt.Horizontal, "Depot File")
        self.baseModel.setHeaderData(1, Qt.Horizontal, "Status")
        self.baseModel.setHeaderData(2, Qt.Horizontal, "Submitter")
        self.baseModel.setHeaderData(3, Qt.Horizontal, "Changelist")
        self.baseModel.setHeaderData(4, Qt.Horizontal, "Action")
        self.baseModel.setHeaderData(5, Qt.Horizontal, "Time")
        self.baseModel.setHeaderData(6, Qt.Horizontal, "Size")
        self.baseModel.setHeaderData(7, Qt.Horizontal, "Desc")
        self.baseModel.setHeaderData(8, Qt.Horizontal, "Client File")
        self.model = QStandardItemModel(0, 8)
        self.model.setHeaderData(0, Qt.Horizontal, "File")
        self.model.setHeaderData(1, Qt.Horizontal, "Status")
        self.model.setHeaderData(2, Qt.Horizontal, "Submitter")
        self.model.setHeaderData(3, Qt.Horizontal, "Changelist")
        self.model.setHeaderData(4, Qt.Horizontal, "Action")
        self.model.setHeaderData(5, Qt.Horizontal, "Time")
        self.model.setHeaderData(6, Qt.Horizontal, "Size")
        self.model.setHeaderData(7, Qt.Horizontal, "Desc")

        view.setModel(self.model)
        view.setColumnWidth(0,200)
        view.setColumnWidth(1, 50)
        view.setColumnWidth(2, 70)
        view.setColumnWidth(3, 60)
        view.setColumnWidth(4, 50)
        view.setColumnWidth(5, 110)
        view.setColumnWidth(6, 80)

        openButton = QPushButton("Open")
        def openButtonCallBack():
            openFileRet = QFileDialog.getOpenFileNames()
            openFileName = ""
            if len(openFileRet) > 0 and len(openFileRet[0]) > 0:
                openFileName = openFileRet[0][0]
            else:
                return
            if openFileName:
                workBook = xlrd.open_workbook(openFileName)
                sheet = workBook.sheet_by_index(0)
                if sheet.ncols <7:
                    return
                for row in range(sheet.nrows):
                    if row == 0:
                        continue
                    self.baseModel.insertRow(0)

                    if sheet.cell(row,0).value.startswith("//"):
                        self.baseModel.setData(self.baseModel.index(0, 0), sheet.cell(row, 0).value)
                        self.baseModel.setData(self.baseModel.index(0, 8), "not in file")
                    elif "/" in sheet.cell(row,0).value or "\\" in sheet.cell(row,0).value:
                        self.baseModel.setData(self.baseModel.index(0, 8), sheet.cell(row, 0).value)
                        self.baseModel.setData(self.baseModel.index(0, 0), "not in file")
                    else:
                        self.baseModel.setData(self.baseModel.index(0, 8), sheet.cell(row, 0).value)
                        self.baseModel.setData(self.baseModel.index(0, 0), sheet.cell(row, 0).value)
                    self.baseModel.setData(self.baseModel.index(0, 1), sheet.cell(row,1).value)
                    self.baseModel.setData(self.baseModel.index(0, 2), sheet.cell(row,2).value)
                    self.baseModel.setData(self.baseModel.index(0, 3), sheet.cell(row,3).value)
                    self.baseModel.setData(self.baseModel.index(0, 4), sheet.cell(row,4).value)
                    self.baseModel.setData(self.baseModel.index(0, 5), sheet.cell(row,5).value)
                    self.baseModel.setData(self.baseModel.index(0, 6), sheet.cell(row,6).value)
                    self.baseModel.setData(self.baseModel.index(0, 7), sheet.cell(row, 7).value)

                filterButtonCallBack()
        openButton.clicked.connect(openButtonCallBack)
        fileViewGBoxLayout.addWidget(openButton, 2, 0)

        exportButton = QPushButton("Export")
        def exportButtonCallBack():
            self.exportDialog = QMessageBox()
            ret = self.exportDialog.question(self, "Warning", "If you are overwritting existing file,make sure the file is not opened by other apllication", self.exportDialog.Ok | self.exportDialog.Cancel)
            if not ret == self.exportDialog.Ok:
                return

            saveFileName = QFileDialog.getSaveFileName()[0]
            if saveFileName:
                if not saveFileName.endswith(".xls"):
                    saveFileName += ".xls"
                wb = xlwt.Workbook()
                sheet = wb.add_sheet("sheet1")
                sheet.col(0).width = 12000
                sheet.col(5).width = 5000
                sheet.col(6).width = 3500
                sheet.col(7).width = 8000
                sheet.write(0, 0, "File name")
                sheet.write(0, 1, "Status")
                sheet.write(0, 2, "Submitter")
                sheet.write(0, 3, "Changelist")
                sheet.write(0, 4, "Action")
                sheet.write(0, 5, "Time")
                sheet.write(0, 6, "Size")
                sheet.write(0, 7, "Description")
                for index in range(self.model.rowCount()):
                    sheet.write(index+1, 0, str(self.model.data(self.model.index(index, 0))))
                    sheet.write(index+1, 1, str(self.model.data(self.model.index(index, 1))))
                    sheet.write(index+1, 2, str(self.model.data(self.model.index(index, 2))))
                    sheet.write(index+1, 3, str(self.model.data(self.model.index(index, 3))))
                    sheet.write(index+1, 4, str(self.model.data(self.model.index(index, 4))))
                    sheet.write(index+1, 5, str(self.model.data(self.model.index(index, 5))))
                    sheet.write(index+1, 6, str(self.model.data(self.model.index(index, 6))))
                    sheet.write(index+1, 7, str(self.model.data(self.model.index(index, 7))))
                wb.save(saveFileName)
        exportButton.clicked.connect(exportButtonCallBack)
        fileViewGBoxLayout.addWidget(exportButton, 2, 1)

    def closeEvent(self, event):
        print("window closing")
        print("disconnecting p4")
        if  self.p4.connected():
            self.p4.disconnect()


if __name__ == '__main__':
    import sys
    App = QApplication(sys.argv)
    MainWindow = MainWindow()
    MainWindow.show()
    ret = App.exec_()
    sys.exit(ret)