#!/usr/bin/env python
# coding: utf-8

import os.path
import sys
from PySide2.QtCore import Slot
from PySide2.QtGui import QIcon
from PySide2.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QStatusBar,
    QTableView,
    QTabWidget,
    QToolBar,
    QToolButton,
    QVBoxLayout,
    QWidget,
)
from office import ExcelSPC
from worksheet import SheetMaster


class SPCMaster(QMainWindow):
    # Application information
    app_name = 'SPC Master'
    app_ver = '0.4 (alpha)'

    # initialize instances
    notebook = None
    statusbar = None
    grid_master = None
    sheets = None
    chart = None
    num_param = 0

    # icons
    icon_logo = 'images/logo.ico'
    icon_excel = 'images/excel.png'
    icon_warn = 'image/warning.png'

    # filter for file extentions
    filters = 'Excel file (*.xlsx *.xlsm);; All (*.*)'

    def __init__(self):
        super().__init__()
        # super(SPCMaster, self).__init__()

        self.initUI()
        self.setWindowIcon(QIcon(self.icon_logo))
        self.setAppTitle()
        self.setGeometry(100, 100, 800, 600)
        self.show()

    # -------------------------------------------------------------------------
    #  initUI - UI initialization
    # -------------------------------------------------------------------------
    def initUI(self):
        # Create toolbar
        toolbar: QToolBar = QToolBar()
        self.addToolBar(toolbar)

        # Add buttons to toolbar
        tool_excel: QToolButton = QToolButton()
        tool_excel.setIcon(QIcon(self.icon_excel))
        tool_excel.setStatusTip('Open Excel macro file for SPC')
        tool_excel.clicked.connect(self.openFile)
        toolbar.addWidget(tool_excel)

        # Tab widget
        self.notebook: QTabWidget = QTabWidget()
        self.setCentralWidget(self.notebook)

        # Status Bar
        self.statusbar: QStatusBar = QStatusBar()
        self.setStatusBar(self.statusbar)

    # -------------------------------------------------------------------------
    #  setAppTitle
    #  set application title
    #
    #  argument
    #    filename : filename already read, otherwise None as default
    #
    #  return
    #    title string
    # -------------------------------------------------------------------------
    def setAppTitle(self, filename=None):
        app_title = self.app_name + ' ' + self.app_ver
        if filename is not None:
            app_title = app_title + ' - ' + os.path.basename(filename)

        self.setWindowTitle(app_title)

    # -------------------------------------------------------------------------
    #  readExcel
    #  Aggregation from Excel for SPC
    #
    #  argument
    #    filename : Excel file to read
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    @Slot()
    def readExcel(self, filename):
        if self.sheets is not None:
            del self.sheets
        self.sheets: ExcelSPC = ExcelSPC(filename)
        if self.sheets.valid is not True:
            QMessageBox.critical(self, 'Error', 'Not appropriate format!')
            self.sheets = None
            return

        self.setAppTitle(filename)
        self.createTabs()

    # -------------------------------------------------------------------------
    #  createTabs
    #  create tab instances
    #
    #  argument
    #    sheet :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def createTabs(self):
        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  'Master' tab
        self.createTabMaster()

    # -------------------------------------------------------------------------
    #  createTabMaster
    #  creating 'Master' tab
    #
    #  argument
    #    sheet : object of Excel sheet
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def createTabMaster(self):
        tab_master: SheetMaster = SheetMaster(self.sheets)
        self.notebook.addTab(tab_master, 'Master')

    # -------------------------------------------------------------------------
    #  showDialog
    # -------------------------------------------------------------------------
    @Slot()
    def openFile(self):
        dialog: QFileDialog = QFileDialog()
        dialog.setNameFilter(self.filters)
        if dialog.exec_():
            filename = dialog.selectedFiles()[0]
            self.readExcel(filename)

    # -------------------------------------------------------------------------
    #  closeEvent
    # -------------------------------------------------------------------------
    def closeEvent(self, event):
        reply: QMessageBox = QMessageBox.warning(
            self,
            'Quit App',
            'Are you sure you want to quit?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()


# =============================================================================
#  MAIN
# =============================================================================
def main():
    app: QApplication = QApplication(sys.argv)
    ex: SPCMaster = SPCMaster()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
# ---
#  END OF PROGRAM
