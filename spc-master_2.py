#!/usr/bin/env python
# coding: utf-8

import os.path
import sys
from PySide2.QtCore import Slot
from PySide2.QtGui import QIcon
from PySide2.QtWidgets import (
    QApplication,
    QFileDialog,
    QHeaderView,
    QMainWindow,
    QMessageBox,
    QStatusBar,
    QTabWidget,
    QToolBar,
    QToolButton,
)
from office import ExcelSPC
from worksheet import SheetMaster


class SPCMaster(QMainWindow):
    # Application information
    app_name: str = 'SPC Master'
    app_ver: str = '0.4 (alpha)'

    # initialize instances
    notebook: QTabWidget = None
    statusbar: QStatusBar = None
    sheets: ExcelSPC = None
    chart = None

    # icons
    icon_logo: str = 'images/logo.ico'
    icon_excel: str = 'images/excel.png'
    icon_warn: str = 'image/warning.png'

    # filter for file extentions
    filters: str = 'Excel file (*.xlsx *.xlsm);; All (*.*)'

    def __init__(self):
        super().__init__()

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
    def setAppTitle(self, filename: str = None):
        app_title: str = self.app_name + ' ' + self.app_ver
        if filename is not None:
            app_title = app_title + ' - ' + os.path.basename(filename)

        self.setWindowTitle(app_title)

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

        row_header: QHeaderView = tab_master.verticalHeader()
        row_header.sectionDoubleClicked.connect(self.handleRowHeaderDblClicked)

        self.notebook.addTab(tab_master, 'Master')

    @Slot()
    def handleRowHeaderDblClicked(self, row: int):
        print('Row %d is selected' % row)

    # -------------------------------------------------------------------------
    #  openFile
    # -------------------------------------------------------------------------
    @Slot()
    def openFile(self):
        # file selection dialog
        dialog: QFileDialog = QFileDialog()
        dialog.setNameFilter(self.filters)
        if not dialog.exec_():
            return

        # read selected file
        filename: str = dialog.selectedFiles()[0]
        if self.sheets is not None:
            del self.sheets
        self.sheets: ExcelSPC = ExcelSPC(filename)

        # check if sheets have valid format or not
        if self.sheets.valid is not True:
            QMessageBox.critical(self, 'Error', 'Not appropriate format!')
            self.sheets = None
            return

        # update application title
        self.setAppTitle(filename)

        # create new tab
        self.createTabs()

    # -------------------------------------------------------------------------
    #  closeEvent
    # -------------------------------------------------------------------------
    def closeEvent(self, event):
        reply: QMessageBox.StandardButton = QMessageBox.warning(
            self,
            'Quit App',
            'Are you sure you want to quit?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        print(type(reply))
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
