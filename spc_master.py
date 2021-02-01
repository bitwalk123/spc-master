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
    QTableView,
    QTabWidget,
    QToolBar,
    QToolButton,
    QWidget,
)
from office import ExcelSPC
from worksheet import SheetMaster
from spc_chart import ChartWin


# _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
class SPCMaster(QMainWindow):
    # Application information
    app_name: str = 'SPC Master'
    app_ver: str = '0.5 (alpha)'

    # initial windows position and size
    x_init: int = 100
    y_init: int = 100
    w_init: int = 800
    h_init: int = 600

    # initialize instances for main GUI
    tabwidget: QTabWidget = None
    statusbar: QStatusBar = None
    sheets: ExcelSPC = None
    chart = None

    # icons
    icon_book: str = 'images/book.png'
    #icon_excel: str = 'images/x-office-spreadsheet.png'
    icon_excel: str = 'images/File-Spreadsheet-icon.png'
    icon_logo: str = 'images/logo.ico'
    icon_warn: str = 'image/warning.png'

    # filter for file extentions to read
    filters: str = 'Excel file (*.xlsx *.xlsm);; All (*.*)'

    def __init__(self):
        super().__init__()

        self.initUI()
        self.setWindowIcon(QIcon(self.icon_logo))
        self.setAppTitle()
        self.setGeometry(self.x_init, self.y_init, self.w_init, self.h_init)

    # -------------------------------------------------------------------------
    #  initUI
    #  UI initialization
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
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
        self.tabwidget: QTabWidget = QTabWidget()
        self.tabwidget.setTabPosition(QTabWidget.South)
        self.setCentralWidget(self.tabwidget)

        # Status Bar
        self.statusbar: QStatusBar = QStatusBar()
        self.setStatusBar(self.statusbar)

        self.show()

    # -------------------------------------------------------------------------
    #  setAppTitle
    #  Set application title
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
    #  Create tab instances
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def createTabs(self):
        # delete contents of tab if exist
        for idx in range(self.tabwidget.count() - 1, -1, -1):
            tabContent: QWidget = self.tabwidget.widget(idx)
            self.tabwidget.removeTab(idx)
            tabContent.destroy()

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  'Master' tab
        self.createTabMaster()

    # -------------------------------------------------------------------------
    #  createTabMaster
    #  Create 'Master' tab
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def createTabMaster(self):
        # create Master sheet
        self.sheet_master: SheetMaster = SheetMaster(self.sheets)
        self.num_param: int = self.sheet_master.get_num_param()
        icon_master: QIcon = QIcon(self.icon_book)

        # double click event at row header
        header_row: QHeaderView = self.sheet_master.verticalHeader()
        header_row.sectionDoubleClicked.connect(self.handleRowHeaderDblClick)

        # add Master sheet to Tab widget
        self.tabwidget.addTab(self.sheet_master, icon_master, 'Master')

    # -------------------------------------------------------------------------
    #  handleRowHeaderDblClick
    #  Handle event when double clicked row header
    #
    #  argument
    #    row : row number where is clicked
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    @Slot()
    def handleRowHeaderDblClick(self, row: int):
        if self.chart is not None:
            self.chart.close()
            self.chart.deleteLater()

        self.chart = ChartWin(self, self.sheets, self.num_param, row)

    # -------------------------------------------------------------------------
    #  setRowSelect
    #  Set row selection
    #
    #  argument
    #    row : row to be selected
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def setMasterRowSelect(self, row: int):
        self.sheet_master.selectRow(row)

    # -------------------------------------------------------------------------
    #  openFile
    #  Open file dialog
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
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
    #  Dialog for close confirmation
    #
    #  argument
    #    event
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def closeEvent(self, event):
        reply: QMessageBox.StandardButton = QMessageBox.warning(
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
