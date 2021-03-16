#!/usr/bin/env python
# coding: utf-8

import configparser
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
    QSizePolicy,
    QTabWidget,
    QToolBar,
    QToolButton,
    QWidget,
)
from database import SqlDB
from db_manager import DBManWin
from office import ExcelSPC
from resource import Icons
from spc_chart import ChartWin
from worksheet import SheetMaster


# _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
class SPCMaster(QMainWindow):
    # Application information
    app_name: str = 'SPC Master'
    app_ver: str = '0.6 (alpha)'

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
    db = None

    # filter for file extentions to read
    filters: str = 'Excel file (*.xlsx *.xlsm);; All (*.*)'

    # configuraion file
    confFile: str = 'spc_master.ini'
    config: configparser.ConfigParser = None

    def __init__(self):
        super().__init__()
        self.icons = Icons()

        # CONFIGURATION FILE READ
        self.config = configparser.ConfigParser()
        self.config.read(self.confFile, 'UTF-8')
        self.initDB()

        self.initUI()
        self.setWindowIcon(QIcon(self.icons.LOGO))
        self.setAppTitle()
        self.setGeometry(self.x_init, self.y_init, self.w_init, self.h_init)

    # -------------------------------------------------------------------------
    #  initDB
    # -------------------------------------------------------------------------
    def initDB(self):
        # ---------------------------------------------------------------------
        #  DATABASE CONNECTION
        # ---------------------------------------------------------------------
        # Config for Database
        config_db = self.config['Database']
        dbname = config_db['DBNAME']

        if len(dbname) == 0:
            # empty value
            print('Empty!')
            pass
        elif os.path.exists(dbname):
            # make SqlDB instance
            self.db = SqlDB(dbname)
        else:
            # delete dbname in config file
            self.config.set('Database', 'DBNAME', '')
            with open(self.confFile, 'w') as file:
                self.config.write(file)

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
        # --------------
        # Create toolbar
        toolbar: QToolBar = QToolBar()
        self.addToolBar(toolbar)

        # Add Excel read buttons to toolbar
        tool_excel: QToolButton = QToolButton()
        tool_excel.setIcon(QIcon(self.icons.EXCEL))
        tool_excel.setStatusTip('Open Excel macro file for SPC')
        tool_excel.clicked.connect(self.openFile)
        toolbar.addWidget(tool_excel)

        # spacer
        spacer: QWidget = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        toolbar.addWidget(spacer)

        # Add Excel read buttons to toolbar
        tool_db: QToolButton = QToolButton()
        tool_db.setIcon(QIcon(self.icons.DB))
        tool_db.setStatusTip('DB setting')
        tool_db.clicked.connect(self.dbMan)
        toolbar.addWidget(tool_db)

        # Add Excel read buttons to toolbar
        tool_exit: QToolButton = QToolButton()
        tool_exit.setIcon(QIcon(self.icons.EXIT))
        tool_exit.setStatusTip('Exit application')
        tool_exit.clicked.connect(self.closeEvent)
        toolbar.addWidget(tool_exit)

        # --------------
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
        icon_master: QIcon = QIcon(self.icons.DB)

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
    #  dbMan
    #  database manager
    #
    #  argument
    #    event
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def dbMan(self, event):
        self.db_man =DBManWin(self, self.db)

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
        sender = self.sender()

        reply: QMessageBox.StandardButton = QMessageBox.warning(
            self,
            'Quit App',
            'Are you sure you want to quit?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if sender is not None:
            # Exit button is clicked
            if reply == QMessageBox.Yes:
                QApplication.quit()
            return
        else:
            # x on thw window is clicked
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
