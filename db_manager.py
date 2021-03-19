import os.path
import re
from PySide2.QtCore import Slot
from PySide2.QtGui import QIcon
from PySide2.QtWidgets import (
    QComboBox,
    QFileDialog,
    QGridLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStatusBar,
    QTabWidget,
    QToolBar,
    QToolButton,
    QWidget,
)
from database import SqlDB
from resource import Icons


# _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
class DBManWin(QMainWindow):
    parent = None
    db = None
    flag_db = False
    config = None
    confFile = None

    w_init: int = 600
    h_init: int = 200

    # Regular Expression
    pattern1: str = re.compile(r'([a-zA-Z0-9\s]+).*SPC.*')

    def __init__(self, parent: QMainWindow):
        super().__init__(parent=parent)
        self.icons = Icons()

        # copy parent values
        self.parent = parent
        self.db = parent.db
        self.config = parent.config
        self.confFile = parent.confFile

        self.initUI()
        self.setWindowIcon(QIcon(self.icons.DB))
        self.setWindowTitle('DB Manager')

    # -------------------------------------------------------------------------
    #  initUI - UI initialization
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def initUI(self):
        # ----------------
        #  Create toolbar
        toolbar = QToolBar()
        self.addToolBar(toolbar)

        # spacer
        spacer: QWidget = QWidget()
        spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        toolbar.addWidget(spacer)

        # Close Window
        tool_close: QToolButton = QToolButton()
        tool_close.setIcon(QIcon(self.icons.CLOSE))
        tool_close.setStatusTip('close this window')
        tool_close.clicked.connect(self.closeEvent)
        toolbar.addWidget(tool_close)

        area = QScrollArea()
        area.setWidgetResizable(True)
        self.setCentralWidget(area)

        base = QWidget()
        base.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        area.setWidget(base)

        grid = QGridLayout()
        base.setLayout(grid)

        row = 0

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        # SQLite file
        lab_name_db = QLabel('SQLite file')
        lab_name_db.setStyleSheet("QLabel {font-size:10pt; padding: 0 2px;}")
        lab_name_db.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        ent_name_db = QLineEdit()
        config_db = self.config['Database']
        dbname = config_db['DBNAME']
        ent_name_db.setText(dbname)
        ent_name_db.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        but_name_db = QPushButton()
        but_name_db.setIcon(QIcon(self.icons.FOLDER))
        but_name_db.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        but_name_db.setStatusTip('select SQLite file')
        but_name_db.clicked.connect(lambda: self.openFile(ent_name_db))

        grid.addWidget(lab_name_db, row, 0)
        grid.addWidget(ent_name_db, row, 1)
        grid.addWidget(but_name_db, row, 2)

        row += 1

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        # Excel macro file
        lab_name_excel = QLabel('Excel macro file')
        lab_name_excel.setStyleSheet("QLabel {font-size:10pt; padding: 0 2px;}")
        lab_name_excel.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        ent_name_excel = QLineEdit()
        if self.parent.sheets is not None:
            filename = self.parent.sheets.get_filename()
            self.flag_db = True
        else:
            filename = ''
            self.flag_db = False

        ent_name_excel.setText(filename)
        ent_name_excel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        grid.addWidget(lab_name_excel, row, 0)
        grid.addWidget(ent_name_excel, row, 1)

        row += 1

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        # Supplier List
        lab_name_supplier = QLabel('Supplier')
        lab_name_supplier.setStyleSheet("QLabel {font-size:10pt; padding: 0 2px;}")
        lab_name_supplier.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        combo_name_supplier = QComboBox()
        self.add_supplier_list_to_combo(combo_name_supplier)
        combo_name_supplier.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        combo_name_supplier.setStyleSheet("QComboBox:disabled {color:black; background-color:white;}");

        but_db_add = QPushButton()
        but_db_add.setIcon(QIcon(self.icons.DBADD))
        but_db_add.setEnabled(self.flag_db)
        but_db_add.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        but_db_add.setStatusTip('add Excel data to database')

        grid.addWidget(lab_name_supplier, row, 0)
        grid.addWidget(combo_name_supplier, row, 1)
        grid.addWidget(but_db_add, row, 2)

        row += 1

        # Status Bar
        self.statusbar: QStatusBar = QStatusBar()
        self.setStatusBar(self.statusbar)

        self.resize(self.w_init, self.h_init)
        self.show()

    # -------------------------------------------------------------------------
    #  add_supplier_list_to_combo
    #  add supplier list to specified combobox
    #
    #  argument
    #    combo: QComboBox  instance of QComboBox
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def add_supplier_list_to_combo(self, combo: QComboBox):
        # clear QComboBox
        combo.clear()
        combo.clearEditText()
        # DB Query and update QConboBox
        sql = "SELECT name_supplier_short FROM supplier;"
        out = self.db.get(sql)
        for supplier in out:
            combo.addItem(supplier[0])

        name = self.get_supplier_name()
        index = combo.findText(name)
        if index >= 0:
            combo.setCurrentIndex(index)
            combo.setEnabled(False)
        else:
            combo.setEnabled(True)

    # -------------------------------------------------------------------------
    #  get_supplier_name
    #  get supplier name from Excel filename
    #
    #  argument
    #    (none)
    #
    #  return
    #    Supplier name
    # -------------------------------------------------------------------------
    def get_supplier_name(self):
        if self.parent.sheets is None:
            return 'Unknown'
        else:
            name_excel = os.path.basename(self.parent.sheets.get_filename())

        print(name_excel)
        match: bool = self.pattern1.match(name_excel)
        if match:
            name = match.group(1).strip()

            print(name)
            # exception
            if name == 'FerroTech':
                return 'Ferrotec'

            return name

    # -------------------------------------------------------------------------
    #  openFile
    #  Open file dialog for selecting / opening SQLite file
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    @Slot()
    def openFile(self, ent: QLineEdit):
        # file selection dialog
        dialog: QFileDialog = QFileDialog()
        filters: str = 'SQLite file (*.sqlite *.sqlite3);; All (*.*)'
        dialog.setNameFilter(filters)

        if not dialog.exec_():
            return

        # read selected file
        dbname: str = dialog.selectedFiles()[0]
        if not os.path.exists(dbname):
            return
        else:
            # make SqlDB instance
            self.db = SqlDB(dbname)

            # set dbname in config file
            self.config.set('Database', 'DBNAME', dbname)
            with open(self.confFile, 'w') as file:
                self.config.write(file)

        ent.setText(dbname)

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
            'Close this Window',
            'Are you sure you want to close?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if sender is not None:
            # Exit button is clicked
            if reply == QMessageBox.Yes:
                self.destroy()
        else:
            # x on the window is clicked
            if reply == QMessageBox.Yes:
                event.accept()
            else:
                event.ignore()
