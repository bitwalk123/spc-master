import os.path
from PySide2.QtCore import Slot
from PySide2.QtGui import QIcon
from PySide2.QtWidgets import (
    QFileDialog,
    QGridLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QPushButton,
    QSizePolicy,
    QStatusBar,
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
    config = None
    confFile = None

    def __init__(self, parent: QMainWindow):
        super().__init__(parent=parent)
        self.icons = Icons()

        # copy parent values
        self.parent = parent
        self.db = parent.db
        self.config = parent.config
        self.confFile = parent.confFile

        self.initUI()

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
        # tool_db.clicked.connect(self.dbMan)
        toolbar.addWidget(tool_close)

        base = QWidget()
        base.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setCentralWidget(base)
        grid = QGridLayout()
        base.setLayout(grid)
        row = 0

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        # DB Label
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
        but_name_db.clicked.connect(lambda: self.openFile(ent_name_db))

        grid.addWidget(lab_name_db, row, 0)
        grid.addWidget(ent_name_db, row, 1)
        grid.addWidget(but_name_db, row, 2)

        row += 1

        # Status Bar
        self.statusbar: QStatusBar = QStatusBar()
        self.setStatusBar(self.statusbar)

        self.show()

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
