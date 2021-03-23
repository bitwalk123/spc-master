import math
import os.path
import pandas as pd
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
    QProgressBar,
    QPushButton,
    QScrollArea,
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
    flag_db = False
    config = None
    confFile = None

    w_init: int = 600
    h_init: int = 200

    # Regular Expression
    pattern1: str = re.compile(r'([a-zA-Z0-9\s]+).*SPC.*')
    pattern2: str = re.compile(r'([0-9]{4}-[0-9]{3}-[0-9]{2}).*')

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
        if len(dbname) > 0:
            self.flag_db = True
        else:
            self.flag_db = False

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
        but_db_add.clicked.connect(lambda: self.updateDB(combo_name_supplier))

        grid.addWidget(lab_name_supplier, row, 0)
        grid.addWidget(combo_name_supplier, row, 1)
        grid.addWidget(but_db_add, row, 2)

        row += 1

        # for status bar
        statusLabel = QLabel("Showing Progress")
        statusLabel.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)

        self.progressbar = QProgressBar()
        self.progressbar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.progressbar.setMinimum(0)
        self.progressbar.setMaximum(100)
        self.progressbar.setValue(10)

        # statusBar = QStatusBar()

        # Status Bar
        self.statusbar: QStatusBar = QStatusBar()
        self.statusbar.addWidget(statusLabel, 1)
        self.statusbar.addWidget(self.progressbar, 2)
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

        if self.parent.sheets is None:
            combo.setEnabled(False)
            return

        if self.db is None:
            self.flag_db = False
            combo.setEnabled(False)
            return

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
    #  updateDB
    #  updating dB
    #
    #  argument
    #    combo: QComboBox for suppliers
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def updateDB(self, combo: QComboBox):
        name_supplier = combo.currentText()

        # id_supplier
        id_supplier = self.select_id_supplier(name_supplier)
        if id_supplier is None:
            return

        print('Suuplier :', name_supplier, ', id_supplier =', id_supplier)

        dict_header = {
            'Key Parameter': 'num_key',
            'Parameter Name': 'name_param',
            'LSL': 'lsl',
            'Target': 'target',
            'USL': 'usl',
            'Chart Type': 'charttype',
            'Metrology': 'metrology',
            'Multiple': 'multiple',
            'Spec Type': 'spectype',
            'CL Frozen': 'frozen',
            'LCL': 'lcl',
            'Avg': 'mean',
            'UCL': 'ucl',
        }

        list_part = self.parent.sheets.get_unique_part_list()
        for num_part_excel in list_part:

            match: bool = self.pattern2.match(num_part_excel)
            if match:
                num_part = match.group(1)
            else:
                num_part = ''

            print(num_part_excel, num_part)

            # id_part
            id_part = self.select_id_part(id_supplier, num_part)
            if id_part is None:
                print('id_part NOT FOUND!')
                continue

            print('PART# :', num_part, ', id_part =', id_part)

            list_param = self.parent.sheets.get_param_list(num_part_excel)
            for name_param in list_param:
                # print(num_part_excel, name_param)
                metrics = self.parent.sheets.get_metrics(num_part_excel, name_param)
                param = self.get_param(metrics)

                # id_param
                id_param = self.select_id_param(id_part, id_supplier, name_param)
                if id_param is None:
                    self.insert_param(id_part, id_supplier, name_param, num_part_excel, param)
                else:
                    print('### SPC Data')
                    print('> PART# :', num_part, ', Parameter :', name_param, ', id_param =', id_param)
                    df: pd.DataFrame = self.parent.sheets.get_part_all(num_part_excel)
                    for r in range(len(df)):
                        df_r: pd.Series = df.iloc[r]
                        # batch
                        # id_batch INTEGER
                        # 'Sample' => sample INTEGER,
                        # 'Date' => timestamp INTEGER,
                        # 'Job ID or Lot ID' => id_lot TEXT,
                        # 'Serial Number' => serial,
                        sample = df_r['Sample']
                        timestamp = int(pd.to_datetime(df_r['Date']).timestamp())

                        id_lot = df_r['Job ID or Lot ID']
                        if math.isnan(id_lot):
                            id_lot = ''

                        serial = df_r['Serial Number']

                        measured = df_r[name_param]
                        if math.isnan(measured):
                            measured = 'NULL'

                        # print(sample, timestamp, id_lot, serial, value)

                        # id_batch
                        id_batch = self.select_id_batch(sample, serial, timestamp)
                        if id_batch is None:
                            self.insert_batch(id_lot, sample, serial, timestamp)
                            id_batch = self.select_id_batch(sample, serial, timestamp)

                        # id_measure
                        id_measure = self.select_id_measure(id_batch, id_param)
                        if id_measure is None:
                            self.insert_measure(id_batch, id_param, measured)

    def get_param(self, metrics):
        param = {}
        # param_num_key = metrics['Key Parameter']
        param['num_key'] = metrics['Key Parameter']
        # param_lsl = metrics['LSL']
        param['lsl'] = metrics['LSL']
        if math.isnan(param['lsl']):
            param['lsl'] = 'NULL'
        # param_target = metrics['Target']
        param['target'] = metrics['Target']
        if math.isnan(param['target']):
            param['target'] = 'NULL'
        # param_usl = metrics['USL']
        param['usl'] = metrics['USL']
        if math.isnan(param['usl']):
            param['usl'] = 'NULL'
        # param_charttype = metrics['Chart Type']
        param['charttype'] = metrics['Chart Type']
        # param_metrology = metrics['Metrology']
        param['metrology'] = metrics['Metrology']
        # param_multiple = metrics['Multiple']
        param['multiple'] = metrics['Multiple']
        # param_spectype = metrics['Spec Type']
        param['spectype'] = metrics['Spec Type']
        # param_frozen = metrics['CL Frozen']
        param['frozen'] = metrics['CL Frozen']
        # param_lcl = metrics['LCL']
        param['lcl'] = metrics['LCL']
        if math.isnan(param['lcl']):
            param['lcl'] = 'NULL'
        # param_mean = metrics['Avg']
        param['mean'] = metrics['Avg']
        if math.isnan(param['mean']):
            param['mean'] = 'NULL'
        # param_ucl = metrics['UCL']
        param['ucl'] = metrics['UCL']
        if math.isnan(param['ucl']):
            param['ucl'] = 'NULL'
        return param

    def select_id_supplier(self, name_supplier):
        sql = self.db.sql("SELECT id_supplier FROM supplier WHERE name_supplier_short = '?';", [name_supplier])
        # print(sql1)
        out = self.db.get(sql)
        id_supplier = None
        for id in out:
            id_supplier = id[0]
        return id_supplier

    def select_id_part(self, id_supplier, num_part):
        sql = self.db.sql("SELECT id_part FROM part WHERE num_part = '?' AND id_supplier = ?;", [num_part, id_supplier])
        # print(sql2)
        out = self.db.get(sql)
        id_part = None
        for id in out:
            id_part = id[0]
        return id_part

    def select_id_param(self, id_part, id_supplier, name_param):
        sql = self.db.sql("SELECT id_param FROM param WHERE id_supplier = ? AND id_part = ? AND name_param = '?';", [id_supplier, id_part, name_param])
        # print(sql3)
        out = self.db.get(sql)
        id_param = None
        for id in out:
            id_param = id[0]
        return id_param

    def insert_param(self, id_part, id_supplier, name_param, num_part_excel, param):
        sql = self.db.sql("INSERT INTO param VALUES(NULL, ?, ?, '?', '?', '?', ?, ?, ?, '?', '?', '?', '?', '?', ?, ?, ?);",
                          [
                              id_supplier,
                              id_part,
                              num_part_excel,
                              name_param,
                              param['num_key'],
                              param['lsl'],
                              param['target'],
                              param['usl'],
                              param['charttype'],
                              param['metrology'],
                              param['multiple'],
                              param['spectype'],
                              param['frozen'],
                              param['lcl'],
                              param['mean'],
                              param['ucl']
                          ]
                          )
        # print(sql4)
        self.db.put(sql)

    def select_id_batch(self, sample, serial, timestamp):
        sql = self.db.sql("SELECT id_batch FROM batch WHERE sample = ? AND timestamp = ? AND serial = '?';", [sample, timestamp, serial])
        out = self.db.get(sql)
        id_batch = None
        for id in out:
            id_batch = id[0]
        return id_batch

    def insert_batch(self, id_lot, sample, serial, timestamp):
        sql = self.db.sql("INSERT INTO batch VALUES(NULL, ?, ?, '?', '?');", [sample, timestamp, id_lot, serial])
        # print(sql)
        self.db.put(sql)

    def select_id_measure(self, id_batch, id_param):
        sql = self.db.sql("SELECT id_measure FROM measure WHERE id_param = ? AND id_batch = ?;", [id_param, id_batch])
        # print(sql)
        out = self.db.get(sql)
        id_measure = None
        for id in out:
            id_measure = id[0]

        return id_measure

    def insert_measure(self, id_batch, id_param, measured):
        sql = self.db.sql("INSERT INTO measure VALUES(NULL, ?, ?, ?);", [id_param, id_batch, measured])
        # print(sql)
        self.db.put(sql)

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
