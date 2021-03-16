from PySide2.QtWidgets import (
    QMainWindow,
    QToolBar,
)

# _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
class DBManWin(QMainWindow):
    parent = None
    def __init__(self, parent: QMainWindow, db):
        super().__init__(parent=parent)

        self.parent = parent
        print(type(db))

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

        self.show()
