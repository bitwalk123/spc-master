from typing import Any
from PySide2.QtCore import (
    Qt,
    QModelIndex,
    QAbstractTableModel
)
# For Sample
from PySide2.QtWidgets import (
    QTableView,
)


class MasterTableModel(QAbstractTableModel):
    def __init__(self):
        QAbstractTableModel.__init__(self)
        self.colheader_master = ["NAME", "AGE", "COUNTRY"]
        self.data_master = [
            ["Taro", 24, "Japan"],
            ["Jiro", 20, "Japan"],
            ["David", 32, "USA"],
            ["Wattson", 15, "US"]
        ]

    def rowCount(self, parent=QModelIndex()) -> int:
        return len(self.data_master)

    def columnCount(self, parent=QModelIndex()) -> int:
        return len(self.colheader_master)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.colheader_master[section]
        else:
            return "{}".format(section + 1)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        if role == Qt.DisplayRole:
            return self.data_master[index.row()][index.column()]


class SheetMaster(QTableView):
    def __init__(self):
        super().__init__()
        self.setModel(MasterTableModel())
