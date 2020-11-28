import pandas as pd
import numpy as np
import math
from office import ExcelSPC

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


class SPCTableModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame, col_headers: list):
        QAbstractTableModel.__init__(self)
        self.df: pd.DataFrame = df
        self.col_headers: list = col_headers

    def rowCount(self, parent=QModelIndex()) -> int:
        return len(self.df)

    def columnCount(self, parent=QModelIndex()) -> int:
        return len(self.df.columns)

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return self.col_headers[section]
        else:
            return "{}".format(section + 1)

    def data(self, index: QModelIndex, role: int = Qt.DisplayRole) -> Any:
        if role == Qt.DisplayRole:
            value = self.df.iat[index.row(), index.column()]
            if type(value) is np.int64:
                value = value.tolist()
            elif type(value) is np.float64:
                value = value.tolist()
            return value


class SheetMaster(QTableView):
    def __init__(self, sheets: ExcelSPC):
        super().__init__()
        df: pd.DataFrame = sheets.get_master()
        self.setModel(SPCTableModel(df, sheets.get_header_master()))
