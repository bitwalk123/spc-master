from PySide2.QtWidgets import (
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
)


# _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
# BComboBox
#
# description
#   label + combobox
class BComboBox(QFrame):
    def __init__(self, parent):
        super().__init__(parent=parent)

        #self.setFrameStyle(QFrame.StyledPanel)
        self.hbox = QHBoxLayout()

        self.lab = QLabel()
        self.hbox.addWidget(self.lab)
        self.combo = QComboBox(self)
        self.hbox.addWidget(self.combo)
        self.setLayout(self.hbox)

    # -------------------------------------------------------------------------
    #  setText - set text to label
    #
    #  argument
    #    str_label: str string to be set to label
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def setText(self, str_label: str):
        self.lab.setText(str_label)

    # -------------------------------------------------------------------------
    #  addItems - set list to combobox
    #
    #  argument
    #    list_element: list string list to be displayed on the combobox
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def addItems(self, list_element: list):
        self.combo.addItems(list_element)

    # -------------------------------------------------------------------------
    #  currentIndexChanged - event when combobox selection changed
    #
    #  argument
    #    name_method: method to be executed when selection changed
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def currentIndexChanged(self, name_method):
        self.combo.currentIndexChanged.connect(name_method)
