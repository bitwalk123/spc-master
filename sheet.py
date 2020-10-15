import wx
import wx.grid


# =============================================================================
#  WorkSheet
# =============================================================================
class SpreadSheet(wx.Panel):
    grid = None

    def __init__(self, parent, row, col):
        super(SpreadSheet, self).__init__(parent)

        self.grid = wx.grid.Grid(self)
        self.grid.CreateGrid(row, col)
        self.grid.EnableEditing(False)
        self.grid.AutoSize()

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.grid, 1, wx.EXPAND)
        self.SetSizer(sizer)

    def get_grid(self):
        return (self.grid)

    def update(self):
        self.grid.AutoSize()


