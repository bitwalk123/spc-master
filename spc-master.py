#!/usr/bin/env python
import logging
import math
import os.path
import wx
import wx.grid
from office import ExcelSPC
from pcs import ChartWin


# formatter = '%(asctime)s:%(message)s'
# logging.basicConfig(format=formatter, filename='error.log', level=logging.DEBUG)
# logging.critical('critical message')
# logging.error('error message')


# =============================================================================
#  SPCMaster
# =============================================================================
class SPCMaster(wx.Frame):
    # Application information
    app_name = 'SPC Master'
    app_ver = '0.2 (alpha)'

    # initialize instances
    notebook = None
    statusbar = None
    grid_master = None
    sheets = None
    chart = None
    num_param = 0

    def __init__(self):
        super(SPCMaster, self).__init__(parent=None, id=wx.ID_ANY)

        self.Bind(wx.EVT_CLOSE, self.OnCloseFrame)
        self.SetSize(800, 600)
        self.SetIcon(wx.Icon('images/logo.ico', wx.BITMAP_TYPE_ICO))
        self.set_app_title()

        toolbar = self.CreateToolBar()
        self.statusbar = self.CreateStatusBar()

        tool_excel = toolbar.AddTool(
            toolId=wx.ID_ANY,
            label='Excel',
            bitmap=wx.Bitmap('images/excel.png')
        )
        self.Bind(wx.EVT_TOOL, self.OnOpen, tool_excel)
        toolbar.Realize()

        self.notebook = wx.Notebook(self, wx.ID_ANY, style=wx.NB_BOTTOM)

    # -------------------------------------------------------------------------
    #  set_app_title
    # -------------------------------------------------------------------------
    def set_app_title(self, filename=None):
        app_title = self.app_name + ' ' + self.app_ver
        if filename is not None:
            app_title = app_title + ' - ' + os.path.basename(filename)
        self.SetTitle(app_title)

    # -------------------------------------------------------------------------
    #  read_excel
    #  Aggregation from Excel for SPC
    #
    #  argument
    #    filename : Excel file to read
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def read_excel(self, filename):
        if self.sheets is not None:
            del self.sheets
        self.sheets = ExcelSPC(filename)

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        # check if read format is appropriate ot not
        if self.sheets.valid is not True:
            msg = 'Not appropriate format!'
            self.statusbar.SetStatusText(msg)
            dialog = wx.MessageDialog(
                parent=self,
                message=msg,
                caption='Warning',
                pos=wx.DefaultPosition,
                style=wx.OK | wx.ICON_ERROR
            )
            dialog.ShowModal()

            # delete instance
            self.sheets = None
            return

        # cleat statusbar message
        self.statusbar.SetStatusText('')

        # update application title
        self.set_app_title(filename)

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        # create tabs for tables & charts
        self.create_tabs()

    # -------------------------------------------------------------------------
    #  delete_current - delete current page (tab) of notebook
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def delete_current(self):
        # Notebook contents
        n = self.notebook.GetPageCount()
        for i in range(n):
            self.notebook.DeletePage(0)
            self.notebook.SendSizeEvent()

        # Chart check
        if self.chart is not None:
            self.chart.Destroy()
            self.chart = None

        # update
        self.Update()

    # -------------------------------------------------------------------------
    #  create_tabs
    #  create tab instances
    #
    #  argument
    #    sheet :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def create_tabs(self):
        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  'Master' tab
        self.create_tab_master()

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  PART tab(s)
        list_part = self.sheets.get_unique_part_list()
        for name_part in list_part:
            self.create_tab_part(name_part)

    # -------------------------------------------------------------------------
    #  create_tab_master
    #  creating 'Master' tab
    #
    #  argument
    #    sheet : object of Excel sheet
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def create_tab_master(self):
        df = self.sheets.get_master()
        r = len(df)
        c = len(df.columns)
        panel_master = SpreadSheet(self.notebook, row=r, col=c)
        self.notebook.InsertPage(0, panel_master, 'Master')

        self.grid_master = panel_master.get_grid()
        # double click event definition for opening plot window
        self.grid_master.Bind(
            wx.grid.EVT_GRID_LABEL_LEFT_DCLICK,
            self.OnHeaderDblClicked
        )
        self.num_param = self.gen_table(df, self.grid_master)

        panel_master.update()

    # -------------------------------------------------------------------------
    #  create_tab_part - creating 'Master' tab
    #
    #  argument
    #    sheet     : object of Excel sheet
    #    name_part : part name
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def create_tab_part(self, name_part):
        df = self.sheets.get_part(name_part)
        r = len(df)
        c = len(df.columns)
        panel_part = SpreadSheet(self.notebook, row=r, col=c)
        n = self.notebook.GetPageCount()
        self.notebook.InsertPage(n, panel_part, name_part)

        grid = panel_part.get_grid()
        self.gen_table(df, grid)
        panel_part.update()

    # -------------------------------------------------------------------------
    #  gen_table
    #
    #  argument
    #    df   :
    #    grid :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def gen_table(self, df, grid):
        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  table header
        x = 0
        for item in df.columns.values:
            grid.SetColLabelValue(x, str(item))
            x += 1

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  table contents
        y = 0
        for row in df.itertuples(name=None):
            x = 0
            for item in list(row):
                if x == 0:
                    x += 1
                    continue

                if (type(item) is float) or (type(item) is int):
                    # right align on the widget
                    xalign = wx.ALIGN_RIGHT
                    if math.isnan(item):
                        item = ''
                else:
                    # left align on the widget
                    xalign = wx.ALIGN_LEFT

                grid.SetCellValue(y, x - 1, str(item))
                grid.SetCellAlignment(y, x - 1, xalign, wx.ALIGN_CENTER)

                x += 1

            y += 1

        return y

    # -------------------------------------------------------------------------
    #  setRowSelect
    #
    #  argument
    #    row : row to be selected
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def setMasterRowSelect(self, row):
        self.grid_master.SelectRow(row)
        self.grid_master.MakeCellVisible(row, 0)

    # -------------------------------------------------------------------------
    #  OnCloseFrame - Makes sure user was intending to quit the application
    #
    #  argument
    #    event :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def OnCloseFrame(self, event):
        dialog = wx.MessageDialog(
            parent=self,
            message='Are you sure you want to quit?',
            caption='Warning',
            pos=wx.DefaultPosition,
            style=wx.YES_NO | wx.NO_DEFAULT | wx.ICON_WARNING
        )
        response = dialog.ShowModal()

        if (response == wx.ID_YES):
            self.OnExitApp(event)
        else:
            event.StopPropagation()

    # -------------------------------------------------------------------------
    #  OnExitApp - Destroys the main frame which quits the wxPython application
    #
    #  argument
    #    event :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def OnExitApp(self, event):
        self.Destroy()

    # -------------------------------------------------------------------------
    #  OnHeaderDblClicked - double click event on row header of grid
    #
    #  argument
    #    event :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def OnHeaderDblClicked(self, event):
        row = event.GetRow()
        if self.chart is not None:
            self.chart.Destroy()

        self.chart = ChartWin(self, self.sheets, self.num_param, row)

    # -------------------------------------------------------------------------
    #  OnOpen
    #
    #  argument
    #    event :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def OnOpen(self, event):
        self.statusbar.SetStatusText('')
        dialog = wx.FileDialog(
            parent=self,
            message='open Excel file',
            defaultDir='',
            defaultFile='',
            wildcard='*.xlsm',
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST
        )
        if dialog.ShowModal() == wx.ID_CANCEL:
            return

        self.delete_current()
        filename = dialog.GetPath()
        self.read_excel(filename)

        # change size of window a bit to show scrollbars on purpose
        size = self.GetSize()
        self.SetSize(size[0] - 1, size[1] - 1)
        self.SetSize(size[0], size[1])


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

    # -------------------------------------------------------------------------
    #  get_grid - return grid instance
    #
    #  argument
    #    (none)
    #
    #  return
    #    grid instance
    # -------------------------------------------------------------------------
    def get_grid(self):
        return (self.grid)

    # -------------------------------------------------------------------------
    #  update - update grid size
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def update(self):
        self.grid.AutoSize()


# =============================================================================
#  MAIN
# =============================================================================
if __name__ == '__main__':
    app = wx.App()
    win = SPCMaster()
    win.Show()
    app.MainLoop()
# ---
#  END OF PROGRAM
