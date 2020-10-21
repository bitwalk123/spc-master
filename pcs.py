import numpy as np
import pathlib
import subprocess
import tempfile
import wx

from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib import rcParams

from pptx.util import Inches

# =============================================================================
#  ChartWin class
# =============================================================================
from office import PowerPoint


class ChartWin(wx.Frame):
    parent = None
    sheets = None
    num_param = 0
    row = 0

    tool_check = None

    def __init__(self, parent, sheets, num_param, row):
        # super(ChartWin, self).__init__(parent=parent, id=wx.ID_ANY, style= wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX)
        super(ChartWin, self).__init__(parent=parent, id=wx.ID_ANY)

        self.parent = parent
        self.sheets = sheets
        self.num_param = num_param
        self.row = row

        self.Bind(wx.EVT_CLOSE, self.OnCloseFrame)
        self.SetIcon(wx.Icon('images/chart.ico', wx.BITMAP_TYPE_ICO))
        # self.MakeModal(True)

        toolbar = self.CreateToolBar()
        tool_before = toolbar.AddTool(
            toolId=wx.ID_ANY,
            label='previous',
            bitmap=wx.Bitmap('images/before.png')
        )
        tool_after = toolbar.AddTool(
            toolId=wx.ID_ANY,
            label='next',
            bitmap=wx.Bitmap('images/after.png')
        )
        toolbar.AddSeparator()
        self.tool_check = wx.CheckBox(toolbar, label='All parameters')
        toolbar.AddControl(self.tool_check)
        tool_ppt = toolbar.AddTool(
            toolId=wx.ID_ANY,
            label='PowerPoint',
            bitmap=wx.Bitmap('images/powerpoint.png')
        )
        self.Bind(wx.EVT_TOOL, self.OnBefore, tool_before)
        self.Bind(wx.EVT_TOOL, self.OnAfter, tool_after)
        self.Bind(wx.EVT_TOOL, self.OnPPT, tool_ppt)

        toolbar.Realize()
        self.statusbar = self.CreateStatusBar()

        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.SetSizer(self.sizer)
        self.create_chart()
        self.Fit()
        self.Show()

    # -------------------------------------------------------------------------
    #  create_chart
    # -------------------------------------------------------------------------
    def create_chart(self):
        # get Parameter Name & PART Number
        name_part, name_param = self.get_part_param(self.row)
        self.UpdateTitle(name_part, name_param)
        self.canvas = self.gen_chart(name_part, name_param)

        # assign canvas on the widget
        size = self.sizer.GetSize()
        self.sizer.Clear(delete_windows=True)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.sizer.SetDimension(0, 0, size[0], size[1])

        # update row selection of 'Master' sheet
        self.parent.setMasterRowSelect(self.row)

    # -------------------------------------------------------------------------
    #  gen_chart - generate chart
    #
    #  argument
    #    name_part  : PART Number
    #    name_param : Parameter Name
    #    sheet      : data sheet from Excel file
    #
    #  return
    #    canvas : generated chart
    # -------------------------------------------------------------------------
    def gen_chart(self, name_part, name_param):
        # create PowerPoint file
        info = {
            'PART': name_part,
            'PARAM': name_param,
        }
        trend = Trend(self.sheets)
        figure = trend.get(info)
        canvas = FigureCanvas(self, -1, figure)
        # canvas.set_size_request(1500, 500)
        return canvas

    # -------------------------------------------------------------------------
    #  get_part_param - get PART No & Parameter Name from sheet
    #
    #  argument
    #    part  : PART name
    #    param : PARAMETER name
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def UpdateTitle(self, part, param):
        self.SetTitle(part + ' : ' + param)

    # -------------------------------------------------------------------------
    #  get_part_param - get PART No & Parameter Name from sheet
    #
    #  argument
    #    sheets : data sheet from Excel file
    #    row    : row object on the Master Table
    #
    #  return
    #    name_part  :
    #    name_param :
    # -------------------------------------------------------------------------
    def get_part_param(self, row):
        df_master = self.sheets.get_master()
        df_row = df_master.iloc[row]
        name_part = df_row['Part Number']
        name_param = df_row['Parameter Name']

        return name_part, name_param

    # -------------------------------------------------------------------------
    #  MakeModal
    # -------------------------------------------------------------------------
    def MakeModal(self, modal=True):
        if modal and not hasattr(self, '_disabler'):
            self._disabler = wx.WindowDisabler(self)
        if not modal and hasattr(self, '_disabler'):
            del self._disabler

    # -------------------------------------------------------------------------
    #  OnPPT
    # -------------------------------------------------------------------------
    def OnPPT(self, event):
        template_path = "./template.pptx"
        image_path = tempfile.NamedTemporaryFile(suffix='.png').name
        save_path = tempfile.NamedTemporaryFile(suffix='.pptx').name

        # check box is checked?
        if self.tool_check.GetValue():
            # loop fpr all parameters
            loop = range(self.num_param)
        else:
            # This is single loop
            loop = [self.row]

        for row in loop:
            # get Parameter Name & PART Number
            name_part, name_param = self.get_part_param(row)

            # create PowerPoint file
            info = {
                'PART': name_part,
                'PARAM': name_param,
                'IMAGE': image_path,
                'ileft': Inches(0),
                'itop': Inches(0.84),
                'iheight': Inches(3.5),
            }

            # create chart
            trend = Trend(self.sheets)
            figure = trend.get(info)
            # create PNG file of plot
            figure.savefig(image_path)

            # gen_ppt(template_path, image_path, save_path, info)
            if self.tool_check.GetValue() and row > 0:
                template_path = save_path

            ppt_obj = PowerPoint(template_path)
            ppt_obj.add_slide(self.sheets, info)
            ppt_obj.save(save_path)

        # open created file
        self.open_file_with_app(save_path)

    # -------------------------------------------------------------------------
    #  OnAfter
    # -------------------------------------------------------------------------
    def OnAfter(self, event):
        if self.row >= self.num_param - 1:
            self.row = self.num_param - 1
            return

        self.row += 1
        self.create_chart()

    # -------------------------------------------------------------------------
    #  OnBefore
    # -------------------------------------------------------------------------
    def OnBefore(self, event):
        if self.row <= 0:
            self.row = 0
            return

        self.row -= 1
        self.create_chart()

    # -------------------------------------------------------------------------
    #  open_file_with_app
    #
    #  argument
    #    name_file :  file to open
    # -------------------------------------------------------------------------
    def open_file_with_app(self, name_file):
        link_file = pathlib.PurePath(name_file)
        # Explorer can cover all cases on Windows NT
        subprocess.Popen(['explorer', link_file])

    # -------------------------------------------------------------------------
    #  OnCloseFrame - Makes sure user was intending to quit the application
    # -------------------------------------------------------------------------
    def OnCloseFrame(self, event):
        self.parent.chart = None
        self.Destroy()


# =============================================================================
#  Trend class
# =============================================================================
class Trend():
    sheets = None
    splot = None

    size_point = 10
    size_oos = 30
    size_ooc = 50

    x_left = 0
    x_right = 0
    unit = 0
    nchar = 93

    def __init__(self, sheets):
        self.sheets = sheets

    # -------------------------------------------------------------------------
    #  get
    # -------------------------------------------------------------------------
    def get(self, info):
        name_part = info['PART']
        name_param = info['PARAM']

        metrics = self.sheets.get_metrics(name_part, name_param)
        df = self.sheets.get_part(name_part)
        x = df['Sample']
        y = df[name_param]

        rcParams['font.family'] = 'monospace'
        fig = Figure(dpi=100, figsize=(10, 3.5))
        self.splot = fig.add_subplot(111, title=name_param, ylabel='Value')
        self.splot.grid(True)

        if metrics['Spec Type'] == 'Two-Sided':
            self.axhline_two_sided(metrics)
        elif metrics['Spec Type'] == 'One-Sided':
            self.axhline_one_sided(metrics)

        # Avg
        self.splot.axhline(y=metrics['Avg'], linewidth=1, color='green', label='Avg')

        # Line
        self.splot.plot(x, y, linewidth=1, color="gray")

        # Out Of Limits
        if metrics['Spec Type'] == 'Two-Sided':
            self.violation_two_sided(df, metrics, name_param, x, y)
        elif metrics['Spec Type'] == 'One-Sided':
            self.violation_one_sided(df, metrics, name_param, x, y)

        # Histric data
        dataType = 'Historic'
        color_point = 'gray'
        self.draw_points(color_point, dataType, df, x, y)
        # Recent data
        dataType = 'Recent'
        color_point = 'black'
        self.draw_points(color_point, dataType, df, x, y)

        # plot dimension for labels
        self.x_left = self.splot.get_xlim()[0]
        self.x_right = self.splot.get_xlim()[1]
        self.unit = self.get_unit(self.x_left, self.x_right)
        # Label for Limits
        if metrics['Spec Type'] == 'Two-Sided':
            self.limits_label_two_sided(metrics)
        elif metrics['Spec Type'] == 'One-Sided':
            self.limits_label_one_sided(metrics)

        #y = metrics['USL']
        #label='123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123'
        #self.splot.text(x=self.x_left, y=y, s=label, color='red')

        return fig

    # -------------------------------------------------------------------------
    #  draw_points
    # -------------------------------------------------------------------------
    def draw_points(self, color, type, df, x, y):
        x_historic = x[df['Data Type'] == type]
        y_historic = y[df['Data Type'] == type]
        self.splot.scatter(x_historic, y_historic, s=self.size_point, c=color, marker='o', label=type)

    # -------------------------------------------------------------------------
    #  axhline_one_sided
    # -------------------------------------------------------------------------
    def axhline_one_sided(self, metrics):
        if not np.isnan(metrics['USL']):
            self.splot.axhline(y=metrics['USL'], linewidth=1, color='blue', label='USL')
        if not np.isnan(metrics['UCL']):
            self.splot.axhline(y=metrics['UCL'], linewidth=1, color='red', label='UCL')
        if not np.isnan(metrics['RUCL']):
            self.splot.axhline(y=metrics['RUCL'], linewidth=1, color='black', label='UCL')

    # -------------------------------------------------------------------------
    #  axhline_two_sided
    # -------------------------------------------------------------------------
    def axhline_two_sided(self, metrics):
        self.axhline_one_sided(metrics)
        if not np.isnan(metrics['Target']):
            self.splot.axhline(y=metrics['Target'], linewidth=1, color='purple', label='Target')
        if not np.isnan(metrics['RLCL']):
            self.splot.axhline(y=metrics['RLCL'], linewidth=1, color='black', label='LCL')
        if not np.isnan(metrics['LCL']):
            self.splot.axhline(y=metrics['LCL'], linewidth=1, color='red', label='LCL')
        if not np.isnan(metrics['LSL']):
            self.splot.axhline(y=metrics['LSL'], linewidth=1, color='blue', label='LSL')

    # -------------------------------------------------------------------------
    #  violation_one_sided
    # -------------------------------------------------------------------------
    def violation_one_sided(self, df, metrics, name_param, x, y):
        # OOC check
        x_ooc = x[df[name_param] > metrics['UCL']]
        y_ooc = y[df[name_param] > metrics['UCL']]
        self.splot.scatter(x_ooc, y_ooc, s=self.size_ooc, c='orange', marker='o', label="Recent")
        # OOS check
        x_oos = x[df[name_param] > metrics['USL']]
        y_oos = y[df[name_param] > metrics['USL']]
        self.splot.scatter(x_oos, y_oos, s=self.size_oos, c='red', marker='o', label="Recent")

    # -------------------------------------------------------------------------
    #  violation_two_sided
    # -------------------------------------------------------------------------
    def violation_two_sided(self, df, metrics, name_param, x, y):
        # OOC check
        x_ooc = x[(df[name_param] < metrics['LCL']) | (df[name_param] > metrics['UCL'])]
        y_ooc = y[(df[name_param] < metrics['LCL']) | (df[name_param] > metrics['UCL'])]
        self.splot.scatter(x_ooc, y_ooc, s=self.size_ooc, c='orange', marker='o', label="Recent")
        # OOS check
        x_oos = x[(df[name_param] < metrics['LSL']) | (df[name_param] > metrics['USL'])]
        y_oos = y[(df[name_param] < metrics['LSL']) | (df[name_param] > metrics['USL'])]
        self.splot.scatter(x_oos, y_oos, s=self.size_oos, c='red', marker='o', label="Recent")

    def get_unit(self, x_left, x_right):
        n = 87.5
        d = (x_right - x_left) / n
        return d

    # -------------------------------------------------------------------------
    #  limits_label_one_sided
    # -------------------------------------------------------------------------
    def limits_label_one_sided(self, metrics):
        label = 'USL'
        lcolor = 'blue'
        self.put_label_right(label, lcolor, metrics)

        label = 'UCL'
        lcolor = 'red'
        self.put_label_left(label, lcolor, metrics)

        label = 'RUCL'
        lcolor = 'black'
        self.put_label_right(label, lcolor, metrics)

        label = 'Avg'
        lcolor = 'green'
        self.put_label_right(label, lcolor, metrics)

    # -------------------------------------------------------------------------
    #  limits_label_two_sided
    # -------------------------------------------------------------------------
    def limits_label_two_sided(self, metrics):
        self.limits_label_one_sided(metrics)

        label = 'Target'
        lcolor = 'purple'
        self.put_label_left(label, lcolor, metrics)

        label = 'RLCL'
        lcolor = 'black'
        self.put_label_right(label, lcolor, metrics)

        label = 'LCL'
        lcolor = 'red'
        self.put_label_left(label, lcolor, metrics)

        label = 'LSL'
        lcolor = 'blue'
        self.put_label_right(label, lcolor, metrics)

    # -------------------------------------------------------------------------
    #  put_label_right
    # -------------------------------------------------------------------------
    def put_label_right(self, label, lcolor, metrics):
        if not np.isnan(metrics[label]):
            x = self.x_right + self.unit / 2
            y = metrics[label]
            self.splot.text(x=x, y=y, s=label, color=lcolor)

    # -------------------------------------------------------------------------
    #  put_label_left
    # -------------------------------------------------------------------------
    def put_label_left(self, label, lcolor, metrics):
        if not np.isnan(metrics[label]):
            x = self.x_left - len(label) * self.unit
            y = metrics[label]
            self.splot.text(x=x, y=y, s=label, color=lcolor)

# ---
# PROGRAM END
