import numpy as np
import pathlib
import re
import subprocess
import tempfile
import wx
import matplotlib.pyplot as plt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib import rcParams
from pptx.util import Inches
from office import PowerPoint


# =============================================================================
#  ChartWin class
# =============================================================================
class ChartWin(wx.Frame):
    ID_KNOB = 100
    ID_UPDATE = 101

    parent = None
    sheets = None
    num_param = 0
    row = 0

    check_all_slides = None
    canvas = None

    width_spin = 80

    def __init__(self, parent, sheets, num_param, row):
        super().__init__(parent=parent, id=wx.ID_ANY)

        self.parent = parent
        self.sheets = sheets
        self.num_param = num_param
        self.row = row

        self.Bind(wx.EVT_CLOSE, self.OnCloseFrame)
        self.SetIcon(wx.Icon('images/chart.ico', wx.BITMAP_TYPE_ICO))
        # self.MakeModal(True)

        toolbar = self.CreateToolBar()
        self.statusbar = self.CreateStatusBar()
        self.sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.SetSizer(self.sizer)

        # ---------------------------------------------------------------------
        # toolbar contents
        # ---------------------------------------------------------------------
        # previous parameter
        but_param_before = toolbar.AddTool(
            toolId=wx.ID_ANY,
            label='previous',
            bitmap=wx.Bitmap('images/before.png')
        )
        # post parameter
        but_param_after = toolbar.AddTool(
            toolId=wx.ID_ANY,
            label='next',
            bitmap=wx.Bitmap('images/after.png')
        )
        toolbar.AddSeparator()
        self.check_update = wx.CheckBox(toolbar, label='No SL displayed')
        self.check_update.SetValue(self.sheets.get_SL_flag(self.row))
        toolbar.AddControl(self.check_update)
        # button for updating chart
        but_update = toolbar.AddTool(
            toolId=wx.ID_ANY,
            label='Update',
            bitmap=wx.Bitmap('images/update.png')
        )
        toolbar.AddSeparator()
        self.check_all_slides = wx.CheckBox(toolbar, label='All parameters')
        toolbar.AddControl(self.check_all_slides)
        # button for generating PowerPoint slide(s)
        but_ppt = toolbar.AddTool(
            toolId=wx.ID_ANY,
            label='PowerPoint',
            bitmap=wx.Bitmap('images/powerpoint.png')
        )
        toolbar.Realize()

        self.create_chart()

        self.Fit()
        self.Show()

        self.Bind(wx.EVT_TOOL, self.OnBefore, but_param_before)
        self.Bind(wx.EVT_TOOL, self.OnAfter, but_param_after)
        self.Bind(wx.EVT_TOOL, self.OnUpdate, but_update)
        self.Bind(wx.EVT_TOOL, self.OnPPT, but_ppt)

    def OnUpdate(self, event):
        flag = self.WithoutSL()
        flag_old = self.sheets.get_SL_flag(self.row)
        if flag is not flag_old:
            self.sheets.set_SL_flag(self.row, flag)
            self.create_chart()

    def WithoutSL(self):
        return self.check_update.IsChecked()

    # -------------------------------------------------------------------------
    #  create_chart
    # -------------------------------------------------------------------------
    def create_chart(self):
        # get Parameter Name & PART Number
        name_part, name_param = self.get_part_param(self.row)
        self.UpdateTitle(name_part, name_param)

        # _/_/_/_/_/_/_/_/_/_/_/
        # delete old children
        size = self.sizer.GetSize()
        self.sizer.Clear(delete_windows=True)

        # Drawer
        # self.drawer = wx.BoxSizer(wx.HORIZONTAL)
        # self.sizer.Add(self.drawer, 0, wx.EXPAND, 0)

        # Canvas for SPC Chart
        if self.canvas is not None:
            del self.canvas
        self.canvas = self.gen_chart(name_part, name_param)
        self.sizer.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.sizer.SetDimension(0, 0, size[0], size[1])

        # Left Drawer
        # self.gen_drawer()
        # self.Layout()
        # self.Fit()

        # update row selection of 'Master' sheet
        self.parent.setMasterRowSelect(self.row)

    # -------------------------------------------------------------------------
    #  gen_drawer
    # -------------------------------------------------------------------------
    def gen_drawer(self):
        # ---------------------------------------------------------------------
        #  Left Drawer
        # ---------------------------------------------------------------------
        self.panel = wx.BoxSizer(wx.VERTICAL)
        self.drawer.Add(self.panel, 0, wx.EXPAND, 0)
        # Knob for Drawer
        self.knob = wx.Button(self, self.ID_KNOB, '▒', size=(10, 0), name='Hide')
        self.knob.SetBackgroundColour('#FFFFFF')
        self.drawer.Add(self.knob, 0, wx.EXPAND, 0)
        # ---------------------------------------------------------------------
        #  Drawer Contents
        # ---------------------------------------------------------------------
        # axis high
        self.spin_high = wx.SpinCtrlDouble(self, wx.ID_ANY, size=(self.width_spin, -1))
        self.spin_high.SetDigits(3)
        self.spin_high.SetValue(0.90)
        self.spin_high.SetMin(0.50)
        self.spin_high.SetMax(0.90)
        self.spin_high.SetIncrement(0.01)
        self.panel.Add(self.spin_high)
        # padding (1)
        pad1 = wx.Panel(self)
        self.panel.Add(pad1, 1, wx.EXPAND)
        # update button
        self.but_update = wx.Button(self, self.ID_UPDATE, 'update', name='Update')
        self.panel.Add(self.but_update, 0, wx.EXPAND)
        # padding (2)
        pad2 = wx.Panel(self)
        self.panel.Add(pad2, 1, wx.EXPAND)
        # axis low
        self.spin_low = wx.SpinCtrlDouble(self, wx.ID_ANY, size=(self.width_spin, -1))
        self.spin_low.SetDigits(3)
        self.spin_low.SetValue(0.10)
        self.spin_low.SetMin(0.10)
        self.spin_low.SetMax(0.50)
        self.spin_low.SetIncrement(0.01)
        self.panel.Add(self.spin_low)
        self.drawer.Hide(self.panel)

        # ---------------------------------------------------------------------
        #  Binding
        # ---------------------------------------------------------------------
        self.Bind(wx.EVT_BUTTON, self.click_knob)

    # -------------------------------------------------------------------------
    #  Event Handler for Knob/Drawer
    # -------------------------------------------------------------------------
    def click_knob(self, event):
        if event.GetEventObject().GetId() == self.ID_KNOB:
            if self.knob.GetName() == 'Hide':
                self.drawer.Show(self.panel)
                self.knob.SetName('Show')
            elif self.knob.GetName() == 'Show':
                self.drawer.Hide(self.panel)
                self.knob.SetName('Hide')
            else:
                return

            self.Layout()

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
        trend = Trend(self.sheets, self.row)
        figure = trend.get(info)
        canvas = FigureCanvas(self, -1, figure)
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
        if self.check_all_slides.GetValue():
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
            trend = Trend(self.sheets, row)
            figure = trend.get(info)
            # create PNG file of plot
            figure.savefig(image_path)

            # gen_ppt(template_path, image_path, save_path, info)
            if self.check_all_slides.GetValue() and row > 0:
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
        self.update_chart()

    # -------------------------------------------------------------------------
    #  OnBefore
    # -------------------------------------------------------------------------
    def OnBefore(self, event):
        if self.row <= 0:
            self.row = 0
            return

        self.row -= 1
        self.update_chart()

    # -------------------------------------------------------------------------
    #  update_chart
    # -------------------------------------------------------------------------
    def update_chart(self):
        self.check_update.SetValue(self.sheets.get_SL_flag(self.row))
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
    row = 0
    ax = None

    size_point = 10
    size_oos = 30
    size_ooc = 50

    SL = 'blue'
    CL = 'red'
    RCL = 'black'
    TG = 'purple'
    AVG = 'green'

    # pattern for Regular Expression
    pattern = re.compile(r'.*\.(.*)')

    def __init__(self, sheets, row):
        self.sheets = sheets
        self.row = row

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
        fig = plt.figure(dpi=100, figsize=(10, 3.5))

        self.ax = fig.add_subplot(111, title=name_param)
        plt.subplots_adjust(left=0.17, right=0.83)
        self.ax.grid(True)

        if metrics['Spec Type'] == 'Two-Sided':
            self.axhline_two_sided(metrics)
        elif metrics['Spec Type'] == 'One-Sided':
            self.axhline_one_sided(metrics)

        # Avg
        if not np.isnan(metrics['Avg']):
            self.ax.axhline(y=metrics['Avg'], linewidth=1, color=self.AVG, label='Avg')

        # _/_/_/_/_/_/_/
        # Line
        self.ax.plot(x, y, linewidth=1, color="gray")

        # Axis color
        self.ax.xaxis.label.set_color('gray')
        self.ax.yaxis.label.set_color('gray')
        self.ax.tick_params(axis='x', colors='gray')
        self.ax.tick_params(axis='y', colors='gray')

        # Out Of Limits
        if metrics['Spec Type'] == 'Two-Sided':
            self.violation_two_sided(df, metrics, name_param, x, y)
        elif metrics['Spec Type'] == 'One-Sided':
            self.violation_one_sided(df, metrics, name_param, x, y)

        # DATA POINTS
        # _/_/_/_/_/_/_/
        # Histric data
        dataType = 'Historic'
        color_point = 'gray'
        self.draw_points(color_point, dataType, df, x, y)
        # _/_/_/_/_/_/_/
        # Recent data
        dataType = 'Recent'
        color_point = 'black'
        self.draw_points(color_point, dataType, df, x, y)

        # ---------------------------------------------------------------------
        # Label for HORIZONTAL LINE
        # ---------------------------------------------------------------------
        self.add_y_axis_labels(fig, metrics)
        # fig.canvas.draw();

        return fig

    def __del__(self):
        plt.close()

    # -------------------------------------------------------------------------
    #  draw_points
    # -------------------------------------------------------------------------
    def draw_points(self, color, type, df, x, y):
        x_historic = x[df['Data Type'] == type]
        y_historic = y[df['Data Type'] == type]
        self.ax.scatter(x_historic, y_historic, s=self.size_point, c=color, marker='o', label=type)

    # -------------------------------------------------------------------------
    #  axhline_one_sided
    # -------------------------------------------------------------------------
    def axhline_one_sided(self, metrics):
        if self.sheets.get_SL_flag(self.row) is False:
            if not np.isnan(metrics['USL']):
                self.ax.axhline(y=metrics['USL'], linewidth=1, color=self.SL, label='USL')
        if not np.isnan(metrics['UCL']):
            self.ax.axhline(y=metrics['UCL'], linewidth=1, color=self.CL, label='UCL')
        if not np.isnan(metrics['RUCL']):
            self.ax.axhline(y=metrics['RUCL'], linewidth=1, color=self.RCL, label='RUCL')

    # -------------------------------------------------------------------------
    #  axhline_two_sided
    # -------------------------------------------------------------------------
    def axhline_two_sided(self, metrics):
        self.axhline_one_sided(metrics)
        if not np.isnan(metrics['Target']):
            self.ax.axhline(y=metrics['Target'], linewidth=1, color=self.TG, label='Target')
        if not np.isnan(metrics['RLCL']):
            self.ax.axhline(y=metrics['RLCL'], linewidth=1, color=self.RCL, label='RLCL')
        if not np.isnan(metrics['LCL']):
            self.ax.axhline(y=metrics['LCL'], linewidth=1, color=self.CL, label='LCL')
        if self.sheets.get_SL_flag(self.row) is False:
            if not np.isnan(metrics['LSL']):
                self.ax.axhline(y=metrics['LSL'], linewidth=1, color=self.SL, label='LSL')

    # -------------------------------------------------------------------------
    #  violation_one_sided
    # -------------------------------------------------------------------------
    def violation_one_sided(self, df, metrics, name_param, x, y):
        # OOC check
        x_ooc = x[df[name_param] > metrics['UCL']]
        y_ooc = y[df[name_param] > metrics['UCL']]
        self.ax.scatter(x_ooc, y_ooc, s=self.size_ooc, c='orange', marker='o', label="Recent")
        # OOS check
        x_oos = x[df[name_param] > metrics['USL']]
        y_oos = y[df[name_param] > metrics['USL']]
        self.ax.scatter(x_oos, y_oos, s=self.size_oos, c='red', marker='o', label="Recent")

    # -------------------------------------------------------------------------
    #  violation_two_sided
    # -------------------------------------------------------------------------
    def violation_two_sided(self, df, metrics, name_param, x, y):
        # OOC check
        x_ooc = x[(df[name_param] < metrics['LCL']) | (df[name_param] > metrics['UCL'])]
        y_ooc = y[(df[name_param] < metrics['LCL']) | (df[name_param] > metrics['UCL'])]
        self.ax.scatter(x_ooc, y_ooc, s=self.size_ooc, c='orange', marker='o', label="Recent")
        # OOS check
        x_oos = x[(df[name_param] < metrics['LSL']) | (df[name_param] > metrics['USL'])]
        y_oos = y[(df[name_param] < metrics['LSL']) | (df[name_param] > metrics['USL'])]
        self.ax.scatter(x_oos, y_oos, s=self.size_oos, c='red', marker='o', label="Recent")

    # -------------------------------------------------------------------------
    #  add_y_axis_labels
    # -------------------------------------------------------------------------
    def add_y_axis_labels(self, fig, metrics):
        list_labels_left = []
        list_labels_right = []
        if metrics['Spec Type'] == 'Two-Sided':
            labels_left = ['LCL', 'Target', 'UCL']
            if self.sheets.get_SL_flag(self.row) is False:
                labels_left.extend(['LSL', 'USL'])
            labels_right = ['RLCL', 'Avg', 'RUCL']
        elif metrics['Spec Type'] == 'One-Sided':
            labels_left = ['UCL']
            if self.sheets.get_SL_flag(self.row) is False:
                labels_left.extend(['USL'])
            labels_right = ['Avg', 'RUCL']
        else:
            labels_left = []
            labels_right = ['Avg']

        for label in labels_left:
            if not np.isnan(metrics[label]):
                list_labels_left.append(label)
        for label in labels_right:
            if not np.isnan(metrics[label]):
                list_labels_right.append(label)

        self.add_y_axis_labels_at_left(fig, list_labels_left, metrics)
        self.add_y_axis_labels_at_right(fig, list_labels_right, metrics)

    # -------------------------------------------------------------------------
    #  add_y_axis_labels_at_left
    # -------------------------------------------------------------------------
    def add_y_axis_labels_at_left(self, fig, list_labels, metrics):
        if len(list_labels) > 0:
            # Left Axis: add extra ticks
            self.add_extra_tick_values(self.ax, fig, list_labels, metrics)

            # Left Axis: extra labels
            labels = [item.get_text() for item in self.ax.get_yticklabels()]
            nformat = self.get_tick_label_format(labels)
            n = len(labels)
            m = len(list_labels)
            for i in range(m):
                k = n - m + i
                label_new = list_labels[i]
                value = metrics[label_new]
                labels[k] = label_new + ' = ' + nformat.format(value)
            self.ax.set_yticklabels(labels)

            # Left Axis: color
            yticklabels = self.ax.get_yticklabels()
            n = len(yticklabels)
            m = len(list_labels)
            for i in range(m):
                k = n - m + i
                label = list_labels[i]
                if label == 'USL' or label == 'LSL':
                    color = self.SL
                elif label == 'UCL' or label == 'LCL':
                    color = self.CL
                elif label == 'Target':
                    color = self.TG
                else:
                    color = 'black'

                yticklabels[k].set_color(color)

    # -------------------------------------------------------------------------
    #  add_y_axis_labels_at_right
    # -------------------------------------------------------------------------
    def add_y_axis_labels_at_right(self, fig, list_labels, metrics):
        if len(list_labels) > 0:
            # -----------------------------------------------------------------
            # add second y axis wish same range as first y axis
            ax2 = self.ax.twinx()
            ax2.set_ylim(self.ax.get_ylim())
            ax2.tick_params(axis='y', colors='gray')

            # Right Axis: add extra ticks
            self.add_extra_tick_values(ax2, fig, list_labels, metrics)

            # Right Axis: labels
            labels = [item.get_text() for item in ax2.get_yticklabels()]
            nformat = self.get_tick_label_format(labels)
            n = len(labels)
            m = len(list_labels)
            for i in range(m):
                k = n - m + i
                label_new = list_labels[i]
                value = metrics[label_new]
                labels[k] = nformat.format(value) + ' = ' + label_new
            ax2.set_yticklabels(labels)

            # Right Axis: color
            yticklabels = ax2.get_yticklabels()
            n = len(yticklabels)
            m = len(list_labels)
            for i in range(m):
                k = n - m + i
                label = list_labels[i]
                if label == 'RUCL' or label == 'RLCL':
                    color = self.RCL
                elif label == 'Avg':
                    color = self.AVG
                else:
                    color = 'black'

                yticklabels[k].set_color(color)

    # -------------------------------------------------------------------------
    #  add_extra_tick_values
    # -------------------------------------------------------------------------
    def add_extra_tick_values(self, ax, fig, list_labels, metrics):
        extraticks = []
        for label in list_labels:
            extraticks.append(metrics[label])

        ax.set_yticks(list(ax.get_yticks()) + extraticks)
        # update drawing to reflect new ticks
        fig.canvas.draw();

    # -------------------------------------------------------------------------
    #  get_tick_label_format
    # -------------------------------------------------------------------------
    def get_tick_label_format(self, labels):
        digit = 0
        for label in labels:
            match = self.pattern.match(label)
            if match:
                n = len(match.group(1))
                if n > digit:
                    digit = n
        nformat = '{:.' + str(digit) + 'f}'

        return nformat

# ---
# PROGRAM END
