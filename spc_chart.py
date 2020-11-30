import numpy as np
import re
import matplotlib as mpl
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
import matplotlib.pyplot as plt

from PySide2.QtCore import Qt
from PySide2.QtGui import QIcon
from PySide2.QtWidgets import (
    QDockWidget,
    QMainWindow,
    QMessageBox,
    QStatusBar,
    QToolBar,
    QToolButton,
)

from office import ExcelSPC


class ChartWin(QMainWindow):
    parent = None
    sheets = None
    num_param = 0
    row = 0

    canvas = None

    # icons
    icon_chart: str = 'images/chart.ico'
    icon_before: str = 'images/before.png'
    icon_after: str = 'images/after.png'

    def __init__(self, parent: QMainWindow, sheets: ExcelSPC, num_param: int, row: int):
        super().__init__(parent=parent)

        self.parent = parent
        self.sheets = sheets
        self.num_param = num_param
        self.row = row

        self.initUI()
        self.setWindowIcon(QIcon(self.icon_chart))

    # -------------------------------------------------------------------------
    #  initUI - UI initialization
    # -------------------------------------------------------------------------
    def initUI(self):
        # Create toolbar
        toolbar: QToolBar = QToolBar()
        self.addToolBar(toolbar)

        # Add buttons to toolbar
        tool_param_before: QToolButton = QToolButton()
        tool_param_before.setIcon(QIcon(self.icon_before))
        tool_param_before.setStatusTip('before PARAMETER')
        # tool_param_before.clicked.connect(self.openFile)
        toolbar.addWidget(tool_param_before)

        # Add buttons to toolbar
        tool_param_after: QToolButton = QToolButton()
        tool_param_after.setIcon(QIcon(self.icon_after))
        tool_param_after.setStatusTip('after PARAMETER')
        # tool_param_after.clicked.connect(self.openFile)
        toolbar.addWidget(tool_param_after)

        # Status Bar
        self.statusbar: QStatusBar = QStatusBar()
        self.setStatusBar(self.statusbar)

        self.create_chart()

        self.show()

    # -------------------------------------------------------------------------
    #  create_chart
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def create_chart(self):
        # get Parameter Name & PART Number
        name_part, name_param = self.get_part_param(self.row)
        self.updateTitle(name_part, name_param)

        # Canvas for SPC Chart
        if self.canvas is not None:
            del self.canvas
        self.canvas: FigureCanvas = self.gen_chart(name_part, name_param)
        navtoolbar: NavigationToolbar = NavigationToolbar(self.canvas, self)
        dock: QDockWidget = QDockWidget()
        dock.setWidget(navtoolbar)
        self.setCentralWidget(self.canvas)
        self.addDockWidget(Qt.TopDockWidgetArea, dock)

        # update row selection of 'Master' sheet
        # self.parent.setMasterRowSelect(self.row)

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
    #  updateTitle - update window title
    #
    #  argument
    #    part  : PART name
    #    param : PARAMETER name
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def updateTitle(self, part: str, param: str):
        self.setWindowTitle(part + ' : ' + param)

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
        trend = Trend(self, self.sheets, self.row)
        figure = trend.get(info)
        canvas = FigureCanvas(figure)

        return canvas


# =============================================================================
#  Trend class
# =============================================================================
class Trend():
    # initial value of instances
    parent = None
    sheets = None
    row = 0
    ax1 = None
    ax2 = None

    # plot margin
    margin_plot_left = 0.17
    margin_plot_right = 0.83

    # font family to display
    font_family = 'monospace'

    # tick color
    color_tick = '#c0c0c0'

    # circle size of OOC, OOS
    size_point = 10
    size_oos_out = 100
    size_oos_in = 50
    size_ooc_out = 80
    size_ooc_in = 40

    # color of OOC, OOS
    color_ooc_out = 'red'
    color_ooc_in = 'white'
    color_oos_out = 'red'
    color_oos_in = 'white'

    # color of metrics
    SL = 'blue'
    CL = 'red'
    RCL = 'black'
    TG = 'purple'
    AVG = 'green'

    # Regular Expression
    pattern1 = re.compile(r'.*_(Max|Min)')  # check whether parameter name includes Max/Min
    pattern2 = re.compile(r'.*\.(.*)')  # ___ extract right side from floating point in mumber

    flag_no_CL = False
    date_last = None

    def __init__(self, parent, sheets, row):
        plt.close()
        self.parent = parent
        self.sheets = sheets
        self.row = row

    # -------------------------------------------------------------------------
    #  get - obtain SPC chart
    #
    #  argument
    #    info : dictionary including parameter specific information
    #
    #  return
    #    plt.figure instance with SPC chart
    # -------------------------------------------------------------------------
    def get(self, info):
        mpl.rcParams['font.family'] = self.font_family

        name_part = info['PART']
        name_param = info['PARAM']

        # check whether parameter name includes Max/Min
        match = self.pattern1.match(name_param)
        if match:
            self.flag_no_CL = True
        else:
            self.flag_no_CL = False

        metrics = self.sheets.get_metrics(name_part, name_param)
        df = self.sheets.get_part(name_part)
        x = df['Sample']
        try:
            y = df[name_param]
        except KeyError:
            return self.KeyErrorHandle(name_param)

        date = df['Date']
        # print(date[len(date)])

        if len(date) == 0:
            self.date_last = 'n/a'
        else:
            self.date_last = list(date)[len(date) - 1]

        fig = plt.figure(dpi=100, figsize=(10, 3.5))

        # if self.ax1 is not None:
        #    self.ax1.clear()
        # if self.ax2 is not None:
        #    self.ax2.clear()

        # -----------------------------------------------------------------
        # add first y axis
        self.ax1 = fig.add_subplot(111, title=name_param)
        plt.subplots_adjust(left=self.margin_plot_left, right=self.margin_plot_right)
        self.ax1.grid(False)

        # -----------------------------------------------------------------
        # add second y axis wish same range as first y axis
        self.ax2 = self.ax1.twinx()

        if metrics['Spec Type'] == 'Two-Sided':
            self.axhline_two_sided(metrics)
        elif metrics['Spec Type'] == 'One-Sided':
            self.axhline_one_sided(metrics)

        # Avg
        if not np.isnan(metrics['Avg']):
            self.ax1.axhline(y=metrics['Avg'], linewidth=1, color=self.AVG, label='Avg')

        # _/_/_/_/_/_/_/
        # Line
        self.ax1.plot(x, y, linewidth=1, color='gray')
        self.ax2.plot(x, y, linewidth=0, color='red')  # for debug

        # Axis color
        self.ax1.xaxis.label.set_color('gray')
        self.ax1.yaxis.label.set_color(self.color_tick)
        self.ax2.yaxis.label.set_color(self.color_tick)

        # default tick color
        self.ax1.tick_params(axis='x', colors='gray')
        self.ax1.tick_params(axis='y', colors='gray')
        self.ax2.tick_params(axis='y', colors='gray')

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

        self.ax2.set_ylim(self.ax1.get_ylim())

        # ---------------------------------------------------------------------
        # Label for HORIZONTAL LINE
        # ---------------------------------------------------------------------
        self.add_y_axis_labels(fig, metrics)
        # fig.canvas.draw();

        # print(self.ax1.get_ylim())
        # print(self.ax2.get_ylim())

        return fig

    # -------------------------------------------------------------------------
    #  KeyErrorHandle - Error handring for no name_param matching
    #
    #  argument
    #    name_param : Parameter Name
    #
    #  return
    #    plt.figure instance with blank plot frame & parameter name
    # -------------------------------------------------------------------------
    def KeyErrorHandle(self, name_param):
        msg = 'Oops!  There is no value associate with the parameter name, \'' \
              + name_param + '\'.  Please check the Excel macro/sheet.'
        # dialog = wx.MessageDialog(
        #    parent=self.parent,
        #    message=msg,
        #    caption='Error',
        #    pos=wx.DefaultPosition,
        #    style=wx.OK | wx.ICON_ERROR
        # )
        # dialog.ShowModal()

        QMessageBox.critical(self.parent, 'Error', msg)

        # return blank figure
        fig = plt.figure(dpi=100, figsize=(10, 3.5))
        fig.add_subplot(111, title=name_param)
        plt.subplots_adjust(left=self.margin_plot_left, right=self.margin_plot_right)
        return fig

    # -------------------------------------------------------------------------
    #  get_last_date
    #
    #  argument
    #    (none)
    #
    #  return
    #    latest data information of the data
    # -------------------------------------------------------------------------
    def get_last_date(self):
        return self.date_last

    # -------------------------------------------------------------------------
    #  draw_points
    #
    #  argument
    #    color :
    #    type  :
    #    df    :
    #    x     :
    #    y     :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def draw_points(self, color, type, df, x, y):
        x_historic = x[df['Data Type'] == type]
        y_historic = y[df['Data Type'] == type]
        self.ax1.scatter(x_historic, y_historic, s=self.size_point, c=color, marker='o', label=type)

    # -------------------------------------------------------------------------
    #  axhline_one_sided
    #
    #  argument
    #    metrics :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def axhline_one_sided(self, metrics):
        if self.sheets.get_SL_flag(self.row) is False:
            if not np.isnan(metrics['USL']):
                self.ax1.axhline(y=metrics['USL'], linewidth=1, color=self.SL, label='USL')
                self.ax2.axhline(y=metrics['USL'], linewidth=0, color=self.SL, label='USL')
        if self.flag_no_CL is False:
            if not np.isnan(metrics['UCL']):
                self.ax1.axhline(y=metrics['UCL'], linewidth=1, color=self.CL, label='UCL')
            if not np.isnan(metrics['RUCL']):
                self.ax1.axhline(y=metrics['RUCL'], linewidth=1, color=self.RCL, label='RUCL')

    # -------------------------------------------------------------------------
    #  axhline_two_sided
    #
    #  argument
    #    metrics :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def axhline_two_sided(self, metrics):
        self.axhline_one_sided(metrics)

        if not np.isnan(metrics['Target']):
            self.ax1.axhline(y=metrics['Target'], linewidth=1, color=self.TG, label='Target')

        if self.flag_no_CL is False:
            if not np.isnan(metrics['RLCL']):
                self.ax1.axhline(y=metrics['RLCL'], linewidth=1, color=self.RCL, label='RLCL')
            if not np.isnan(metrics['LCL']):
                self.ax1.axhline(y=metrics['LCL'], linewidth=1, color=self.CL, label='LCL')

        if self.sheets.get_SL_flag(self.row) is False:
            if not np.isnan(metrics['LSL']):
                self.ax1.axhline(y=metrics['LSL'], linewidth=1, color=self.SL, label='LSL')
                self.ax2.axhline(y=metrics['LSL'], linewidth=0, color=self.SL, label='LSL')

    # -------------------------------------------------------------------------
    #  violation_one_sided
    #
    #  argument
    #    df         :
    #    metrics    :
    #    name_param :
    #    x          :
    #    y          :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def violation_one_sided(self, df, metrics, name_param, x, y):
        # OOC check
        if self.flag_no_CL is False:
            x_ooc = x[(df[name_param] > metrics['UCL']) & (df['Data Type'] == 'Recent')]
            y_ooc = y[(df[name_param] > metrics['UCL']) & (df['Data Type'] == 'Recent')]
            self.draw_circle(self.ax1, x_ooc, y_ooc, self.size_ooc_out, self.size_ooc_in, self.color_ooc_out, self.color_ooc_in)

        # OOS check
        x_oos = x[(df[name_param] > metrics['USL']) & (df['Data Type'] == 'Recent')]
        y_oos = y[(df[name_param] > metrics['USL']) & (df['Data Type'] == 'Recent')]
        self.draw_circle(self.ax1, x_oos, y_oos, self.size_oos_out, self.size_oos_in, self.color_oos_out, self.color_oos_in)

    # -------------------------------------------------------------------------
    #  violation_two_sided
    #
    #  argument
    #    df         :
    #    metrics    :
    #    name_param :
    #    x          :
    #    y          :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def violation_two_sided(self, df, metrics, name_param, x, y):
        # OOC check
        if self.flag_no_CL is False:
            x_ooc = x[((df[name_param] < metrics['LCL']) | (df[name_param] > metrics['UCL'])) & (df['Data Type'] == 'Recent')]
            y_ooc = y[((df[name_param] < metrics['LCL']) | (df[name_param] > metrics['UCL'])) & (df['Data Type'] == 'Recent')]
            self.draw_circle(self.ax1, x_ooc, y_ooc, self.size_ooc_out, self.size_ooc_in, self.color_ooc_out, self.color_ooc_in)

        # OOS check
        x_oos = x[((df[name_param] < metrics['LSL']) | (df[name_param] > metrics['USL'])) & (df['Data Type'] == 'Recent')]
        y_oos = y[((df[name_param] < metrics['LSL']) | (df[name_param] > metrics['USL'])) & (df['Data Type'] == 'Recent')]
        self.draw_circle(self.ax1, x_oos, y_oos, self.size_oos_out, self.size_oos_in, self.color_oos_out, self.color_oos_in)

    # -------------------------------------------------------------------------
    #  add_y_axis_labels
    #
    #  argument
    #    ax        :
    #    x         :
    #    y         :
    #    size_out  :
    #    size_in   :
    #    color_out :
    #    color_in  :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def draw_circle(self, ax, x, y, size_out, size_in, color_out, color_in):
        ax.scatter(x, y, s=size_out, c=color_out, marker='o')
        ax.scatter(x, y, s=size_in, c=color_in, marker='o')

    # -------------------------------------------------------------------------
    #  add_y_axis_labels
    #
    #  argument
    #    fig     :
    #    metrics :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def add_y_axis_labels(self, fig, metrics):
        list_labels_left = []
        list_labels_right = []
        if metrics['Spec Type'] == 'Two-Sided':
            # LEFT
            if self.flag_no_CL is False:
                labels_left = ['LCL', 'Target', 'UCL']
            else:
                labels_left = ['Target']

            if self.sheets.get_SL_flag(self.row) is False:
                labels_left.extend(['LSL', 'USL'])

            # RIGHT
            if self.flag_no_CL is False:
                labels_right = ['RLCL', 'Avg', 'RUCL']
            else:
                labels_right = ['Avg']
        elif metrics['Spec Type'] == 'One-Sided':
            # LEFT
            if self.flag_no_CL is False:
                labels_left = ['UCL']
            else:
                labels_left = []

            if self.sheets.get_SL_flag(self.row) is False:
                labels_left.extend(['USL'])

            # RIGHT
            if self.flag_no_CL is False:
                labels_right = ['Avg', 'RUCL']
            else:
                labels_right = ['Avg']
        else:
            labels_left = []
            labels_right = ['Avg']

        # Check whether defined label has number or not
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
    #
    #  argument
    #    fig         :
    #    list_labels :
    #    metrics     :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def add_y_axis_labels_at_left(self, fig, list_labels, metrics):
        if len(list_labels) > 0:
            # Left Axis: add extra ticks
            self.add_extra_tick_values(self.ax1, fig, list_labels, metrics)

            # Left Axis: extra labels
            labels = [item.get_text() for item in self.ax1.get_yticklabels()]
            nformat = self.get_tick_label_format(labels)
            n = len(labels)
            m = len(list_labels)
            for i in range(m):
                k = n - m + i
                label_new = list_labels[i]
                value = metrics[label_new]
                labels[k] = label_new + ' = ' + nformat.format(value)
            self.ax1.set_yticklabels(labels)

            # Left Axis: color
            yticklabels = self.ax1.get_yticklabels()
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
    #
    #  argument
    #    fig         :
    #    list_labels :
    #    metrics     :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def add_y_axis_labels_at_right(self, fig, list_labels, metrics):
        if len(list_labels) > 0:
            # fig.canvas.draw();

            # Right Axis: add extra ticks
            self.add_extra_tick_values(self.ax2, fig, list_labels, metrics)

            # Right Axis: labels
            labels = [item.get_text() for item in self.ax2.get_yticklabels()]
            nformat = self.get_tick_label_format(labels)
            n = len(labels)
            m = len(list_labels)
            for i in range(m):
                k = n - m + i
                label_new = list_labels[i]
                value = metrics[label_new]
                labels[k] = nformat.format(value) + ' = ' + label_new
            self.ax2.set_yticklabels(labels)

            # Right Axis: color
            yticklabels = self.ax2.get_yticklabels()
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

            # set axis
            # print(self.ax.get_ylim())
            # self.ax2.set_ylim(self.ax.get_ylim())
            # print(self.ax2.get_ylim())

    # -------------------------------------------------------------------------
    #  add_extra_tick_values
    #
    #  argument
    #    ax          :
    #    fig         :
    #    list_labels :
    #    metrics     :
    #
    #  return
    #    (none)
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
    #
    #  argument
    #    labels :
    #
    #  return
    #    nformat - formatted string
    # -------------------------------------------------------------------------
    def get_tick_label_format(self, labels):
        digit = 0
        for label in labels:
            match = self.pattern2.match(label)
            if match:
                n = len(match.group(1))
                if n > digit:
                    digit = n
        nformat = '{:.' + str(digit) + 'f}'

        return nformat

# ---
# PROGRAM END
