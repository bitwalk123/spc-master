import datetime
import math
import matplotlib
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pathlib
import platform
import re
import subprocess
import sys
import tempfile

from PySide2.QtCore import Qt
from PySide2.QtGui import QIcon
from PySide2.QtWidgets import (
    QCheckBox,
    QDockWidget,
    QMainWindow,
    QMessageBox,
    QStatusBar,
    QToolBar,
    QToolButton,
)

from office import ExcelSPC, PowerPoint


class ChartWin(QMainWindow):
    parent = None
    sheets = None
    num_param = 0
    row = 0

    canvas = None

    # icons
    icon_chart: str = 'images/chart.ico'
    icon_before: str = 'images/go-previous.png'
    icon_after: str = 'images/go-next.png'
    icon_ppt: str = 'images/x-office-presentation.png'

    NavigationToolbar.toolitems = (
        ('Home', 'Reset original view', 'home', 'home'),
        # ('Back', 'Back to previous view', 'back', 'back'),
        # ('Forward', 'Forward to next view', 'forward', 'forward'),
        (None, None, None, None),
        ('Pan', 'Pan axes with left mouse, zoom with right', 'move', 'pan'),
        ('Zoom', 'Zoom to rectangle', 'zoom_to_rect', 'zoom'),
        # ('Subplots', 'Configure subplots', 'subplots', 'configure_subplots'),
        (None, None, None, None),
        ('Save', 'Save the figure', 'filesave', 'save_figure'),
    )

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
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def initUI(self):
        # Create toolbar
        toolbar = QToolBar()
        self.addToolBar(toolbar)

        # Add buttons to toolbar
        btn_before: QToolButton = QToolButton()
        btn_before.setIcon(QIcon(self.icon_before))
        btn_before.setStatusTip('goto previous PARAMETER')
        btn_before.clicked.connect(self.prev_chart)
        toolbar.addWidget(btn_before)

        # Add buttons to toolbar
        btn_after: QToolButton = QToolButton()
        btn_after.setIcon(QIcon(self.icon_after))
        btn_after.setStatusTip('go to next PARAMETER')
        btn_after.clicked.connect(self.next_chart)
        toolbar.addWidget(btn_after)

        toolbar.addSeparator()

        self.check_update: QCheckBox = QCheckBox('Hide Spec Limit(s)', self)
        self.checkbox_state()
        self.check_update.stateChanged.connect(self.update_status)
        toolbar.addWidget(self.check_update)

        toolbar.addSeparator()

        self.check_all_slides: QCheckBox = QCheckBox('All parameters', self)
        toolbar.addWidget(self.check_all_slides)

        # PowerPoint
        but_ppt: QToolButton = QToolButton()
        but_ppt.setIcon(QIcon(self.icon_ppt))
        but_ppt.setStatusTip('generate PowerPoint slide(s)')
        but_ppt.clicked.connect(self.OnPPT)
        toolbar.addWidget(but_ppt)

        # Status Bar
        self.statusbar = QStatusBar()
        self.setStatusBar(self.statusbar)

        self.create_chart()

        self.show()

    # -------------------------------------------------------------------------
    #  checkbox_state
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def checkbox_state(self):
        if self.sheets.get_SL_flag(self.row):
            self.check_update.setCheckState(Qt.Checked)
        else:
            self.check_update.setCheckState(Qt.Unchecked)

    # -------------------------------------------------------------------------
    #  update_status
    #
    #  argument
    #    state :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def update_status(self, state):
        sender = self.sender()
        if sender.checkState() == Qt.Checked:
            flag_new: bool = True
        else:
            flag_new: bool = False
        flag_old = self.sheets.get_SL_flag(self.row)
        if flag_new is not flag_old:
            self.sheets.set_SL_flag(self.row, flag_new)
            self.create_chart()

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
        # Canvas for SPC Chart
        if self.canvas is not None:
            # CentralWidget
            self.takeCentralWidget()
            del self.canvas
            # DockWidget
            self.removeDockWidget(self.dock)
            self.dock.deleteLater()
            del self.navtoolbar

        # PART Number & PARAMETER Name
        name_part, name_param = self.get_part_param(self.row)
        self.updateTitle(name_part, name_param)
        # CentralWidget
        self.canvas: FigureCanvas = self.gen_chart(name_part, name_param)
        self.setCentralWidget(self.canvas)

        # DockWidget
        self.navtoolbar: NavigationToolbar = NavigationToolbar(self.canvas, self)
        self.dock: QDockWidget = QDockWidget('Navigation Toolbar')
        self.dock.setFeatures(QDockWidget.NoDockWidgetFeatures)
        self.dock.setWidget(self.navtoolbar)
        self.addDockWidget(Qt.BottomDockWidgetArea, self.dock)

        # update row selection of 'Master' sheet
        self.parent.setMasterRowSelect(self.row)

    # -------------------------------------------------------------------------
    #  get_part_param - get PART No & PARAMETER Name from sheet
    #
    #  argument
    #    row   : row object on the Master Table
    #
    #  return
    #    part  : PART Name
    #    param : PPARAMETER Name
    # -------------------------------------------------------------------------
    def get_part_param(self, row: int):
        df_master: pd.DataFrame = self.sheets.get_master()
        df_row: int = df_master.iloc[row]
        part: str = df_row['Part Number']
        param: str = df_row['Parameter Name']

        return part, param

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
    #    part  : PART Number
    #    param : Parameter Name
    #    sheet : data sheet from Excel file
    #
    #  return
    #    canvas : generated chart
    # -------------------------------------------------------------------------
    def gen_chart(self, part: str, param: str):
        # create PowerPoint file
        info = {
            'PART': part,
            'PARAM': param,
        }
        trend: Trend = Trend(self, self.sheets, self.row)
        figure = trend.get(info)
        canvas: FigureCanvas = FigureCanvas(figure)

        return canvas

    # -------------------------------------------------------------------------
    #  next_chart
    #
    #  argument
    #    event :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def next_chart(self, event: bool):
        if self.row >= self.num_param - 1:
            self.row = self.num_param - 1
            return

        self.row += 1
        self.update_chart()

    # -------------------------------------------------------------------------
    #  prev_chart
    #
    #  argument
    #    event :
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def prev_chart(self, event: bool):
        if self.row <= 0:
            self.row: int = 0
            return

        self.row -= 1
        self.update_chart()

    # -------------------------------------------------------------------------
    #  update_chart
    #
    #  argument
    #    (none)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def update_chart(self):
        self.checkbox_state()
        self.create_chart()

    # -------------------------------------------------------------------------
    #  OnPPT
    # -------------------------------------------------------------------------
    def OnPPT(self, event):
        template_path: str = 'template/template.pptx'
        image_path: str = tempfile.NamedTemporaryFile(suffix='.png').name
        save_path: str = tempfile.NamedTemporaryFile(suffix='.pptx').name

        # check box is checked?
        if self.check_all_slides.checkState() == Qt.Checked:
            # loop fpr all parameters
            loop = range(self.num_param)
        else:
            # This is single loop
            loop = [self.row]

        for row in loop:
            # get Parameter Name & PART Number
            name_part, name_param = self.get_part_param(row)
            # print(row + 1, name_part, name_param)

            # create PowerPoint file
            info = {
                'PART': name_part,
                'PARAM': name_param,
                'IMAGE': image_path,
            }

            # create chart
            trend = Trend(self, self.sheets, row)
            figure = trend.get(info)

            dateObj = trend.get_last_date()
            # print(dateObj)
            # print(type(dateObj))
            if dateObj is None:
                info['Date of Last Lot Received'] = 'n/a'
            elif type(dateObj) is float:
                if math.isnan(dateObj):
                    info['Date of Last Lot Received'] = 'n/a'
                else:
                    info['Date of Last Lot Received'] = str(dateObj)
            elif type(dateObj) is str:
                info['Date of Last Lot Received'] = dateObj
            else:
                info['Date of Last Lot Received'] = dateObj.strftime('%m/%d/%Y')

            # create PNG file of plot
            figure.savefig(image_path)

            # gen_ppt(template_path, image_path, save_path, info)
            if self.check_all_slides.checkState() == Qt.Checked and row > 0:
                template_path = save_path

            ppt_obj = PowerPoint(template_path)
            ppt_obj.add_slide(self.sheets, info)
            ppt_obj.save(save_path)

        # open created file
        self.open_file_with_app(save_path)

    # -------------------------------------------------------------------------
    #  open_file_with_app
    #
    #  argument
    #    name_file : filename to open with associated application
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def open_file_with_app(self, name_file):
        path = pathlib.PurePath(name_file)

        if platform.system() == 'Linux':
            app_open = 'xdg-open'
        else:
            # Windows Explorer can cover all cases to start application with file
            app_open = 'explorer'

        subprocess.Popen([app_open, path])


# =============================================================================
#  Trend class
# =============================================================================
class Trend():
    # initial value of instances
    parent = None
    sheets = None
    row: int = 0
    ax1 = None
    ax2 = None

    # plot margin
    margin_plot_left: float = 0.17
    margin_plot_right: float = 0.83

    # font family to display
    font_family: str = 'monospace'

    # tick color
    color_tick: str = '#c0c0c0'

    # circle size of OOC, OOS
    size_point: int = 10
    size_oos_out: int = 100
    size_oos_in: int = 50
    size_ooc_out: int = 80
    size_ooc_in: int = 40

    # color of OOC, OOS
    color_ooc_out: str = 'red'
    color_ooc_in: str = 'white'
    color_oos_out: str = 'red'
    color_oos_in: str = 'white'

    # color of metrics
    SL: str = 'blue'
    CL: str = 'red'
    RCL: str = 'black'
    TG: str = 'purple'
    AVG: str = 'green'

    # Regular Expression
    pattern1: str = re.compile(r'.*_(Max|Min)')  # check whether parameter name includes Max/Min
    pattern2: str = re.compile(r'.*\.(.*)')  # ___ extract right side from floating point in mumber

    flag_no_CL: bool = False
    date_last = None

    def __init__(self, parent: ChartWin, sheets: ExcelSPC, row: int):
        plt.close()
        self.parent: ChartWin = parent
        self.sheets: ExcelSPC = sheets
        self.row: int = row

    # -------------------------------------------------------------------------
    #  get - obtain SPC chart
    #
    #  argument
    #    info : dictionary including parameter specific information
    #
    #  return
    #    plt.figure instance with SPC chart
    # -------------------------------------------------------------------------
    def get(self, info: dict):
        matplotlib.rcParams['font.family'] = self.font_family

        name_part: str = info['PART']
        name_param: str = info['PARAM']

        # check whether parameter name includes Max/Min
        match: bool = self.pattern1.match(name_param)
        if match:
            self.flag_no_CL: bool = True
        else:
            self.flag_no_CL: bool = False

        metrics: dict = self.sheets.get_metrics(name_part, name_param)
        df: pd.DataFrame = self.sheets.get_part(name_part)

        x: pd.Series = df['Sample']
        if len(x.index) != len(x.unique()):
            # copy() is for preventing from following warning:
            # ------------------------------------------------
            # SettingWithCopyWarning:
            # A value is trying to be set on a copy of a slice from a DataFrame
            x_copy: pd.Series = x.copy()
            for i in x.index:
                x_copy.loc[i] = i
            x: pd.Series = x_copy

        try:
            y: pd.Series = df[name_param]
        except KeyError:
            return self.KeyErrorHandle(name_param)

        date: pd.Series = df['Date']

        if len(date) == 0:
            self.date_last = 'n/a'
        else:
            self.date_last: datetime = list(date)[len(date) - 1]

        fig = plt.figure(dpi=100, figsize=(10, 3.5))

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
        data_type: str = 'Historic'
        color_point: str = 'gray'
        self.draw_points(color_point, data_type, df, x, y)

        # _/_/_/_/_/_/_/
        # Recent data
        data_type: str = 'Recent'
        color_point: str = 'black'
        self.draw_points(color_point, data_type, df, x, y)

        # reflect ax1 limits to ax2 limits
        self.ax2.set_ylim(self.ax1.get_ylim())

        # ---------------------------------------------------------------------
        # Label for HORIZONTAL LINE
        # ---------------------------------------------------------------------
        self.add_y_axis_labels(fig, metrics)

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
    def KeyErrorHandle(self, name_param: str):
        msg: str = 'Oops!  There is no value associate with the parameter name, \'' \
                   + name_param + '\'.  Please check the Excel macro/sheet.'
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
    def draw_points(self, color: str, type: str, df, x, y):
        x_historic: pd.Series = x[df['Data Type'] == type]
        y_historic: pd.Series = y[df['Data Type'] == type]
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
            x_ooc: pd.Series = x[(df[name_param] > metrics['UCL']) & (df['Data Type'] == 'Recent')]
            y_ooc: pd.Series = y[(df[name_param] > metrics['UCL']) & (df['Data Type'] == 'Recent')]
            self.draw_circle(self.ax1, x_ooc, y_ooc, self.size_ooc_out, self.size_ooc_in, self.color_ooc_out, self.color_ooc_in)

        # OOS check
        x_oos: pd.Series = x[(df[name_param] > metrics['USL']) & (df['Data Type'] == 'Recent')]
        y_oos: pd.Series = y[(df[name_param] > metrics['USL']) & (df['Data Type'] == 'Recent')]
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
            x_ooc: pd.Series = x[((df[name_param] < metrics['LCL']) | (df[name_param] > metrics['UCL'])) & (df['Data Type'] == 'Recent')]
            y_ooc: pd.Series = y[((df[name_param] < metrics['LCL']) | (df[name_param] > metrics['UCL'])) & (df['Data Type'] == 'Recent')]
            self.draw_circle(self.ax1, x_ooc, y_ooc, self.size_ooc_out, self.size_ooc_in, self.color_ooc_out, self.color_ooc_in)

        # OOS check
        x_oos: pd.Series = x[((df[name_param] < metrics['LSL']) | (df[name_param] > metrics['USL'])) & (df['Data Type'] == 'Recent')]
        y_oos: pd.Series = y[((df[name_param] < metrics['LSL']) | (df[name_param] > metrics['USL'])) & (df['Data Type'] == 'Recent')]
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
        list_labels_left: list[str] = []
        list_labels_right: list[str] = []
        if metrics['Spec Type'] == 'Two-Sided':
            # LEFT
            if self.flag_no_CL is False:
                labels_left: list[str] = ['LCL', 'Target', 'UCL']
            else:
                labels_left: list[str] = ['Target']

            if self.sheets.get_SL_flag(self.row) is False:
                labels_left.extend(['LSL', 'USL'])

            # RIGHT
            if self.flag_no_CL is False:
                labels_right: list[str] = ['RLCL', 'Avg', 'RUCL']
            else:
                labels_right: list[str] = ['Avg']
        elif metrics['Spec Type'] == 'One-Sided':
            # LEFT
            if self.flag_no_CL is False:
                labels_left: list[str] = ['UCL']
            else:
                labels_left: list[str] = []

            if self.sheets.get_SL_flag(self.row) is False:
                labels_left.extend(['USL'])

            # RIGHT
            if self.flag_no_CL is False:
                labels_right: list[str] = ['Avg', 'RUCL']
            else:
                labels_right: list[str] = ['Avg']
        else:
            labels_left: list[str] = []
            labels_right: list[str] = ['Avg']

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
            labels: list = [item.get_text() for item in self.ax1.get_yticklabels()]
            nformat: str = self.get_tick_label_format(labels)
            n: int = len(labels)
            m: int = len(list_labels)
            for i in range(m):
                k: int = n - m + i
                label_new: str = list_labels[i]
                value: float = metrics[label_new]
                labels[k] = label_new + ' = ' + nformat.format(value)
            self.ax1.set_yticklabels(labels)

            # Left Axis: color
            yticklabels: list = self.ax1.get_yticklabels()
            n: int = len(yticklabels)
            m: int = len(list_labels)
            for i in range(m):
                k: int = n - m + i
                label: str = list_labels[i]
                if label == 'USL' or label == 'LSL':
                    color: str = self.SL
                elif label == 'UCL' or label == 'LCL':
                    color: str = self.CL
                elif label == 'Target':
                    color: str = self.TG
                else:
                    color: str = 'black'

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
            labels: list = [item.get_text() for item in self.ax2.get_yticklabels()]
            nformat: str = self.get_tick_label_format(labels)
            n: int = len(labels)
            m: int = len(list_labels)
            for i in range(m):
                k: int = n - m + i
                label_new: str = list_labels[i]
                value: float = metrics[label_new]
                labels[k] = nformat.format(value) + ' = ' + label_new
            self.ax2.set_yticklabels(labels)

            # Right Axis: color
            yticklabels: list = self.ax2.get_yticklabels()
            n: int = len(yticklabels)
            m: int = len(list_labels)
            for i in range(m):
                k: int = n - m + i
                label: str = list_labels[i]
                if label == 'RUCL' or label == 'RLCL':
                    color: str = self.RCL
                elif label == 'Avg':
                    color: str = self.AVG
                else:
                    color: str = 'black'

                yticklabels[k].set_color(color)

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
        extraticks: list = []
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
    def get_tick_label_format(self, labels: list) -> str:
        digit: int = 0
        for label in labels:
            match: bool = self.pattern2.match(label)
            if match:
                n: int = len(match.group(1))
                if n > digit:
                    digit: int = n
        nformat: str = '{:.' + str(digit) + 'f}'

        return nformat

# ---
# PROGRAM END
