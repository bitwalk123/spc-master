import numpy as np
from gi.repository import Gtk
from matplotlib.backends.backend_gtk3agg import FigureCanvasGTK3Agg as FigureCanvas
from matplotlib.figure import Figure

from module import utils


# =============================================================================
#  CharWin - single chart window
# =============================================================================
class ChartWin(Gtk.Window):

    def __init__(self, info_master, widget, sheet):
        Gtk.Window.__init__(self)
        self.connect("delete-event", self.on_delete)
        self.info_master = info_master

        row = utils.register(int(widget.get_label()), 1, self.info_master.get_rows() - 1)
        self.info_master.select_row(row.get())

        name_param, name_part = self.get_part_param(row, sheet)

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  HeaderBar
        self.hbar = Gtk.HeaderBar()
        self.hbar.set_show_close_button(True)
        self.set_hbar_title(name_param, name_part)

        box_header = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)

        image_left = Gtk.Image.new_from_icon_name('go-previous', Gtk.IconSize.BUTTON)
        but_left = Gtk.Button()
        but_left.set_image(image_left)
        box_header.add(but_left)

        image_right = Gtk.Image.new_from_icon_name('go-next', Gtk.IconSize.BUTTON)
        but_right = Gtk.Button()
        but_right.set_image(image_right)
        box_header.add(but_right)

        self.hbar.pack_start(box_header)

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  SPC chart
        # canvas = self.generate_spc_plot(sheet, name_part, name_param)
        self.box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.add(self.box)
        self.gen_chart(name_param, name_part, sheet)

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  binding for clicking on arrow button
        but_left.connect('clicked', self.on_arrow_left_clicked, row, sheet)
        but_right.connect('clicked', self.on_arrow_right_clicked, row, sheet)

        self.show_all()

    def set_hbar_title(self, name_param, name_part):
        self.hbar.props.title = name_part + " - " + name_param
        self.set_titlebar(self.hbar)

    def gen_chart(self, name_param, name_part, sheet):
        canvas = make_trend_chart(sheet, name_part, name_param)
        canvas.set_size_request(1500, 500)
        self.box.pack_start(canvas, expand=True, fill=True, padding=0)

    def get_part_param(self, row, sheet):
        df_master = sheet.get_master()
        df_row = (df_master.iloc[row.get() - 1])
        name_part = df_row['Part Number']
        name_param = df_row['Parameter Name']

        return name_param, name_part

    def on_arrow_left_clicked(self, widget, row, sheet):
        if row.isMin():
            return

        row.dec()
        self.update_chart(row, sheet)

    def on_arrow_right_clicked(self, widget, row, sheet):
        if row.isMax():
            return

        row.inc()
        self.update_chart(row, sheet)

    def update_chart(self, row, sheet):
        for child in self.box.get_children():
            self.box.remove(child)
            child.destroy()

        self.info_master.select_row(row.get())
        name_param, name_part = self.get_part_param(row, sheet)
        self.set_hbar_title(name_param, name_part)
        self.gen_chart(name_param, name_part, sheet)
        self.show_all()

    def on_delete(self, widget, foo):
        self.__del__()

    # destructor
    def __del__(self):
        # print('DEBUG')
        self.info_master.deselect_row()
        self.close()
        self.destroy()


# =============================================================================
#  GLOBAL FUNCTIONS
# =============================================================================
# -----------------------------------------------------------------------------
#  make_trend_chart
# -----------------------------------------------------------------------------
def make_trend_chart(sheet, name_part, name_param):
    metrics = sheet.get_metrics(name_part, name_param)
    df = sheet.get_part(name_part)
    x = df['Sample']
    y = df[name_param]
    fig = Figure(dpi=100)
    splot = fig.add_subplot(111, title=name_param, ylabel='Value')
    splot.grid(True)
    if metrics['Spec Type'] == 'Two-Sided':
        if not np.isnan(metrics['USL']):
            splot.axhline(y=metrics['USL'], linewidth=1, color='blue', label='USL')
        if not np.isnan(metrics['UCL']):
            splot.axhline(y=metrics['UCL'], linewidth=1, color='red', label='UCL')
        if not np.isnan(metrics['Target']):
            splot.axhline(y=metrics['Target'], linewidth=1, color='purple', label='Target')
        if not np.isnan(metrics['LCL']):
            splot.axhline(y=metrics['LCL'], linewidth=1, color='red', label='LCL')
        if not np.isnan(metrics['LSL']):
            splot.axhline(y=metrics['LSL'], linewidth=1, color='blue', label='LSL')
    elif metrics['Spec Type'] == 'One-Sided':
        if not np.isnan(metrics['USL']):
            splot.axhline(y=metrics['USL'], linewidth=1, color='blue', label='USL')
        if not np.isnan(metrics['UCL']):
            splot.axhline(y=metrics['UCL'], linewidth=1, color='red', label='UCL')
    # Avg
    splot.axhline(y=metrics['Avg'], linewidth=1, color='green', label='Avg')
    # Line
    splot.plot(x, y, linewidth=1, color="gray")
    size_oos = 60
    size_ooc = 100
    if metrics['Spec Type'] == 'Two-Sided':
        # OOC check
        x_ooc = x[(df[name_param] < metrics['LCL']) | (df[name_param] > metrics['UCL'])]
        y_ooc = y[(df[name_param] < metrics['LCL']) | (df[name_param] > metrics['UCL'])]
        splot.scatter(x_ooc, y_ooc, s=size_ooc, c='orange', marker='o', label="Recent")
        # OOS check
        x_oos = x[(df[name_param] < metrics['LSL']) | (df[name_param] > metrics['USL'])]
        y_oos = y[(df[name_param] < metrics['LSL']) | (df[name_param] > metrics['USL'])]
        splot.scatter(x_oos, y_oos, s=size_oos, c='red', marker='o', label="Recent")
    elif metrics['Spec Type'] == 'One-Sided':
        # OOC check
        x_ooc = x[(df[name_param] > metrics['UCL'])]
        y_ooc = y[(df[name_param] > metrics['UCL'])]
        splot.scatter(x_ooc, y_ooc, s=size_ooc, c='orange', marker='o', label="Recent")
        # OOS check
        x_oos = x[(df[name_param] > metrics['USL'])]
        y_oos = y[(df[name_param] > metrics['USL'])]
        splot.scatter(x_oos, y_oos, s=size_oos, c='red', marker='o', label="Recent")
    splot.scatter(x, y, s=20, c='black', marker='o', label="Recent")
    x_label = splot.get_xlim()[1]
    if metrics['Spec Type'] == 'Two-Sided':
        if not np.isnan(metrics['USL']):
            splot.text(x_label, y=metrics['USL'], s=' USL', color='blue')
        if not np.isnan(metrics['UCL']):
            splot.text(x_label, y=metrics['UCL'], s=' UCL', color='red')
        if not np.isnan(metrics['Target']):
            splot.text(x_label, y=metrics['Target'], s=' Target', color='purple')
        if not np.isnan(metrics['LCL']):
            splot.text(x_label, y=metrics['LCL'], s=' LCL', color='red')
        if not np.isnan(metrics['LSL']):
            splot.text(x_label, y=metrics['LSL'], s=' LSL', color='blue')
    elif metrics['Spec Type'] == 'One-Sided':
        if not np.isnan(metrics['USL']):
            splot.text(x_label, y=metrics['USL'], s=' USL', color='blue')
        if not np.isnan(metrics['UCL']):
            splot.text(x_label, y=metrics['UCL'], s=' UCL', color='red')
    # Avg
    splot.text(x_label, y=metrics['Avg'], s=' Avg', color='green')

    canvas = FigureCanvas(fig)
    return canvas

# ---
# PROGRAM END
