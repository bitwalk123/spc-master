import gi
import math

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk, Gdk

from module import excel, dlg, mbar, pcs, utils


# =============================================================================
#  SPCMaster - main class of this application
# =============================================================================
class SPCMaster(Gtk.Window):
    mainpanel = None
    info_master = None
    chart = None

    def __init__(self):
        Gtk.Window.__init__(self, title="SPC Master")
        self.set_icon_from_file("./img/logo.png")
        self.set_default_size(800, 600)

        # container
        box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.add(box)

        ### menubar
        self.menubar = mbar.main()
        box.pack_start(self.menubar, expand=False, fill=True, padding=0)

        # folder button clicked event
        (self.menubar.get_obj('excel')).connect('clicked', self.on_file_clicked)

        # exit button clicked event
        (self.menubar.get_obj('exit')).connect('clicked', self.on_click_app_exit)

        # main pabel
        self.mainpanel = Gtk.Notebook()
        self.mainpanel.set_tab_pos(Gtk.PositionType.BOTTOM)
        box.pack_start(self.mainpanel, expand=True, fill=True, padding=0)

        # master tab
        grid_master = Gtk.Grid()
        page_master = Gtk.ScrolledWindow()
        page_master.add(grid_master)
        page_master.set_policy(
            Gtk.PolicyType.AUTOMATIC,
            Gtk.PolicyType.AUTOMATIC
        )
        self.mainpanel.append_page(page_master, Gtk.Label(label="Master"))

        # create instance to store master page information
        self.info_master = utils.info_page(grid_master)

    # -------------------------------------------------------------------------
    #  calc
    #  Aggregation from Excel for SPC
    #
    #  argument
    #    filename : Excel file to read
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def calc(self, filename):
        sheets = excel.ExcelSPC(filename)

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        # check if read format is appropriate ot not
        if sheets.valid is not True:
            title = 'Error'
            text = 'Not appropriate format!'
            dialog = dlg.ok(self, title, text, 'dialog-error')
            dialog.run()
            dialog.destroy()

            # delete instance
            del sheets
            return

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        # create tabs for tables & charts
        # sheets.create_tabs(self.mainpanel)
        self.create_tabs(sheets)

        # update GUI
        self.show_all()

    # -------------------------------------------------------------------------
    #  create_panel_part
    #  creating 'Master' page
    #
    #  argument
    #    (none)
    #
    #  return
    #    instance of container
    # -------------------------------------------------------------------------
    def create_page_part(self, tabname):
        # DATA tab
        grid_data = Gtk.Grid()
        scrwin_data = Gtk.ScrolledWindow()
        scrwin_data.add(grid_data)
        scrwin_data.set_policy(
            Gtk.PolicyType.AUTOMATIC,
            Gtk.PolicyType.AUTOMATIC
        )
        self.mainpanel.append_page(scrwin_data, Gtk.Label(label=tabname))

        return grid_data

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
    def create_tabs(self, sheet):
        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  'Master' tab

        # create 'Master' tab
        self.create_tab_master(sheet)

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  PART tab

        # obtain unique part list
        list_part = sheet.get_unique_part_list()

        # create tab for etch part
        for name_part in list_part:
            # create initial tab for part
            grid_part_data = self.create_page_part(name_part)

            # get dataframe of part data
            df_part = sheet.get_part(name_part)

            # create tab to show part data
            self.create_tab_part_data(grid_part_data, df_part)

    # -------------------------------------------------------------------------
    #  create_tab_master
    #  creating 'Master' tab
    #
    #  argument
    #    sheet   : object of Excel sheet
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def create_tab_master(self, sheet):
        grid = self.info_master.grid
        df = sheet.get_master()

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  table header

        x = 0;  # column
        y = 0;  # row

        # first column
        widget = Gtk.Label(name='LabelHead', label='#')
        widget.set_hexpand(True)
        widget.get_style_context().add_class("header")
        grid.attach(widget, x, y, 1, 1)
        x += 1

        # rest of columns
        for item in df.columns.values:
            widget = Gtk.Label(name='LabelHead', label=item)
            widget.set_hexpand(True)
            widget.get_style_context().add_class("header")
            grid.attach(widget, x, y, 1, 1)
            x += 1

        y += 1

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  table contents
        for row in df.itertuples(name=None):
            x = 0
            for item in list(row):
                if (type(item) is float) or (type(item) is int):
                    # the first column '#' starts from 0,
                    # change to start from 1
                    if x == 0:
                        item += 1

                    # right align on the widget
                    xpos = 1.0
                    if math.isnan(item):
                        item = ''
                else:
                    # left align on the widget
                    xpos = 0.0

                item = str(item)

                if x == 0:
                    widget = Gtk.Button(label=item)
                    widget.connect('clicked', self.on_param_clicked, sheet)
                else:
                    widget = Gtk.Label(name='Label', label=item, xalign=xpos)
                    widget.set_hexpand(True)

                widget.get_style_context().add_class("sheet")
                grid.attach(widget, x, y, 1, 1)

                x += 1

            y += 1

    # -------------------------------------------------------------------------
    #  create_tab_part_data
    #  creating DATA tab in (Part Number) tab
    #
    #  argument
    #    grid : grid container where creating table
    #    df   : dataframe for specified (Part Number)
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def create_tab_part_data(self, grid, df):
        x = 0;  # column
        y = 0;  # row

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  table header

        # first column
        lab = Gtk.Label(name='LabelHead', label='#')
        lab.set_hexpand(True)
        lab.get_style_context().add_class("header");
        grid.attach(lab, x, y, 1, 1)
        x += 1

        # rest of columns
        for item in df.columns.values:
            lab = Gtk.Label(name='LabelHead', label=item)
            lab.set_hexpand(True)
            lab.get_style_context().add_class("header");
            grid.attach(lab, x, y, 1, 1)
            x += 1

        y += 1

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  table contents
        for row in df.itertuples(name=None):
            x = 0
            for item in list(row):
                if (type(item) is float) or (type(item) is int):
                    # right align on the widget
                    xpos = 1.0
                    if math.isnan(item):
                        item = ''
                else:
                    # left align on the widget
                    xpos = 0.0

                item = str(item)

                lab = Gtk.Label(name='Label', label=item, xalign=xpos)
                lab.set_hexpand(True)
                # lab.set_alignment(xalign=xpos, yalign=0.5)
                lab.get_style_context().add_class("sheet");
                grid.attach(lab, x, y, 1, 1)
                x += 1

            y += 1

    # -------------------------------------------------------------------------
    #  on_click_app_exit - Exit Application, emitting 'destroy' signal
    #
    #  argument
    #    widget : clicked widget, automatically added from caller
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def on_click_app_exit(self, widget):
        self.emit('destroy')

    # -------------------------------------------------------------------------
    #  on_file_clicked - read Exel file
    #
    #  argument
    #    widget : clicked widget, automatically added from caller
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def on_file_clicked(self, widget):
        filename = dlg.file_chooser.get(parent=self, flag='excel')
        if filename is not None:
            if self.info_master.hasChild():
                self.init_app()

            # read Excel file
            self.calc(filename)

    # -------------------------------------------------------------------------
    #  on_param_clicked - create plot for specified parameter
    #
    #  argument
    #    widget : clicked widget, automatically added from caller
    #    sheet  : object of Excel sheet
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def on_param_clicked(self, widget, sheet):
        self.init_chart()
        self.chart = pcs.ChartWin(self.info_master, widget, sheet)

    # -------------------------------------------------------------------------
    #  init_chart -- initiqalize chart object (to display new data)
    # -------------------------------------------------------------------------
    def init_chart(self):
        if self.chart is not None:
            self.chart.close()
            self.chart.destroy()
            self.chart = None

    # -------------------------------------------------------------------------
    #  init_app -- initiqalize application to read new data
    # -------------------------------------------------------------------------
    def init_app(self):
        # delete children of master grid
        self.info_master.delChildren()

        # initialize chart
        self.init_chart()

        # initialize mainpanel
        for child in self.mainpanel.get_children():
            name_page = self.mainpanel.get_tab_label_text(child)
            if name_page != 'Master':
                self.mainpanel.detach_tab(child)
                child.destroy()

        # update GUI
        self.show_all()


# =============================================================================
#  MAIN
# =============================================================================
# CSS handling
provider = Gtk.CssProvider()
provider.load_from_path('./spc-master.css')
Gtk.StyleContext.add_provider_for_screen(
    Gdk.Screen.get_default(),
    provider,
    Gtk.STYLE_PROVIDER_PRIORITY_APPLICATION
)
# main class
win = SPCMaster()
win.connect("destroy", Gtk.main_quit)
win.show_all()

# event loop
Gtk.main()

# ---
# PROGRAM END
