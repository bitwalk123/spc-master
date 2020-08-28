# -----------------------------------------------------------------------------
#  dlg.py
#  dialog class for SDE Tool
# -----------------------------------------------------------------------------
import gi
import os
import pathlib

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk, Gdk

# -----------------------------------------------------------------------------
#  ok dialog
# -----------------------------------------------------------------------------
class ok(Gtk.Dialog):

    def __init__(self, parent, title, text, image):
        Gtk.Dialog.__init__(self, parent=parent, title=title)
        self.add_button(Gtk.STOCK_OK, Gtk.ResponseType.OK)
        self.set_default_size(200, 0)
        self.set_resizable(False)

        msg = Gtk.TextBuffer()
        msg.set_text(text)
        tview = Gtk.TextView()
        tview.set_wrap_mode(wrap_mode=Gtk.WrapMode.WORD)
        tview.set_buffer(msg)
        tview.set_editable(False)
        tview.set_can_focus(False)
        tview.set_top_margin(10)
        tview.set_bottom_margin(10)
        tview.set_left_margin(10)
        tview.set_right_margin(10)
        tview.override_background_color(
            Gtk.StateFlags.NORMAL,
            Gdk.RGBA(0, 0, 0, 0)
        )

        content = self.get_content_area()
        content.add(tview)

        self.show_all()



# -----------------------------------------------------------------------------
#  file chooser
# -----------------------------------------------------------------------------
class file_chooser():
    basedir = ''

    # -------------------------------------------------------------------------
    #  get
    #  get filename with dialog (class method)
    #
    #  argument
    #    cls : this class object for this class method
    # -------------------------------------------------------------------------
    @classmethod
    def get(cls, parent, flag='default'):
        dialog = Gtk.FileChooserDialog(
            title='Select File',
            parent=parent,
            action=Gtk.FileChooserAction.OPEN
        )
        dialog.add_buttons(
            Gtk.STOCK_CANCEL,
            Gtk.ResponseType.CANCEL,
            Gtk.STOCK_OPEN,
            Gtk.ResponseType.OK
        )

        if os.path.exists(cls.basedir):
            dialog.set_current_folder(str(cls.basedir))

        if flag == 'excel':
            cls.add_filters_excel(cls, dialog)
        else:
            cls.add_filters_all(cls, dialog)

        response = dialog.run()

        if response == Gtk.ResponseType.OK:
            p = pathlib.Path(dialog.get_filename())
            cls.basedir = os.path.dirname(p)
            dialog.destroy()
            # change path separator '\' to '/' to avoid unexpected errors
            name_file = str(p.as_posix())
            return name_file
        elif response == Gtk.ResponseType.CANCEL:
            dialog.destroy()
            return None

    # -------------------------------------------------------------------------
    #  filename_filter_all
    #  filter for ALL
    #
    #  argument
    #    dialog : instance of Gtk.FileChooserDialog to attach this file filter
    # -------------------------------------------------------------------------
    def add_filters_all(self, dialog):
        filter_any = Gtk.FileFilter()
        filter_any.set_name('All File')
        filter_any.add_pattern('*')
        dialog.add_filter(filter_any)

    # -------------------------------------------------------------------------
    #  File Open Filter for Excel
    # -------------------------------------------------------------------------
    def add_filters_excel(self, dialog):
        filter_xls = Gtk.FileFilter()
        filter_xls.set_name('Excel')
        filter_xls.add_pattern('*.xls')
        filter_xls.add_pattern('*.xlsx')
        filter_xls.add_pattern('*.xlsm')
        dialog.add_filter(filter_xls)

        filter_any = Gtk.FileFilter()
        filter_any.set_name('All types')
        filter_any.add_pattern('*')
        dialog.add_filter(filter_any)

# ---
# PROGRAM END
