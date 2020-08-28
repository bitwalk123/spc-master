import gi

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk

from module import mbar


class SPCMaster(Gtk.Window):
    def __init__(self):
        Gtk.Window.__init__(self, title="SPC Master")
        self.set_default_size(600, 400)

        # container
        box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self.add(box)

        ### menubar
        self.menubar = mbar.main()
        box.pack_start(self.menubar, expand=False, fill=True, padding=0)

        # exit button clicked event
        (self.menubar.get_obj('exit')).connect('clicked', self.on_click_app_exit)

    # -------------------------------------------------------------------------
    #  on_click_app_exit
    #  Exit Application, emitting 'destroy' signal
    #
    #  argument
    #    widget : clicked widget, automatically added from caller
    #
    #  return
    #    (none)
    # -------------------------------------------------------------------------
    def on_click_app_exit(self, widget):
        self.emit('destroy')


win = SPCMaster()
win.connect("destroy", Gtk.main_quit)
win.show_all()
Gtk.main()
