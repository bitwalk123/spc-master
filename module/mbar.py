# -----------------------------------------------------------------------------
#  mbar.py --- widget class for SDE Tool
# -----------------------------------------------------------------------------
import gi

gi.require_version('Gtk', '3.0')
from gi.repository import Gtk


# =============================================================================
#  MenuBar
#  menubar class (template)
# =============================================================================
class MenuBar(Gtk.Frame):
    def __init__(self):
        Gtk.Frame.__init__(self)
        self.set_shadow_type(Gtk.ShadowType.ETCHED_IN)
        self.box = Gtk.Box()
        self.add(self.box)

    # -------------------------------------------------------------------------
    #  get_box
    #  get container instance for layouting widgets on it
    #
    #  argument:
    #    (none)
    #
    #  return
    #    Gtk.Box() layout instance
    # -------------------------------------------------------------------------
    def get_box(self):
        return self.box


# -----------------------------------------------------------------------------
#  menubar_button
#  button class for menubar class
# -----------------------------------------------------------------------------
class menubar_button(Gtk.Button):
    def __init__(self, icon_name):
        Gtk.Button.__init__(self)
        self.add(Gtk.Image.new_from_icon_name(icon_name, Gtk.IconSize.DND))


# =============================================================================
#  implementation
# =============================================================================

class main(MenuBar):
    def __init__(self):
        MenuBar.__init__(self)
        box = self.get_box()

        # excel button
        self.but_excel = menubar_button(icon_name='x-office-spreadsheet')
        box.pack_start(self.but_excel, expand=False, fill=True, padding=0)

        # powerpoint button
        self.but_ppt = menubar_button(icon_name='x-office-presentation')
        box.pack_start(self.but_ppt, expand=False, fill=True, padding=0)

        # exit button
        self.but_exit = menubar_button(icon_name='application-exit')
        box.pack_end(self.but_exit, expand=False, fill=True, padding=0)

        # info button
        self.but_info = menubar_button(icon_name='dialog-information')
        box.pack_end(self.but_info, expand=False, fill=True, padding=0)

    # -------------------------------------------------------------------------
    #  get_obj
    #  get object instance of button
    #
    #  argument:
    #    image : image name of button
    # -------------------------------------------------------------------------
    def get_obj(self, name_image):
        if name_image == 'excel':
            return self.but_excel
        if name_image == 'ppt':
            return self.but_ppt
        if name_image == 'info':
            return self.but_info
        if name_image == 'exit':
            return self.but_exit

# ---
# PROGRAM END
