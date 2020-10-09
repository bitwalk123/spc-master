# =============================================================================
#  info_page - control selection on grid table
# =============================================================================
class info_page():
    grid = None

    def __init__(self, grid):
        self.grid = grid

    def get_columns(self):
        cols = 0
        for child in self.grid.get_children():
            x = self.grid.child_get_property(child, 'left-attach')
            width = self.grid.child_get_property(child, 'width')
            cols = max(cols, x + width)
        return cols

    def get_rows(self):
        rows = 0
        for child in self.grid.get_children():
            y = self.grid.child_get_property(child, 'top-attach')
            height = self.grid.child_get_property(child, 'height')
            rows = max(rows, y + height)
        return rows

    def select_row(self, row):
        for child in self.grid.get_children():
            x = self.grid.child_get_property(child, 'left-attach')
            y = self.grid.child_get_property(child, 'top-attach')
            if x > 0:
                if y == row:
                    if child.get_style_context().has_class("sheet"):
                        child.get_style_context().remove_class("sheet")
                    child.get_style_context().add_class("select")
                else:
                    if y > 0:
                        if child.get_style_context().has_class("select"):
                            child.get_style_context().remove_class("select")
                        child.get_style_context().add_class("sheet")

    def deselect_row(self):
        for child in self.grid.get_children():
            y = self.grid.child_get_property(child, 'top-attach')
            if y > 0:
                if child.get_style_context().has_class("select"):
                    child.get_style_context().remove_class("select")
                child.get_style_context().add_class("sheet")

    def hasChild(self):
        if len(self.grid.get_children()) > 0:
            return True
        else:
            return False

    def delChildren(self):
        for child in self.grid.get_children():
            self.grid.remove(child)
            child.destroy()



# =============================================================================
#  register
# =============================================================================
class register():
    value = None
    value_min = None
    value_max = None

    def __init__(self, value, value_min, value_max):
        self.value_min = value_min
        self.value_max = value_max
        self.set(value)

    def inc(self):
        self.value += 1
        if self.value > self.value_max:
            self.value = self.value_max

    def dec(self):
        self.value -= 1
        if self.value < self.value_min:
            self.value = self.value_min

    def get(self):
        return self.value

    def set(self, value):
        self.value = value
        if self.value > self.value_max:
            self.value = self.value_max
        if self.value < self.value_min:
            self.value = self.value_min

    def isMin(self):
        if self.value == self.value_min:
            return True
        else:
            return False

    def isMax(self):
        if self.value == self.value_max:
            return True
        else:
            return False


# ---
# PROGRAM END
