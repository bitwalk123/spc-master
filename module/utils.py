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
            if  x > 0:
                if y == int(row):
                    if child.get_style_context().has_class("sheet"):
                        child.get_style_context().remove_class("sheet")
                    child.get_style_context().add_class("select")
                else:
                    if y > 0:
                        if child.get_style_context().has_class("select"):
                            child.get_style_context().remove_class("select")
                        child.get_style_context().add_class("sheet")
