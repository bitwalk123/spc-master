import pandas as pd
import numpy as np
from pptx import Presentation


class ExcelSPC():
    filename = None
    sheets = None
    valid = False

    # CONSTRUCTOR
    def __init__(self, filename):
        self.filename = filename
        self.sheets = self.read(filename)
        self.valid = self.check_valid_sheet(self.sheets)

    # -------------------------------------------------------------------------
    #  check_valid_sheet
    #  check if read file (sheets) has 'Master' tab
    #
    #  argument
    #    sheets : dataframe containing Excel contents
    #
    #  return
    #    True if dataframe is valid for SPC, otherwise False
    # -------------------------------------------------------------------------
    def check_valid_sheet(self, sheets):
        # check if 'Master' tab exists
        if 'Master' in sheets.keys():
            return True
        else:
            return False

    # -------------------------------------------------------------------------
    #  get_master
    #  get dataframe of 'Master' tab
    #
    #  argument
    #    (none)
    #
    #  return
    #    pandas dataframe of 'Master' tab
    # -------------------------------------------------------------------------
    def get_master(self):
        df = self.sheets['Master']
        # drop row if column 'Part Number' is NaN
        df = df.dropna(subset=['Part Number'])

        return df

    # -------------------------------------------------------------------------
    #  get_metrics
    # -------------------------------------------------------------------------
    def get_metrics(self, name_part, param):
        df = self.get_master()
        df1 = df[(df['Part Number'] == name_part) & (df['Parameter Name'] == param)]
        # print(df2)
        dict = {}
        dict['LSL'] = list(df1['LSL'])[0]
        dict['Target'] = list(df1['Target'])[0]
        dict['USL'] = list(df1['USL'])[0]
        dict['Chart Type'] = list(df1['Chart Type'])[0]
        dict['Metrology'] = list(df1['Metrology'])[0]
        dict['Multiple'] = list(df1['Multiple'])[0]
        dict['Spec Type'] = list(df1['Spec Type'])[0]
        dict['CL Frozen'] = list(df1['CL Frozen'])[0]
        dict['LCL'] = list(df1['LCL'])[0]
        dict['Avg'] = list(df1['Avg'])[0]
        dict['UCL'] = list(df1['UCL'])[0]

        return dict

    # -------------------------------------------------------------------------
    #  get_param_list
    #  get list of 'Parameter Name' of specified 'Part Number'
    #
    #  argument
    #    name_part : part name
    #
    #  return
    #    list of 'Parameter Name' of specified 'Part Number'
    # -------------------------------------------------------------------------
    def get_param_list(self, name_part):
        df = self.get_master()

        return list(df[df['Part Number'] == name_part]['Parameter Name'])

    # -------------------------------------------------------------------------
    #  get_part
    #  get dataframe of specified name_part tab
    #
    #  argument
    #    (none)
    #
    #  return
    #    pandas dataframe of specified name_part tab
    # -------------------------------------------------------------------------
    def get_part(self, name_part):
        # dataframe of specified name_part tab
        df = self.sheets[name_part]

        # delete row including NaN
        df = df.dropna(how='all')

        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_
        #  the first row od data sheet is used for 'Create Charts' button for
        #  the Excel macro
        #
        #  So, new dataframe is created for this application
        # _/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_

        # obtain number of rows on this dataframe
        row_size = len(df)

        # extract data rows
        df1 = df[1:row_size]

        # extract column name used for this dataframe
        list_colname = list(df.loc[0])
        df1.columns = list_colname

        # eliminate 'Hide' data
        df2 = df1[df1['Data Type'] != 'Hide']

        return df2

    # -------------------------------------------------------------------------
    #  get_sheets
    #  get dataframe containing Excel contents
    #
    #  argument
    #    (none)
    #
    #  return
    #    array of pandas dataframe containing Excel tab/data
    # -------------------------------------------------------------------------
    def get_sheets(self):
        return self.sheets

    # -------------------------------------------------------------------------
    #  get_unique_part_list
    #  get unique part list found in 'Part Number' column in 'Master' tab
    #
    #  argument
    #    (none)
    #
    #  return
    #    list of unique 'Part Number'
    # -------------------------------------------------------------------------
    def get_unique_part_list(self):
        df = self.get_master()
        list_part = list(np.unique(df['Part Number']))

        return list_part

    # -------------------------------------------------------------------------
    #  read
    #  read specified Excel file
    #
    #  argument
    #    filename : Excel file
    #
    #  return
    #    array of pandas dataframe including all Excel sheets
    # -------------------------------------------------------------------------
    def read(self, filename):
        # read specified filename as Excel file including all tabs
        return pd.read_excel(filename, sheet_name=None)

# =============================================================================
#  PowerPoint class
# =============================================================================
class PowerPoint():
    ppt = None

    def __init__(self, template):
        # insert empty slide
        self.ppt = Presentation(template)

    def add_slide(self, info):
        # ---------------------------------------------------------------------
        #  refer layout from original master
        # ---------------------------------------------------------------------
        slide_layout = self.ppt.slide_layouts[1]
        slide = self.ppt.slides.add_slide(slide_layout)

        # ---------------------------------------------------------------------
        #  slide title
        # ---------------------------------------------------------------------
        shapes = slide.shapes
        shapes.title.text = info['PART']

        # ---------------------------------------------------------------------
        # insert image
        # ---------------------------------------------------------------------
        slide.shapes.add_picture(
            info['IMAGE'], left=info['left'], top=info['top'], height=info['height']
        )

    # -------------------------------------------------------------------------
    #  save PowerPoint file
    # -------------------------------------------------------------------------
    def save(self, save_path):
        self.ppt.save(save_path)

# ---
# PROGRAM END
