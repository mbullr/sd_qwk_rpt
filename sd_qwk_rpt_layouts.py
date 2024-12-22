'''
  guiBuilder_layout.py part of guiBuilder
  A minimalist freesimplegui builder
'''
import FreeSimpleGUI as sg
import sd_qwk_rpt_constants as gc


cell_type_list = [
    'Date',
    'Time',
    'Text',
    'Query',
    'Query Col',
    'Replace',
    'Lookup',
    'Image',
    'Workbook']

def cell_edit_window(col,row,edef_line):
    ''' generate the layout for cell edit pop upt'''
    # if previosly set, set list box default value
    if len(edef_line[gc.DF_DATA_TYPE]) == 0:
        data_type = ''
    else:
          defListv = edef_line[gc.DF_DATA_TYPE]
          if defListv == "R":
            if edef_line[gc.DF_DATA_PARAM] == '%%RUNDATE%%':
                data_type = 'Date'
            else:
                data_type = 'Time'

          elif defListv == "T":
            data_type = 'Text'
               
          elif defListv == "Q":
            data_type = 'Query'

          elif defListv == "QC":
            data_type = 'Query Col'

          elif defListv == "R":
            data_type = 'Replace'

          elif defListv == "L":
            data_type = 'Lookup'

          elif defListv == "I":
            data_type == 'Image'

          elif defListv == "W":
            data_type == 'Workbook'


            
    sg.set_options(font=('Consolas', 14))
    cell_edit_layout = [
    [sg.Text('Row: '+str(row)),sg.Text('Col: '+str(col))],
    [sg.Text('Cell Name '),sg.Input(default_text=edef_line[gc.DF_ITEM_NAME], size=(20,1),key= '-CELL_NAME-', enable_events=True )],
   
 #   [sg.Text('Data Type '),sg.Input(size=(20,1),key= '-DATA_TYPE-' )],
    [sg.Text('Data Type '),sg.Listbox(default_values= [data_type], values=cell_type_list,size=(20,4), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_DATA_TYPE-')],
    [sg.Button('Font',size=(12,1))],
    [sg.Button('Alignment',size=(12,1))],
    [sg.Button('Border',size=(12,1))],
    [sg.Button('Fill',size=(12,1))],
    [sg.Button('Save Cell'),sg.Button('Delete Cell')]
    ]

    return cell_edit_layout

def edit_font(edef_line):
    font_uln_type = ['none','single','double','singleAccounting','doubleAccounting']
    font_list = sg.Text.fonts_installed_list()
    font_layout = [
        [sg.Text('Font Name ',size = (10,1)),sg.Listbox(default_values= [edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_NAME]],values=font_list,size=(20,5), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_FONT_NAME-')],
        [sg.Text('Size',size = (10,1)), sg.Input(default_text=edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_SZ], size = (15,1), key = '-FONT_SZ-')],
        [sg.Checkbox('Bold', default=edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_BOLD], key = '-FONT_BOLD-'), sg.Checkbox('Italic',default=edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_ITLC], key = '-FONT_ITLC-'), sg.Checkbox('Strike',default=edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_STKE], key = '-FONT_STKE-')],
        [sg.Text('Underline ',size = (10,1)),sg.Listbox(default_values= [edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_UNLN]],values=font_uln_type,size=(20,3), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_FONT_UNLN-')],
        [sg.Text('Font Color',size = (10,1)),sg.Input(default_text=edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_COLOR],enable_events=True, key='-FONT_COLOR-', size=(15,1)),sg.ColorChooserButton("Font Color")],
 
        [sg.Button('Save')]
 
    ]
    return font_layout

def edit_border(edef_line):
    bdr_type = ['dashDot', 
    'dashDotDot', 
    'dashed', 
    'dotted', 
    'double', 
    'hair',
    'medium', 
    'mediumDashDot', 
    'mediumDashDotDot', 
    'mediumDashed', 
    'slantDashDot', 
    'thick', 
    'thin']

    bdr_layout = [
        
        [sg.Checkbox('Left', default=edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_LEFT], key = '-LEFT_BDR-'), sg.Checkbox('Right',default=edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_RIGHT], key = '-RIGHT_BDR-'), sg.Checkbox('Top',default=edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_TOP], key = '-TOP_BDR-'),sg.Checkbox('Bottom',default=edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_BOTTOM], key = '-BOTTOM_BDR-')],
        [sg.Text('Style ',size = (10,1)),sg.Listbox(default_values= [edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_STYLE]],values=bdr_type,size=(20,3), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_BDR_STYLE-')],
        [sg.Text('Border Color',size = (10,1)),sg.Input(default_text=edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_COLOR],enable_events=True, key='-BDR_COLOR-', size=(15,1)),sg.ColorChooserButton("Border Color")],
 
        [sg.Button('Save')]
 
    ]
    return bdr_layout

def edit_align(edef_line):
    align_horz = ['fill', 'left', 'distributed', 'justify', 'center', 'general', 'centerContinuous', 'right']
    align_vert = ['distributed', 'justify', 'center', 'bottom', 'top']
    align_layout = [
        
        [sg.Text('Align Horz ',size = (10,1)),sg.Listbox(default_values= [edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_HORZ]],values=align_horz,size=(20,3), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_ALIGN_HORZ-')],
        [sg.Text('Align Vert ',size = (10,1)),sg.Listbox(default_values= [edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_VERT]],values=align_vert,size=(20,3), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_ALIGN_VERT-')],
        [sg.Checkbox('Wrap Text', default=edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_WRAP], key = '-WRAP_TEXT-')],
        [sg.Checkbox('Shrink To Fit', default=edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_SHRK], key = '-SHRINK_TEXT-')],
        [sg.Button('Save')]
 
    ]
    return align_layout

def edit_fill(edef_line):
    fill_type = ['darkGray', 'darkUp', 'lightDown', 'darkGrid', 'darkHorizontal', 'lightTrellis', 'lightVertical', 'gray0625', 'gray125', 'lightGray', 'lightUp', 'darkDown', 'darkTrellis', 'lightGrid', 'mediumGray', 'solid', 'darkVertical', 'lightHorizontal']
    fill_layout = [
        
        [sg.Text('Fill Type ',size = (10,1)),sg.Listbox(default_values= [edef_line[gc.DF_CELL_FILL][gc.DF_FILL_TYPE]],values=fill_type,size=(20,3), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_FILL_TYPE-')],
        [sg.Text('Fill Color',size = (10,1)),sg.Input(default_text=edef_line[gc.DF_CELL_FILL][gc.DF_FILL_BGCLR],enable_events=True, key='-FILL_COLOR-', size=(15,1)),sg.ColorChooserButton("Fill Color")],
 
        [sg.Button('Save')]
 
    ]
    return fill_layout



def query(edef_line,qfilename,dict_list):
    # look at MVQuery
    # suggest use dict_list in mult lists: sort by, selection(s) with (need to add n/a to list to signal not used)
    # will create popup list for add field (genreated outside of this routine??)
    sort_by = ['BY','BY.DSND','BY.EXP','BY.EXP.DSND']
    with_logic = ['EQ','NE','LT','LE','GT','GE','LIKE','UNLIKE']
    sel_logic  = ['AND','OR']
    col_values= []
    query_layout = [
        [sg.Text('File: ' + qfilename)],
        [sg.Text('Sort By Clause(s) ',size = (15,1)),sg.Listbox(values=sort_by,default_values= [sort_by[0]], size=(15,1), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SORT_BY_LOGIC-'),
         sg.Listbox(values=dict_list,size=(25,1), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SORT_BY_DICT-')],
        [sg.Text('Selection(s)')],
        [sg.Text('WITH'),sg.Listbox(values=dict_list,size=(25,1), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_DICT1-'),
         sg.Listbox(values=with_logic, default_values=[with_logic[0]], size=(6,1), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_LOGIC1-'),
         sg.Input(enable_events=True, key='-SEL_VALUE1-', size=(25,1))],
        [sg.Listbox(values=sel_logic, default_values=[sel_logic[0]], enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_AND_OR-')],
        [sg.Text('WITH'),sg.Listbox(values=dict_list,size=(25,1), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_DICT2-'),
         sg.Listbox(values=with_logic, default_values=[with_logic[0]], size=(6,1), enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_LOGIC2-'),
         sg.Input(enable_events=True, key='-SEL_VALUE2-', size=(25,1))],
        [sg.Text('Columns'),sg.Table(col_values,headings=['Column_Value'],  max_col_width=25,pad=0,
                        auto_size_columns=False,
                        text_color='black',
                        background_color='White',
                        font=('Consolas', 14),
                        def_col_width=25,
                        # cols_justification=('left','center','right','c', 'l', 'bad'),       # Added on GitHub only as of June 2022
                        display_row_numbers=False,
                        starting_row_number= 1,
                        justification='left',
                        num_rows=5,
                        #alternating_row_color= gc.CELL_TABLE_COLOR,
                        key='-COL_TABLE-',
                        selected_row_colors=gc.SELECT_COLOR,
                        enable_events=True,
                        expand_x=False,
                        expand_y=False,
                        vertical_scroll_only=False,
                        hide_vertical_scroll = False,
                        enable_click_events=True,           # Comment out to not enable header and other clicks
                        tooltip=None),
                          sg.Button('Add Column'),
                          sg.Listbox(values=dict_list,size=(25,7),enable_events=True,select_mode = 'LISTBOX_SELECT_MODE_SINGLE',  key ='-SEL_ADD_COL-') ], 
        
        [sg.Checkbox('Auto Create Columns', default=True, key = '-AUTO_COLUMNS-'), sg.Checkbox('Auto Create Headings', default=True, key = '-AUTO_HEADINGS-') ],
        [sg.Button('Save')]
 
    ]
    return query_layout

def data_line():
    data_layout = [
        [sg.Text('Static Data / Text',size = (10,1)), sg.Input(size = (15,1), key = '-DATA_PARAM-')],

        [sg.Button('Save')]
 
    ]
    return data_layout

def sd_dsn_window(table_col_count, cell_values):

    ### prebuild our layout table
    tableheadings = [f'{chr(col)}' for col in range(65,(65+table_col_count))]

    
    # arrow images from the Tango Desktop Project http://tango.freedesktop.org/Tango_Desktop_Project
    down_image = b'iVBORw0KGgoAAAANSUhEUgAAABYAAAAWCAYAAADEtGw7AAAABHNCSVQICAgIfAhkiAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAL8SURBVDiNtZRPbBRVHMc/781Md1bTGMCkhYNtOIiAnkw0bSFpNCGeOEAwMRyMMaReNCY0JXpFozQWT3JWExsEIzaxJkSREGiJIglRNCqVDSEIxlKqsN3ZeX9+HnZ2O112Wznwkl/e78283+f33nd+v1Eiwv0Y+r5QgbDdi4GR6AaermWjNX9NjZruewLj6Xp76COsNzhvsT7FeYvzBusNqUs4dPittonbgwFQnP7jCFW7QNUukLqk4T+3cc8Kl7nHoVD/a98KYEFYrBqR2lrEs1I1LQtehHgEX3signHVFU/c0HjLSHRSPIONjKFKRHws4vF4vNQsdRUSW27sGRiOFo+uOTU1agaXgL2TsZ7uR5/es/3NYhhEOLGx8xYvFrRHh4L3CQm3iAoaHcAru96IRcCYKhMnxhdmb/15sM5Tea0GhqO9m9Y/uX/H4MvFC9e+oWLnsVSp+jvMVa+R2H9BAQIiEOtOtvbs5uTUVwuXr8/snx4177bUeOo9M/brlR/Gvz13LNm8rp8gVBhV5u/qDFbfIYo1HbEmijXFYszW9bs4f/FscuXGpYk89C4wwNqSG5r+5evvf7z0ndm4tp/ZtITq8EQZMIoDOuKAvp4dlK5etj/9fu7nyrx7sZmjWpXNln2qU6nowvZnX+gNH6roi/PHCUKNDhVKweOrtlGZ0zIxefS6M/aJ6YMy18xoWW5nDshtbczg5Imj/xTcKjas6SfsUEQFzYaH+yi61Xx5/IvbQvBMK2hbMMCpMblqU7vt2OSn5e5oE+sefKxmhc0c/ny8XE3NzjMHkt/axbeUIj8GRsKdqzvXfLz7+ZceCCLFJ599WL45e3Pf6dH0g+XiloCVUgoIcqYB3fdasPeR3t7hoKBUaaZ05Oz77nXA58wBTkTsEnAGjKg1TJjzg2zWT70avBMEuuv8ITOUppgMaDOoBUw2W8DUwVEO2Ow3r00OYtqs03pL16/Q/E+sX1NnlvVdQ4LmE6dAKiL+ro/XTufMryerw10OLpKD/QceAIZcIO5IFQAAAABJRU5ErkJggg=='
    up_image = b'iVBORw0KGgoAAAANSUhEUgAAABYAAAAWCAYAAADEtGw7AAAABHNCSVQICAgIfAhkiAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAL/SURBVDiNtZRNaFxVFMd/5368eZkxRmQsbTEGxbZC0IWMpU5SbCmxILapFCO2LgSlWbhLbFJBcKErP0BcCIqCdeEiUAviQoJIwIgW3biwCoYGjTImTIXWTjsz7+O4yLzwUkuSgjlwuPdd7vmd/z33vCuqymaY2RTqZoLdRjfunXTvg2zfNh8PT01p8r+AByf9S9vv6DuGQWr627vA6Hox65ZiYMIdva1Ufvlg9WhpaM+RYs+t5eN7J4IX1ouTtbpicCyoFEulmacOniidr8+SpBG7ynv4dPrjxpVW48i3b0Rf3rTiR8al13o3/cSBZ0uLV+f46/IFapfmWLh8nsf2PVnyzp0ZnAx33RR4cFK6U2dnDu97pie1LS5c/BFnHNZ4fl36nkZSZ6h6qBuNv6qOye0bAo+MiFXc54/uHundUt5qfln6Dmt8xx2I8MPCF5RuKcrAg0NbjHfTlVHx64Jrd9v3qv1Du/t3VPxPta+xYvHOUwgCfGBQ28Z45dzCWfruvMfdf+9D/V099vSa4IEX/fh9d1WO7a8Mhz/XZjFOCMMCYVggCD2RaeAKQlAw4CPO/XmGygMPh33bdg5XJ/ypPGulK6rj7vG+rTumnjt0qss5j5KikjAzf5pYWtSvzXMt+Wc5SuHwzpNoqmgKzXaLs9OfXK1f/OPp2bfizyD3gxgr478vznW98sHzAFgvzVdPfBj6wNFoLxLZBt4bBFAF44R3PnqtGbc17CCKOBkDVoNnX4/2X1cWNdbiA8+V1hK+YDAGEEFVsVaI2xp+82Yk19d3FfhGZozQiP9GbYwLBOMEEUFTRewNeRsEi6WZXkIMGCs4bzBOiNspsjZ3dVfIsjkRKWRrzbgBCppCmihJpGgC6EpMUUQKIrJKpMuAgO98u84cVTjQO4pKirEgIizfHqTpyhvTDcRAJCJxNne5BDYHdUD95NvHy2sfmHomImcKJKKqmeI8NBuzZKbjHb2kHY+BJFMJtIG2qqb/eTY7SWzOM6jtbElz8CQHV83B/gXtSQriGSyg6AAAAABJRU5ErkJggg=='
   
    w_arrows = [[sg.Button('',image_data=up_image,key='-CELL_UP-',pad = 0)],[sg.Button('',image_data=down_image,key='-CELL_DOWN-',pad = 0)]]
    cell_table = [
        [sg.Column(w_arrows,pad=0),
        sg.Table(cell_values, headings=tableheadings, max_col_width=25,pad=0,
                        auto_size_columns=False,
                        font=('Consolas', 14),
                        text_color = 'black',
                        #col_widths = 25,
                        def_col_width = 15, 
                        # cols_justification=('left','center','right','c', 'l', 'bad'),       # Added on GitHub only as of June 2022
                        display_row_numbers=True,
                        starting_row_number= 1,
                        justification='center',
                        num_rows=20,
                        alternating_row_color= gc.CELL_TABLE_COLOR,
                        key='-CELL_TABLE-',
                        selected_row_colors=gc.SELECT_COLOR,
                        enable_events=True,
                        expand_x=False,
                        expand_y=False,
                        vertical_scroll_only=False,
                        hide_vertical_scroll = False,
                        enable_click_events=True,           # Comment out to not enable header and other clicks
                        tooltip=None)]
        ]
    #
    # main menu 
    menu_layout = [['File', ['Load', 'Save', '---', 'Exit']],
                ['Help', ['About']]]
    
    # guibuilder main form layout
    layout = [
    [sg.Menu(menu_layout)],
    [cell_table],
    [sg.Text('Folder'), sg.In(size=(25,1), enable_events=True ,key='-FOLDER-'), sg.FolderBrowse(),
     sg.Text('Report File Name'),sg.Input(size=(25,1),key= '-REPORT_FNAME-'),
     sg.Text('Description'),sg.Input(size=(55,1),key= '-REPORT_DESC-'),
     sg.Button(button_text='Write Definition', key='WRTIE_DEF')],
    [sg.Text('Cell Definition Data'),sg.Input(size=(80,1),key= '-DEF_LINE-' )]
    ]
    return layout