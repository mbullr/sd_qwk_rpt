# sd_qwk_rpt_dsgn.py
#
# SD Quick Report is a simple spreadsheet based report design and creation package.
# The original intent was to test SD's embedded python and play with FreeSimpleGUI as a basis for GUI's with SD.
# This script is used to create the report definition file item that is then used by the SD program sd_qwk_rpt to generate a spreadsheet report.
# sd_qwk_rpt either creates a complete python script file to be executed outside of sd (via os.execute) 
#  or spoon feeds the script statements to the python interpreter via SD's embedded python feature.
# The python package openpyxl is used to create the spreadsheet file. 
# helpful links:
# https://openpyxl.readthedocs.io/en/stable/index.html
# https://openpyxl.readthedocs.io/en/stable/styles.html
#
# This script's user interface is not very polished, and somewhat clunky, remember it was delevloped as a learning experience.
#
# Copyright (c)2024 The SD Developers, All Rights Reserved
#
# This program is free software; you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation; either version 3, or (at your option)
# any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program; if not, write to the Free Software Foundation,
# Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.
#
# 
# to do list:
#
# add page format
# add number formating (or default to directory info and add in sd_qwk_rpt??) 
# add Image
# add some documentation (lol)
#
#
## note!! tkinter issue 46180 Button clicked failed when mouse hover tooltip and tooltip destroyed
## work around for now is to remove all tool tips from sg.button calls
# 
#
#   Arrow images used for table row move buttons from the Tango Desktop Project http://tango.freedesktop.org/Tango_Desktop_Project
#   Demo_Base64_Single_Image_Encoder.py  used to convert into Base64
#

import argparse
import FreeSimpleGUI as sg
import os
import json
import subprocess
import uuid
import copy

import sd_qwk_rpt_layouts as gl
import sd_qwk_rpt_constants as gc

sg.change_look_and_feel('BrownBlue') # change style

# all important field separators
SDME_FM  = chr(254)
SDME_VM  = chr(253)
SDME_SVM = chr(252)

# def for SD vs QM vs SCARLETDME
SD_FILE_BOLB = '%0'
SD_EXE = '/usr/local/sdsys/bin/sd'

## global var (i know, i shouldn't)
window = None   # the all important freesimplegui window created in main()

# note column sizing is kind of a copout, still cannot get table to update size when additional columns are added.  
# these can now be set by caller with simple_rpt -r rowcount -c col count
table_col_count = 15  # number of columns to create for layout table
table_row_count = 25  # and rows


cellvalues = []       # the list which will hold the cell value names
defdict = {}          # dictionary of cell definitions
# key - name  #DF_ITEM_NAME	- Item Name / Ref
# value  -  list of:
#
#	DF_CELL_REF	- Cell Ref   
#	DF_DATA_TYPE	- Data Type 
#	DF_FONT_STYLE	- Cell Font (data sub value mark separated)
#		DF_FONT_NAME  
#		DF_FONT_SZ    
#     	DF_FONT_BOLD  
#		DF_FONT_ITLC  
#     	DF_FONT_UNLN  
#     	DF_FONT_STKE
#		DF_FONT_COLOR 
#	DF_ALIGNMENT	- Cell Alignment (data sub value mark separated)
#		DF_ALIGN_HORZ 
#     	DF_ALIGN_VERT 
#     	DF_ALIGN_ROT  
#		DF_ALIGN_WRAP 
#		DF_ALIGN_SHRK 
#     	DF_ALIGN_IDNT 
#	DF_CELL_BORDR	- Cell Border (data sub value mark separated) 
#		DF_BORDR_STYLE 
#		DF_BORDR_COLOR 
#	DF_CELL_FILL 	- Cell Fill (data sub value mark separated)
#		DF_FILL_PATTRN 
#		DF_FILL_FGCLR  
#		DF_FILL_BGCLR 
#		DF_FILL_TYPE   	
#	DF_DATA_PARAM – Actual Data Value(s) 

fpath = '' # file path to save report definition to
selected_row = 0
selected_col = 0
################# support functions #############################################
def int_or_float(val):
    # returns integer or float value of string, if not able to convert, None
    try:
        # all int like strings can be converted to float so int tries first    
        return int(val)
    except ValueError:
        pass
    try:
        return float(val)
    except:
        return None
################# table functions ###############################################
def refresh_table():
    window['-CELL_TABLE-'].update(values = cellvalues)

def move_row_up(t_key,t_values,t_row):
    '''move selected row - t_row for t_key / t_values - table up one row'''
    if t_row == 0:
        pass  # already at the top!
    else:
        temp_row = t_values[t_row - 1]
        t_values[t_row-1] = t_values[t_row]
        t_values[t_row] = temp_row
        window[t_key].update(values = t_values)

def move_row_down(t_key,t_values,t_row):
    '''move selected row - t_row for t_key / t_values - table down one row'''
    if t_row == len(t_values)-1:
        pass  # already at the bottom!
    else:
        temp_row = t_values[t_row + 1]
        t_values[t_row+1] = t_values[t_row]
        t_values[t_row] = temp_row
        window[t_key].update(values = t_values)     

########################################## editor windows        ##################################
def editor_dialog(ecol,erow,edef_line):
    '''main cell def editor dialog / code'''
    rtn_action = 'CANCEL'
    layout = gl.cell_edit_window(ecol,erow,edef_line)
    edef_line[gc.DF_CELL_REF] = [ecol, erow]  # when building dsn we save col and row index, change to letter, row when writing def file
    ewindow= sg.Window('Cell Edit',layout,modal=True)
    while True:
    #   windowid, event, values = sg.read_all_windows()
        event, values = ewindow.read()
    #   print(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':
            rtn_action = 'CANCEL'
            break
        
        elif event == '-CELL_NAME-':
            cell_name = values['-CELL_NAME-']
            if edef_line[gc.DF_ITEM_NAME]  == cell_name:
                # name in text field same as what we entered with, ok
                pass
            elif cell_name in defdict:
                sg.popup('Cell Edit', 'Cell Name Must Be Unique')
            else:
                edef_line[gc.DF_ITEM_NAME]  = cell_name

        elif event == '-SEL_DATA_TYPE-':
            data_type = values['-SEL_DATA_TYPE-'][0]
            if data_type == 'Date':
                edef_line[gc.DF_DATA_TYPE] = "R"
                edef_line[gc.DF_DATA_PARAM] = '%%RUNDATE%%'

            elif data_type == 'Time':
                edef_line[gc.DF_DATA_TYPE] = "R"
                edef_line[gc.DF_DATA_PARAM] = '%%RUNTIME%%' 

            elif data_type == 'Text':
                get_param(edef_line)
                edef_line[gc.DF_DATA_TYPE] = "T"
               
            elif data_type == 'Query':
                query_dialog(edef_line)
                edef_line[gc.DF_DATA_TYPE] = "Q" 

            elif data_type == 'Query Col':
                edef_line[gc.DF_DATA_TYPE] = "QC" 

            elif data_type == 'Replace':
                edef_line[gc.DF_DATA_TYPE] = "R"

            elif data_type == 'Lookup':
                edef_line[gc.DF_DATA_TYPE] = "L" 

            elif data_type == 'Image':
                edef_line[gc.DF_DATA_TYPE] = "I"

            elif data_type == 'Workbook':
                edef_line[gc.DF_DATA_TYPE] = "W"      
            
            set_cell_name(edef_line,data_type)
            if edef_line[gc.DF_ITEM_NAME] != '':
                ewindow['-CELL_NAME-'].update(value=edef_line[gc.DF_ITEM_NAME])

        elif event == 'Font':
            edit_font(edef_line)

        elif event == 'Alignment':
            edit_align(edef_line) 

        elif event == 'Border':
            edit_border(edef_line)


        elif event == 'Fill':
            edit_fill(edef_line) 

        elif event == 'Save Cell':
            if edef_line[gc.DF_ITEM_NAME] == '':
                sg.popup('Cell Edit', 'Cell Name Required') 
               
            else: 
                rtn_action = 'SAVE' 
                break 

        elif event == 'Delete Cell':
            rtn_action = 'DELETE' 
            break      

        window['-DEF_LINE-'].update(value = str(edef_line))   
    ewindow.close()    
    return rtn_action

def set_cell_name(edef_line,data_type):
    '''set default cell name'''
    cell_name = ''
    if edef_line[gc.DF_ITEM_NAME] == '':

        if data_type == 'Date':
            cell_name = 'Date'

        elif data_type == 'Time':
            cell_name = 'Time' 

        elif data_type == 'Text':
            cell_name = edef_line[gc.DF_DATA_PARAM][:15]
               
        elif data_type == 'Query':
            cell_name = 'Query'

        elif data_type == 'Query Col':
            cell_name = 'Query Col'

        elif data_type == 'Replace':
            cell_name = 'Replace'

        elif data_type == 'Lookup':
            cell_name = 'Lookup'

        elif data_type == 'Image':
            cell_name = 'Image'

        elif data_type == 'Workbook':
            cell_name == 'Workbook' 

        # make cell_name unique
        name_cnt = 0
        test_name = cell_name
        while True:
            if test_name in defdict:
                name_cnt += 1
                test_name = cell_name + str(name_cnt)
            else:
                break 
        edef_line[gc.DF_ITEM_NAME] = test_name

    return

def get_param(edef_line):
    layout = gl.data_line()
    ewindow= sg.Window('Static Data / Text',layout,modal=True)
    while True:
    #   windowid, event, values = sg.read_all_windows()
        event, values = ewindow.read()
    #   print(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Save':    
            if values['-DATA_PARAM-']:
                edef_line[gc.DF_DATA_PARAM] = values['-DATA_PARAM-']
                
            break
    ewindow.close()    
    return

def query_dialog(edef_line):
    '''create a query cell definition'''
    qfilename = get_sd_file()
    if qfilename == '':
        sg.popup('Query', 'No File Selected')
        return
    sdfilename =  os.path.split(qfilename) 
    dict_list, hdr_dict = get_sd_file_dict(sdfilename[0],sdfilename[1])
    by_clause = ''
    by_clause_dict = ''
    with_dict1 = ''
    with_logic1 = ''
    with_val1 = ''
    sel_logic = ''
    with_dict2 = ''
    with_logic2 = ''
    with_val2 = ''
    rpt_cols = ''

    layout = gl.query(edef_line,qfilename,dict_list)
    ewindow= sg.Window('Query',layout,modal=True)
    cols_list = []
    while True:
    #   windowid, event, values = sg.read_all_windows()
        event, values = ewindow.read()
        #print(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':
            break

        if event == 'Add Column':
            if values['-SEL_ADD_COL-']:
                rpt_cols = rpt_cols + ' ' + values['-SEL_ADD_COL-'][0]
                cols_list.append(values['-SEL_ADD_COL-'][0])
                ewindow['-COL_TABLE-'].update(values = cols_list)

        if event == 'Save':
            if values['-SORT_BY_DICT-']:
                by_clause_dict = values['-SORT_BY_DICT-'][0]
                by_clause = values['-SORT_BY_LOGIC-'][0]

            if values['-SEL_DICT1-']:
                with_dict1 = values['-SEL_DICT1-'][0]
                with_logic1 =  values['-SEL_LOGIC1-'][0]
                with_val1 = values['-SEL_VALUE1-']  

            if values['-SEL_DICT2-']:
                sel_logic = values['-SEL_AND_OR-'][0]
                with_dict2 = values['-SEL_DICT2-'][0]
                with_logic2 =  values['-SEL_LOGIC2-'][0]
                with_val2 = values['-SEL_VALUE2-']   

#def_rec<12,gc.DF_DATA_PARAM> = '"CUSTOMERS","BY CM_NAME","WITH CM_STATUS =' : "'A'" : '"' : ',"CM_NAME CM_ADDR CM_CITY CM_STATE CM_ZIP CM_STATUS"'
            query_text = '"' + sdfilename[1] +'","' + by_clause + ' ' + by_clause_dict + '",'
            if  with_val1 != '':
                with_val1 = '"WITH ' + with_dict1 + ' ' + with_logic1 + """ '""" + with_val1 + """'"""
            else:
                with_val1 = '"'

            if  with_val2 != '':
                with_val2 = ' ' + sel_logic +' WITH ' + with_dict2 + ' ' + with_logic2 + """ '""" + with_val2 + """'"""

                query_text = query_text + with_val1 + with_val2 +  '",'
            else:
                query_text = query_text + with_val1 +  '",'

            edef_line[gc.DF_DATA_PARAM] = query_text + '"' + rpt_cols + '"'

            if values['-AUTO_COLUMNS-']:
                auto_create_query_cols(edef_line,cols_list)
                if values['-AUTO_HEADINGS-']:
                    auto_create_query_hdrs(edef_line,cols_list,hdr_dict)
            break
        
    ewindow.close()    
    return

def auto_create_query_cols(edef_line,rpt_cols):
    '''auto create query column cell defs'''
    global cellvalues, defdict
    # first check for empty cells to right
    table_col_count = len(cellvalues[0])
    col_idx = int(edef_line[gc.DF_CELL_REF][0])
    cols_to_create = len(rpt_cols)
    row_idx = int(edef_line[gc.DF_CELL_REF][1])
    # do we have room to auto create columns?
    if cols_to_create + col_idx > table_col_count:
        sg.popup('Query', 'Report Columns Will Exceed Table, Cannot Auto Create')
        return
    # are there cells already defined in the columns we are thinking about using?
    for col_tst in range(col_idx+1,table_col_count):
        if cellvalues[row_idx][col_tst] != gc.EMPTYCELL:
            sg.popup('Query', 'Report Column: '+ str(col_tst) + ' Not Empty, Cannot Auto Create')
            return
    # how many query's are defined in this report
    # should make column ids unique
    q_cnt = 1
    for def_item in defdict:
        if defdict[def_item][gc.DF_DATA_TYPE] == 'Q':
            q_cnt +=1    
    # looks good, create the columns
    for i in range(1, cols_to_create):
        new_col_idx =  col_idx + i
        new_col_id =  'Q' + str(q_cnt) + '_' + str(rpt_cols[i])  

        if new_col_id in defdict:
            sg.popup('Query', 'Duplicate Column Id: ' + new_col_id + ' Cannot Complete Auto Create')
            return    
        
        cellvalues[row_idx][new_col_idx] = new_col_id
        new_def = copy.deepcopy(edef_line)
        new_def[gc.DF_ITEM_NAME] = new_col_id
        new_def[gc.DF_CELL_REF] = [new_col_idx,row_idx]
        new_def[gc.DF_DATA_TYPE] = 'QC'
        new_def[gc.DF_DATA_PARAM] = ''
        defdict[new_col_id] = new_def

    return

def auto_create_query_hdrs(edef_line,rpt_cols,hdr_dict):
    '''auto create query column headings cell defs'''
    global cellvalues, defdict
    # first check for empty cells to right
    table_col_count = len(cellvalues[0])
    col_idx = int(edef_line[gc.DF_CELL_REF][0])
    cols_to_create = len(rpt_cols)
    row_idx = int(edef_line[gc.DF_CELL_REF][1]) - 1  # rem headers go above query
    if row_idx < 0:
        sg.popup('Query', 'No open row for Report Headings, Cannot Auto Create')
        return        
    # do we have room to auto create columns?
    if cols_to_create + col_idx > table_col_count:
        sg.popup('Query', 'Report Headings Will Exceed Table, Cannot Auto Create')
        return
    # are there cells already defined in the columns we are thinking about using?
    for col_tst in range(col_idx+1,table_col_count):
        if cellvalues[row_idx][col_tst] != gc.EMPTYCELL:
            sg.popup('Query', 'Heading Column: '+ str(col_tst) + ' Not Empty, Cannot Auto Create')
            return
    
    # how many query's are defined in this report
    # should make column ids unique
    q_cnt = 1
    for def_item in defdict:
        if defdict[def_item][gc.DF_DATA_TYPE] == 'Q':
            q_cnt +=1

    # looks good, create the columns
    for i in range(0, cols_to_create):
        new_col_idx =  col_idx + i
        
        new_col_header = str(hdr_dict[rpt_cols[i]])
        new_col_id =  'Q' + str(q_cnt) + '_' + new_col_header[:15]  

        if new_col_id in defdict:
            sg.popup('Query', 'Duplicate Column Id: ' + new_col_id + ' Cannot Complete Auto Create')
            return    
        
        cellvalues[row_idx][new_col_idx] = new_col_id
        new_def = copy.deepcopy(edef_line)
        new_def[gc.DF_ITEM_NAME] = new_col_id
        new_def[gc.DF_CELL_REF] = [new_col_idx,row_idx]
        new_def[gc.DF_DATA_TYPE] = 'T'
        new_def[gc.DF_DATA_PARAM] = new_col_header
        defdict[new_col_id] = new_def
    return

def edit_font(edef_line):
    '''add / edit font styles to cell def'''
    layout = gl.edit_font(edef_line)
    ewindow= sg.Window('Font Edit',layout,modal=True)
    while True:
    #   windowid, event, values = sg.read_all_windows()
        event, values = ewindow.read()
    #   print(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Save':
            if values['-SEL_FONT_NAME-']:
                edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_NAME] = values['-SEL_FONT_NAME-'][0]
 
            if values['-FONT_SZ-']:
                if int_or_float(values['-FONT_SZ-']):
                    edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_SZ] = values['-FONT_SZ-']
                    
            if values['-FONT_BOLD-']:
                edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_BOLD] = values['-FONT_BOLD-']

            if values['-FONT_ITLC-']:
                edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_ITLC] = values['-FONT_ITLC-']

            if values['-FONT_STKE-']:
                edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_STKE] = values['-FONT_STKE-']
 
            if values['-SEL_FONT_UNLN-']:
                edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_UNLN] = values['-SEL_FONT_UNLN-'][0]
                
            if values['-FONT_COLOR-']:
                edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_COLOR] = values['-FONT_COLOR-']
            
            break
    ewindow.close()    
    return

def edit_border(edef_line):
    '''add / edit border styles to cell def '''
    layout = gl.edit_border(edef_line)
    ewindow= sg.Window('Border Edit',layout,modal=True)
    while True:
    #   windowid, event, values = sg.read_all_windows()
        event, values = ewindow.read()
    #   print(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Save':
            if values['-SEL_BDR_STYLE-']:
                edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_STYLE] = values['-SEL_BDR_STYLE-'][0]
 
            if values['-LEFT_BDR-']:
                edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_LEFT] = values['-LEFT_BDR-']

            if values['-RIGHT_BDR-']:
                edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_RIGHT] = values['-RIGHT_BDR-']

            if values['-TOP_BDR-']:
                edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_TOP] = values['-TOP_BDR-']
            
            if values['-BOTTOM_BDR-']:
                edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_BOTTOM] = values['-BOTTOM_BDR-']
               
            if values['-BDR_COLOR-']:
                edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_COLOR] = values['-BDR_COLOR-']
            
            break
    ewindow.close()    
    return

def edit_align(edef_line):
    '''add / edit alignment styles to cell def '''
    layout = gl.edit_align(edef_line)
    ewindow= sg.Window('Alignment Edit',layout,modal=True)
    while True:
    #   windowid, event, values = sg.read_all_windows()
        event, values = ewindow.read()
    #   print(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Save':
            if values['-SEL_ALIGN_HORZ-']:
                edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_HORZ] = values['-SEL_ALIGN_HORZ-'][0]

            if values['-SEL_ALIGN_VERT-']:
                edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_VERT] = values['-SEL_ALIGN_VERT-'][0]    
 
            if values['-WRAP_TEXT-']:
                edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_WRAP] = values['-WRAP_TEXT-']

            if values['-SHRINK_TEXT-']:
                edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_SHRK] = values['-SHRINK_TEXT-']

           
            break
    ewindow.close()    
    return

def edit_fill(edef_line):
    '''add / edit fill styles to cell def '''
    layout = gl.edit_fill(edef_line)
    ewindow= sg.Window('Fill Edit',layout,modal=True)
    while True:
    #   windowid, event, values = sg.read_all_windows()
        event, values = ewindow.read()
    #   print(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':
            break
        if event == 'Save':
            if values['-SEL_FILL_TYPE-']:
                edef_line[gc.DF_CELL_FILL][gc.DF_FILL_TYPE] = values['-SEL_FILL_TYPE-'][0]
               
            if values['-FILL_COLOR-']:
                edef_line[gc.DF_CELL_FILL][gc.DF_FILL_BGCLR] = values['-FILL_COLOR-']
            
            break
    ewindow.close()    
    return

def get_sd_file():
    '''dialog for SD database file name, rem file name is actually the folder'''
    filename = ''
    foldername = sg.popup_get_folder('Select MV File Folder', no_window=True)
    foldertest = os.path.join(foldername, SD_FILE_BOLB)

    if os.path.isfile(foldertest):
        filename = foldername
    else:
        sg.popup('Select MV File','File does not appear to be a SD file')
  
    return filename

def get_sd_file_dict(filepath,filename):
    '''get database file directory item names and description'''
    unique_filename = str(uuid.uuid4())
    save_cwd = ''
    # make sure we are in the correct directory for the QM / SD command to execute (the account dir)
    #print ('Current Dir: ' + os.getcwd())
    if filepath != os.getcwd():
        save_cwd = os.getcwd()
        os.chdir(filepath)
        #print (os.getcwd())
    # use program sd_qwk_dict to generate a file with the directory names    
    cmd = "SD_QWK_DICT " + ' ' + filename + ' ' + filepath + ' ' + unique_filename
    #print (cmd)
    result = subprocess.run([SD_EXE,'-QUIET', cmd],stdout = subprocess.PIPE, text=True)
    list_data = []
    list_data_file = os.path.join(filepath, unique_filename)
    if os.path.exists(list_data_file):
        with open(list_data_file) as f:
            list_data = f.read().splitlines()
        f.close()
        os.remove(list_data_file)    
    if save_cwd != '':
        os.chdir(save_cwd) 
    # convert lists
    rtn_ids = []
    rtn_hdrs = {}
    for dict_item in list_data:
        tmp_list = dict_item.split(",")
        rtn_ids.append(tmp_list[0])
        rtn_hdrs[tmp_list[0]] = tmp_list[1]  
    return rtn_ids, rtn_hdrs



########################################## main menu functions  ##################################

def about_me():
    #sg.popup_no_wait('simple_rpt 0.0')
    try:
        #open text file in read mode
        text_file = open("README.md", "r")
        notes = text_file.read()
        text_file.close()
    except:
        notes = 'Could not read README.md?'

    choice, _ = sg.Window('About',[[sg.Multiline(default_text = notes,font=('Consolas', 12),size=(132,25))]]).read(close=True) 
########################################## main menu functions for saving and loading report def ##################################
def save_data(formname):
    '''quick and dirty way to save a report definition, its kind of a hack, but allows work to be saved and edited'''
    mysave_data = {}
    mysave_data['cellsvalues'] = cellvalues
    mysave_data['defdict'] = defdict
    if len(fpath) == 0: # if user has not set a folder path use current directory
        spath = os.path.join(os.getcwd(), str(formname))
    else:    
        spath = os.path.join(fpath, str(formname)) 
    try:
        with open(spath+'.json','w') as fp:
            json.dump(mysave_data, fp, sort_keys=True, indent=4) 
            fp.close()
            sg.popup('Save Data', 'Data saved to: ' + spath+'.json')
    except OSError as e:
        sg.popup(f"{type(e)}: {e}" + spath+'.json')

def load_data():
    '''load a saved report definition (saved via save_data)'''
    global cellvalues, defdict, fpath
   
    filename = sg.popup_get_file('Open', no_window=True)
    if filename:
        fname = os.path.splitext(str(filename))
        if fname[1] != '.json':
            sg.popup(filename + ' Does not appear to be a saved report work file')
            return  
          
        fpath, fname = os.path.split(fname[0])

        try:
            with open(str(filename), 'r') as fp:
                try:
                    save_data = json.load(fp)
                    cellvalues = save_data['cellsvalues'] 
                    defdict = save_data['defdict']
                    refresh_table()
                    
                    window['-FOLDER-'].update(value = fpath)    
                    window['-REPORT_FNAME-'].update(value = fname)
                except Exception as e:
                    sg.popup('Loaded error', f"{type(e)}: {e}")   

        except OSError as e:
            sg.popup(f"{type(e)}: {e}" + str(filename)+'.json')

def write_def(formname,values):
    '''write the sd_qwk_rpt report definition file item'''

    if len(fpath) == 0: # if user has not set a folder path use current directory
        spath = os.path.join(os.getcwd(), str(formname))
    else:    
        spath = os.path.join(fpath, str(formname)) 
    try:
        with open(spath,'w', encoding='ISO-8859-1') as fp:
            # fld 1 desc
            if '-REPORT_DESC-' in values:
                desc = values['-REPORT_DESC-']
                fp.write(desc)    
            else:
                fp.write(str(formname))
            fp.write('\n')

            # fld 2 page layout - to be added
            fp.write('\n')  

            # flds 3 - 10 tbd
            for x in range(3, 10+1):
                fp.write('\n')

            # cell definitions flds 11 and on    
            for row_idx in range(len(cellvalues)):
                row = cellvalues[row_idx]
                for col_idx  in range(len(row)):
                    cell = row[col_idx]
                    if cell == gc.EMPTYCELL:
                        pass
                    else:
                        edef_line = defdict[cell]
                        wr_ln = ''
                        wr_ln = wr_ln + edef_line[gc.DF_ITEM_NAME] + SDME_VM    #  cell name
                        edef_line[gc.DF_CELL_REF] = chr(col_idx+65) + str(row_idx+1)   # change to letter row address
                        wr_ln = wr_ln + edef_line[gc.DF_CELL_REF] + SDME_VM     #  cell address
                        wr_ln = wr_ln + edef_line[gc.DF_DATA_TYPE] + SDME_VM    #  data type
                    
                        # font styles
                        # name
                        fs = edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_NAME] + SDME_SVM
                        # size    
                        fs = fs + edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_SZ] + SDME_SVM 
                        # bold
                        if edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_BOLD]:
                            fs = fs + 'True' + SDME_SVM 
                        else:
                            fs = fs + '' + SDME_SVM 
                        # italic
                        if edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_ITLC]:
                            fs = fs + 'True' + SDME_SVM 
                        else:
                            fs = fs + '' + SDME_SVM    
                        # underline
                        fs = fs + edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_UNLN] + SDME_SVM
                        # strike
                        if edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_STKE]:
                            fs = fs + 'True' + SDME_SVM 
                        else:
                            fs = fs + '' + SDME_SVM                        
                        # color, rem to remove the '#'
                        mycolor = edef_line[gc.DF_FONT_STYLE][gc.DF_FONT_COLOR].replace('#','')
                        fs = fs + mycolor 
                        wr_ln = wr_ln + fs + SDME_VM

                        # Alignment
                        al = ''
                        # Horz Alignment
                        al = al + edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_HORZ] + SDME_SVM
                        # Vert Alignment
                        al = al + edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_VERT] + SDME_SVM
                        # Rot (currently not implymented)
                        al = al + '' + SDME_SVM   
                        # Wrap Alignment
                        if edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_WRAP]:
                            al = al + 'True' + SDME_SVM 
                        else:
                            al = al + '' + SDME_SVM
                        # Shrink Alignment
                        if edef_line[gc.DF_ALIGNMENT][gc.DF_ALIGN_SHRK]:
                            al = al + 'True' + SDME_SVM 
                        else:
                            alr = al + '' + SDME_SVM
                        # Indent (currently not implymented)
                        al = al + '' + SDME_SVM 

                        wr_ln = wr_ln + al + SDME_VM  # Align
                        
                        # Border
                        br = ''
                        # border left?
                        if edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_LEFT]:
                            br = br + 'True' + SDME_SVM 
                        else:
                            br = br + '' + SDME_SVM
                        
                        # border right?
                        if edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_RIGHT]:
                            br = br + 'True' + SDME_SVM 
                        else:
                            br = br + '' + SDME_SVM 

                        # border top?
                        if edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_TOP]:
                            br = br + 'True' + SDME_SVM 
                        else:
                            br = br + '' + SDME_SVM  

                        # border bottom?
                        if edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_BOTTOM]:
                            br = br + 'True' + SDME_SVM 
                        else:
                            br = br + '' + SDME_SVM                                                           

                        # border style
                        br = br + edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_STYLE] + SDME_SVM

                        # color, rem to remove the '#'
                        mycolor = edef_line[gc.DF_CELL_BORDER][gc.DF_BORDER_COLOR].replace('#','')
                        br = br + mycolor 

                        wr_ln = wr_ln + br + SDME_VM  # Border 

                        #Fill  (to be added)
                        fl = ''
                        # Fill Type
                        fl = fl + edef_line[gc.DF_CELL_FILL][gc.DF_FILL_TYPE] + SDME_SVM
                        # FILL FG Color (not used)
                        fl = fl + '' + SDME_SVM 
                        # FILL BG Color 
                        mycolor = edef_line[gc.DF_CELL_FILL][gc.DF_FILL_BGCLR].replace('#','')
                        fl = fl + mycolor 
                        wr_ln = wr_ln + fl + SDME_VM  # Fill

                        # Parameter
                        wr_ln = wr_ln + edef_line[gc.DF_DATA_PARAM] 
                        fp.write(wr_ln+'\n')

            fp.close()
            sg.popup('Write Report Def', 'Report Definition saved to: ' + spath)
    except OSError as e:
        sg.popup(f"{type(e)}: {e}" + spath)    

def getcell(row,col,cellvalues,defdict):
    '''load cell def data at row, col into editor window'''
    # print('clicked row: ' + str(row) +' col: ' + str(col))
    # lookup data for cell col,row and place in edef_line
    # for now its empty, need to add code

    
    def_key = cellvalues[row][col]
    if def_key == gc.EMPTYCELL:
        # note to prevent index error, fill out definition line structure
        edef_line =  copy.deepcopy(gc.DF_STRUCTURE) 
    else:
        edef_line = defdict[def_key]
    rtn_action = editor_dialog(col,row,edef_line)
    # propcess returned action

    if rtn_action == 'SAVE':
        # if this was an existing def, delete it and save the latest info
        if def_key == gc.EMPTYCELL:
            def_key = edef_line[gc.DF_ITEM_NAME]
            defdict[def_key] = edef_line
            cellvalues[row][col] = def_key  
        else:
            del defdict[def_key]    
            def_key = edef_line[gc.DF_ITEM_NAME]
            defdict[def_key] = edef_line
            cellvalues[row][col] = def_key

    elif rtn_action == 'DELETE':
        del defdict[def_key]
        cellvalues[row][col] = gc.EMPTYCELL 

    window['-CELL_TABLE-'].update(values = cellvalues)    

############################################## Windows                ############################################################
def main():
    global window,  table_col_count, table_row_count, cellvalues, defdict, fpath
    global selected_row, selected_col

    # user can specify starting row and col count for tables

    parser = argparse.ArgumentParser()
    parser.add_argument('-c', '--columns', help='Initial table column count', type=int, default=table_col_count)
    parser.add_argument('-r', '--rows', help='Initial table row count', type=int, default=table_row_count)
    args = parser.parse_args()
    
    table_col_count = args.columns
    table_row_count = args.rows

    # layout table values, start as empty cell
    cellvalues = [[gc.EMPTYCELL for col in range(table_col_count)] for count in range(table_row_count)]
            
    layout = gl.sd_dsn_window(table_col_count,cellvalues)
    window = sg.Window('SD_Quick Report Designer', layout=layout, margins=(0, 0),size=gc.WINDOW_SIZE, resizable=True,  finalize=True)
    pos_x, pos_y = window.current_location()
    win_width, win_height = window.size

    
    ############################################## all important event loop ############################################################
    while True:
    #   windowid, event, values = sg.read_all_windows()
        event, values = window.read()
    #   print(event, values)
        if event == sg.WIN_CLOSED or event == 'Exit':

            break

        # table events
        elif isinstance(event, tuple):
            # TABLE CLICKED Event has value in format ('-TABLE=', '+CLICKED+', (row,col))
            # You can also call Table.get_last_clicked_position to get the cell clicked
            #if event[2][0] == -1 and event[2][1] != -1:           # Header was clicked and wasn't the "row" column
            if event[0] == '-CELL_TABLE-':
                if event[2][0] == None:           # column resize was attempted
                    pass
                elif (event[2][0] == -1) or (event[2][1] == -1):
                    pass  # header or "row" column was clicked
                else:
                    # cell was clicked, get cell data (if any)
                    selected_row = event[2][0] 
                    getcell(event[2][0],event[2][1],cellvalues,defdict)


    # table move events move selected row up or down
        elif event == '-CELL_UP-':
            move_row_up('-CELL_TABLE-',cellvalues,selected_row)
        elif event == '-CELL_DOWN-':
            move_row_down('-CELL_TABLE-',cellvalues,selected_row)

    # Menu events

        elif event == 'Load':
            if sg.popup_ok_cancel('Warning, existing data will be overwritten') == 'OK':
                load_data()

        elif event == 'Save':
            if not values['-REPORT_FNAME-']:
                sg.popup('Save Data','No Report Name')
            else:
                save_data(values['-REPORT_FNAME-'])

        elif event == 'About':
            about_me()
   
        elif event == '-FOLDER-':
            fpath = values['-FOLDER-']

        elif event == 'WRTIE_DEF':
            write_def(values['-REPORT_FNAME-'],values)            

    window.close()

if __name__ == "__main__":
   main()    