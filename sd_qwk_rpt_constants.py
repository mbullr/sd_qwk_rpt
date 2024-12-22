'''
  simple_rpt_constants.py part of guiBuilder

'''
# starting window size
WINDOW_SIZE = (1200,600)

# name we use to create our test view for the layouts
WIN_W = 90
WIN_H = 25
MAX_PROPERTIES = 60    # max number of properties we are setup to display
#

EMPTYCELL = '-             -'
CONTINUE_LAYOUT = '+++++++++'
# text for selected layout table cell text box
SELECT_CELL_TEXT = 'Selected Cell (r/c): '

#define some colors
CELL_TABLE_COLOR = 'light steel blue'

SELECT_COLOR = 'red on yellow'
#  SELECT_COLOR = 'yellow'

DF_STRUCTURE =  ["",
                  "",
                  "",
                  ['','','','','','',''],
                  ['','','','','',''],
                  ['','','','','',''],
                  ['','',''],
                  ''] 

# define report definition line layout (elements within list)
DF_ITEM_NAME  = 0
DF_CELL_REF   = 1
DF_DATA_TYPE  = 2
#
DF_FONT_STYLE = 3
# list element within list element of DF_FONT_STYLE
DF_FONT_NAME  = 0
DF_FONT_SZ    = 1
DF_FONT_BOLD  = 2
DF_FONT_ITLC  = 3
DF_FONT_UNLN  = 4
DF_FONT_STKE  = 5
DF_FONT_COLOR = 6
#
DF_ALIGNMENT  = 4
# list element within list element of DF_ALIGNMENT
DF_ALIGN_HORZ = 0
DF_ALIGN_VERT = 1
DF_ALIGN_ROT  = 2
DF_ALIGN_WRAP = 3
DF_ALIGN_SHRK = 4
DF_ALIGN_IDNT = 5
#
DF_CELL_BORDER  = 5
# list element within list element of DF_CELL_BORDR
DF_BORDER_LEFT   = 0
DF_BORDER_RIGHT  = 1
DF_BORDER_TOP    = 2
DF_BORDER_BOTTOM = 3
DF_BORDER_STYLE  = 4
DF_BORDER_COLOR  = 5
#
DF_CELL_FILL  = 6
# list element within list element of DF_CELL_FILL
DF_FILL_TYPE   = 0
DF_FILL_FGCLR  = 1
DF_FILL_BGCLR  = 2

#	
DF_DATA_PARAM = 7