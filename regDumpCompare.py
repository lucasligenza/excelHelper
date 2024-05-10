import win32gui, win32con
from win32com.client import Dispatch
from tkinter import Tk, filedialog


prot = 'PCIE'               # PCIE, ETH, ...
targetCol = 'Y'             # hardcode the column you want to compare to the dump file

if prot == 'PCIE':
    cols = [targetCol, 'U', 'N', 'L', 'J']
elif prot == 'ETH':
    cols = [targetCol, 'AP', 'AG', 'AF', 'L', 'J']
elif prot == 'JESD':
    cols = [targetCol, 'BR', 'BE', 'BD', 'L', 'J']
elif prot == 'CPRI':
    cols = [targetCol, 'CI', 'CH', 'L', 'J']
else:
    print('Invalid protocol')


regdumps = [['RXTX (vA2)', 'T5RATB0.' , 12, 6473],
            ]

def col2n(col):
    num = 0
    for c in col:
        num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

def GetInt(value):
    try:
        if value.lower()[:2] == '0x':
            return int(value,16)
    except: pass
    return int(value)

def getExpVal(row):
    sbval = None
    for col in cols:
        coln = col2n(col)
        entry = xlSht.Cells(row, col).Value
        if entry not in (None, 'None', 'none', ' '):
            sbval = GetInt(entry)
            break
    return sbval


def hex2int(num_or_hexstr):
    if isinstance(num_or_hexstr, str) or isinstance(num_or_hexstr, type(u"")):
        if ('x' in num_or_hexstr) or ('X' in num_or_hexstr):
            hexstr = num_or_hexstr.lower().split('x')[1]
        else:
            hexstr = num_or_hexstr
        try:
            res = int(hexstr, 16)
        except:
            res = None
    else:
        try:
            res = int(num_or_hexstr)
        except:
            res = None
    return res

            
HWdef_only = False
hwdef_col = col2n('J') 
symb_col = col2n('F')
type_col = col2n('CW')

color_red = 3
color_green = 4
color_blue = 5
color_yellow = 6
color_magenta = 7
color_cyan = 8
color_maroon = 9
color_orange = 45
color_nofill = 0

   
####################################################################
##############  Main Loop
####################################################################

root = Tk.Tk()
root.withdraw()

print ('Select dump file')
myFormats = [('TXT File','*.txt',),]
dumpfilename = filedialog.askopenfilename(parent=root,filetypes=myFormats ,title = "Select Arctuc reg dump file (.txt)")

print ('Select register xlsx file')
myFormats = [('Excel File','*.xlsx'),]
IOfile = filedialog.askopenfilename(parent=root,filetypes=myFormats ,title = "Select register XLSX file ...")
xlApp = Dispatch("Excel.Application")
xlWb = xlApp.Workbooks.Open(IOfile)
xlApp.Visible = 1

val_col = input('Enter output value column (A - ZZ): ')
val_col = col2n(val_col)


print ("Creating reg:val dict from dump file")
# Read the reg dump for the selected lane / CMU into dicts
prefixes = []
for rd in regdumps:
    prefixes.append(rd[1].lower())

regval_dict = {}
with open(dumpfilename, 'r') as dumpfile:
    lines = dumpfile.readlines()
    for line in lines:
        #print line
        for pf in prefixes:
            if pf in line.lower():
                lineitems = line.split('=')
                symbol = ((lineitems[0].split('.')[1]).strip()).lower()
                val = hex2int(((lineitems[1].split('#'))[0]).strip())
                regval_dict[symbol] = val
                break

for regdump in regdumps:
    ws = regdump[0]    
    prefix_ = regdump[1]
    first_r = regdump[2]
    last_r = regdump[3]
    
    xlSht = xlWb.Worksheets(ws)

    row = first_r
    while row <= last_r:
        print(row)
        try:
            symb = xlSht.Cells(row,symb_col).Value
            if (symb not in (None, 'None', ''))  and (symb.lower() != 'reserved'):
                print ('Checking ', symb)
                devval = regval_dict[symb]
                print ('DUT %s = %d' % symb, devval)
                xlSht.Cells(row, val_col).Value = '0x%x' % (devval)
                expVal = getExpVal(row)
                if devval != expVal:
                    xlSht.Cells(row,val_col).Interior.ColorIndex = color_red            # Value seems wrong
        except:
            pass
        row += 1
# change colors to only red if the value differs from expected

# Save and close Excel spreadshseet
#xlWb.Close(SaveChanges=True)
#xlApp.Quit()
#lApp.Visible = 0
#del xlApp

