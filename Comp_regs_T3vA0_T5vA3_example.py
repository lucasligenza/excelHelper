import win32com.client

color_red = 3
color_green = 4
color_blue = 5
color_yellow = 6
color_magenta = 7
color_cyan = 8
color_pale_blue = 20
color_pale_green = 35
color_purple = 39 
color_olive = 43
color_orange = 45

wb_name = 'Talon3_default_vs_T5vA3_ETH.xlsx'

T3 = ['Talon3 (vA0)', 12, 5245]                 # worksheet name, first row, last row 
T5 = ['T5 RXTX (vA3) 031524', 12, 4803] 


def col2n(col):
    num = 0
    for c in col:
        num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


cfg_col = col2n('B')
domain_col = col2n('E')
symb_col = col2n('F')
bitfield_col = col2n('H')
hwdef_col = col2n('J')

anyanyany_col = col2n('L')

# just for ETH; should also include the refclk specific any/any column, but don't need in this case because it would not contain any of the settings we are interested it 
ETH_any_col = col2n('AF')

T3_eff_col = col2n('DO')        # output columns for the deterined "effective" values
T5_eff_col = col2n('DP')


def GetInt(value):
    try:
        if value.lower()[:2] == '0x':
            return int(value,16)
    except: pass
    try:
        return int(value)
    except:
        return None


def dict_symb_row(wb, shtdescr, symb_col):    # creates the sheet obeject and a dict that maps {symbol:row} 
    shtname = shtdescr[0]
    sht = wb.Worksheets(shtname)
    first_row = shtdescr[1]
    last_row = shtdescr[2]
    sdict = {}
    for row in range(first_row, last_row+1):
        symb = sht.Cells(row, symb_col).Value.lower().strip()
        sdict[symb]=row
    return [sht, sdict]         # worksheet and {symb:row} dict


def comp_by_symbol(t3_sht_sdict, t5_sht_sdict):
    t5_sht = t5_sht_sdict[0]
    t5_dict = t5_sht_sdict[1]
    t3_sht = t3_sht_sdict[0]
    t3_dict = t3_sht_sdict[1]
    
    for t5_symb, t5_row in t5_dict.iteritems():            
        if t5_sht.Cells(t5_row, domain_col).Value not in ('RX', 'rx'):        # we only care about RX here   
            continue
        if t5_sht.Cells(t5_row, cfg_col).Value not in (None, 'None', 'none', '', ' '):        # and only about non-cfgs   
            continue
        
        if t5_symb not in t3_dict:              # symbol only exists in T5?
            t5_sht.Cells(t5_row, symb_col).interior.colorindex = color_orange       # just flag this symbol
            continue

        # we have a RX symbol that exists in both T5 and T3

        t3_row = t3_dict[t5_symb]
        t3_sht.Cells(t3_row, symb_col).interior.colorindex = color_pale_green    # found it!

        # determine "effective" value on T5 sheet
        t5_eff_val = t5_sht.Cells(t5_row, hwdef_col).Value
        t5_ETH_any_val = t5_sht.Cells(t5_row, ETH_any_col).Value
        if t5_ETH_any_val not in (None, 'None', 'none', '', ' '):
            t5_eff_val = t5_ETH_any_val
        else:
            t5_anyanyany_val = t5_sht.Cells(t5_row, anyanyany_col).Value
            if t5_anyanyany_val not in (None, 'None', 'none', '', ' '):
                t5_eff_val = t5_anyanyany_val

        # determine "effective" value on T3 sheet
        t3_eff_val = t3_sht.Cells(t3_row, hwdef_col).Value
        t3_ETH_any_val = t3_sht.Cells(t3_row, ETH_any_col).Value
        if t3_ETH_any_val not in (None, 'None', 'none', '', ' '):
            t3_eff_val = t3_ETH_any_val
        else:
            t3_anyanyany_val = t3_sht.Cells(t3_row, anyanyany_col).Value
            if t3_anyanyany_val not in (None, 'None', 'none', '', ' '):
                t3_eff_val = t3_anyanyany_val

        t3_sht.Cells(t3_row, T3_eff_col).Value = t3_eff_val
        t3_sht.Cells(t3_row, T5_eff_col).Value = t5_eff_val

        # if effective values don't match, color cell red
        if GetInt(t3_eff_val) != GetInt(t5_eff_val):
            t3_sht.Cells(t3_row, T3_eff_col).interior.colorindex = color_red 
            t3_sht.Cells(t3_row, T5_eff_col).interior.colorindex = color_red

                      

if __name__ == '__main__': 

    xlApp = win32com.client.GetActiveObject('Excel.Application')
    xlWb = xlApp.Workbooks
    # wb_names = [wb.Name for wb in xlWb]
    wb = xlApp.Workbooks[wb_name]

    print('Generating rxtx dicts')
    t5_sht_sdict = dict_symb_row(wb, T5, symb_col)      
    t3_sht_sdict = dict_symb_row(wb, T3, symb_col)

    print('Comparing effective values')
    comp_by_symbol(t3_sht_sdict, t5_sht_sdict)


