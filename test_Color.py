import openpyxl as xl
from openpyxl.styles import colors
from openpyxl.styles import Color, PatternFill

wb = xl.load_workbook('colors.xlsx')
sh = wb['Sheet1']

test_Pass_Color = (
    colors.COLOR_INDEX[sh['A1'].fill.fgColor.index],
    colors.COLOR_INDEX[sh['A3'].fill.fgColor.index]
)

test_Fail_Color = (
    colors.COLOR_INDEX[sh['A2'].fill.fgColor.index],
    # colors.COLOR_INDEX[sh['A4'].fill.fgColor.index]
    None
)

test_Blocked_Color = (
    colors.COLOR_INDEX[sh['A7'].fill.fgColor.index],
    None
)


# print(test_Pass_Color[0])

# if sh['A4'].fill.fgColor.index != 'FFFFC000':
#     print('egal')
# else:
#     print('nu')