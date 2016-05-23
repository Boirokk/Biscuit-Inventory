import os, pyautogui, time, xlrd
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True

WORKBOOK_LOCATION = r"\\T13d-02\c$\Users\jblaner\Desktop\ESTHEC CUT OFFS.xlsx"
ESTHEC_CUTOFF_LOCATION = "S:\Yachts\Drawings\Finished CadFiles\_CNC DATA\_ESTHEC-CUT-OFFS\BISCUIT.dwg"

# Import information to a variable
worksheet = xlrd.open_workbook(WORKBOOK_LOCATION)
sheet = worksheet.sheet_by_index(0)

# Create lists for coordinates
x = []
y = []

# add excel to list
for index in range(sheet.nrows):
    # Check to see if the cell has a number in it
    if type(sheet.cell_value(index, 8)) == float or type(sheet.cell_value(index, 8)) == int:
        # If the cell has a number in it the make sure its not 0
        if sheet.cell_value(index, 8) == 0.0:
            pass
        else:
             x.append(sheet.cell_value(index, 8))
             y.append(sheet.cell_value(index, 10))

movex = 300
movey = 547

os.startfile(ESTHEC_CUTOFF_LOCATION)
time.sleep(10)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')

pyautogui.click(260, movey)

for numx in range(len(x)):
    pyautogui.typewrite('rectangle')
    pyautogui.press('enter')
    pyautogui.click(movex, movey)
    pyautogui.typewrite('d')
    pyautogui.press('enter')
    movex += 10
    pyautogui.typewrite(str(x[numx]))
    pyautogui.press('enter')
    pyautogui.typewrite(str(y[numx]))
    pyautogui.press('enter')
    pyautogui.click(movex, movey)
    movex += 10
pyautogui.hotkey('ctrl', 's')
pyautogui.press('enter')

print('There are', len(x), 'pieces of cut-offs.')
print('Done...')
input()

