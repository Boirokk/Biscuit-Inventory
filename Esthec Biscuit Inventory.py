import os, pyautogui, time, xlrd, sys
import time
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True


DATE = time.strftime("%x")
fileDate = DATE.replace('/','_')

WORKBOOK_LOCATION = r"S:\Product Development\Design Dept. Transfer\ESTHEC CUT OFFS.xlsx"
ESTHEC_CUTOFF_LOCATION = r"S:\Yachts\Drawings\Finished CadFiles\_CNC DATA\_ESTHEC-CUT-OFFS\Master Do Not Delete.dwg"

# Import information to a variable
worksheet = xlrd.open_workbook(WORKBOOK_LOCATION)
sheet = worksheet.sheet_by_index(1)

# Create lists for coordinates

x = []

y = []

# add excel to list
for index in range(sheet.nrows):
    # Check to see if the cell has a number in it
    if type(sheet.cell_value(index, 0)) == float or type(sheet.cell_value(index, 0)) == int:
        # If the cell has a number in it the make sure its not 0
        if sheet.cell_value(index, 0) == 0.0:
            pass
        else:
             x.append(sheet.cell_value(index, 0))
             y.append(sheet.cell_value(index, 2))



movex = 500
movey = 250
count = 0

try:
    os.startfile(ESTHEC_CUTOFF_LOCATION)
    time.sleep(15)

    pyautogui.moveTo(movex,movey)

except PermissionError:
    pass

try:
    for numx in range(len(x)):
        if count == 14:
            count = 0
            movex = 500
            movey += 50
        pyautogui.typewrite('rectangle')
        pyautogui.press('enter')

        try:
            pyautogui.click(movex, movey)
        except PermissionError:
            pass

        pyautogui.typewrite('d')

        pyautogui.press('enter')
        pyautogui.press('enter')

        print('Drawing rectangle: ' + str(int(x[numx])) + ' X ' + str(int(y[numx])))

        pyautogui.typewrite(str(x[numx]))
        pyautogui.press('enter')
        pyautogui.typewrite(str(y[numx]))
        pyautogui.press('enter')

        
        try:
            pyautogui.click(movex, movey)
        except PermissionError:
            pass

        movex += 105
        count +=1
    pyautogui.hotkey('ctrl', 's')
    time.sleep(3)
    pyautogui.typewrite('Biscuit Cut Offs ' + fileDate)
    pyautogui.press('enter')

except Exception as ex:
    print(ex)



print('There are', len(x), 'pieces of cut-offs.')
print('Done...')

exit = input()

