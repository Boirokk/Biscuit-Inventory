import os, pyautogui, time, xlrd
pyautogui.PAUSE = 1
pyautogui.FAILSAFE = True



# Import information to a variable
worksheet = xlrd.open_workbook(r"\\T13d-02\c$\Users\jblaner\Desktop\ESTHEC CUT OFFS.xlsx")
sheet = worksheet.sheet_by_index(0)

# Create lists for coordinates
x = []
y = []

# add excel to list
for index in range(sheet.nrows):
    if type(sheet.cell_value(index, 8)) == float or type(sheet.cell_value(index, 8)) == int:
        if sheet.cell_value(index, 8) == 0.0:
            pass
        else:
             x.append(sheet.cell_value(index, 8))
             y.append(sheet.cell_value(index, 10))
print(x,y)


move = 400
os.startfile("S:\Yachts\Drawings\Finished CadFiles\_CNC DATA\_ESTHEC-CUT-OFFS\BISCUIT.dwg")
time.sleep(10)
pyautogui.hotkey('ctrl', 'a')
pyautogui.press('delete')



for numx in range(24):

    move += 50
    pyautogui.typewrite('rectangle')
    pyautogui.press('enter')
    pyautogui.click(move, 269)
    pyautogui.typewrite('d')
    pyautogui.press('enter')
    pyautogui.typewrite(str(x[numx]))
    pyautogui.press('enter')
    pyautogui.typewrite(str(y[numx]))
    pyautogui.press('enter')
    pyautogui.click(move, 269)

pyautogui.hotkey('ctrl', 's')
pyautogui.press('enter')

print('Done...')
