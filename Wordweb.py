import pywinauto, time,sys
import pyautogui as pg
from pywinauto.application  import Application
from pywinauto import mouse
from pywinauto.keyboard import send_keys
import pyperclip as pc

"""app=Application().start(cmd_line=u'C:\Program Files (x86)\WordWeb\wweb32.exe',wait_for_idle=False)
time.sleep(5)# u can increase delay here, if you are experiencing issue to 10 sec.

app=Application().connect(title=u'WordWeb')# Put application title in xxxx

#app.WorWeb.print_control_identifiers()
app.WordWeb.maximize()
time.sleep(5)
app.WordWeb.wait('visible', timeout=20)
app.WordWeb.set_focus()"""

pg.hotkey('winleft')
time.sleep(2)
pg.typewrite("WordWeb\n",0.5)
time.sleep(2)
pg.hotkey('winleft','up')
time.sleep(2)
#pyautogui.moveTo(10,100)
#pyautogui.click()
#app.WordWeb.set_focus() # to set focus to wordweb window
#pywinauto.mouse.click(button='left',coords=(10,100))
app=Application().connect(title=u'WordWeb')
time.sleep(2)
app.WordWeb.menu_item(u'&Help->&About').click()
app.About.ok.click_input()

app.WordWeb[u'ComboBox:Edit'].type_keys("Help")

send_keys("{ENTER}")
time.sleep(5)
#pywinauto.mouse.press(button='right', coords=(30, 710))
pywinauto.mouse.press(button='left', coords=(35, 222))
pywinauto.mouse.release(button='left', coords=(350, 222))
time.sleep(1)
send_keys('^c')

paste=pc.paste()
with open("result.txt",'a') as file:
    file.write("\n")
    file.write(paste)

time.sleep(3)
pg.moveTo(1326, 703)
pg.click()
#app.WordWeb.TPanel.click_input() # [Dipon ]-Here I would like to click on the "close" button to exit  the wordweb window
