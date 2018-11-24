import xlwings as xw
import pyautogui
from appJar import gui
from win32com.client import Dispatch
import re
import xlrd
import clipboard
from datetime import timedelta
from datetime import datetime
from dateutil import  parser


wb = xw.Book('EricssonBSS V7.6.xlsm')
sht = wb.sheets['Ericsson']
workbook = xlrd.open_workbook("warid.xlsx")
worksheet = workbook.sheet_by_index(0)

#########################################

reason = []
down_time = []
BSC = []
Site_id = []
city = []
rbu = []
sites = []


##################################################

screenWidth, screenHeight = pyautogui.size()
currentMouseX, currentMouseY = pyautogui.position()
wsh = Dispatch("WScript.Shell")

#############################################################################

for x in range(1, 16480):
    rbu.append(worksheet.cell(x, 17).value)
    city.append(worksheet.cell(x, 5).value)
    sites.append(worksheet.cell(x, 0).value)

for y in range(0, 16479):
    temp = sites[y]
    sites[y] = temp[-6:]


#########################################

reason = []
down_time = []
BSC = []
Site_id = []
city = []
rbu = []
sites = []


##################################################

screenWidth, screenHeight = pyautogui.size()
currentMouseX, currentMouseY = pyautogui.position()
wsh = Dispatch("WScript.Shell")

#############################################################################

for x in range(1, 16480):
    rbu.append(worksheet.cell(x, 17).value)
    city.append(worksheet.cell(x, 5).value)
    sites.append(worksheet.cell(x, 0).value)

for y in range(0, 16479):
    temp = sites[y]
    sites[y] = temp[-6:]

############################################################################


def press2G(btn):
    count = 5
    check = True
    while check == True:
        c = str(count)
        if (sht.range('P' + c).value == None):
            reason = sht.range('O' + c).value
            down_time = sht.range('Q' + c).value
            Site_id = sht.range('S' + c).value
            check = False
        elif (sht.range('P' + c).value != None):
            count = count + 1
            check = True

    start_time = parser.parse(down_time)
    esc_time = start_time + timedelta(minutes=5)
    start_time = datetime.strftime(start_time, '%d/%m/%Y %H:%M')
    esc_time = datetime.strftime(esc_time, '%d/%m/%Y %H:%M')

    for z in range(0, 16479):
        if Site_id == sites[z]:
            city_name = city[z]
        else:
            temp1 = 0

    ###################click on task_bar################################

    pyautogui.click(x=int(app.getTextArea(title='Task bar x')), y=int(app.getTextArea(title='Task bar y')), button='left',interval = 0.1,tween=pyautogui.easeInOutQuad)


    ###################click on region_dropdown###############################

    pyautogui.click(x=722,y=286,button='left',tween=pyautogui.easeInOutQuad)

    ###################click on clear################################

    pyautogui.click(x=621, y=505, button='left',interval = 0.8,tween=pyautogui.easeInOutQuad)

    ###################click on site_id################################

    pyautogui.click(x=669, y=356, button='left',clicks=2,interval = 0.1,tween=pyautogui.easeInOutQuad)
    pyautogui.hotkey('ctrlleft', 'a')
    pyautogui.typewrite(message=Site_id)
    wsh.Sendkeys("{ENTER}")

    ###################click on title_site_id_is_down################################

    pyautogui.click(x=347, y=157, button='left',tween=pyautogui.easeInOutQuad)
    pyautogui.hotkey('ctrlleft', 'a')
    pyautogui.press('delete')
    z = Site_id + " is up now"
    pyautogui.typewrite(message=z)
    ###################click on description################################

    pyautogui.click(x=346, y=326, button='left')
    pyautogui.hotkey('ctrlleft', 'a')
    pyautogui.press('delete')
    pyautogui.typewrite(message=reason)

    ###################click on scroll_down################################

    pyautogui.click(x=1593, y=837, button='left',clicks=10,tween=pyautogui.easeInOutQuad)

    ###################click on details################################

    pyautogui.click(x=113, y=309, button='left',interval = 0.1)

    ###################click on responsible_stakeholder################################

    pyautogui.click(x=898, y=238, button='left',clicks=2,interval = 0.1,tween=pyautogui.easeInOutQuad)

    ###################click on type_NOSS################################

    pyautogui.click(x=879, y=510, button='left',interval = 1.0,tween=pyautogui.easeInOutQuad)
    pyautogui.typewrite(message="NOSS")
    wsh.Sendkeys("{ENTER}")

    ###################click on child_category################################

    pyautogui.click(x=58, y=419, clicks = 2, button='left', interval=0.1)

    ###################click on type_CP################################

    pyautogui.click(x=879, y=251092, button='left',interval = 1.0,tween=pyautogui.easeInOutQuad)
    pyautogui.typewrite(message="CP")
    wsh.Sendkeys("{ENTER}")

    ###################click on close_actions################################

    pyautogui.click(x=170, y=312, button='left',tween=pyautogui.easeInOutQuad)

    ###################click on esc_to_noss_and_fme_and_RCA################################

    pyautogui.click(x=418, y=526, button='left',interval = 0.2)
    pyautogui.hotkey('ctrlleft', 'a')
    pyautogui.press('delete')
    m = "Escalated to NOSS and FME" + "\r\n" \
        + "RCA Will be shared later by NOSS"
    pyautogui.typewrite(message=m)

    ###################click on action_code_type_Power_Issue_Fixed################################

    pyautogui.click(x=477, y=394, button='left',tween=pyautogui.easeInOutQuad)
    pyautogui.typewrite(message="Power Issue Fixed")

    ###################click on close_code_type_Resolved################################

    pyautogui.click(x=633, y=414, button='left',tween=pyautogui.easeInOutQuad)
    pyautogui.typewrite(message ="Resolved")

    ###################click on time_stamps################################

    pyautogui.click(x=577, y=283, button='left',interval = 0.1,tween=pyautogui.easeInOutQuad)

    ###################click on start_time################################

    pyautogui.click(x=170, y=339, button='left',interval = 0.1,tween=pyautogui.easeInOutQuad)
    pyautogui.typewrite(message = start_time)


    ###################click on esc_time################################

    pyautogui.click(x=170, y=454, button='left',interval = 0.1,tween=pyautogui.easeInOutQuad)
    pyautogui.typewrite(message = esc_time)

    ###################click on scroll_up################################

    pyautogui.click(x=1591, y=114, button='left',clicks=10,tween=pyautogui.easeInOutQuad)

    ###################click on save################################

    pyautogui.click(x=15, y=85, button='left',tween=pyautogui.easeInOutQuad)


    ###################click on state_drop_down################################

    pyautogui.click(x=584, y=264, button='left',interval= 2.0)

    ###################click on Open################################

    pyautogui.click(x=587, y=303, button='left',interval = 0.5,tween=pyautogui.easeInOutQuad)

    ###################click on save################################

    pyautogui.click(x=15, y=85, button='left',interval = 0.1, tween=pyautogui.easeInOutQuad)

    ###################click on TT and copy################################

    pyautogui.click(x=97, y=116, button='left', interval=2, tween=pyautogui.easeInOutQuad)
    pyautogui.hotkey('ctrlleft', 'a')
    pyautogui.hotkey('ctrlleft', 'c')

    ############################################################
    sht.range('P' + c).value = clipboard.paste()



###################################################
app = gui("1TM")
app.addButton("Create 2G TT", press2G,0,1)
app.setButtonWidhts("Create 2G TT",2)


app.enableEnter(press2G)

############################################
app.addLabel("l1", "Task bar x",1,0)
app.addTextArea("Task bar x",2,0)

############################################
app.addLabel("l2", "Task bar y",1,2)
app.addTextArea("Task bar y",2,2)

###########################################

app.setAllTextAreaHeights(1)

app.go()