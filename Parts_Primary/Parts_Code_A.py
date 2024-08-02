#Install py -m pip install pyautogui
#Approx. 7 seconds/part

import pyautogui
import time
import pyperclip

#Beginning actions
target_main = pyautogui.locateCenterOnScreen("Capture_Main.png",confidence=0.85)
pyautogui.click(target_main)     #Move the mouse to target_main coordinates and click

#Creating new parts
try:
    #New
    target_1 = pyautogui.locateCenterOnScreen("Capture_01.png",confidence=0.85)
    pyautogui.click(target_1)     #Move the mouse to target_1 coordinates and click

    #Part ID
    pyperclip.copy("04")
    time.sleep(1)          #Sleep for 1 second
    target_2 = pyautogui.locateCenterOnScreen("Capture_02.png",confidence=0.85)
    pyautogui.click(target_2)     #Move the mouse to target_2 coordinates and click
    pyautogui.write(pyperclip.paste())

    #Part Suffix
    time.sleep(0.1)          #Sleep for 0.1 second
    target_3 = pyautogui.locateCenterOnScreen("Capture_03.png",confidence=0.85)
    pyautogui.click(target_3)     #Move the mouse to target_3 coordinates and click

    #Clear popup
    target_0 = pyautogui.locateCenterOnScreen("Capture_00.png",confidence=0.85)
    pyautogui.click(target_0)     #Move the mouse to target_0 coordinates and click

    #Keyword
    time.sleep(0.1)          #Sleep for 0.1 second
    target_4 = pyautogui.locateCenterOnScreen("Capture_04.png",confidence=0.85)
    pyautogui.click(target_4)     #Move the mouse to target_4 coordinates and click
    pyautogui.write(pyperclip.paste())

    #Clear popup
    pyautogui.click(target_0)     #Move the mouse to target_0 coordinates and click

    #Short Description
    time.sleep(0.1)          #Sleep for 0.1 seconds
    target_5 = pyautogui.locateCenterOnScreen("Capture_05.png",confidence=0.85)
    pyautogui.click(target_5)     #Move the mouse to target_5 coordinates and click
    pyautogui.write(pyperclip.paste())

    #Product Category ID
    time.sleep(0.1)          #Sleep for 0.1 seconds
    target_6 = pyautogui.locateCenterOnScreen("Capture_06.png",confidence=0.85)
    pyautogui.click(target_6)     #Move the mouse to target_6 coordinates and click       
    pyautogui.write(pyperclip.paste())

    #Part Classification ID
    time.sleep(0.1)          #Sleep for 0.1 seconds
    target_7 = pyautogui.locateCenterOnScreen("Capture_07.png",confidence=0.85)
    pyautogui.click(target_7)     #Move the mouse to target_7 coordinates and click       
    pyautogui.write(pyperclip.paste())

    #Cancel/Save
    time.sleep(0.1)          #Sleep for 0.1 seconds
    target_8 = pyautogui.locateCenterOnScreen("Capture_08A.png",confidence=0.85)
    pyautogui.click(target_8)     #Move the mouse to target_7 coordinates and click 
    #pyautogui.press("enter")     #Press the enter key
    time.sleep(1)          #Sleep for 1 seconds
except:
    print("ERROR")
