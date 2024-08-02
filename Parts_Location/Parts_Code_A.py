#Install py -m pip install pyautogui
#Approx. X seconds/part

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
    pyperclip.copy("FILTER-1000X")
    time.sleep(1)          #Sleep for 1 second
    pyautogui.write(pyperclip.paste())

    #Inventory Location ID
    pyperclip.copy("CY")
    time.sleep(0.2)          #Sleep for 0.2 second
    target_3 = pyautogui.locateCenterOnScreen("Capture_03.png",confidence=0.85)
    pyautogui.click(target_3)     #Move the mouse to target_3 coordinates and click
    pyautogui.write(pyperclip.paste())

    #Unit of Measure
    pyperclip.copy("EA")
    time.sleep(0.2)          #Sleep for 0.2 second
    target_4 = pyautogui.locateCenterOnScreen("Capture_04.png",confidence=0.85)
    pyautogui.click(target_4)     #Move the mouse to target_4 coordinates and click
    pyautogui.write(pyperclip.paste())

    #Bin ID
    pyperclip.copy("BIN-1000")
    time.sleep(0.2)          #Sleep for 0.2 seconds
    target_5 = pyautogui.locateCenterOnScreen("Capture_05.png",confidence=0.85)
    pyautogui.click(target_5)     #Move the mouse to target_5 coordinates and click
    pyautogui.write(pyperclip.paste())

    #Primary
    time.sleep(1)          #Sleep for 1 seconds
    target_6 = pyautogui.locateCenterOnScreen("Capture_06.png",confidence=0.85)
    pyautogui.click(target_6)     #Move the mouse to target_6 coordinates and click

    #Purchasing Info
    time.sleep(0.2)          #Sleep for 0.2 seconds
    target_10 = pyautogui.locateCenterOnScreen("Capture_10.png",confidence=0.85)
    pyautogui.click(target_10)     #Move the mouse to target_10 coordinates and click

    #Manufacturer Name
    pyperclip.copy("Wix")
    time.sleep(0.2)          #Sleep for 0.2 seconds
    target_11 = pyautogui.locateCenterOnScreen("Capture_11.png",confidence=0.85)
    pyautogui.click(target_11)     #Move the mouse to target_11 coordinates and click       
    pyautogui.write(pyperclip.paste())

    #Preferred Vendor ID
    pyperclip.copy("13857")
    time.sleep(0.2)          #Sleep for 0.2 seconds
    target_12 = pyautogui.locateCenterOnScreen("Capture_12.png",confidence=0.85)
    pyautogui.click(target_12)     #Move the mouse to target_12 coordinates and click       
    pyautogui.write(pyperclip.paste())     

    #Cancel/Save
    time.sleep(0.2)          #Sleep for 0.2 seconds
    target_17 = pyautogui.locateCenterOnScreen("Capture_17.png",confidence=0.85)
    pyautogui.click(target_17)     #Move the mouse to target_17 coordinates and click 
    #pyautogui.press("enter")     #Press the enter key
    time.sleep(1)          #Sleep for 1 seconds
except:
    print("ERROR")
