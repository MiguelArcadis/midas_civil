import pywinauto as pwa
from pywinauto import keyboard
from time import sleep
import pandas as pd

ELEMENTS_TO_EXTRACT = [1970, 1971, 1972, 1973]
COMBS_TO_EXTRACT = ['SW - Steel Only(ST)', 'SW - Piers(ST)', 'Wind X +ve(ST)', 'Wind Z +ve(ST)']
ELEMENT_PARTS_TO_EXTRACT = ['Part 2/4']

windows = pwa.Desktop(backend="win32").windows()
running_windows = [window.window_text() for window in windows]

midas_title = ""
target_title = "MIDAS"

for window in running_windows:
	if target_title in window:
		midas_title += window
	else:
		pass

app = pwa.Application().connect(title=midas_title)
app[midas_title].set_focus()

keyboard.send_keys('rrrrrrrrrrrr') # Finds Results Tables
keyboard.send_keys('{ENTER}') # Press enter on results tables
keyboard.send_keys('{ENTER}') # Press enter on beam
keyboard.send_keys('f') # Finds forces for results
keyboard.send_keys('{ENTER}') # Press enter on force

for e in ELEMENTS_TO_EXTRACT:
	keyboard.send_keys(str(e))
	keyboard.send_keys('{SPACE}')

for comb in COMBS_TO_EXTRACT:
	sleep(0.1)
	app.RecordsActivationDialog.ListBox2.select(comb).set_focus().type_keys('{VK_SPACE}')

for part in ELEMENT_PARTS_TO_EXTRACT:
	app.RecordsActivationDialog.ListBox3.select(part).set_focus().type_keys('{VK_SPACE}')

app.RecordsActivationDialog.OK.click()

app.MidasGenMainFrmClass.set_focus().click_input(button='right')

for _ in range(9):
	keyboard.send_keys('{VK_DOWN}')

keyboard.send_keys('{ENTER}')

sleep(1)
forces_to_extract = ['Axial', 'Shear-y', 'Shear-z', 'Torsion', 'Moment-y', 'Moment-z']

for fte in forces_to_extract:
	app.ResultViewItems.ListBox.select(fte).set_focus().type_keys('{VK_SPACE}')

app.ResultViewItems.OK.set_focus().click_input(button='left')

sleep(2)

app.MidasGenMainFrmClass['Result by Max Value-[Beam Force]'].GxMdiFrame.GXWND.set_focus()

keyboard.send_keys("{VK_RIGHT}")
sleep(0.5)
keyboard.send_keys("{VK_SHIFT down}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_RIGHT}{VK_SHIFT up}")
keyboard.send_keys("{VK_CONTROL down}{VK_SHIFT down}{VK_DOWN}{VK_CONTROL up}{VK_SHIFT up}")

keyboard.send_keys(r'^c')

column_names = ["Elem", "Load", "Part", "Component", "Axial (kN)", "Shear-y (kN)", "Shear-z (kN)", "Torsion (kN*m)", "Moment-y (kN*m)", "Moment-z (kN*m)"]
df_results = pd.read_clipboard(header=None, names=column_names)

print(df_results.head())
