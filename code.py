import win32api, win32con, win32clipboard, win32gui, win32process
import ImageGrab
import os
import time
import ImageOps
from numpy import *

screen = {}
screen['top'] = 60
screen['left'] = 0
screen['width'] = 1920
screen['height'] = 980

class Stats:
	treasures_found_total = 0
	treasures_found_brown = 0
	treasures_found_gold = 0
	treasures_found = [
		0,
		0,
		0,
		0,
		0,
		0,
		0,
		0,
		0,
		0
	]

class Coords:
	ui_sell_center = (440, 576)
	ui_sell_center_close = (1405, 246)
	ui_hire_center = (834, 618)
	ui_drill_center = (1330, 559)
	ui_go_top = (109, 170)
	ui_go_up = (106, 246)
	ui_go_down = (106, 838)
	ui_go_bottom = (109, 915)
	ui_up_down_button_enabled = (118, 227, 116)
	ui_up_down_button_disabled = (60, 115, 59)

	ui_export = (187, 232)
	ui_export_ok = (1002, 151)
	ui_export_ok_additional = (1000, 180)

	ui_save = (235, 183)
	ui_save_ok = (1085, 102)
	ui_save_ok_additional = (1080, 135)

	brown_treasure = (77, 33, 19)
	treasure_color = (839, 572)

	worker_working = (32, 26, 9)
	worker_holding_treasure = (255, 255, 255)
	w = [
		(187 + 1, 895), # updated
		(325 + 1, 895),
		(463 + 1, 895),
		(601 + 1, 895),
		(740 + 1, 895),
		(878 + 1, 895), # updated
		(1016 + 1, 895), # updated
		(1154 + 1, 895), # updated
		(1293 + 1, 895),
		(1431 + 1, 895)
	]
	
	s_coal = (598, 273)
	s_copper = (595, 303)
	s_silver = (598, 331)
	s_gold = (596, 361)
	s_platinum = (592, 391)
	s_diamond = (604, 422)
	s_coltan = (606, 450)
	s_painite = (606, 480)
	s_black_opal = (613, 510)
	s_red_diamond = (613, 534)
	s_obsidian = (595, 570)
	s_californium = (595, 601)

def almost_the_same(c0, c1):
	n = 10
	if abs(c0[0] - c1[0]) < n and abs(c0[1] - c1[1]) < n and abs(c0[2] - c1[2]) < n:
		   return True
	return False

def screen_grab():
	box = (screen['left'], screen['top'], screen['left'] + screen['width'], screen['top'] + screen['height'])
	im = ImageGrab.grab(box)
	return im

def save_screen_grab(filename):
	im = ImageGrab.grab()
	im.save(os.getcwd() + '\\' + str(int(time.time())) + '__' + filename + '.png', 'PNG')

def left_click_at(coords):
	set_mouse_pos(coords)
	left_click()

def left_click():
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)

def left_down():
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
	time.sleep(0.1)

def left_up():
	win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
	time.sleep(0.1)

def set_mouse_pos(coordinates):
	win32api.SetCursorPos((screen['left'] + coordinates[0], screen['top'] + coordinates[1]))

def get_mouse_pos():
	x, y = win32api.GetCursorPos()
	x = x - screen['left']
	y = y - screen['top']
	return x, y

def open_chest(coords):
	left_click_at(coords)
	time.sleep(0.1)

	time.sleep(0.5)

	brown_treasure = True
	imm = ImageGrab.grab()
	time.sleep(0.5)
	if almost_the_same(Coords.brown_treasure, imm.getpixel(Coords.treasure_color)):
		print('Found a brown treasure.')
		Stats.treasures_found_brown += 1
	else:
		print('Found a gold treasure.')
		Stats.treasures_found_gold += 1
		brown_treasure = False
	
	left_click_at((screen['width'] / 2, screen['height'] / 2))
	time.sleep(0.5)
	
	# Save our victory!
	if brown_treasure:
		#save_screen_grab('brown')
		time.sleep(0.1)
	else:
		save_screen_grab('gold')
	
	time.sleep(0.5)
	left_click()

def look_for_chests(im):
	for i in range(0, 10):
		c0 = im.getpixel(Coords.w[i])
		c1 = Coords.worker_holding_treasure
		if c0[0] > 200 and c0[1] > 200 and c0[2] > 200:
			Stats.treasures_found_total += 1
			Stats.treasures_found[i] += 1
			open_chest(Coords.w[i])

def save_game():
	left_click_at(Coords.ui_save)
	time.sleep(0.1)
	left_click_at(Coords.ui_save_ok)
	left_click_at(Coords.ui_save_ok_additional)

def export():
	left_click_at(Coords.ui_export)
	time.sleep(0.1)

	win32api.keybd_event(win32con.VK_LCONTROL, 0x1d, 0, 0)
	win32api.keybd_event(win32api.VkKeyScan('C'), 0x1e, 0, 0)
	win32api.keybd_event(win32api.VkKeyScan('C'), 0x9e, win32con.KEYEVENTF_KEYUP, 0)
	win32api.keybd_event(win32con.VK_LCONTROL, 0x9d, win32con.KEYEVENTF_KEYUP, 0)
	time.sleep(0.1)

	win32clipboard.OpenClipboard()
	data = win32clipboard.GetClipboardData()
	win32clipboard.CloseClipboard()
	print data

	f = open(os.getcwd() + '\\' + str(int(time.time())) + '__export.txt', 'w')
	f.write(data)
	f.close()

	time.sleep(0.1)
	left_click_at(Coords.ui_export_ok)
	left_click_at(Coords.ui_export_ok_additional)
	
def go_to_bottom():
	while True:
		im = screen_grab()
		look_for_chests(im)
		if im.getpixel(Coords.ui_go_down) != Coords.ui_up_down_button_enabled:
			break
		left_click_at(Coords.ui_go_down)
		time.sleep(0.5)
	left_click_at(Coords.ui_go_top)

def print_stats():
	print 
	print 'TREASURES FOUND'
	print '---------------'
	print '  Total treasures found: ' + str(Stats.treasures_found_total)
	print '    Brown: ' + str(Stats.treasures_found_brown)
	print '    Gold: ' + str(Stats.treasures_found_gold)
	print Stats.treasures_found
	print 

def sell():
	left_click_at(Coords.ui_sell_center)
	left_click_at(Coords.s_coal)
	left_click_at(Coords.s_copper)
	left_click_at(Coords.s_silver)
	left_click_at(Coords.s_gold)
	left_click_at(Coords.s_platinum)
	left_click_at(Coords.s_diamond)
	left_click_at(Coords.s_coltan)
	left_click_at(Coords.s_painite)
	left_click_at(Coords.s_black_opal)
	left_click_at(Coords.s_red_diamond)
	left_click_at(Coords.s_obsidian)
	left_click_at(Coords.s_californium)
	left_click_at(Coords.ui_sell_center_close)

def force_always_on_top():
	hwnd = win32gui.GetForegroundWindow()
	win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 1430, 150, 500, 700, 0)

def main():
	force_always_on_top()
	export()
	print 'WELCOME MINION.'
	print 'Newb bot v0.1'
	print '-------------'
	print 'Starting in...'

	time_until_start = 10
	for i in range(0, time_until_start):
		print ' ' + str(time_until_start - i)
		time.sleep(1)

	runs_to_export = 10
	while True:
		print('-- Searching for treasures...')
		go_to_bottom()

		print('-- Selling')
		sell()

		print('-- Saving')
		save_game()

		if runs_to_export < 2:
			runs_to_export = 10
			print('-- Exporting')
			export()
		else:
			runs_to_export -= 1
			print('-- ' + str(runs_to_export) + ' runs left to export.')

		print_stats()
		print('-- Sleeping for 60 seconds...')
		for i in range(0, 6):
			print str(60 - i * 10) + ' seconds left until the next run.'
			time.sleep(10)
			  
if __name__ == '__main__':
	main()
