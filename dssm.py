import os.path
import datetime
import hashlib
import zipfile
import time
import re

# Screenshot functionality
import win32gui
import win32ui
import win32con
import win32api

# BMP -> PNG
from PIL import Image



""" compresses files, returns None on errors else output file name """
def zip_file(input_files, output_file, compress_type=zipfile.ZIP_DEFLATED):
	try:
		out_file = zipfile.ZipFile(output_file, 'w', compress_type)

		for file in input_files:
			filename = os.path.basename(os.path.normpath(file))
			out_file.write(file, filename)

		out_file.close()
		return output_file
	except Exception as e:
		print("Error zipping files " + str(input_files))
		print(e)
		return None


""" Gets a files md5 hash, naive approach. Mod date is faster most likely """
def get_file_md5(file_dir):
	try:
		with open(file_dir, 'rb') as f:
			data = f.read()
			md5 = hashlib.md5(data).hexdigest()
			return md5
	except FileNotFoundError:
		print("Exception: File Not Found")
		return None
	except Exception as e:
		print(e)
		return None


""" Get the users base save directory and user numbers to find saves """
def get_directories(save_dir='Documents/NBGI/', game_name='DARK SOULS REMASTERED'):
	home_dir = os.path.expanduser('~')
	base_save_dir = os.path.join(home_dir, save_dir, game_name)

	# clean up mixed seps
	base_save_dir = base_save_dir.replace('/', '\\')

	# make sure that the requested save directory exists
	if not os.path.exists(base_save_dir):
		raise Exception("Error: Save base folder not found " + str(base_save_dir))

	# get user number(s) from directory
	user_numbers = [k for k in os.listdir(base_save_dir) if os.path.isdir(os.path.join(base_save_dir, k)) ]
	
	# List full paths with user as key
	dirs = {n : os.path.join(base_save_dir, n) for n in user_numbers}

	return {
		'base_save_dir': base_save_dir,
		'dirs': dirs,
	}


""" Get all save files, filterable by user """
def get_all_saves(dirs, save_ext='.sl2', use_hash=False):
	saves = []
	for user in dirs['dirs'].keys():
		for file in os.listdir(dirs['dirs'][user]):
			if file.endswith(save_ext):
				#only calculate hash if flag set
				file_hash = get_file_md5(os.path.join(dirs['dirs'][user], file)) if use_hash else None

				saves.append({
					'user': user,
					'base_save_dir': dirs['dirs'][user],
					'save_file': os.path.join(dirs['dirs'][user], file),
					'mod_date': os.path.getmtime(os.path.join(dirs['dirs'][user], file)),
					'hash': file_hash,
				})
	return saves


""" Get all backed up saves  """
def get_all_backups(dirs, backup_ext='.zip'):
	backups = []
	for user in dirs['dirs'].keys():
		for file in os.listdir(dirs['dirs'][user]):
			if file.endswith(backup_ext):

				#Get time from backup name, depends on fmt
				#TODO separate all this to its on fn
				#DRAKS0005_20180528_224011.zip
				base = os.path.basename(file)
				m = re.search(r'(\d{4}\d{2}\d{2}_\d{2}\d{2}\d{2})', base)
				dt = datetime.datetime.strptime(m[1], '%Y%m%d_%H%M%S')


				backups.append({
					'user': user,
					'base_save_dir': dirs['dirs'][user],
					'backup_zip': os.path.join(dirs['dirs'][user], file),
					'backup_time': dt,
				})
	return backups


""" Checks for save updates, either using hash or mod date """
def check_if_changed(saves, use_hash=False, screenshot=True):
	for save in saves:

		# Use hash to check if changed
		if use_hash:
			new_hash = get_file_md5(save['save_file'])
			old_hash = save['hash']
			if old_hash != new_hash:
				backup_save(save, screenshot)
				print("old:" + old_hash + " new: " + new_hash)
				return True
			else:
				
				return False
		# look at mod date to check if changed
		else:
			old_mod = save['mod_date']
			new_mod = os.path.getmtime(save['save_file'])
			if old_mod != new_mod:
				backup_save(save, screenshot)
				print("old:" + str(old_mod) + " new:" + str(new_mod))
				return True
			else:
				print("no saves changed.")
				return False


""" zips up save file, can also take a screenshot of screen"""
def backup_save(save, screenshot=True):
	dt = get_current_datetime()
	ss_fn = os.path.splitext(save['save_file'])[0] + '_' + dt + '.bmp'
	png_fn = os.path.splitext(save['save_file'])[0] + '_' + dt + '.png'
	zip_fn = os.path.splitext(save['save_file'])[0]+ '_' + dt +'.zip'

	
	if screenshot:
		#Take screenshot
		take_screenshot(ss_fn)
		# Convert to jpeg
		img = Image.open(ss_fn)
		img.save(png_fn, 'png')

		#zip up save + ss
		zip_file([save['save_file'], png_fn], zip_fn)
		#delete screenshots
		os.remove(ss_fn)
		os.remove(png_fn)
	else:
		zip_file([save['save_file']], zip_fn)


""" Makes inital copy of all save files """
def backup_all_saves(saves, screenshot=False):
	for save in saves:
		backup_save(save, screenshot=screenshot)


""" Get the current timestamp as a string in the format requested """
def get_current_datetime(fmt='%Y%m%d_%H%M%S'):
	dt = datetime.datetime.now()
	dt_str = dt.strftime(fmt)
	return dt_str


""" Takes a screenshot (bitmap) of foreground window, save to outfile """
def take_screenshot(out_file):
	# get active window & size
	fg_window = win32gui.GetForegroundWindow()

	# left, top, right, bottom
	fg_l, fg_t, fg_r, fg_b = win32gui.GetWindowRect(fg_window)

	# calculate width & height of fg window
	fg_w = fg_r - fg_l
	fg_h = fg_b - fg_t

	# create device context
	fg_window_dc = win32gui.GetWindowDC(fg_window)
	img_dc = win32ui.CreateDCFromHandle(fg_window_dc)

	# create memory based dc
	mem_dc = img_dc.CreateCompatibleDC()

	# create bitmap
	screenshot = win32ui.CreateBitmap()
	screenshot.CreateCompatibleBitmap(img_dc, fg_w, fg_h)
	mem_dc.SelectObject(screenshot)

	# copy screen into memory dc
	mem_dc.BitBlt((0, 0), (fg_w, fg_h), img_dc, (fg_l, fg_t), win32con.SRCCOPY)

	# save bitmap
	screenshot.SaveBitmapFile(mem_dc, out_file)

	# clean up
	mem_dc.DeleteDC()
	win32gui.DeleteObject(screenshot.GetHandle())


""" Deletes old backups if we have newer saves"""
def delete_old_backups(backups, last_n_to_keep=10):

	if len(backups) <= last_n_to_keep:
		return None

	# Get all time / files then sort by time 
	backups = [ {'dt': backup['backup_time'], 'file':backup['backup_zip']} for backup in backups]
	backups = sorted(backups, key=lambda k: k['dt'])

	to_delete = backups[:last_n_to_keep-1]

	for backup in to_delete:
		print("deleting backup: %s", backup['file'])
		os.remove(os.path.basename(backup['file']))



	


""" Main runtime """
if __name__ == "__main__":
	dirs = get_directories()
	saves = get_all_saves(dirs)
	backup_all_saves(saves)

	interval = 60 * 5


	while True:
		time.sleep(interval)
		if check_if_changed(saves, hash=False, screenshot=True):
			saves = get_all_saves(dirs)
			backups = get_all_backups(dirs)
			delete_old_backups(backups)