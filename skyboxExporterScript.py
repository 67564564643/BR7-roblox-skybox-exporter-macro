import os
from ahk import AHK
import time
import win32gui
import win32con
import win32process
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import psutil
import subprocess


DELAY = 1
BRYCE_PROCESS_NAME = "Bryce.exe"
EXPECTED_OPEN_COUNT = 6
index = -1
ahk = AHK()
pos = ahk.mouse_position
move = ahk.mouse_move
color = ahk.pixel_get_color
press = ahk.key_press
views = ['Up', 'Down', 'Front', 'Back', 'Right', 'Left']
xPos = ['-90', '90', '0', '0', '0,', '0']
yPos = ['270', '270', '0', '180', '270', '90']
opened_files = []
up_file = []
down_file = []
ahk_script_path = ""


def count_bryce_processes():
    count = 0
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] and proc.info['name'].lower() == BRYCE_PROCESS_NAME.lower():
                count += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return count
def close_bryce_processes():
    for proc in psutil.process_iter(['name']):
        try:
            if proc.info['name'] and proc.info['name'].lower() == BRYCE_PROCESS_NAME.lower():
                proc.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    time.sleep(2)
def browse_br7_file():
    """Select a .br7 file."""
    file_path = filedialog.askopenfilename(
        title="Select a Bryce 7 scene file",
        filetypes=[("Bryce 7 Scene Files", "*.br7")]
    )
    if file_path:
        entry_var.set(file_path)
        global opened_files
        opened_files = [] 
def open_file_six_times():
    file_path = entry_var.get()
    if not file_path or not os.path.isfile(file_path):
        messagebox.showerror("Error", "Please select a valid .br7 file first.")
        return

    def worker():
        global opened_files
        # Close existing Bryce if any
        if count_bryce_processes() > 0:
            confirm = messagebox.askyesno(
                "Bryce is running",
                "Bryce 7 is currently running.\n\nDo you want to close it? MAKE SURE TO SAVE ANY UNSAVED WORK!"
            )
            if not confirm:
                return
            close_bryce_processes()
        for _ in range(EXPECTED_OPEN_COUNT):
            os.startfile(file_path)
            time.sleep(DELAY)
        opened_files = [file_path] * EXPECTED_OPEN_COUNT
        messagebox.showinfo("Done", f"{EXPECTED_OPEN_COUNT} Bryce files opened.")
    threading.Thread(target=worker, daemon=True).start()
def enum_windows_callback(hwnd, hwnd_list):
    if win32gui.IsWindowVisible(hwnd):
        title = win32gui.GetWindowText(hwnd)
        if "Bryce 7" in title:  
            hwnd_list.append(hwnd)
def get_app_windows():
    hwnd_list = []
    win32gui.EnumWindows(enum_windows_callback, hwnd_list)
    return hwnd_list
def restore_next_window():
    global index
    hwnds = []
    hwnds = get_app_windows()
    if not hwnds:
        print("No Bryce 7 windows found")
        return

    index = (index + 1)
    hwnd = hwnds[index]

    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)
def movement():
    restore_next_window()
    time.sleep(1)
    press('1')
    move(20,150, speed=0, relative = False)
    ahk.click()
    time.sleep(.175)
    move(100,180, speed = 0, relative = False)
    ahk.click()
    time.sleep(.175)
    move(20, 150, speed = 0, relative = False)
    ahk.click()
    time.sleep(.175)
    move(170, 260, speed = 0, relative = False)
    ahk.click()
    ahk.click()
    ahk.send('{Backspace 2}')
    ahk.type('112.5')
    time.sleep(.175)
    move(170, 280, speed = 0, relative = False)
    ahk.click()
    ahk.click()
    ahk.send('{Backspace 2}')
    ahk.type('100')
    time.sleep(.175)
def position(x, y, w):
    move(170, 230, speed = 0, relative = False)
    ahk.click()
    ahk.click()
    ahk.send('{Backspace 2}')
    ahk.type('0')
    time.sleep(.175)
    move(140, 230, speed = 0, relative = False)
    ahk.click()
    ahk.click()
    ahk.send('{Backspace 2}')
    ahk.type(y)
    time.sleep(.175)
    move(90, 230, speed = 0, relative = False)
    ahk.click()
    ahk.click()
    ahk.send('{Backspace 2}')
    ahk.type(x)
    time.sleep(.175)
    move(220, 310, speed = 0, relative = False)
    ahk.click()
    time.sleep(.175)
    move(5, 30, speed = 0, relative = False)
    ahk.click()
    time.sleep(.175)
    move(10, 445, speed = 0, relative = False)
    ahk.click()
    time.sleep(.175)
    move(250, 240, speed = 0, relative = False)
    ahk.click()
    time.sleep(.3)
    move(270, 240, speed = 0, relative = False)
    time.sleep(.3)
    ahk.click()
    time.sleep(.3)
    move(270, 330, speed = 0, relative = False)
    ahk.click()
    time.sleep(.3)
    move(270, 210, speed = 0, relative = False)
    ahk.click()
    ahk.click()
    ahk.send('{Backspace 2}')
    time.sleep(.175)
    ahk.type(w + '.png')
    move(350, 210, speed = 0, relative = False)
    ahk.click()
    time.sleep(.5)
    ahk.send('{Left}{Left}{Left}{Enter}')
    time.sleep(.5)
def macro(num):
    movement()
    position(xPos[num],yPos[num], views[num])
def run_ahk_macro():
    global index
    running = 1
    if running == 2:
        pass
    count = count_bryce_processes()
    if count != EXPECTED_OPEN_COUNT:
        messagebox.showerror(
            "Cannot run AHK Macro",
            f"Expected {EXPECTED_OPEN_COUNT} Bryce files to be open.\nCurrently {count} are open."
        )
        return
    file_path = entry_var.get()
    if not opened_files or any(f != file_path for f in opened_files):
        messagebox.showerror(
            "Cannot run AHK Macro",
            "The 6 Bryce files currently open do not match the file that was opened."
        )
        return
    running = 2
    for i in range(6):
        macro(i)
    index = -1
    running = 1
root = tk.Tk()
root.title("BR7 File Opener")
root.geometry("460x180")
entry_var = tk.StringVar()
entry_var2 = tk.StringVar()
entry_var3 = tk.StringVar()
ahk_label_var = tk.StringVar(value="No AHK script selected")
frame = tk.Frame(root, padx=10, pady=10)
frame.pack(fill="both", expand=True)
tk.Label(frame, text="Bryce 7 File:").grid(row=0, column=0, sticky="w")
tk.Entry(frame, textvariable=entry_var, width=40).grid(row=1, column=0, padx=5, pady=5)
tk.Button(frame, text="Browse BR7 File", command=browse_br7_file).grid(row=1, column=1, padx=5, pady=5)
tk.Button(frame, text="Open BR7 scenes", command=open_file_six_times, width=20).grid(row=2, column=0, columnspan=2, pady=10)
tk.Button(frame, text='Run export macro', command = run_ahk_macro, width=18).grid(row=3, column = 0, columnspan=2, pady=10)
root.mainloop()

