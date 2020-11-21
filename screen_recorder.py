#!/usr/bin/env python
# coding: utf-8

import pyautogui
import os
import time
import tkinter.filedialog
import tkinter.messagebox
import tkinter as tk
import threading
import shlex
import signal
from subprocess import Popen,DEVNULL,STDOUT
import sys
sys.coinit_flags = 2
import pywinauto
from tkinter.ttk import *
from docx import Document
from docx.shared import Inches
from PIL import Image
from PyQt5.QtMultimedia import QAudioDeviceInfo,QAudio

#screenshot count
counter = 0                     

def take_screenshot_func():
    window.withdraw()
    image=pyautogui.screenshot('new1.png')
    window.deiconify()
    global counter
    #increment counter when screenshot is taken
    counter += 1 
    #display the counter
    count_text.set("Screenshot taken -> "+str(counter)) 
    #add screenshot to a document
    document.add_picture('new1.png', width=Inches(5.8))

def browse_func():
    path =tk.filedialog.askdirectory(initialdir=default_path)
    #display path in text field
    entry_text.set(path)

def savefile_func():
    global screenshot_filename
    global counter
    #check if any screenshot is taken
    if counter>0:
        #save document at the specified path
        document.save(os.path.join(entry_text.get(),screenshot_filename + doc_suffix))
        #display message with screenshot count,filename and location
        tk.messagebox.showinfo('Info','Total Screenshot taken -> '+str(counter)+'\nFile name ->'+screenshot_filename+'\n'+'Saved at '+entry_text.get())
        screenshot_filename='screenshots_'+time.strftime("%Y%m%d-%H%M%S")
        #reinitialize counter and label
        counter =0
        count_text.set("")
        #now clear the document object for a new file
        document.element.body.clear()
    else:
        #Display message
        tk.messagebox.showinfo('Info','No screenshot captured')

def startrecording_func():
	#disable all buttons and activate stop button
    screen_shot_btn.configure(state=tk.DISABLED)
    start_record_btn.configure(state=tk.DISABLED)
    browse_btn.configure(state=tk.DISABLED)
    save_btn.configure(state=tk.DISABLED)
    stop_record_btn.configure(state=tk.NORMAL)
    global video_location
    global video_filename
    video_filename='Recording_'+time.strftime("%Y%m%d-%H%M%S")
    video_location=os.path.join(entry_text.get(),video_filename+video_suffix)
    audio = audio_devices_combobox.get()
    pipeline=f"""ffmpeg -f gdigrab  -r 10  -i desktop -f dshow -i audio="{audio}" -c:v libx264 -r 10 -c:a aac -strict -2 -b:a 128k  "{video_location}" """
    global p
    p=Popen(shlex.split(pipeline))

def stoprecording_func():
    p.send_signal(signal.CTRL_C_EVENT)
    start_record_btn.configure(state=tk.NORMAL)
    screen_shot_btn.configure(state=tk.NORMAL)
    browse_btn.configure(state=tk.NORMAL)
    save_btn.configure(state=tk.NORMAL)
    tk.messagebox.showinfo('Info','Recording file ->'+video_filename+'\nSaved at ->'+entry_text.get())   

screenshot_filename='screenshots_'+time.strftime("%Y%m%d-%H%M%S")
default_path='C:/Users/vikesh/Desktop'
doc_suffix='.docx'
video_suffix='.mp4'
input_audio_devices = QAudioDeviceInfo.availableDevices(QAudio.AudioInput)
devices=[]
for device in input_audio_devices:
    if device.deviceName() not in devices:
        devices.append(device.deviceName())

document = Document()
window=tk.Tk()
window.configure(background='#767676')
window.resizable(width = False, height = False)
window.title('Recorder')

style = Style()
style.theme_use('clam')
style.configure('TButton', font = ('Segoe',9, 'bold'),bordercolor='white',background='#00a2ed',foreground='white')
style.map('TButton',background=[('active','#66ffc2')])
style.configure('TLabel', font = ('Segoe',9, 'bold'),background='#767676',foreground='white')

entry_text = tk.StringVar()
count_text = tk.StringVar()

start_icon=tk.PhotoImage(file='start.png')
stop_icon=tk.PhotoImage(file='stop.png')
screen_icon=tk.PhotoImage(file='screen.png')

screen_shot_btn = Button(window,text='Capture screen',image=screen_icon,command=take_screenshot_func, width=20, compound='bottom')
screen_shot_btn.grid(row=0, column=0)
start_record_btn = Button(window,text='Start recording',command=startrecording_func,image=start_icon, width=20,compound='bottom')
start_record_btn.grid(row=0, column=1)
stop_record_btn = Button(window,text='Stop recording',command=stoprecording_func,image=stop_icon, width=20, compound='bottom',state=tk.DISABLED)
stop_record_btn.grid(row=0, column=2)
audio_device_lbl = Label(window,  text = "Audio device to listen ->",font = ('Segoe',9, 'bold'))
audio_device_lbl.grid(row=1, column=0)
audio_devices_combobox=Combobox(window,values=devices,width=36,state="readonly")
audio_devices_combobox.grid(row=1, column=1,columnspan=2)
audio_devices_combobox.current(0)
path_lbl = Label(window,  text = "Saving Location ->",font = ('Segoe',9, 'bold'))
path_lbl.grid(row=2, column=0)
ent1 = Entry(window,width=38,textvariable=entry_text)
ent1.grid(row=2,column=1,columnspan=2)
ent1.insert(0,default_path)
count_lbl = Label(window,  text = " ->",font = ('Segoe',9, 'bold'),textvariable=count_text)
count_lbl.grid(row=3, column=0)
browse_btn = Button(window,text="Change",style='TButton',command=browse_func,width=20)
browse_btn.grid(row=3,column=1)
save_btn = Button(window,text='Save Screenshot file',command=savefile_func,width=20)
save_btn.grid(row=3, column=2)

window.mainloop()



