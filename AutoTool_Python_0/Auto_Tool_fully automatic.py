# -*- coding: cp949 -*-
import os
import time
import inspect
import win32gui
import clipboard
import tkinter as tk
import natsort as ns
import pyautogui as py

from tkinter import messagebox
from tkinter import filedialog
from tqdm import tqdm


def enumHandler(hwnd, lParam):
    if win32gui.IsWindowVisible(hwnd):
        if appname in win32gui.GetWindowText(hwnd):
            win32gui.MoveWindow(hwnd, xpos, ypos, width, length, True)

def get_image_file_list(path):
    imageList = []
    directory = os.listdir(path)
    for file in directory:
        if file.endswith('.jpg') or file.endswith('.miss'):
            imageList.append(file)

    return ns.natsorted(imageList)

def copy_and_paste(data):
    clipboard.copy(data)
    clipboard.paste()
 
def image_button_click(image):
    while True: # ��ư�� ã�� ���� �� ���� �ݺ�
        try:
            time.sleep(0.7) # �˾Ƽ� �߸��߱�
            btnImage = py.locateOnScreen(image)  # ȭ�鿡�� image�� ����       
            point = py.center(btnImage) # ã�� �̹����� ��ǥ�� �߾� ���� ����
            print(image + ' ' + str(point)) # �ܼ�â�� image ���, ã�� ��ǥ�� ���
            time.sleep(0.7) # �˾Ƽ� �߸��߱�
            py.click(point) # ���� ��ǥ�� Ŭ��
            break # �ݺ��� ����
        except: # ��ư�� ��ã�� ���
            print('Error: ' + image + ' not found.') # �ܼ�â�� image�� ��ã�Ҵٴ� ���� ���
            time.sleep(0.7)
            continue # ��ݺ�

def load_filedialog(initdir, tit, filetype):
    root = tk.Tk()
    root.withdraw()
    root.filename = filedialog.askopenfilename(initialdir=initdir, title=tit, filetypes=filetype)
    root.destroy()
    return root.filename

def load_dirdialog(initdir, tit):
    root = tk.Tk()
    root.withdraw()
    dir_path = filedialog.askdirectory(parent=root, initialdir=initdir, title=tit) + '/'
    root.destroy()
    return dir_path

def load_check_point():
    msgbox = messagebox.askquestion('Checkpoint file load', 'Do you want to load checkpoint file?\nIf you want to load then press the "YES" button.')
    if msgbox == 'yes':
        file_name = load_filedialog('./', 'Select checkpoint file', (('checkpoint file','*.chkpt'),('all files','*.*')))
        print('Checkpoint file loaded: ' + file_name)
        with open(file_name, 'r') as file:
            return file.read()
    else:
        print('Checkpoint file not loaded.')
        return ''

def save_check_point(current_position):
    with open('position.chkpt', 'w') as file:
        file.write(current_position)


appname = 'py.exe'
xpos = 1000
ypos = 0
width = 1000
length = 300
win32gui.EnumWindows(enumHandler, None)

image_dir_path = load_dirdialog('/', 'Please select a image file directory')
print("Image directory path : ", image_dir_path)
if not image_dir_path:
    quit()

imageList = get_image_file_list(image_dir_path) # �̹��� ���� ��� ��������

check_point = load_check_point() # üũ ����Ʈ �ҷ�����

btnImagePath = load_dirdialog('/', 'Please select a button image file directory')
print("Button image directory path : ", btnImagePath)
if not btnImagePath:
    quit()
    
isLoadedCheckpoint = False
for image in tqdm(imageList):
    # ���������� �۾� �ߴ� ��ġ ã��
    if not check_point in '' and not isLoadedCheckpoint:
        if image in check_point:
            image = check_point
            isLoadedCheckpoint = True
        else:
            continue

    print('Current image: ' + image)
    
    # ���� Ȯ���ڰ� .miss �̸�
    if image.endswith('.miss'):
        print ('Warning: image file missed: ' + image)
        image_button_click(btnImagePath + 'next.png')
        time.sleep(0.2) # ������ �߰�
        continue # �ǳʶٱ�
    else:
        save_check_point(image) # üũ ����Ʈ �����ϱ�

    image_button_click(btnImagePath + 'picture.png')
    time.sleep(0.1) # ������ �߰�
    py.press('Enter')  # ���� �������ε�� ������ �˾��� ����� �ۼ���
    
    image_button_click(btnImagePath + 'register.png')

    image_dir_path = image_dir_path.replace('/', '\\') # �̹��� �ҷ��� �� ��� �ν� ���� �ذ�
    copy_and_paste(image_dir_path + image) # �̹��� ���� ��� ����
    time.sleep(0.1) # ������ �߰�
    py.hotkey('ctrl', 'v') # ������ ��� �ٿ��ֱ�
    time.sleep(0.7) # ������ �߰�
    py.press('Enter') # ����
    time.sleep(0.7) # ������ �߰�
    py.press('Enter') # ���ε� �Ϸ�â Ȯ��
    time.sleep(0.3) # ������ �߰�

    image_button_click(btnImagePath + 'close.png')
    time.sleep(0.2) # ������ �߰�
    image_button_click(btnImagePath + 'next.png')
    time.sleep(0.2) # ������ �߰�

print('Finish !!! ^0^')
