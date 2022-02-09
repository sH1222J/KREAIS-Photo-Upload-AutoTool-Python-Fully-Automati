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
    while True: # 버튼을 찾고 누를 떄 까지 반복
        try:
            time.sleep(0.7) # 알아서 잘맞추기
            btnImage = py.locateOnScreen(image)  # 화면에서 image를 감지       
            point = py.center(btnImage) # 찾은 이미지의 좌표의 중앙 지점 감지
            print(image + ' ' + str(point)) # 콘솔창에 image 경로, 찾은 좌표를 출력
            time.sleep(0.7) # 알아서 잘맞추기
            py.click(point) # 구한 좌표를 클릭
            break # 반복문 종료
        except: # 버튼을 못찾을 경우
            print('Error: ' + image + ' not found.') # 콘솔창에 image를 못찾았다는 내용 출력
            time.sleep(0.7)
            continue # 재반복

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

imageList = get_image_file_list(image_dir_path) # 이미지 파일 목록 가져오기

check_point = load_check_point() # 체크 포인트 불러오기

btnImagePath = load_dirdialog('/', 'Please select a button image file directory')
print("Button image directory path : ", btnImagePath)
if not btnImagePath:
    quit()
    
isLoadedCheckpoint = False
for image in tqdm(imageList):
    # 마지막으로 작업 했던 위치 찾기
    if not check_point in '' and not isLoadedCheckpoint:
        if image in check_point:
            image = check_point
            isLoadedCheckpoint = True
        else:
            continue

    print('Current image: ' + image)
    
    # 파일 확장자가 .miss 이면
    if image.endswith('.miss'):
        print ('Warning: image file missed: ' + image)
        image_button_click(btnImagePath + 'next.png')
        time.sleep(0.2) # 딜레이 추가
        continue # 건너뛰기
    else:
        save_check_point(image) # 체크 포인트 저장하기

    image_button_click(btnImagePath + 'picture.png')
    time.sleep(0.1) # 딜레이 추가
    py.press('Enter')  # 최초 사진업로드시 나오는 팝업을 대비해 작성함
    
    image_button_click(btnImagePath + 'register.png')

    image_dir_path = image_dir_path.replace('/', '\\') # 이미지 불러올 때 경로 인식 문제 해결
    copy_and_paste(image_dir_path + image) # 이미지 파일 경로 복사
    time.sleep(0.1) # 딜레이 추가
    py.hotkey('ctrl', 'v') # 복사한 경로 붙여넣기
    time.sleep(0.7) # 딜레이 추가
    py.press('Enter') # 열기
    time.sleep(0.7) # 딜레이 추가
    py.press('Enter') # 업로드 완료창 확인
    time.sleep(0.3) # 딜레이 추가

    image_button_click(btnImagePath + 'close.png')
    time.sleep(0.2) # 딜레이 추가
    image_button_click(btnImagePath + 'next.png')
    time.sleep(0.2) # 딜레이 추가

print('Finish !!! ^0^')
