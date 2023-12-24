import math
import os
import shutil
import time

import cv2 as cv
import keyboard
import pyautogui
import win32con
import win32gui

from PIL import ImageGrab
from loguru import logger


class WinGUI:
    def __init__(self, window_name):
        self.window_name = window_name
        self.work_screen_path = "./temp/work_screen.png"
        self.app_screen_path = "./temp/app.png"
        self.img_folder_path = "./image" 


    def get_app_screenshot(self):
        original_position, handle = get_window_pos(self.window_name)
 
        # 设为高亮
        win32gui.ShowWindow(handle, True)
        win32gui.SetForegroundWindow(handle)
        time.sleep(1)

        img_ready = ImageGrab.grab(original_position)
        img_ready.save(self.app_screen_path)
        return original_position


    def get_workscreen_screenshot(self):
        screenshot = ImageGrab.grab()

        if screenshot:
            screenshot.save(self.work_screen_path)
            return screenshot
        return None


    def move_and_click(self, x, y):
        pyautogui.moveTo(x, y)
        pyautogui.click(x, y)
        time.sleep(1)


    def click_icon(self, icon_path):
        x, y = self.locate_icon(icon_path)
        self.move_and_click(x, y)


    def locate_icon(
        self,
        img_name,
        x_start_ratio=0,
        x_end_ratio=1.0,
        y_start_ratio=0,
        y_end_ratio=1.0,
        try_number=3,
    ):
        # return the coordinates of the center point of the target icon
        print(img_name)
        obj_path = os.path.join(self.img_folder_path, img_name)

        result_x, result_y = -1, -1
        for i in range(try_number):
            print()
            print(f"第{i + 1}次查找")
            # print("mouse: ", pyautogui.position())
            x_init, y_init = self.get_app_screenshot()[:2]
            source = cv.imread(self.app_screen_path)

            # height x width x RGB
            h, w, d = source.shape
            # print(f"original size: (width {w}, height {h}")

            # x direction: left to right
            # y direction: top to bottom
            x_start = math.floor(w * x_start_ratio)
            x_end = math.ceil(w * x_end_ratio)
            y_start = math.floor(h * y_start_ratio)
            y_end = math.floor(h * y_end_ratio)
            # print(f"crop_location: {x_start, y_start, x_end, y_end}")
            source = source[y_start : y_end + 1, x_start : x_end + 1]
            # print(f"crop size: width {source.shape[1]}, height {source.shape[0]}")

            cv.imwrite(self.app_screen_path, source)

            template = cv.imread(obj_path)
            # print(f"template_size: {template.shape}")

            result = cv.matchTemplate(source, template, cv.TM_CCOEFF_NORMED)

            similarity = cv.minMaxLoc(result)[1]

            if similarity < 0.90:
                logger.info("low similarity")
                logger.info(cv.minMaxLoc(result)[3])
            else:
                pos_start = cv.minMaxLoc(result)[3]

                result_x = (
                    x_init + x_start + int(pos_start[0]) + int(template.shape[1] / 2)
                )
                result_y = (
                    y_init + y_start + int(pos_start[1]) + int(template.shape[0] / 2)
                )
                break
        
        return result_x, result_y


    def check_icon(
        self,
        img_name,
        x_start_ratio=0,
        x_end_ratio=1.0,
        y_start_ratio=0,
        y_end_ratio=1.0,
    ):
        x, y = self.locate_icon(
            img_name,
            x_start_ratio,
            x_end_ratio,
            y_start_ratio,
            y_end_ratio,
            try_number=3,
        )
        if x < 0 or y < 0:
            return False, x, y
        return True, x, y


def get_window_pos(name):
    handle = win32gui.FindWindow(0, name)
    # 获取窗口句柄
    if handle == 0:
        return None
    else:
        # 返回坐标值和handle
        win32gui.SendMessage(handle, win32con.WM_SYSCOMMAND, win32con.SC_RESTORE, 0)
        time.sleep(1)
        return win32gui.GetWindowRect(handle), handle


def move_files(original_folder, target_folder, suffix_list=[]):
    data_list = []
    dir_list = os.listdir(original_folder)
    if len(suffix_list) == 0:
        data_list = dir_list
    else:
        for file_name in dir_list:
            for suffix in suffix_list:
                if file_name.endswith(suffix):
                    data_list.append(file_name)
                    # print(file_name)
    
    if len(data_list) == 0:
        return

    cur_time = time.strftime("%Y-%m-%d_%H_%M_%S")
    logger.info(cur_time)
    target_folder_new = os.path.join(target_folder, cur_time)
    os.mkdir(target_folder_new)
    for file_name in data_list:
        source = os.path.join(original_folder, file_name)
        destination = os.path.join(target_folder_new, file_name)
        shutil.move(source, destination)


# need change
def running_program(window_name, original_folder, target_folder, cycle_number=-1, suffix_list=[]):
    exit_flag = False
    def on_key_event(event):
        nonlocal exit_flag
        if event.name == 'q':
            logger.info("terminated by user")
            exit_flag = True

    keyboard.on_press(on_key_event)

    app = WinGUI(window_name)
    logger.info(window_name)
    
    cycle_count = 0
    while not exit_flag:
        try:
            if is_test_over(app):
                move_files(original_folder, target_folder, suffix_list)
                logger.info(f"Cycle {cycle_count} is finished")
                if cycle_number > 0 and cycle_count >= cycle_number:
                    logger.info(f"finished {cycle_count} cycles!")
                    return
                cycle_count += 1
                # write your operations
                # usually use app functions

                # operation end
        except Exception as err:
            logger.info(err)

        # process abnormal cases
        try:
            print("test")
        except:
            print()
        
        time.sleep(1)


# need change
def is_test_over(app):
    valid1, _, _ = app.check_icon("running_1.png")
    if valid1:
        return False
    
    valid2, _, _ = app.check_icon("running_2.png")
    # if valid2:
    #     return False
    # else:
    #     return True

    return not valid2


if __name__ == "__main__":
    pyautogui.FAILSAFE = False
    pyautogui.PAUSE = 1

    logger.add("dev.log", rotation="10 MB")
    
    # ---------- need change -------------
    # original data folder
    original_folder = "C:/Users/Public/Documents/Data"
    # target data folder
    target_folder = r"C:\Users\Joey\Desktop\data"
    suffix_list = []

    window_name = "window name of your application"
    # cycle_number = -1 means infinite loop
    cycle_number = -1
    # ------------------------------------

    running_program(window_name, original_folder, target_folder, cycle_number, suffix_list)
