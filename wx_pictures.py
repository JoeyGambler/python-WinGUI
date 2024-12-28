import os
import time

import keyboard
import pyautogui
import pyperclip

from loguru import logger

from gui import WinGUI 

# need change
def running_program(window_name, target_folder,  cycle_number=-1, prefix=""):
    exit_flag = False
    def on_key_event(event):
        nonlocal exit_flag
        if event.name == 'q':
            logger.info("terminated by user")
            exit_flag = True

    keyboard.on_press(on_key_event)

    app = WinGUI(window_name)
    logger.info(window_name)

    last_image = ""

    
    cycle_count = 0
    while not exit_flag:
        try:
            if cycle_number > 0 and cycle_count >= cycle_number:
                logger.info(f"finished {cycle_count} cycles!")
                return
            # write your operations
            # usually use app functions
            app.get_app_screenshot()
            
            # 模拟按下 Ctrl+S
            pyautogui.hotkey('ctrl', 's')
            time.sleep(1)  # 等待操作完成

            # 模拟按下 Ctrl+C
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.5)  # 等待复制操作完成

            # 获取剪贴板内容
            copied_content = pyperclip.paste()

            # 检查是否到达了第一张图片
            if copied_content == last_image:
                pyautogui.hotkey('esc')
                logger.info(f"This is the first image!")
                logger.info(f"finished {cycle_count} cycles!")
                return
            
            last_image = copied_content

            # 修改前缀
            if prefix:
                copied_content = copied_content.replace("微信图片_", prefix)

            # 修改文件路径
            image_path = os.path.join(target_folder, copied_content)

            # 将新的文件路径写入剪贴板
            pyperclip.copy(image_path)

            # 模拟 Ctrl+V 粘贴操作
            pyautogui.hotkey('ctrl', 'v')

            # 确认保存
            pyautogui.press('enter')
            time.sleep(0.5)

            # 确认覆盖重复的图片
            pyautogui.press("y")
            logger.info(f"save {image_path}")
            time.sleep(0.5)
            
            # 切换到上一张图片
            pyautogui.hotkey('left')

            cycle_count += 1

        except Exception as err:
            logger.info(err)
        
        time.sleep(1)

if __name__ == "__main__":
    logger.add("dev.log", rotation="10 MB")
    
    # ---------- need change -------------
    # target data folder
    window_name = "图片查看"
    target_folder = r"C:\Users\Joey\Desktop\python-WinGUI" 
    # cycle_number = -1 means infinite loop
    cycle_number = -1

    # ------------------------------------
    running_program(window_name, target_folder, cycle_number, prefix="")
