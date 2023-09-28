# import cv2
# import numpy as np
# import os
# import pyautogui

# output = "video.avi"
# img = pyautogui.screenshot()
# img = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
# #get info from img
# height, width, channels = img.shape
# # Define the codec and create VideoWriter object
# fourcc = cv2.VideoWriter_fourcc(*'mp4v')
# out = cv2.VideoWriter(output, fourcc, 20.0, (width, height))

# while(True):
#  try:
#   img = pyautogui.screenshot()
#   image = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
#   out.write(image)
#   StopIteration(0.5)
#  except KeyboardInterrupt:
#   break

# out.release()
# cv2.destroyAllWindows()


import numpy as np
import cv2
from mss import mss

# 设置捕获区域的坐标和大小
bounding_box = {'top': 100, 'left': 0, 'width': 400, 'height': 300}

# 创建MSS（Python屏幕捕获库）对象
sct = mss()

# 视频写入设置
fourcc = cv2.VideoWriter_fourcc(*'XVID')
out = cv2.VideoWriter('screen_capture.avi', fourcc, 20.0, (400, 300))

while True:
    # 捕获屏幕区域
    sct_img = sct.grab(bounding_box)
    
    # 将捕获的图像转换为OpenCV格式
    frame = np.array(sct_img)
    
    # 写入视频文件
    out.write(frame)

    # 显示屏幕捕获画面（可选）
    cv2.imshow('Screen Capture', frame)

    # 检测键盘输入，如果按下 'q' 键则退出循环
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# 释放VideoWriter对象
out.release()

# 关闭OpenCV窗口
cv2.destroyAllWindows()
