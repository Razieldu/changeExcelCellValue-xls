import pyautogui
import time
import pyperclip
# 打开记事本应用程序（如果尚未打开）
pyautogui.hotkey('win', 'r')
pyautogui.write('notepad')
pyautogui.press('enter')
time.sleep(1)  # 等待记事本打开

# 输入文本
chinesePoem = """
這次我離開你，是風，是雨，是夜晚；
你笑了笑，我擺一擺手
一條寂寞的路便展向兩頭了。
念此際你已回到濱河的家居，
想你在梳理長髮或是濕了的外衣，
而我風雨的歸程還正長；
山退得很遠，平蕪拓得更大，
哎，這世界，怕黑暗已真的成型了……
                                                                               
你說，你真傻，多像那放風箏的孩子，
本不該縛它又放它
風箏去了，留一線斷了的錯誤；
書太厚了，本不該掀開扉頁的；
沙灘太長，本不該走出足印的；
雲出自岫谷，泉水滴自石隙，
一切都開始了，而海洋在何處？
「獨木橋」的初遇已成往事了，
如今又已是廣闊的草原了，
我已失去扶持你專寵的權利；
紅與白揉藍與晚天，錯得多美麗，
而我不錯入金果的園林，
卻誤入維特的墓地……
                                                                               
這次我離開你，便不再想見你了，
念此際你已靜靜入睡。
留我們未完的一切，留給這世界，
這世界，我仍體切地踏著，
而已是你底夢境了……
"""

EnglishPoem = """

"""

# pyautogui.typewrite(text_to_type,0.05)
pyperclip.copy(chinesePoem)

time.sleep(2)

# 模拟按下Ctrl+V粘贴
pyautogui.hotkey('Ctrl', 'V')

# 保存文件
pyautogui.hotkey('ctrl', 's')
time.sleep(1)  # 等待保存对话框打开

# 输入文件名并保存
file_name = "my_text_file.txt"
pyautogui.write(file_name)
pyautogui.press('enter')

# 输入保存路径（替换为你的路径）
save_path = "C:\\Users\\jeremydu\\Desktop\\"
pyautogui.write(save_path)
pyautogui.press('enter')
# pyautogui.press('q')
pyautogui.hotkey('win', 'd')
print("文本已保存到记事本中。")
