import time
import random
import cv2
import numpy as np
import pyautogui
import pyperclip
import io
import asyncio
from PIL import Image
import winrt.windows.media.ocr as ocr
import winrt.windows.graphics.imaging as imaging
import winrt.windows.storage.streams as streams

def calculate_dhash(image_input):
    """计算图片的 dHash 指纹，支持传入 numpy array 或 文件路径字符串"""
    if image_input is None:
        return None
        
    try:
        # 先判断是不是字符串路径，如果是，先把它读取成 numpy 数组
        if isinstance(image_input, str):  
            import os
            if not os.path.exists(image_input):
                return None
            img = Image.open(image_input)
            image_np = np.array(img)
        else:
            image_np = image_input
            
        # 转换为数组后，再进行安全校验
        if image_np is None or image_np.size == 0:
            return None
            
        resized = cv2.resize(image_np, (9, 8))
        if len(resized.shape) == 3:
            gray = cv2.cvtColor(resized, cv2.COLOR_RGB2GRAY)
        else:
            gray = resized
            
        diff = gray[:, 1:] > gray[:, :-1]
        return hex(sum([2 ** i for (i, v) in enumerate(diff.flatten()) if v]))[2:]
    except Exception as e:
        print(f"计算哈希出错: {e}")
        return None

def get_hamming_distance(hash1, hash2):
    """计算两个十六进制哈希字符串的汉明距离"""
    if not hash1 or not hash2 or len(hash1) != len(hash2):
        return 999
    try:
        return bin(int(hash1, 16) ^ int(hash2, 16)).count('1')
    except ValueError:
        return 999

# 常量定义
TITLE_ROI = (320, 10, 500, 50)
CHAT_AREA_PT = (600, 200)
INPUT_BOX_PT = (500, 520)
BUBBLE_ROI = (380, 60, 720, 480)
GREEN_LOWER = np.array([40, 50, 150])
GREEN_UPPER = np.array([75, 255, 255])

class MessageProcessor:
    def __init__(self, hwnd, win_x, win_y, dpi_scale):
        self.hwnd = hwnd
        self.win_x = win_x
        self.win_y = win_y
        self.dpi_scale = dpi_scale
    
    def get_screenshot(self):
        import win32gui
        rect = win32gui.GetWindowRect(self.hwnd)
        screenshot = pyautogui.screenshot(region=(int(rect[0]*self.dpi_scale), 
                                                int(rect[1]*self.dpi_scale), 
                                                int(800*self.dpi_scale), 
                                                int(625*self.dpi_scale)))
        return np.array(screenshot)
    
    def get_chat_title(self, screenshot_np):
        """同步封装异步的 Windows OCR"""
        async def _do_ocr():
            x1, y1, x2, y2 = [int(v * self.dpi_scale) for v in TITLE_ROI]
            roi = screenshot_np[y1:y2, x1:x2]
            pil_img = Image.fromarray(roi)
            pil_img = pil_img.resize((pil_img.width * 2, pil_img.height * 2), Image.LANCZOS)
            
            img_byte_arr = io.BytesIO()
            pil_img.save(img_byte_arr, format='PNG')
            
            stream = streams.InMemoryRandomAccessStream()
            writer = streams.DataWriter(stream)
            writer.write_bytes(img_byte_arr.getvalue())
            await writer.store_async()
            stream.seek(0)
            
            decoder = await imaging.BitmapDecoder.create_async(stream)
            soft_bitmap = await decoder.get_software_bitmap_async()
            
            engine = ocr.OcrEngine.try_create_from_user_profile_languages()
            if not engine: return "未知联系人"
                
            result = await engine.recognize_async(soft_bitmap)
            if result and result.text:
                clean_name = result.text.replace("\n", "").replace(" ", "").strip()
                return clean_name if clean_name else "未知联系人"
            return "未知联系人"
            
        return asyncio.run(_do_ocr())
    
    def get_contact_name_smart(self, screenshot_np, database=None, draw=False):
        if draw:
            # 创建截图的副本用于标记
            marked_screenshot = screenshot_np.copy()
        else:
            marked_screenshot = None
            
        nx1, ny1, nx2, ny2 = [int(v * self.dpi_scale) for v in TITLE_ROI]
        name_img = screenshot_np[max(0, ny1):ny2, max(0, nx1):nx2]
        name_img_hash = calculate_dhash(name_img) or "unknown_name_hash"
        
        _, white_rects, _ = self._get_bubble_rects(screenshot_np, draw=draw)
        avatar_hash = "unknown_avatar_hash"
        
        if white_rects:
            bx, by, bw, bh = white_rects[0]
            left_offset = int(42 * self.dpi_scale)
            top_offset = int(15 * self.dpi_scale)
            avatar_center_x = bx - left_offset
            avatar_center_y = by + top_offset
            half_w = int(20 * self.dpi_scale)
            
            x1, y1 = avatar_center_x - half_w, avatar_center_y - half_w
            x2, y2 = avatar_center_x + half_w, avatar_center_y + half_w
            avatar_img = screenshot_np[max(0, y1):y2, max(0, x1):x2]
            avatar_hash = calculate_dhash(avatar_img) or "unknown_avatar_hash"
            
        # 强制转换为字符串，防止某些情况下 OCR 返回 None
        raw_nickname = str(self.get_chat_title(screenshot_np))
        
        # 严格判断是否需要物理回退
        if (raw_nickname == "未知联系人" or raw_nickname == "None" or not raw_nickname) and white_rects:
            print("👤 OCR失败或未识别，启动物理提取...")
            import win32gui, win32con
            
            # 【保命机制 1】：暂时取消微信置顶，让名片能正常弹到最前面
            win32gui.SetWindowPos(self.hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            
            raw_nickname = self.fallback_get_contact_name(white_rects, draw=draw, screenshot=marked_screenshot)
            
            # 【保命机制 2】：物理提取完后，恢复微信置顶，保证后续截图正常
            win32gui.SetWindowPos(self.hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        
        if draw and marked_screenshot is not None:
            # 保存标记的截图
            import os
            if not os.path.exists("debug"):
                os.makedirs("debug")
            timestamp = int(time.time())
            save_path = f"debug/wechat_screenshot_{timestamp}.png"
            cv2.imwrite(save_path, cv2.cvtColor(marked_screenshot, cv2.COLOR_RGB2BGR))
            print(f"📸 已保存标记的截图到: {save_path}")

        if database:
            contact_id, final_nickname = database.resolve_contact(avatar_hash, name_img_hash, raw_nickname)
            return contact_id, final_nickname
            
        return -1, raw_nickname
    
    def fallback_get_contact_name(self, white_rects):
        """带有 点后状态校验与回退机制 的物理提取"""
        max_retries = 3
        
        for attempt in range(max_retries):
            # 每次重试前，如果是失败重来的，必须重新获取当前最新的屏幕状态
            if attempt > 0:
                current_screenshot = self.get_screenshot()
                _, new_white_rects, _ = self._get_bubble_rects(current_screenshot)
                if not new_white_rects:
                    return "未知联系人"
                white_rects = new_white_rects
                
            bx, by, bw, bh = white_rects[0]
            left_offset = random.uniform(35, 50)
            top_offset = random.uniform(10, 20)
            avatar_x = bx - int(left_offset * self.dpi_scale)
            avatar_y = by + int(top_offset * self.dpi_scale)
            
            click_avatar_x = self.win_x + avatar_x
            click_avatar_y = self.win_y + avatar_y
            
            # 【Action: 执行点击】
            time.sleep(random.uniform(0, 1))
            pyautogui.click(click_avatar_x, click_avatar_y)
            time.sleep(0.8) # 等待名片弹出动画
            
            # 估算名片昵称位置并双击复制
            name_x = avatar_x + int(133 * self.dpi_scale)
            name_y = avatar_y + int(35 * self.dpi_scale) - int(11 * self.dpi_scale)
            abs_name_x = self.win_x + name_x
            abs_name_y = self.win_y + name_y
            
            time.sleep(random.uniform(0, 0.5))
            pyperclip.copy("___EMPTY___") # 塞入探针
            pyautogui.click(abs_name_x, abs_name_y, clicks=2, interval=0.1)
            time.sleep(0.2)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.3)
            
            name = pyperclip.paste().strip()
            
            # 【Verify & Next Step: 点后状态检测】
            if name != "___EMPTY___" and name != "":
                # 状态正确！
                print(f"✅ 第 {attempt+1} 次点击正确，成功提取昵称: {name}")
                # 清理现场：按 ESC 正常关闭名片，进行下一步
                pyautogui.press('esc')
                time.sleep(0.2)
                return name
            else:
                # 【Rollback: 状态错误，执行回退】
                # 如果读到探针，说明刚才点错地方了（比如点到了空白处没弹名片，或者点到了文字弹出了右键菜单）
                print(f"⚠️ 第 {attempt+1} 次点击检测失败（名片未弹出或位置偏移）。")
                print("🔄 执行状态回退：发送 ESC 消除错误弹窗/菜单，恢复初始状态...")
                
                # 发送 ESC 键。如果弹出了右键菜单，ESC会关掉它；如果弹出了别的，ESC也会关掉它。
                # 确保界面恢复到点击前的干净状态。
                pyautogui.press('esc')
                time.sleep(0.5)
                # 继续下一次循环：重新截图 -> 重新定位 -> 再次点击
                
        return "未知联系人"
    
    def _get_bubble_rects(self, screenshot_np, draw=False):
        x1, y1, x2, y2 = [int(v * self.dpi_scale) for v in BUBBLE_ROI]
        roi = screenshot_np[y1:y2, x1:x2]
        
        hsv = cv2.cvtColor(roi, cv2.COLOR_RGB2HSV)
        mask_green = cv2.inRange(hsv, GREEN_LOWER, GREEN_UPPER)
        
        gray = cv2.cvtColor(roi, cv2.COLOR_RGB2GRAY)
        _, mask_white = cv2.threshold(gray, 250, 255, cv2.THRESH_BINARY)
        
        kernel_close = np.ones((15, 45), np.uint8)
        mask_green = cv2.morphologyEx(mask_green, cv2.MORPH_CLOSE, kernel_close)
        mask_white = cv2.morphologyEx(mask_white, cv2.MORPH_CLOSE, kernel_close)
        
        contours_green, _ = cv2.findContours(mask_green, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        max_green_y = -1
        green_rects = []
        for cnt in contours_green:
            if cv2.contourArea(cnt) > 800:
                x, y, w, h = cv2.boundingRect(cnt)
                green_rects.append((x + x1, y + y1, w, h))
                if (y + h) > max_green_y:
                    max_green_y = y + h
                    
        contours_white, _ = cv2.findContours(mask_white, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        white_rects = []
        for cnt in contours_white:
            if cv2.contourArea(cnt) > 200:
                x, y, w, h = cv2.boundingRect(cnt)
                white_rects.append((x + x1, y + y1, w, h))
                if draw:
                    cv2.rectangle(screenshot_np, (x + x1, y + y1), (x + x1 + w, y + y1 + h), (0, 255, 0), 2)
                    cv2.putText(screenshot_np, f"White Bubble", (x + x1, y + y1 - 5), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 1)
                
        white_rects.sort(key=lambda rect: rect[1]) 
        abs_green_y = (max_green_y + y1) if max_green_y != -1 else -1
        
        return abs_green_y, white_rects, green_rects
    
    def extract_latest_messages(self):
        chat_area_x = random.uniform(500, 670)
        chat_area_y = random.uniform(100, 400)
        click_x, click_y = self.win_x + int(chat_area_x * self.dpi_scale), self.win_y + int(chat_area_y * self.dpi_scale)
        
        time.sleep(random.uniform(0, 1))
        pyautogui.click(click_x, click_y)
        time.sleep(0.3)

        max_scrolls = 3
        scroll_up_count = 0
        found_green_bubble = False
        
        for scroll in range(max_scrolls):
            screenshot = self.get_screenshot()
            max_green_y, white_rects, green_rects = self._get_bubble_rects(screenshot)
            
            if max_green_y != -1:
                found_green_bubble = True
                break
                
            prev_rects = white_rects + green_rects
            pyautogui.press('pageup')
            time.sleep(0.8) 
            
            _, new_white_rects, new_green_rects = self._get_bubble_rects(self.get_screenshot())
            if prev_rects == (new_white_rects + new_green_rects):
                break
            scroll_up_count += 1
            
        all_new_messages = []
        
        for page in range(scroll_up_count + 1):
            screenshot = self.get_screenshot()
            max_green_y, white_rects, _ = self._get_bubble_rects(screenshot)
            
            target_bubbles = []
            for bx, by, bw, bh in white_rects:
                if page == 0 and found_green_bubble and max_green_y != -1:
                    if (by + bh) > max_green_y:
                        target_bubbles.append((bx, by, bw, bh))
                else:
                    target_bubbles.append((bx, by, bw, bh))
                    
            page_messages = []
            for bx, by, bw, bh in target_bubbles:
                center_x = bx + (bw / 2.0)
                center_y = by + (bh / 2.0)
                click_x = self.win_x + center_x + random.uniform(-1, 1)
                click_y = self.win_y + center_y + random.uniform(-1, 1)
                
                time.sleep(random.uniform(0.1, 0.3))
                pyautogui.doubleClick(click_x, click_y)
                time.sleep(random.uniform(0.15, 0.35))
                
                pyperclip.copy("___EMPTY___")
                pyautogui.hotkey('ctrl', 'c')
                time.sleep(random.uniform(0.3, 0.5))
                
                text = pyperclip.paste().strip()
                if text == "___EMPTY___":
                    text = "[非文本消息]"
                page_messages.append(text)
                
            if not all_new_messages:
                all_new_messages = page_messages
            else:
                max_overlap = 0
                for i in range(1, min(len(all_new_messages), len(page_messages)) + 1):
                    if all_new_messages[-i:] == page_messages[:i]:
                        max_overlap = i
                all_new_messages.extend(page_messages[max_overlap:])
                
            if page < scroll_up_count:
                pyautogui.click(self.win_x + int(CHAT_AREA_PT[0] * self.dpi_scale), self.win_y + int(CHAT_AREA_PT[1] * self.dpi_scale))
                pyautogui.press('pagedown')
                time.sleep(0.8) 

        return all_new_messages
    
    def send_reply(self, text):
        input_box_x = random.uniform(500, 540)
        input_box_y = random.uniform(520, 560)
        click_x, click_y = self.win_x + int(input_box_x * self.dpi_scale), self.win_y + int(input_box_y * self.dpi_scale)
        
        time.sleep(random.uniform(0, 1))
        pyautogui.click(click_x, click_y)
        time.sleep(0.3)
        pyperclip.copy(text)
        time.sleep(0.2)
        
        time.sleep(random.uniform(0, 1))
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.5)
        
        time.sleep(random.uniform(0, 1))
        pyautogui.press('enter')
        time.sleep(0.5)
        return True
