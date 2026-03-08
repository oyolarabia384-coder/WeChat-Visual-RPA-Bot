import time
import random
import cv2
import numpy as np
import win32gui, win32con, win32api, ctypes
import pyautogui
import pyperclip
import os
import json
from openai import OpenAI
from message import MessageProcessor, calculate_dhash, get_hamming_distance

SEARCH_REF_Y, SEARCH_REF_X = 34, 1869
SB_ICON_X = (1788, 1809)
SB_RED_X = (1798, 1814)
SIDEBAR_SLOTS = [(87, 105, 79, 87), (135, 153, 127, 135), (183, 201, 175, 183), (231, 249, 223, 231), (279, 297, 271, 279), (327, 345, 319, 327), (375, 393, 367, 375), (422, 442, 414, 422), (523, 543, 515, 523), (575, 587, 567, 575)]
ITEM_H = 68
LIST_AV_X = (1841, 1877)
LIST_RED_X = (1876, 1885)
LIST_FIRST_AV_Y = (84, 120)
LIST_FIRST_RED_Y = (76, 91)
SCROLL_CLICK_X1, SCROLL_CLICK_Y1 = 2060, 75
SCROLL_CLICK_X2, SCROLL_CLICK_Y2 = 2066, 126

class WeChatScanner:
    def __init__(self, template_path="search_fingerprint.png"):
        self.dpi_scale = self.get_dpi_scale()
        self.hwnd = None
        self.template = cv2.imread(template_path, cv2.IMREAD_GRAYSCALE)
        self.anchor_abs_x, self.anchor_abs_y, self.offset_y = 0, 0, 0
        self.win_x, self.win_y = 0, 0
        self.first_page_offset_applied = False
        self.running = True
        self.callback = None
        self.keyword_rules = []
        self.reply_strategy = "keyword"
        self.non_text_message_action = "ignore"
        self.user_info = {"nickname": "等待连接...", "avatar_path": None}
        self.openai_config = {"key": "", "url": "https://api.openai.com/v1/chat/completions", "model": "gpt-3.5-turbo", "system": ""}
        self.history_rounds = 10

    def get_dpi_scale(self):
        ctypes.windll.user32.SetProcessDPIAware()
        hdc = ctypes.windll.user32.GetDC(0)
        scale = ctypes.windll.gdi32.GetDeviceCaps(hdc, 88) / 96.0
        ctypes.windll.user32.ReleaseDC(0, hdc)
        return scale

    def find_wechat(self):
        hwnds = []
        def check_window(h, l):
            if win32gui.GetWindowText(h) == "微信": l.append(h)
        win32gui.EnumWindows(check_window, hwnds)
        if not hwnds: return None
        for h in hwnds:
            if win32gui.GetClassName(h) == "WeChatMainWndForPC": return h
        return hwnds[0]

    def initialize(self):
        self.hwnd = self.find_wechat()
        if not self.hwnd: return False
        
        win32gui.ShowWindow(self.hwnd, win32con.SW_RESTORE)
        sw = win32api.GetSystemMetrics(0)
        win32gui.SetWindowPos(self.hwnd, win32con.HWND_TOPMOST, sw - 800, 0, 800, 625, win32con.SWP_SHOWWINDOW)
        time.sleep(0.5)
        
        rect = win32gui.GetWindowRect(self.hwnd)
        self.win_x, self.win_y = rect[0], rect[1]
        screenshot = pyautogui.screenshot(region=(int(rect[0]*self.dpi_scale), int(rect[1]*self.dpi_scale), int(800*self.dpi_scale), int(625*self.dpi_scale)))
        
        anchor = self.find_search_anchor(screenshot)
        if not anchor: return False
        
        ax_local, ay_local, _ = anchor
        self.anchor_abs_x = ax_local + self.win_x
        self.anchor_abs_y = ay_local + self.win_y
        self.offset_y = self.anchor_abs_y - SEARCH_REF_Y
        self.first_page_offset_applied = False
        return True

    def find_search_anchor(self, screenshot):
        gray_screen = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2GRAY)
        res = cv2.matchTemplate(gray_screen, self.template, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)
        if max_val > 0.8: return max_loc[0] / self.dpi_scale, max_loc[1] / self.dpi_scale, max_val
        return None

    def check_red(self, img_np, x1, y1, x2, y2):
        if x1 < 0 or y1 < 0 or x2 > img_np.shape[1] or y2 > img_np.shape[0]: return False, 0
        roi = img_np[y1:y2, x1:x2]
        if roi.size == 0: return False, 0
        hsv = cv2.cvtColor(roi, cv2.COLOR_RGB2HSV)
        mask = cv2.bitwise_or(cv2.inRange(hsv, (0, 100, 100), (10, 255, 255)), cv2.inRange(hsv, (160, 100, 100), (180, 255, 255)))
        red_pixel_count = np.count_nonzero(mask)
        return red_pixel_count > 5, red_pixel_count
    
    def is_new_message(self, red_pixel_count): return red_pixel_count > 60

    def get_screenshot(self):
        rect = win32gui.GetWindowRect(self.hwnd)
        return np.array(pyautogui.screenshot(region=(int(rect[0]*self.dpi_scale), int(rect[1]*self.dpi_scale), int(800*self.dpi_scale), int(625*self.dpi_scale))))

    def scan_sidebar(self):
        img_np = self.get_screenshot()
        has_new_message = False
        for idx, (iy1, iy2, ry1, ry2) in enumerate(SIDEBAR_SLOTS):
            if idx != 0: continue
            cur_y1, cur_y2 = iy1 + self.offset_y, iy2 + self.offset_y
            red_y1, red_y2 = ry1 + self.offset_y, ry2 + self.offset_y
            rpx1, rpy1 = int((SB_RED_X[0] - self.win_x) * self.dpi_scale), int((red_y1 - self.win_y) * self.dpi_scale)
            rpx2, rpy2 = int((SB_RED_X[1] - self.win_x) * self.dpi_scale), int((red_y2 - self.win_y) * self.dpi_scale)
            has_red, count = self.check_red(img_np, rpx1, rpy1, rpx2, rpy2)
            if has_red and self.is_new_message(count): has_new_message = True
        return has_new_message

    def scan_contact_list(self, y_offset=0):
        img_np = self.get_screenshot()
        new_message_info = []
        for i in range(8):
            off = i * ITEM_H
            ry1 = LIST_FIRST_RED_Y[0] + off + self.offset_y + y_offset
            ry2 = LIST_FIRST_RED_Y[1] + off + self.offset_y + y_offset
            av_y1 = LIST_FIRST_AV_Y[0] + off + self.offset_y + y_offset
            av_y2 = LIST_FIRST_AV_Y[1] + off + self.offset_y + y_offset
            lr_px1, lr_py1 = int((LIST_RED_X[0] - self.win_x) * self.dpi_scale), int((ry1 - self.win_y) * self.dpi_scale)
            lr_px2, lr_py2 = int((LIST_RED_X[1] - self.win_x) * self.dpi_scale), int((ry2 - self.win_y) * self.dpi_scale)
            has_red, count = self.check_red(img_np, lr_px1, lr_py1, lr_px2, lr_py2)
            if has_red and self.is_new_message(count): new_message_info.append((i, LIST_AV_X[0], LIST_AV_X[1], av_y1, av_y2))
        return new_message_info

    def click_contact(self, av_x1, av_x2, av_y1, av_y2):
        time.sleep(random.uniform(0, 1))
        pyautogui.click(random.uniform(av_x1, av_x2), random.uniform(av_y1, av_y2))

    def click_scroll_area(self):
        time.sleep(random.uniform(0, 1))
        scroll_x1 = self.anchor_abs_x + (SCROLL_CLICK_X1 - SEARCH_REF_X)
        scroll_y1 = self.anchor_abs_y + (SCROLL_CLICK_Y1 - SEARCH_REF_Y)
        scroll_x2 = self.anchor_abs_x + (SCROLL_CLICK_X2 - SEARCH_REF_X)
        scroll_y2 = self.anchor_abs_y + (SCROLL_CLICK_Y2 - SEARCH_REF_Y)
        pyautogui.click(random.uniform(scroll_x1, scroll_x2), random.uniform(scroll_y1, scroll_y2))

    def scroll_to_top(self):
        self.click_scroll_area()
        time.sleep(0.2)
        pyautogui.press('home')
        time.sleep(0.3)
        self.first_page_offset_applied = False

    def scroll_page_down(self):
        pyautogui.press('pagedown')
        time.sleep(0.3)
    
    def _process_message(self):
        if self.callback: self.callback('log', "📩 开始处理消息...")
        processor = MessageProcessor(self.hwnd, self.win_x, self.win_y, self.dpi_scale)
        screenshot = processor.get_screenshot()
        contact_id, contact_name = processor.get_contact_name_smart(screenshot, self.database)
        
        latest_messages_list = processor.extract_latest_messages()
        if not latest_messages_list: return
        
        for msg in latest_messages_list:
            if self.database: self.database.save_message(contact_id, contact_name, "user", msg)
            if self.callback: self.callback('message', f"{contact_name}|user|{msg}")

        combined_text = "\n".join(latest_messages_list)
        if any("[非文本消息]" in msg for msg in latest_messages_list):
            if self.non_text_message_action == "ignore": return
            reply = "抱歉，我现在看不到图片/听不了语音，能麻烦您打字说一下吗？"
        else:
            if self.reply_strategy == "keyword":
                reply = self._get_keyword_reply(combined_text) or f"收到消息：\n{combined_text}"
            else:
                reply = self._get_openai_reply(combined_text, contact_id)

        processor.send_reply(reply)
        if self.database:
            self.database.save_message(contact_id, contact_name, "我", reply)
            if self.callback:
                self.callback('log', "✅ 机器人回复已保存")
                self.callback('message', f"{contact_name}|bot|{reply}")
            
        self.close_chat_dialog()

    def click_sidebar_icon(self, icon_index):
        time.sleep(random.uniform(0, 1))
        iy1, iy2, _, _ = SIDEBAR_SLOTS[icon_index]
        pyautogui.click(random.uniform(SB_ICON_X[0], SB_ICON_X[1]), random.uniform(iy1 + self.offset_y, iy2 + self.offset_y))

    def stop(self): self.running = False
    def set_keyword_rules(self, rules): self.keyword_rules = rules
    def set_reply_strategy(self, strategy): self.reply_strategy = strategy
    def set_non_text_message_action(self, action): self.non_text_message_action = action
    def set_openai_config(self, config): 
        if config: self.openai_config.update(config)
    def set_history_rounds(self, rounds): self.history_rounds = rounds
    
    def _get_keyword_reply(self, text):
        for keyword, reply in self.keyword_rules:
            if keyword in text: return reply
        return None
    
    def _get_openai_reply(self, text, contact_id):
        try:
            client = OpenAI(api_key=self.openai_config['key'], base_url=self.openai_config['url'])
            messages = [{"role": "system", "content": self.openai_config['system']}]
            if self.database: messages.extend(self.database.get_context(contact_id, limit=self.history_rounds))
            messages.append({"role": "user", "content": text})
            return client.chat.completions.create(model=self.openai_config['model'], messages=messages, temperature=0.7, max_tokens=500).choices[0].message.content.strip()
        except Exception as e:
            return "抱歉，AI回复失败，请稍后再试。"
    
    def close_chat_dialog(self):
        img_np = self.get_screenshot()
        gray_contact_found = False
        for i in range(8):
            off = i * ITEM_H
            av_y1, av_y2 = LIST_FIRST_AV_Y[0] + off + self.offset_y, LIST_FIRST_AV_Y[1] + off + self.offset_y
            upper_gray_y1, upper_gray_y2, lower_gray_y1, lower_gray_y2 = av_y1 - 10, av_y1, av_y2, av_y2 + 10
            av_px1, av_px2 = int((LIST_AV_X[0] - self.win_x) * self.dpi_scale), int((LIST_AV_X[1] - self.win_x) * self.dpi_scale)
            upper_py1, upper_py2 = int((upper_gray_y1 - self.win_y) * self.dpi_scale), int((upper_gray_y2 - self.win_y) * self.dpi_scale)
            lower_py1, lower_py2 = int((lower_gray_y1 - self.win_y) * self.dpi_scale), int((lower_gray_y2 - self.win_y) * self.dpi_scale)
            
            gray_found = False
            if av_px1 >= 0 and upper_py1 >= 0 and av_px2 <= img_np.shape[1] and upper_py2 <= img_np.shape[0]:
                if np.all(np.mean(img_np[upper_py1:upper_py2, av_px1:av_px2], axis=(0, 1)) >= 220): gray_found = True
            if not gray_found and av_px1 >= 0 and lower_py1 >= 0 and av_px2 <= img_np.shape[1] and lower_py2 <= img_np.shape[0]:
                if np.all(np.mean(img_np[lower_py1:lower_py2, av_px1:av_px2], axis=(0, 1)) >= 220): gray_found = True
            
            if gray_found:
                time.sleep(random.uniform(0, 1))
                pyautogui.click(random.uniform(LIST_AV_X[0], LIST_AV_X[1]), random.uniform(av_y1 - 10, av_y2 + 10))
                time.sleep(0.5)
                gray_contact_found = True
                break
        
        if not gray_contact_found:
            i = 7
            off = i * ITEM_H
            av_y1, av_y2 = LIST_FIRST_AV_Y[0] + off + self.offset_y, LIST_FIRST_AV_Y[1] + off + self.offset_y
            time.sleep(random.uniform(0, 1))
            pyautogui.click(random.uniform(LIST_AV_X[0], LIST_AV_X[1]), random.uniform(av_y1 - 10, av_y2 + 10))
            time.sleep(random.uniform(1, 1.5))
            pyautogui.click(random.uniform(LIST_AV_X[0], LIST_AV_X[1]), random.uniform(av_y1 - 10, av_y2 + 10))
            time.sleep(0.5)

    # --- 核心剥离：专门处理用户信息的逻辑 ---
    def check_and_update_user_info(self, callback):
        config_path = "config.json"
        need_get_user_info = True
        current_avatar_hash = None
        
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    if 'user_info' in config and 'avatar_hash' in config['user_info']:
                        self.user_info = config['user_info']
                        current_avatar_hash = self.user_info['avatar_hash']
                        
                        avatar_x, avatar_y = self.win_x + 38, self.win_y + 42
                        avatar_region = (int((avatar_x - 20) * self.dpi_scale), int((avatar_y - 20) * self.dpi_scale), int(40 * self.dpi_scale), int(40 * self.dpi_scale))
                        temp_avatar_path = "temp_avatar.png"
                        pyautogui.screenshot(region=avatar_region).save(temp_avatar_path)
                        
                        new_avatar_hash = calculate_dhash(temp_avatar_path)
                        dist = get_hamming_distance(new_avatar_hash, current_avatar_hash)
                        if callback: callback('log', f"📊 头像哈希距离: {dist}")
                        
                        if dist <= 5:
                            need_get_user_info = False
                        if os.path.exists(temp_avatar_path): os.remove(temp_avatar_path)
            except Exception as e:
                if callback: callback('log', f"检查配置异常: {e}")

        if need_get_user_info:
            if callback: callback('log', "🔄 获取新用户信息...")
            win32gui.SetWindowPos(self.hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            
            avatar_x, avatar_y = self.win_x + 38, self.win_y + 42
            time.sleep(random.uniform(0, 1))
            pyautogui.click(avatar_x + random.uniform(-10, 10), avatar_y + random.uniform(-10, 10))
            time.sleep(1) 
            
            avatar_region = (int((avatar_x - 20) * self.dpi_scale), int((avatar_y - 20) * self.dpi_scale), int(40 * self.dpi_scale), int(40 * self.dpi_scale))
            avatar_path = "user_avatar.png"
            pyautogui.screenshot(region=avatar_region).save(avatar_path)
            
            nickname_point_x, nickname_point_y = self.win_x + 165, self.win_y + 65
            time.sleep(random.uniform(0, 1))
            pyautogui.doubleClick(nickname_point_x + random.uniform(0, 5), nickname_point_y + random.uniform(-5, 0))
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(0.5)
            nickname = pyperclip.paste().strip()
            
            time.sleep(random.uniform(0, 1))
            pyautogui.press('esc')
            time.sleep(0.5)
            win32gui.SetWindowPos(self.hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            
            avatar_hash = calculate_dhash(avatar_path)
            self.user_info = {"nickname": nickname, "avatar_path": avatar_path, "avatar_hash": avatar_hash}
            
            try:
                with open(config_path, 'r', encoding='utf-8') as f: config = json.load(f)
            except: config = {}
            config['user_info'] = self.user_info
            with open(config_path, 'w', encoding='utf-8') as f: json.dump(config, f, ensure_ascii=False, indent=2)
            
        return self.user_info
    
    def run(self, database=None, callback=None):
            self.database = database
            self.callback = callback
            if not self.hwnd: return
            
            while self.running:
                try:
                    # 1. 扫描侧边栏，有红点直接开干
                    if self.scan_sidebar():
                        
                        # 2. 确保切到“聊天”Tab
                        self.click_sidebar_icon(0)
                        time.sleep(0.5)
                        
                        # ==========================================
                        # 🌟 核心优化：先在当前视野内扫一眼！
                        # ==========================================
                        if self.callback: self.callback('log', "👀 先在当前视野内查找红点...")
                        new_info = self.scan_contact_list(y_offset=0)
                        
                        if new_info:
                            if self.callback: self.callback('log', "🎯 当前视野内直接捕获红点！免除回顶操作。")
                            _, av_x1, av_x2, av_y1, av_y2 = new_info[0]
                            self.click_contact(av_x1, av_x2, av_y1, av_y2)
                            time.sleep(1)
                            
                            # 防位移二次校验
                            if self.scan_contact_list(y_offset=0):
                                if self.callback: self.callback('log', "⚠️ 列表发生位移，放弃读取，重新锁定...")
                                continue  # 发生位移，直接切入下一轮大循环
                                
                            self._process_message()
                            continue  # 处理完毕，完美进入下一轮轮询
                        
                        # ==========================================
                        # 如果当前视野没找到，说明列表在下方，执行回顶翻页
                        # ==========================================
                        if self.callback: self.callback('log', "🔍 当前视野未发现，准备回顶地毯式搜索...")
                        
                        # 3. 点击滚轮区域（获取焦点）
                        self.click_scroll_area()
                        time.sleep(0.3)
                        
                        # 4. 回顶，准备重头扫描
                        self.scroll_to_top()
                        
                        # 5. 开始翻页查找流程
                        page, max_pages = 1, 5
                        while page <= max_pages and self.running:
                            y_offset = -4 if (page == 1 and not self.first_page_offset_applied) else 0
                            if page == 1: self.first_page_offset_applied = True
                            
                            found_new_message = False
                            for attempt in range(5):
                                if not self.running: break
                                new_info = self.scan_contact_list(y_offset=y_offset)
                                if new_info:
                                    _, av_x1, av_x2, av_y1, av_y2 = new_info[0]
                                    self.click_contact(av_x1, av_x2, av_y1, av_y2)
                                    time.sleep(1)
                                    
                                    # 防位移二次校验
                                    if self.scan_contact_list(y_offset=y_offset):
                                        if self.callback: self.callback('log', "⚠️ 列表发生位移，放弃读取，重新锁定...")
                                        found_new_message = True
                                        break
                                    
                                    self._process_message()
                                    found_new_message = True
                                    break
                                else:
                                    if attempt < 4: time.sleep(1)
                            
                            if found_new_message: break
                            
                            if page < max_pages:
                                self.scroll_page_down()
                                page += 1
                            else: 
                                if self.callback: self.callback('log', "   已扫描所有页，未发现红点联系人")
                                break
                    
                    time.sleep(2)
                except Exception as e:
                    import traceback
                    if self.callback: self.callback('log', f"\n❌ 错误: {e}\n{traceback.format_exc()}")
                    time.sleep(2)
