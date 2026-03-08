import sys
import os
import json
import threading
import datetime
import traceback
import psutil

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, 
    QTabWidget, QLabel, QSpinBox, QRadioButton, QButtonGroup, 
    QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QLineEdit, 
    QTextEdit, QPushButton, QMenu, QDialog, QFormLayout, QMessageBox,
    QHeaderView, QFileDialog, QComboBox, QProgressDialog
)
from PyQt6.QtGui import QPixmap, QColor, QFont, QIcon
from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QObject

from Scanner import WeChatScanner
from sqlite import WeChatDatabase

try:
    import pandas as pd
except ImportError:
    pd = None

class ScannerSignals(QObject):
    message_received = pyqtSignal(str)
    log_added = pyqtSignal(str)
    count_updated = pyqtSignal()
    user_info_updated = pyqtSignal(dict)
    warning_emitted = pyqtSignal(str, str)
    stopped = pyqtSignal()

class MessageItemWidget(QWidget):
    """自定义消息项widget，网格布局，左侧占两行，右侧各占一行"""
    def __init__(self, contact_name, user_message="", bot_message=""):
        super().__init__()
        self.contact_name = contact_name
        
        self.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 6px;
            }
        """)
        
        layout = QGridLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(5)
        
        self.name_label = QLabel(contact_name)
        self.name_label.setWordWrap(True)
        self.name_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.name_label.setStyleSheet("font-weight: bold; font-size: 13px; color: #333; background: transparent;")
        self.name_label.setFixedWidth(80) 
        
        self.user_message_label = QLabel(f"收到的消息: {user_message}" if user_message else "收到的消息: ")
        self.user_message_label.setWordWrap(True)
        self.user_message_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.user_message_label.setStyleSheet("color: #4CAF50; background: transparent;") 
        
        self.bot_message_label = QLabel(f"发出的消息: {bot_message}" if bot_message else "发出的消息: ")
        self.bot_message_label.setWordWrap(True)
        self.bot_message_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.bot_message_label.setStyleSheet("color: #2196F3; background: transparent;") 
        
        layout.addWidget(self.name_label, 0, 0, 2, 1)          
        layout.addWidget(self.user_message_label, 0, 1, 1, 1)  
        layout.addWidget(self.bot_message_label, 1, 1, 1, 1)   
        
        self.setMinimumHeight(65)

    def update_message(self, message_type, content):
        if message_type == "user":
            self.user_message_label.setText(f"收到的消息: {content}")
        elif message_type == "bot":
            self.bot_message_label.setText(f"发出的消息: {content}")


class WeChatAssistantGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("微信自动化助手")
        self.setGeometry(100, 100, 800, 600)
        
        icon_path = "unnamed.jpg"
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(QPixmap(icon_path)))
        
        self.processed_messages = 0
        self.scanner = None
        self.database = WeChatDatabase()
        self.is_running = False
        self.scan_thread = None
        self.config_file = "config.json"
        
        self.signals = ScannerSignals()
        self.signals.message_received.connect(self.add_message)
        self.signals.log_added.connect(self.add_log)
        self.signals.count_updated.connect(self.update_message_count)
        self.signals.user_info_updated.connect(self.update_user_info)
        self.signals.warning_emitted.connect(lambda title, content: QMessageBox.warning(self, title, content))
        self.signals.stopped.connect(self.handle_scanner_stopped)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        self.create_top_bar(main_layout)
        self.create_tab_widget(main_layout)
        self.create_footer(main_layout)
        
        self.add_rule_button.clicked.connect(self.add_keyword_rule)
        self.edit_rule_button.clicked.connect(self.edit_keyword_rule)
        self.delete_rule_button.clicked.connect(self.delete_keyword_rule)
        self.export_rule_button.clicked.connect(self.export_keyword_rules)
        self.import_rule_button.clicked.connect(self.import_keyword_rules)
        
        self.load_config()
        self.load_user_info()
        self.load_latest_messages()
        
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_system_info)
        self.timer.start(1000)
    
    def create_top_bar(self, layout):
        top_bar = QWidget()
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(10, 10, 10, 10)
        
        left_container = QWidget()
        left_layout = QHBoxLayout(left_container)
        
        self.avatar_label = QLabel()
        self.avatar_label.setFixedSize(40, 40)
        self.avatar_label.setStyleSheet("border-radius: 20px; background-color: #E0E0E0;")
        left_layout.addWidget(self.avatar_label)
        
        self.nickname_label = QLabel("微信昵称")
        self.nickname_label.setFont(QFont("Microsoft YaHei", 12, QFont.Weight.Bold))
        left_layout.addWidget(self.nickname_label)
        
        status_container = QWidget()
        status_layout = QHBoxLayout(status_container)
        status_layout.setContentsMargins(10, 0, 0, 0)
        
        self.status_light = QLabel()
        self.status_light.setFixedSize(12, 12)
        self.status_light.setStyleSheet("border-radius: 6px; background-color: #4CAF50;") 
        status_layout.addWidget(self.status_light)
        
        self.status_label = QLabel("正常")
        self.status_label.setFont(QFont("Microsoft YaHei", 10))
        status_layout.addWidget(self.status_label)
        left_layout.addWidget(status_container)
        top_layout.addWidget(left_container)
        
        self.start_stop_button = QPushButton("开始")
        self.start_stop_button.setCheckable(True)
        self.start_stop_button.clicked.connect(self.toggle_start_stop)
        top_layout.addWidget(self.start_stop_button)
        
        layout.addWidget(top_bar)
    
    def create_tab_widget(self, layout):
        self.tab_widget = QTabWidget()
        
        # 选项卡 1：主界面
        main_tab = QWidget()
        main_layout = QHBoxLayout(main_tab)
        
        settings_widget = QWidget()
        settings_layout = QVBoxLayout(settings_widget)
        settings_layout.setContentsMargins(10, 10, 10, 10)
        
        interval_layout = QHBoxLayout()
        interval_label = QLabel("轮询间隔 (秒):")
        self.interval_spinbox = QSpinBox()
        self.interval_spinbox.setMinimum(1)
        self.interval_spinbox.setMaximum(60)
        self.interval_spinbox.setValue(2)
        interval_layout.addWidget(interval_label)
        interval_layout.addWidget(self.interval_spinbox)
        settings_layout.addLayout(interval_layout)
        
        slot_layout = QHBoxLayout()
        slot_label = QLabel("屏蔽前N个槽位:")
        self.slot_spinbox = QSpinBox()
        self.slot_spinbox.setMinimum(0)
        self.slot_spinbox.setMaximum(10)
        self.slot_spinbox.setValue(0)
        slot_layout.addWidget(slot_label)
        slot_layout.addWidget(self.slot_spinbox)
        settings_layout.addLayout(slot_layout)
        
        strategy_group = QButtonGroup()
        strategy_layout = QVBoxLayout()
        strategy_label = QLabel("回复策略:")
        strategy_layout.addWidget(strategy_label)
        
        self.strategy1 = QRadioButton("关键字回复")
        self.strategy2 = QRadioButton("OpenAI 回复")
        strategy_group.addButton(self.strategy1)
        strategy_group.addButton(self.strategy2)
        self.strategy1.setChecked(True)
        self.strategy2.toggled.connect(self.check_openai_config)
        
        strategy_layout.addWidget(self.strategy1)
        strategy_layout.addWidget(self.strategy2)
        
        self.history_rounds_layout = QHBoxLayout()
        self.history_rounds_label = QLabel("聊天记录轮数:")
        self.history_rounds_spinbox = QSpinBox()
        self.history_rounds_spinbox.setMinimum(1)
        self.history_rounds_spinbox.setMaximum(50)
        self.history_rounds_spinbox.setValue(10)
        self.history_rounds_layout.addWidget(self.history_rounds_label)
        self.history_rounds_layout.addWidget(self.history_rounds_spinbox)
        strategy_layout.addLayout(self.history_rounds_layout)
        settings_layout.addLayout(strategy_layout)
        
        non_text_layout = QHBoxLayout()
        non_text_label = QLabel("非文本消息处理:")
        self.non_text_combobox = QComboBox()
        self.non_text_combobox.addItems(["忽略，不处理", "回复提示信息"])
        self.non_text_combobox.setCurrentIndex(0) 
        non_text_layout.addWidget(non_text_label)
        non_text_layout.addWidget(self.non_text_combobox)
        settings_layout.addLayout(non_text_layout)
        
        settings_layout.addStretch(1)
        
        message_widget = QWidget()
        message_layout = QVBoxLayout(message_widget)
        message_layout.setContentsMargins(10, 10, 10, 10)
        
        message_header = QWidget()
        message_header_layout = QHBoxLayout(message_header)
        message_label = QLabel("实时消息流")
        message_header_layout.addWidget(message_label)
        
        self.clear_message_button = QPushButton("清空")
        self.clear_message_button.clicked.connect(self.clear_messages)
        message_header_layout.addWidget(self.clear_message_button)
        
        self.export_message_button = QPushButton("导出")
        self.export_message_button.clicked.connect(self.export_messages)
        message_header_layout.addWidget(self.export_message_button)
        message_layout.addWidget(message_header)
        
        self.message_list = QListWidget()
        self.message_list.setSelectionMode(QListWidget.SelectionMode.NoSelection)
        self.message_list.setSpacing(6) 
        self.message_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #E0E0E0;
                border-radius: 4px;
                background-color: #F5F5F5; 
                padding: 6px;
            }
            QListWidget::item { background-color: transparent; }
        """)
        self.message_list.itemDoubleClicked.connect(self.show_chat_history)
        message_layout.addWidget(self.message_list)
        
        main_layout.addWidget(settings_widget, 1)
        main_layout.addWidget(message_widget, 3)
        
        # 选项卡 2：关键字规则
        keyword_tab = QWidget()
        keyword_layout = QVBoxLayout(keyword_tab)
        keyword_layout.setContentsMargins(10, 10, 10, 10)
        
        keyword_label = QLabel("关键字规则")
        keyword_layout.addWidget(keyword_label)
        
        button_widget = QWidget()
        button_layout = QHBoxLayout(button_widget)
        button_layout.setContentsMargins(0, 0, 0, 10)
        
        self.add_rule_button = QPushButton("增加规则")
        self.edit_rule_button = QPushButton("修改规则")
        self.delete_rule_button = QPushButton("删除规则")
        self.export_rule_button = QPushButton("导出规则")
        self.import_rule_button = QPushButton("导入规则")
        
        button_layout.addWidget(self.add_rule_button)
        button_layout.addWidget(self.edit_rule_button)
        button_layout.addWidget(self.delete_rule_button)
        button_layout.addStretch()
        button_layout.addWidget(self.export_rule_button)
        button_layout.addWidget(self.import_rule_button)
        keyword_layout.addWidget(button_widget)
        
        self.keyword_table = QTableWidget()
        self.keyword_table.setColumnCount(2)
        self.keyword_table.setHorizontalHeaderLabels(["关键字", "回复内容"])
        self.keyword_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.keyword_table.customContextMenuRequested.connect(self.show_keyword_menu)
        self.keyword_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.keyword_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.keyword_table.resizeEvent = lambda event: self.adjust_column_widths()
        keyword_layout.addWidget(self.keyword_table)
        
        # 选项卡 3：OpenAI 配置
        openai_tab = QWidget()
        openai_layout = QVBoxLayout(openai_tab)
        openai_layout.setContentsMargins(10, 10, 10, 10)
        
        openai_label = QLabel("OpenAI 配置")
        openai_layout.addWidget(openai_label)
        
        openai_form = QFormLayout()
        self.openai_key = QLineEdit()
        self.openai_key.setEchoMode(QLineEdit.EchoMode.Password)
        openai_form.addRow("API Key:", self.openai_key)
        
        self.openai_url = QLineEdit()
        self.openai_url.setText("https://api.openai.com/v1/chat/completions")
        openai_form.addRow("API URL:", self.openai_url)
        
        self.openai_model = QLineEdit()
        self.openai_model.setText("gpt-3.5-turbo")
        openai_form.addRow("Model:", self.openai_model)
        
        self.openai_system = QTextEdit()
        self.openai_system.setPlainText("你是一个智能微信助手，帮助用户回复微信消息。请保持回复友好、自然，符合日常聊天的语气。\n\n规则：\n1. 回复要简洁明了，不要太长\n2. 保持口语化，避免过于正式的表达\n3. 针对用户的问题提供有用的信息\n4. 如果不清楚问题，可以礼貌地询问\n5. 保持专业和礼貌的态度")
        openai_form.addRow("System Prompt:", self.openai_system)
        
        openai_layout.addLayout(openai_form)
        
        self.save_openai_button = QPushButton("保存配置")
        self.save_openai_button.clicked.connect(self.save_config)
        openai_layout.addWidget(self.save_openai_button)
        
        self.test_openai_button = QPushButton("测试 OpenAI API")
        self.test_openai_button.clicked.connect(self.test_openai_api)
        openai_layout.addWidget(self.test_openai_button)
        
        # 选项卡 4：运行日志
        log_tab = QWidget()
        log_layout = QVBoxLayout(log_tab)
        log_layout.setContentsMargins(10, 10, 10, 10)
        
        log_label = QLabel("运行日志")
        log_layout.addWidget(log_label)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setStyleSheet("border: 1px solid #E0E0E0; border-radius: 4px;")
        log_layout.addWidget(self.log_text)
        
        self.tab_widget.addTab(main_tab, "主界面")
        self.tab_widget.addTab(keyword_tab, "关键字规则")
        self.tab_widget.addTab(openai_tab, "OpenAI 配置")
        self.tab_widget.addTab(log_tab, "运行日志")
        
        layout.addWidget(self.tab_widget)
    
    def create_footer(self, layout):
        footer = QWidget()
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(10, 10, 10, 10)
        
        self.message_count_label = QLabel(f"已处理消息数: {self.processed_messages}")
        footer_layout.addWidget(self.message_count_label)
        
        self.system_info_label = QLabel("CPU: 0% | 内存: 0%")
        footer_layout.addWidget(self.system_info_label, 1, Qt.AlignmentFlag.AlignRight)
        
        layout.addWidget(footer)
    
    def update_system_info(self):
        try:
            cpu_percent = psutil.cpu_percent()
            memory = psutil.virtual_memory()
            memory_percent = memory.percent
            self.system_info_label.setText(f"CPU: {cpu_percent:.1f}% | 内存: {memory_percent:.1f}%")
        except KeyboardInterrupt:
            pass
    
    def show_keyword_menu(self, position):
        menu = QMenu()
        add_action = menu.addAction("添加规则")
        edit_action = menu.addAction("编辑规则")
        delete_action = menu.addAction("删除规则")
        add_action.triggered.connect(self.add_keyword_rule)
        edit_action.triggered.connect(self.edit_keyword_rule)
        delete_action.triggered.connect(self.delete_keyword_rule)
        menu.exec(self.keyword_table.mapToGlobal(position))
    
    def add_keyword_rule(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("添加关键字规则")
        layout = QFormLayout(dialog)
        keyword_edit, reply_edit = QLineEdit(), QLineEdit()
        layout.addRow("关键字:", keyword_edit)
        layout.addRow("回复内容:", reply_edit)
        
        button_box = QHBoxLayout()
        ok_button, cancel_button = QPushButton("确定"), QPushButton("取消")
        ok_button.clicked.connect(lambda: self.save_keyword_rule(dialog, keyword_edit, reply_edit))
        cancel_button.clicked.connect(dialog.reject)
        button_box.addWidget(ok_button)
        button_box.addWidget(cancel_button)
        layout.addRow(button_box)
        dialog.exec()
    
    def edit_keyword_rule(self):
        selected_items = self.keyword_table.selectedItems()
        if not selected_items: return QMessageBox.warning(self, "警告", "请选择要编辑的规则")
        
        row = selected_items[0].row()
        keyword = self.keyword_table.item(row, 0).text()
        reply = self.keyword_table.item(row, 1).text()
        
        dialog = QDialog(self)
        dialog.setWindowTitle("编辑关键字规则")
        layout = QFormLayout(dialog)
        keyword_edit, reply_edit = QLineEdit(keyword), QLineEdit(reply)
        layout.addRow("关键字:", keyword_edit)
        layout.addRow("回复内容:", reply_edit)
        
        button_box = QHBoxLayout()
        ok_button, cancel_button = QPushButton("确定"), QPushButton("取消")
        ok_button.clicked.connect(lambda: self.update_keyword_rule(dialog, row, keyword_edit, reply_edit))
        cancel_button.clicked.connect(dialog.reject)
        button_box.addWidget(ok_button)
        button_box.addWidget(cancel_button)
        layout.addRow(button_box)
        dialog.exec()
    
    def delete_keyword_rule(self):
        selected_items = self.keyword_table.selectedItems()
        if not selected_items: return QMessageBox.warning(self, "警告", "请选择要删除的规则")
        self.keyword_table.removeRow(selected_items[0].row())
    
    def save_keyword_rule(self, dialog, keyword_edit, reply_edit):
        keyword, reply = keyword_edit.text().strip(), reply_edit.text().strip()
        if not keyword: return QMessageBox.warning(self, "警告", "关键字不能为空")
        row_position = self.keyword_table.rowCount()
        self.keyword_table.insertRow(row_position)
        self.keyword_table.setItem(row_position, 0, QTableWidgetItem(keyword))
        self.keyword_table.setItem(row_position, 1, QTableWidgetItem(reply))
        dialog.accept()
    
    def update_keyword_rule(self, dialog, row, keyword_edit, reply_edit):
        keyword, reply = keyword_edit.text().strip(), reply_edit.text().strip()
        if not keyword: return QMessageBox.warning(self, "警告", "关键字不能为空")
        self.keyword_table.setItem(row, 0, QTableWidgetItem(keyword))
        self.keyword_table.setItem(row, 1, QTableWidgetItem(reply))
        dialog.accept()
    
    def add_message(self, message):
            """添加消息到消息流（彻底修复 C++ 底层指针闪退问题）"""
            if "|" not in message: return
            parts = message.split("|", 2)
            if len(parts) != 3: return
            contact_name, msg_type, content = parts

            existing_item = None
            user_msg = ""
            bot_msg = ""
            
            # 1. 查找是否存在该联系人的卡片，如果存在，先提取它的历史文本
            for i in range(self.message_list.count()):
                item = self.message_list.item(i)
                widget = self.message_list.itemWidget(item)
                if widget and hasattr(widget, "contact_name") and widget.contact_name == contact_name:
                    existing_item = item
                    # 提取旧卡片上的文本记录
                    user_msg = widget.user_message_label.text().replace("收到的消息: ", "")
                    bot_msg = widget.bot_message_label.text().replace("发出的消息: ", "")
                    break

            # 2. 如果存在旧卡片，直接从列表中安全抹除它
            if existing_item:
                row = self.message_list.row(existing_item)
                self.message_list.takeItem(row)
                # 此时旧的 item 和 widget 已经被 PyQt 安全回收，不要再引用它们

            # 3. 更新最新的消息内容
            if msg_type == "user":
                user_msg = content
            elif msg_type == "bot":
                bot_msg = content

            # 4. 创建一个全新的干净的 Widget 和 Item 插到第 0 行
            new_widget = MessageItemWidget(contact_name, user_message=user_msg, bot_message=bot_msg)
            new_item = QListWidgetItem()
            new_item.setSizeHint(new_widget.sizeHint())
            
            self.message_list.insertItem(0, new_item)
            self.message_list.setItemWidget(new_item, new_widget)

            # 5. 防止内存溢出，最多保留 50 个联系人卡片
            if self.message_list.count() > 50:
                self.message_list.takeItem(self.message_list.count() - 1)
    
    def clear_messages(self):
            """清空消息流及数据库中的历史消息（带有防误触确认）"""
            # 1. 弹出确认对话框
            reply = QMessageBox.question(
                self, 
                '确认清空', 
                '确定要清空当前的实时消息流，并删除所有历史聊天记录吗？\n\n注：已识别的联系人特征（头像/昵称等指纹库）将被保留，以保证机器人后续的识别速度。',
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, 
                QMessageBox.StandardButton.No  # 默认高亮“否”，防误触键盘回车
            )
            
            # 2. 如果用户点击了“是”，才执行清空操作
            if reply == QMessageBox.StandardButton.Yes:
                # 清空 UI 列表
                self.message_list.clear()
                
                # 清空数据库里的 messages 表
                if self.database:
                    self.database.clear_all_messages()
                    self.add_log("✅ 消息流已清空，所有历史聊天记录已被清除")
                    
                    # 顺便重置一下左下角的计数器
                    self.processed_messages = 0
                    self.message_count_label.setText(f"已处理消息数: {self.processed_messages}")
    
    def export_messages(self):
        """【重构】封装数据库访问"""
        if pd is None: return QMessageBox.warning(self, "警告", "未安装 pandas 库，无法导出消息")
        file_path, _ = QFileDialog.getSaveFileName(self, "导出消息", "", "Excel 文件 (*.xlsx);;CSV 文件 (*.csv)")
        if not file_path: return
        
        messages = self.database.get_all_messages_for_export()
        df = pd.DataFrame(messages, columns=["联系人", "发送者", "内容", "时间戳"])
        df["时间戳"] = pd.to_datetime(df["时间戳"], unit='s')
        
        try:
            if file_path.endswith('.xlsx'): df.to_excel(file_path, index=False)
            else: df.to_csv(file_path, index=False, encoding='utf-8-sig')
            QMessageBox.information(self, "成功", f"消息已导出到 {file_path}")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"导出失败: {e}")
    
    def show_chat_history(self, item):
        """【重构】封装数据库访问"""
        widget = self.message_list.itemWidget(item)
        contact_name = None
        if widget and hasattr(widget, "contact_name"):
            contact_name = widget.contact_name
            if contact_name == "我": return
        else:
            message_text = item.text()
            if ": " in message_text:
                sender = message_text.split(": ")[0]
                contact_name = sender if sender != "我" else None
        
        if not contact_name: return
        
        messages = self.database.get_chat_history_by_name(contact_name)
        
        dialog = QDialog(self)
        dialog.setWindowTitle(f"{contact_name} 的聊天记录")
        dialog.setGeometry(200, 200, 600, 400)
        layout = QVBoxLayout(dialog)
        chat_text = QTextEdit()
        chat_text.setReadOnly(True)
        
        chat_content = ""
        for sender, content, timestamp in messages:
            time_str = datetime.datetime.fromtimestamp(timestamp).strftime("%Y-%m-%d %H:%M:%S")
            display_sender = contact_name if sender == "user" else "我"
            chat_content += f"[{time_str}] {display_sender}: {content}\n\n"
        
        chat_text.setPlainText(chat_content)
        layout.addWidget(chat_text)
        
        button_box = QHBoxLayout()
        close_button = QPushButton("关闭")
        close_button.clicked.connect(dialog.reject)
        button_box.addStretch()
        button_box.addWidget(close_button)
        layout.addLayout(button_box)
        
        dialog.exec()
    
    def add_log(self, log):
        try:
            self.log_text.append(log)
            self.log_text.verticalScrollBar().setValue(self.log_text.verticalScrollBar().maximum())
        except KeyboardInterrupt: pass
    
    def check_openai_config(self, checked):
        if checked and not self.openai_key.text().strip():
            reply = QMessageBox.question(
                self, "提示", "OpenAI API Key 未配置，是否前往配置页面？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.tab_widget.setCurrentIndex(2) 
    
    def test_openai_api(self):
        from openai import OpenAI
        api_key, api_url, model = self.openai_key.text().strip(), self.openai_url.text().strip(), self.openai_model.text().strip()
        
        if not api_key: return QMessageBox.warning(self, "警告", "请先配置 OpenAI API Key")
        
        progress = QProgressDialog("正在测试 OpenAI API...", None, 0, 1, self)
        progress.setWindowTitle("测试中")
        progress.setModal(True)
        progress.show()
        
        try:
            client = OpenAI(api_key=api_key, base_url=api_url)
            completion = client.chat.completions.create(
                model=model, messages=[{"role": "user", "content": "你好"}], temperature=0.7, max_tokens=50
            )
            if completion.choices and len(completion.choices) > 0:
                QMessageBox.information(self, "成功", "测试通过，OpenAI API 可正常使用")
                self.add_log("✅ OpenAI API 测试通过")
            else:
                QMessageBox.warning(self, "错误", "API 响应格式异常")
                self.add_log("❌ OpenAI API 响应格式异常")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"连接异常: {str(e)}")
            self.add_log(f"❌ OpenAI API 测试失败: {str(e)}")
        finally:
            progress.close()
    
    def toggle_start_stop(self, checked):
        if checked:
            if self.strategy1.isChecked() and self.keyword_table.rowCount() == 0:
                QMessageBox.warning(self, "警告", "请先添加关键字规则再开始")
                self.start_stop_button.setChecked(False)
                return
            elif self.strategy2.isChecked() and not self.openai_key.text().strip():
                QMessageBox.warning(self, "警告", "OpenAI API Key 未配置，请修改配置")
                self.start_stop_button.setChecked(False)
                self.tab_widget.setCurrentIndex(2) 
                return
            
            self.start_stop_button.setText("停止")
            self.status_light.setStyleSheet("border-radius: 6px; background-color: #4CAF50;") 
            self.status_label.setText("正常")
            self.add_log("系统已启动")
            self.start_scanner()
        else:
            self.start_stop_button.setText("开始")
            self.status_light.setStyleSheet("border-radius: 6px; background-color: #F44336;") 
            self.status_label.setText("停止")
            self.add_log("系统已停止")
            self.stop_scanner()
    
    def start_scanner(self):
        """【重构】纯同步线程逻辑，完全剥离 RPA 脏活"""
        if not self.scanner: self.scanner = WeChatScanner()
        else: self.scanner.running = True
        
        keyword_rules = []
        for row in range(self.keyword_table.rowCount()):
            k, r = self.keyword_table.item(row, 0), self.keyword_table.item(row, 1)
            if k and r: keyword_rules.append([k.text(), r.text()])
        
        self.scanner.set_keyword_rules(keyword_rules)
        self.scanner.set_reply_strategy("openai" if self.strategy2.isChecked() else "keyword")
        self.scanner.set_non_text_message_action("reply" if self.non_text_combobox.currentIndex() == 1 else "ignore")
        self.scanner.set_openai_config({
            "key": self.openai_key.text(), "url": self.openai_url.text(),
            "model": self.openai_model.text(), "system": self.openai_system.toPlainText()
        })
        self.scanner.set_history_rounds(self.history_rounds_spinbox.value())
        
        def callback(msg_type, content):
            if msg_type == 'message':
                self.signals.message_received.emit(content)
                self.signals.count_updated.emit()
            elif msg_type == 'log':
                self.signals.log_added.emit(content)
        
        def run_scanner_sync():
            try:
                callback('log', "🔄 开始初始化扫描器...")
                if not self.scanner.initialize():
                    self.signals.warning_emitted.emit("警告", "未找到微信窗口，请确认已打开微信")
                    self.signals.stopped.emit()
                    return
                
                callback('log', "✅ 扫描器初始化成功")
                # 完全委托给 Scanner 处理用户信息，代码极致干净
                user_info = self.scanner.check_and_update_user_info(callback)
                self.signals.user_info_updated.emit(user_info)
                
                callback('log', "🔄 启动消息监控循环...")
                self.scanner.run(self.database, callback)
            except Exception as e:
                self.signals.log_added.emit(f"❌ 扫描器错误: {e}\n{traceback.format_exc()}")
                self.signals.stopped.emit()
        
        self.scan_thread = threading.Thread(target=run_scanner_sync, daemon=True)
        self.scan_thread.start()
        self.is_running = True
    
    def stop_scanner(self):
        self.is_running = False
        if self.scanner: self.scanner.stop()
        self.add_log("扫描器已停止")
    
    def adjust_column_widths(self):
        total_width = self.keyword_table.width() - 20
        keyword_width = int(total_width * 1/3)
        self.keyword_table.setColumnWidth(0, keyword_width)
        self.keyword_table.setColumnWidth(1, total_width - keyword_width)
    
    def update_message_count(self):
        self.processed_messages += 1
        self.message_count_label.setText(f"已处理消息数: {self.processed_messages}")
    
    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                self.strategy2.toggled.disconnect(self.check_openai_config)
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                if 'interval' in config: self.interval_spinbox.setValue(config['interval'])
                if 'slot' in config: self.slot_spinbox.setValue(config['slot'])
                if 'history_rounds' in config: self.history_rounds_spinbox.setValue(config['history_rounds'])
                if 'non_text_action' in config: self.non_text_combobox.setCurrentIndex(config['non_text_action'])
                
                if 'openai' in config:
                    o = config['openai']
                    if 'key' in o: self.openai_key.setText(o['key'])
                    if 'url' in o: self.openai_url.setText(o['url'])
                    if 'model' in o: self.openai_model.setText(o['model'])
                    if 'system' in o: self.openai_system.setPlainText(o['system'])
                
                if 'strategy' in config:
                    if config['strategy'] == 'openai': self.strategy2.setChecked(True)
                    else: self.strategy1.setChecked(True)
                
                if 'keywords' in config:
                    for keyword, reply in config['keywords']:
                        row_position = self.keyword_table.rowCount()
                        self.keyword_table.insertRow(row_position)
                        self.keyword_table.setItem(row_position, 0, QTableWidgetItem(keyword))
                        self.keyword_table.setItem(row_position, 1, QTableWidgetItem(reply))
                
                self.strategy2.toggled.connect(self.check_openai_config)
                self.add_log("✅ 配置加载成功")
            except Exception as e:
                try: self.strategy2.toggled.connect(self.check_openai_config)
                except: pass
                self.add_log(f"❌ 配置加载失败: {e}")
    
    def save_config(self):
        existing_config = {}
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f: existing_config = json.load(f)
            except: pass
        
        config = {
            'interval': self.interval_spinbox.value(),
            'slot': self.slot_spinbox.value(),
            'strategy': 'openai' if self.strategy2.isChecked() else 'keyword',
            'history_rounds': self.history_rounds_spinbox.value(),
            'non_text_action': self.non_text_combobox.currentIndex(),
            'openai': {
                'key': self.openai_key.text(), 'url': self.openai_url.text(),
                'model': self.openai_model.text(), 'system': self.openai_system.toPlainText()
            },
            'keywords': []
        }
        
        if 'user_info' in existing_config: config['user_info'] = existing_config['user_info']
        
        for row in range(self.keyword_table.rowCount()):
            k, r = self.keyword_table.item(row, 0), self.keyword_table.item(row, 1)
            if k and r: config['keywords'].append([k.text(), r.text()])
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            self.add_log("✅ 配置保存成功")
        except Exception as e: self.add_log(f"❌ 配置保存失败: {e}")
    
    def export_keyword_rules(self):
        if pd is None: return QMessageBox.warning(self, "警告", "未安装 pandas 库，无法导出规则")
        file_path, _ = QFileDialog.getSaveFileName(self, "导出关键字规则", "", "Excel 文件 (*.xlsx);;CSV 文件 (*.csv)")
        if not file_path: return
        
        rules = []
        for row in range(self.keyword_table.rowCount()):
            k, r = self.keyword_table.item(row, 0), self.keyword_table.item(row, 1)
            if k and r: rules.append([k.text(), r.text()])
        
        df = pd.DataFrame(rules, columns=["关键字", "回复内容"])
        try:
            if file_path.endswith('.xlsx'): df.to_excel(file_path, index=False)
            else: df.to_csv(file_path, index=False, encoding='utf-8-sig')
            QMessageBox.information(self, "成功", f"关键字规则已导出到 {file_path}")
        except Exception as e: QMessageBox.warning(self, "错误", f"导出失败: {e}")
    
    def import_keyword_rules(self):
        if pd is None: return QMessageBox.warning(self, "警告", "未安装 pandas 库，无法导入规则")
        file_path, _ = QFileDialog.getOpenFileName(self, "导入关键字规则", "", "Excel 文件 (*.xlsx);;CSV 文件 (*.csv)")
        if not file_path: return
        
        try:
            if file_path.endswith('.xlsx'): df = pd.read_excel(file_path)
            else: df = pd.read_csv(file_path, encoding='utf-8-sig')
            self.keyword_table.setRowCount(0)
            for _, row in df.iterrows():
                if len(row) >= 2:
                    keyword = str(row[0]) if pd.notna(row[0]) else ""
                    reply = str(row[1]) if pd.notna(row[1]) else ""
                    if keyword:
                        rp = self.keyword_table.rowCount()
                        self.keyword_table.insertRow(rp)
                        self.keyword_table.setItem(rp, 0, QTableWidgetItem(keyword))
                        self.keyword_table.setItem(rp, 1, QTableWidgetItem(reply))
            QMessageBox.information(self, "成功", f"关键字规则已从 {file_path} 导入")
        except Exception as e: QMessageBox.warning(self, "错误", f"导入失败: {e}")
    
    def load_user_info(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                user_info = config.get('user_info', {"nickname": "等待连接...", "avatar_path": None})
                self.update_user_info(user_info)
                self.add_log("✅ 用户信息加载成功")
            except Exception as e: self.add_log(f"❌ 用户信息加载失败: {e}")
        else:
            self.update_user_info({"nickname": "等待连接...", "avatar_path": None})

    def load_latest_messages(self):
        """【重构】封装数据库访问"""
        try:
            latest_msgs = self.database.get_latest_messages_all_contacts()
            for msg in latest_msgs:
                widget = MessageItemWidget(msg['contact_name'], user_message=msg['user_message'], bot_message=msg['bot_message'])
                item = QListWidgetItem()
                item.setSizeHint(widget.sizeHint())
                self.message_list.addItem(item) 
                self.message_list.setItemWidget(item, widget)
            self.add_log(f"✅ 从数据库加载了 {len(latest_msgs)} 个联系人的最新消息")
        except Exception as e:
            self.add_log(f"❌ 加载消息失败: {e}")
    
    def update_user_info(self, user_info):
        self.nickname_label.setText(user_info.get("nickname", "等待连接..."))
        avatar_path = user_info.get("avatar_path")
        if avatar_path and os.path.exists(avatar_path):
            pixmap = QPixmap(avatar_path).scaled(40, 40)
            self.avatar_label.setPixmap(pixmap)
    
    def handle_scanner_stopped(self):
        self.is_running = False
        self.start_stop_button.setChecked(False)
        self.start_stop_button.setText("开始")
        self.status_light.setStyleSheet("border-radius: 6px; background-color: #F44336;") 
        self.status_label.setText("停止")

    def closeEvent(self, event):
        self.save_config()
        if self.is_running: self.stop_scanner()
        if self.database: self.database.close()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet("""
        QMainWindow { background-color: #F5F5F5; }
        QWidget { font-family: Microsoft YaHei; }
        QTabWidget::pane { border: 1px solid #E0E0E0; background-color: white; }
        QTabBar::tab { background-color: #F0F0F0; padding: 8px 16px; margin-right: 2px; border-radius: 4px 4px 0 0; }
        QTabBar::tab:selected { background-color: white; border-bottom: 2px solid #4CAF50; }
        QSpinBox, QLineEdit, QTextEdit { border: 1px solid #E0E0E0; border-radius: 4px; padding: 4px 8px; }
        QSpinBox:focus, QLineEdit:focus, QTextEdit:focus { border-color: #4CAF50; }
        QListWidget, QTableWidget { border: 1px solid #E0E0E0; border-radius: 4px; }
        QTableWidget::header { background-color: #F0F0F0; }
    """)
    window = WeChatAssistantGUI()
    window.show()
    sys.exit(app.exec())
