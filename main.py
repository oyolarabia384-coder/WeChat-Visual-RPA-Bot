import sys
from PyQt6.QtWidgets import QApplication
from gui import WeChatAssistantGUI

def main():
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

if __name__ == "__main__":
    main()
