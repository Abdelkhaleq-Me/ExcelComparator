import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont
from gui import ExcelComparatorApp

def main():
    app = QApplication(sys.argv)
    
    # تعيين الخط الافتراضي لدعم اللغة العربية بشكل جيد
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    window = ExcelComparatorApp()
    window.showMaximized()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
