import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont
from PySide6.QtCore import QLocale
from gui import ExcelComparatorApp

def main():
    # تعيين الـ Locale الافتراضي صراحةً لتوحيد تنسيق الأرقام ولغة أزرار الحوار القياسية
    QLocale.setDefault(QLocale(QLocale.English, QLocale.UnitedStates))

    app = QApplication(sys.argv)
    
    # تعيين الخط الافتراضي مع الخطوط البديلة لدعم جميع أنظمة التشغيل
    font = QFont()
    font.setFamilies(["Segoe UI", "Tahoma", "Arial", "Geeza Pro", "DejaVu Sans", "sans-serif"])
    font.setPointSizeF(10)
    app.setFont(font)
    
    window = ExcelComparatorApp()
    window.showMaximized()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
