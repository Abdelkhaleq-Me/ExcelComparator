import sys
import os
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont
from PySide6.QtCore import QLocale, Qt

# ── ضمان ترميز UTF-8 لمسارات الملفات العربية على أي نظام ──
os.environ.setdefault('PYTHONIOENCODING', 'utf-8')

def main():
    # تعيين الـ Locale الافتراضي صراحةً لتوحيد تنسيق الأرقام ولغة أزرار الحوار القياسية
    QLocale.setDefault(QLocale(QLocale.English, QLocale.UnitedStates))

    # ── دعم شاشات High-DPI على ويندوز 10/11 ──
    # PySide6 >= 6.5 يتعامل مع هذا تلقائياً في الغالب،
    # لكن نعيّنه صراحةً لتغطية الإصدارات الأقدم والحالات الخاصة
    if hasattr(Qt, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)

    # سياسة تقريب DPI — تمنع تشويش الواجهة عند نسب تكبير كسرية (125%, 150%)
    if hasattr(Qt, 'HighDpiScaleFactorRoundingPolicy'):
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )

    # تعيين الخط الافتراضي مع الخطوط البديلة لدعم جميع أنظمة التشغيل
    font = QFont()
    font.setFamilies(["Segoe UI", "Tahoma", "Arial", "Geeza Pro", "DejaVu Sans", "sans-serif"])
    font.setPointSizeF(10)
    app.setFont(font)

    from gui import ExcelComparatorApp
    window = ExcelComparatorApp()
    window.showMaximized()

    sys.exit(app.exec())

if __name__ == "__main__":
    main()

