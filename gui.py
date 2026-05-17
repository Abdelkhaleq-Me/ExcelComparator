import sys
import os
import subprocess
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QGroupBox, QLabel, QLineEdit, QPushButton, QComboBox, QSpinBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QFileDialog,
    QProgressBar, QMessageBox, QAbstractItemView, QStackedWidget,
    QFrame, QSpacerItem, QSizePolicy, QScrollArea
)
from PySide6.QtCore import Qt, QThread, Signal, QSize, QStandardPaths
from PySide6.QtGui import QIcon, QFont, QColor, QPalette, QScreen
from core import get_excel_info, get_sheet_columns, run_comparison, auto_detect_header, get_file_info

def get_icon(name, color='white'):
    import qtawesome as qta
    return qta.icon(name, color=color)

# Worker thread to avoid freezing GUI
class ComparisonWorker(QThread):
    progress = Signal(int, str)
    finished = Signal(bool, str, dict)

    def __init__(self, params):
        super().__init__()
        self.params = params

    def run(self):
        success, msg, stats = run_comparison(
            self.params['file1'], self.params['sheet1'], self.params['header1'], self.params['key1'], self.params['name1'],
            self.params['file2'], self.params['sheet2'], self.params['header2'], self.params['key2'], self.params['name2'],
            self.params['mappings'], self.params['output'], self.progress.emit
        )
        self.finished.emit(success, msg, stats)

class ExcelComparatorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("المطابق الذكي لملفات Excel Pro")
        self.setLayoutDirection(Qt.RightToLeft)
        self.setWindowIcon(get_icon('fa5s.file-excel', color='#1A7A3C'))
        self.current_step = 0
        self.file1_headers = {}
        self.file2_headers = {}
        self.file1_path = ""
        self.file2_path = ""
        self.step_titles = {
            0: "الخطوة 1: إعداد الملف المرجعي",
            1: "الخطوة 2: إعداد الملف المقارن",
            2: "الخطوة 3: مطابقة الحقول",
            3: "الخطوة 4: البدء والمخرجات"
        }
        self.setup_ui()
        self.apply_styles()
        self.center_on_screen()

    def center_on_screen(self):
        screen = QApplication.primaryScreen()
        sg = screen.availableGeometry()
        width = int(sg.width() * 0.65)
        # ترك مساحة كافية لشريط التحكم في الأعلى والأسفل
        margin_top = 35
        margin_bottom = 5
        height = sg.height() - (margin_top + margin_bottom)
        x = sg.x() + (sg.width() - width) // 2
        y = sg.y() + margin_top
        self.setGeometry(x, y, width, height)

    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.main_layout = QVBoxLayout(central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # Top Banner / Header
        self.header_frame = QFrame()
        self.header_frame.setObjectName("headerFrame")
        self.header_frame.setFixedHeight(100)
        header_layout = QHBoxLayout(self.header_frame)
        
        icon_label = QLabel()
        icon_label.setPixmap(get_icon('fa5s.clipboard-check', color='#FFFFFF').pixmap(QSize(48, 48)))
        
        title_container = QVBoxLayout()
        self.title_label = QLabel("المطابق الذكي للبيانات")
        self.title_label.setObjectName("headerTitle")
        self.title_label.setFont(QFont("Segoe UI", 18, QFont.Bold))
        
        self.subtitle_label = QLabel("الخطوة 1: إعداد الملف المرجعي")
        self.subtitle_label.setObjectName("headerSubtitle")
        self.subtitle_label.setFont(QFont("Segoe UI", 10))
        
        title_container.addWidget(self.title_label)
        title_container.addWidget(self.subtitle_label)
        
        header_layout.addWidget(icon_label)
        header_layout.addLayout(title_container)
        header_layout.addStretch()
        
        self.main_layout.addWidget(self.header_frame)

        # Wizard Content (Scrollable Stacked Widget)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.scroll_area.setObjectName("contentScrollArea")
        
        self.stack = QStackedWidget()
        self.scroll_area.setWidget(self.stack)
        
        self.main_layout.addWidget(self.scroll_area, 1)

        # Step Widgets
        self.step1_widget = self.create_step1()
        self.step2_widget = self.create_step2()
        self.step3_widget = self.create_step3()
        self.step4_widget = self.create_step4()

        self.stack.addWidget(self.step1_widget)
        self.stack.addWidget(self.step2_widget)
        self.stack.addWidget(self.step3_widget)
        self.stack.addWidget(self.step4_widget)

        # Bottom Navigation
        nav_frame = QFrame()
        nav_frame.setObjectName("navFrame")
        nav_layout = QHBoxLayout(nav_frame)
        nav_layout.setContentsMargins(20, 15, 20, 15)
        
        self.btn_back = QPushButton(get_icon('fa5s.arrow-right', color='#1E3A5F'), " السابق")
        self.btn_back.clicked.connect(self.go_back)
        self.btn_back.setEnabled(False)
        
        self.btn_next = QPushButton(" التالي")
        self.btn_next.setIcon(get_icon('fa5s.arrow-left', color='white'))
        self.btn_next.setObjectName("nextBtn")
        self.btn_next.clicked.connect(self.go_next)
        
        self.btn_open_report = QPushButton(" فتح ملف التقرير")
        self.btn_open_report.setIcon(get_icon('fa5s.external-link-alt', color='white'))
        self.btn_open_report.setObjectName("runBtn")
        self.btn_open_report.setVisible(False)
        self.btn_open_report.clicked.connect(self.open_result_file)
        
        nav_layout.addWidget(self.btn_back)
        nav_layout.addStretch()
        nav_layout.addWidget(self.btn_open_report)
        nav_layout.addWidget(self.btn_next)
        
        self.main_layout.addWidget(nav_frame)

    def create_step_container(self, title):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(40, 30, 40, 30)
        layout.setSpacing(20)
        
        lbl = QLabel(title)
        lbl.setObjectName("stepTitle")
        lbl.setFont(QFont("Segoe UI", 14, QFont.Bold))
        layout.addWidget(lbl)
        
        return widget, layout

    def create_file_inputs(self, layout, idx):
        # Display Name
        name_layout = QVBoxLayout()
        name_layout.addWidget(QLabel(f"اسم الملف في التقرير (مثلاً: {'البرنامج' if idx==1 else 'اليدوي'}):"))
        le_name = QLineEdit("الملف الأول" if idx==1 else "الملف الثاني")
        name_layout.addWidget(le_name)
        layout.addLayout(name_layout)

        # File Selection
        file_layout = QVBoxLayout()
        file_layout.addWidget(QLabel("اختر ملف Excel:"))
        h_file = QHBoxLayout()
        le_file = QLineEdit()
        le_file.setPlaceholderText("لم يتم اختيار ملف...")
        le_file.setReadOnly(True)
        btn_file = QPushButton(get_icon('fa5s.file-import'), " استيراد")
        btn_file.clicked.connect(lambda: self.browse_file(idx))
        h_file.addWidget(le_file, 1)
        h_file.addWidget(btn_file)
        file_layout.addLayout(h_file)
        layout.addLayout(file_layout)

        # Options Box
        opt_group = QGroupBox("إعدادات قراءة البيانات")
        opt_layout = QVBoxLayout(opt_group)
        
        h_opt = QHBoxLayout()
        
        sheet_vbox = QVBoxLayout()
        sheet_vbox.addWidget(QLabel("ورقة العمل (Sheet):"))
        cb_sheet = QComboBox()
        cb_sheet.setLayoutDirection(Qt.RightToLeft)
        cb_sheet.setMinimumHeight(45)
        cb_sheet.currentIndexChanged.connect(lambda: self.on_sheet_changed(idx, auto_detect=True))
        sheet_vbox.addWidget(cb_sheet)
        
        header_vbox = QVBoxLayout()
        header_vbox.addWidget(QLabel("صف العناوين (1 = الأول):"))
        sb_header = QSpinBox()
        sb_header.setMinimumHeight(45)
        sb_header.setRange(1, 500)
        sb_header.setValue(1)
        sb_header.valueChanged.connect(lambda: self.on_sheet_changed(idx, auto_detect=False))
        header_vbox.addWidget(sb_header)
        
        h_opt.addLayout(sheet_vbox, 2)
        h_opt.addLayout(header_vbox, 1)
        opt_layout.addLayout(h_opt)
        
        layout.addWidget(opt_group)

        # Key Column
        key_layout = QVBoxLayout()
        key_layout.addWidget(QLabel("العمود المفتاح (ID / الحساب البريدي / إلخ):"))
        cb_key = QComboBox()
        cb_key.setLayoutDirection(Qt.RightToLeft)
        cb_key.setMinimumHeight(45)
        key_layout.addWidget(cb_key)
        layout.addLayout(key_layout)

        if idx == 1:
            self.le_name1, self.le_file1, self.cb_sheet1, self.sb_header1, self.cb_key1 = le_name, le_file, cb_sheet, sb_header, cb_key
        else:
            self.le_name2, self.le_file2, self.cb_sheet2, self.sb_header2, self.cb_key2 = le_name, le_file, cb_sheet, sb_header, cb_key

    def create_step1(self):
        widget, layout = self.create_step_container("إدخال معلومات الملف المرجعي (A)")
        self.create_file_inputs(layout, 1)
        layout.addStretch()
        return widget

    def create_step2(self):
        widget, layout = self.create_step_container("إدخال معلومات الملف المقارن (B)")
        self.create_file_inputs(layout, 2)
        layout.addStretch()
        return widget

    def create_step3(self):
        widget, layout = self.create_step_container("مطابقة أعمدة البيانات")
        
        info_lbl = QLabel("قم باختيار الأعمدة المتقابلة التي تريد مقارنة قيمها في التقرير.")
        info_lbl.setStyleSheet("color: #666;")
        layout.addWidget(info_lbl)

        # Auto Match Button
        self.btn_auto_match = QPushButton(get_icon('fa5s.magic', color='white'), " مطابقة تلقائية")
        self.btn_auto_match.setObjectName("autoMatchBtn")
        self.btn_auto_match.clicked.connect(self.auto_match_columns)
        layout.addWidget(self.btn_auto_match)

        map_box = QGroupBox("إضافة مطابقة يدوية")
        map_h = QHBoxLayout(map_box)
        
        v1 = QVBoxLayout()
        v1.addWidget(QLabel("عمود الملف 1:"))
        self.cb_map1 = QComboBox()
        self.cb_map1.setLayoutDirection(Qt.RightToLeft)
        self.cb_map1.setMinimumHeight(45)
        v1.addWidget(self.cb_map1)
        
        v2 = QVBoxLayout()
        v2.addWidget(QLabel("عمود الملف 2:"))
        self.cb_map2 = QComboBox()
        self.cb_map2.setLayoutDirection(Qt.RightToLeft)
        self.cb_map2.setMinimumHeight(45)
        v2.addWidget(self.cb_map2)
        
        btn_add = QPushButton(get_icon('fa5s.plus', color='white'), " إضافة")
        btn_add.setObjectName("addBtn")
        btn_add.setFixedSize(110, 45)
        btn_add.clicked.connect(self.add_mapping)
        
        map_h.addLayout(v1, 1)
        map_h.addLayout(v2, 1)
        map_h.addWidget(btn_add, 0, Qt.AlignBottom)
        
        layout.addWidget(map_box)

        # Table
        self.table_maps = QTableWidget(0, 3)
        self.table_maps.setHorizontalHeaderLabels(["عمود الملف 1", "عمود الملف 2", "إجراء"])
        self.table_maps.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table_maps.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table_maps.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.table_maps.verticalHeader().setDefaultSectionSize(45)
        self.table_maps.verticalHeader().setVisible(False)
        layout.addWidget(self.table_maps)

        return widget

    def create_step4(self):
        widget, layout = self.create_step_container("بدء المعالجة واستخراج النتائج")
        
        layout.addStretch()
        
        self.status_icon = QLabel()
        self.status_icon.setAlignment(Qt.AlignCenter)
        self.status_icon.setPixmap(get_icon('fa5s.hourglass-half', color='#2E5F8A').pixmap(64, 64))
        layout.addWidget(self.status_icon)
        
        self.lbl_status = QLabel("بانتظار بدء العملية...")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        self.lbl_status.setFont(QFont("Segoe UI", 12))
        layout.addWidget(self.lbl_status)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumHeight(25)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)
        
        # Statistics Table
        self.table_stats = QTableWidget(7, 3)
        self.table_stats.setHorizontalHeaderLabels(["البيان", "القيمة", "النسبة %"])
        self.table_stats.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_stats.verticalHeader().setVisible(False)
        self.table_stats.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_stats.setSelectionMode(QAbstractItemView.NoSelection)
        self.table_stats.setFixedHeight(280)
        self.table_stats.setVisible(False)
        
        # Style the stats table to match the requested image
        self.table_stats.setStyleSheet("""
            QTableWidget { border: 1px solid #DEE2E6; border-radius: 4px; }
            QHeaderView::section { background-color: #2E5F8A; color: white; font-weight: bold; }
        """)
        
        layout.addWidget(self.table_stats)
        
        layout.addStretch()
        return widget

    def apply_styles(self):
        # الحصول على المسار المطلق لمجلد الأيقونات
        base_path = os.path.dirname(os.path.abspath(__file__))
        assets_path = os.path.join(base_path, "assets").replace("\\", "/")
        
        arrow_down = f"{assets_path}/arrow_down.svg"
        arrow_up = f"{assets_path}/arrow_up.svg"

        style = f"""
        QMainWindow {{ background-color: #F8F9FA; font-family: "Segoe UI"; font-size: 14px; }}
        #contentScrollArea {{ background-color: #F8F9FA; border: none; }}
        QStackedWidget {{ background-color: #F8F9FA; }}
        #headerFrame {{ background-color: #1E3A5F; border-bottom: 3px solid #1A7A3C; }}
        #headerTitle {{ color: #FFFFFF; font-size: 22px; font-weight: bold; }}
        #headerSubtitle {{ color: #A0C4FF; font-size: 14px;}}
        
        #navFrame {{ background-color: #FFFFFF; border-top: 1px solid #DEE2E6; }}
        
        #stepTitle {{ color: #1E3A5F; margin-bottom: 10px; font-size: 18px; font-weight: bold; }}
        
        QLabel {{ color: #495057; font-size: 14px; font-weight: bold; }}
        
        QLineEdit {{
            padding: 10px;
            border: 1px solid #CED4DA;
            border-radius: 5px;
            background-color: #FFFFFF;
            color: #212529;
            font-size: 14px;
        }}
        QLineEdit:focus {{ border-color: #1E3A5F; }}

        QComboBox {{
            padding: 8px 10px 8px 35px;
            border: 1px solid #CED4DA;
            border-radius: 5px;
            background-color: #FFFFFF;
            color: #212529;
            font-size: 14px;
        }}
        QComboBox QAbstractItemView {{
            background-color: #FFFFFF;
            color: #212529;
            selection-background-color: #1E3A5F;
            selection-color: #FFFFFF;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 30px;
            border-right: 1px solid #CED4DA;
            border-left: none;
        }}
        QComboBox::down-arrow {{
            image: url("{arrow_down}");
            width: 14px;
            height: 14px;
        }}
        
        QSpinBox {{
            padding: 8px 25px 8px 12px;
            border: 1px solid #CED4DA;
            border-radius: 5px;
            background-color: #FFFFFF;
            color: #212529;
            font-size: 14px;
            qproperty-alignment: 'AlignRight | AlignVCenter';
        }}
        QSpinBox::up-button {{
            subcontrol-origin: border;
            subcontrol-position: top right;
            width: 25px;
            border-right: 1px solid #CED4DA;
            border-bottom: 1px solid #CED4DA;
            border-left: none;
        }}
        QSpinBox::down-button {{
            subcontrol-origin: border;
            subcontrol-position: bottom right;
            width: 25px;
            border-right: 1px solid #CED4DA;
            border-left: none;
        }}
        QSpinBox::up-arrow {{
            image: url("{arrow_up}");
            width: 12px;
            height: 12px;
        }}
        QSpinBox::down-arrow {{
            image: url("{arrow_down}");
            width: 12px;
            height: 12px;
        }}

        QPushButton {{
            padding: 10px 20px;
            border-radius: 5px;
            font-weight: bold;
            background-color: #E9ECEF;
            border: 1px solid #ADB5BD;
            color: #1E3A5F;
            font-size: 14px;
        }}
        QPushButton:hover {{ background-color: #DEE2E6; }}
        
        #nextBtn {{ background-color: #1E3A5F; color: white; border: none; }}
        #nextBtn:hover {{ background-color: #2E5F8A; }}
        
        #addBtn {{ background-color: #2E5F8A; color: white; border: none; }}
        #runBtn {{ background-color: #1A7A3C; color: white; border: none; font-size: 18px; }}
        #autoMatchBtn {{ background-color: #17A2B8; color: white; border: none; font-weight: bold; padding: 12px 20px; }}
        #autoMatchBtn:hover {{ background-color: #138496; }}
        
        QGroupBox {{
            font-weight: bold;
            border: 1px solid #DEE2E6;
            border-radius: 8px;
            margin-top: 20px;
            padding-top: 25px;
            background-color: #FFFFFF;
            font-size: 15px;
            color: #1E3A5F;
        }}
        QGroupBox::title {{ subcontrol-origin: margin; subcontrol-position: top center; padding: 0 10px; color: #1E3A5F; }}
        
        QTableWidget {{
            background-color: white;
            border: 1px solid #DEE2E6;
            border-radius: 5px;
            gridline-color: #F1F3F5;
            color: #212529;
            font-size: 14px;
        }}
        QTableWidget::item {{
            color: #212529;
        }}
        QHeaderView::section {{
            background-color: #F8F9FA;
            padding: 10px;
            border: none;
            border-bottom: 2px solid #DEE2E6;
            font-weight: bold;
            color: #495057;
            font-size: 14px;
        }}
        
        QProgressBar {{
            border: 1px solid #DEE2E6;
            border-radius: 10px;
            text-align: center;
            background-color: #E9ECEF;
            color: #212529;
            font-size: 14px;
            font-weight: bold;
        }}
        QProgressBar::chunk {{
            background-color: #1A7A3C;
            border-radius: 9px;
        }}
        
        QMessageBox {{
            background-color: #FFFFFF;
            border: 1px solid #DEE2E6;
        }}
        QMessageBox QLabel {{
            color: #212529;
            font-size: 14px;
            font-weight: normal;
        }}
        QMessageBox QPushButton {{
            background-color: #1E3A5F;
            color: white;
            border: 1px solid #1E3A5F;
            border-radius: 5px;
            padding: 8px 18px;
            font-weight: bold;
            min-width: 80px;
            font-size: 13px;
        }}
        QMessageBox QPushButton:hover {{
            background-color: #2E5F8A;
            border-color: #2E5F8A;
        }}
        """
        self.setStyleSheet(style)

    def browse_file(self, idx):
        default_dir = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        path, _ = QFileDialog.getOpenFileName(self, "اختر ملف Excel", default_dir, "Excel Files (*.xlsx *.xls)")
        if path:
            le = self.le_file1 if idx == 1 else self.le_file2
            filename = os.path.basename(path)
            le.setText(filename)
            le.setToolTip(path)
            
            if idx == 1: self.file1_path = path
            else: self.file2_path = path
            
            success, sheets, headers = get_file_info(path)
            if success:
                if idx == 1: self.file1_headers = headers
                else: self.file2_headers = headers
                
                cb = self.cb_sheet1 if idx == 1 else self.cb_sheet2
                cb.blockSignals(True)
                cb.clear()
                cb.addItems(sheets)
                cb.blockSignals(False)
                
                self.on_sheet_changed(idx, auto_detect=True)
            else:
                QMessageBox.critical(self, "خطأ", f"فشل قراءة الملف: {sheets}")

    def on_sheet_changed(self, idx, auto_detect=False):
        path = self.file1_path if idx == 1 else self.file2_path
        cb_sheet = self.cb_sheet1 if idx == 1 else self.cb_sheet2
        sb_header = self.sb_header1 if idx == 1 else self.sb_header2
        cb_key = self.cb_key1 if idx == 1 else self.cb_key2
        
        sheet = cb_sheet.currentText()
        if not path or not sheet: return

        # اكتشاف سطر العناوين تلقائياً من البيانات المخزنة مسبقاً
        if auto_detect:
            headers_dict = self.file1_headers if idx == 1 else self.file2_headers
            best_header = headers_dict.get(sheet, 0)
            sb_header.blockSignals(True)
            sb_header.setValue(best_header + 1)
            sb_header.blockSignals(False)

        success, cols = get_sheet_columns(path, sheet, sb_header.value() - 1)
        if success:
            # إذا كان هناك مطابقات مسبقة، نقوم بمسحها لتجنب عدم التطابق
            if self.table_maps.rowCount() > 0:
                self.table_maps.setRowCount(0)
                QMessageBox.information(self, "تنبيه", "تم مسح جدول مطابقات الأعمدة بسبب تغيير ورقة العمل أو صف العناوين، يرجى إعادة المطابقة مجدداً.")

            cb_key.clear()
            cb_key.addItems(cols)
            
            if idx == 1: self.cb_map1.clear(); self.cb_map1.addItems(cols)
            if idx == 2: self.cb_map2.clear(); self.cb_map2.addItems(cols)
        else:
            cb_key.clear()
            QMessageBox.warning(self, "تنبيه", f"تعذر جلب الأعمدة: {cols}")

    def add_mapping(self):
        c1 = self.cb_map1.currentText()
        c2 = self.cb_map2.currentText()
        if not c1 or not c2: return

        # Check dupes
        for r in range(self.table_maps.rowCount()):
            if self.table_maps.item(r,0).text() == c1 and self.table_maps.item(r,1).text() == c2:
                return

        row = self.table_maps.rowCount()
        self.table_maps.insertRow(row)
        self.table_maps.setItem(row, 0, QTableWidgetItem(c1))
        self.table_maps.setItem(row, 1, QTableWidgetItem(c2))
        
        btn_del = QPushButton(get_icon('fa5s.trash-alt', color='#D32F2F'), "")
        btn_del.setToolTip("حذف")
        btn_del.setFixedSize(30, 25)
        btn_del.clicked.connect(self.remove_map_row)
        
        container = QWidget()
        lay = QHBoxLayout(container)
        lay.setContentsMargins(0,0,0,0)
        lay.addWidget(btn_del)
        self.table_maps.setCellWidget(row, 2, container)

    def remove_map_row(self, checked=False):
        button = self.sender()
        if not button: return
        
        # الحصول على الحاوية (container) التي يوجد بها الزر
        container = button.parentWidget()
        for r in range(self.table_maps.rowCount()):
            if self.table_maps.cellWidget(r, 2) == container:
                self.table_maps.removeRow(r)
                break

    def auto_match_columns(self):
        cols1 = [self.cb_map1.itemText(i) for i in range(self.cb_map1.count())]
        cols2 = [self.cb_map2.itemText(i) for i in range(self.cb_map2.count())]
        
        matched_count = 0
        unmatched = []
        
        for c1 in cols1:
            found = False
            clean1 = c1.strip().lower()
            for c2 in cols2:
                clean2 = c2.strip().lower()
                if clean1 == clean2:
                    # Add to table
                    self.cb_map1.setCurrentText(c1)
                    self.cb_map2.setCurrentText(c2)
                    self.add_mapping()
                    matched_count += 1
                    found = True
                    break
            if not found:
                unmatched.append(c1)
        
        if matched_count > 0:
            msg = f"تمت مطابقة {matched_count} عمود تلقائياً بنجاح."
            if unmatched:
                msg += "\n\n⚠️ الأعمدة التي لم يتم إيجاد مثيل لها:\n- " + "\n- ".join(unmatched[:10])
                if len(unmatched) > 10:
                    msg += f"\n... (+{len(unmatched)-10} أعمدة أخرى)"
            QMessageBox.information(self, "نتائج المطابقة", msg)
        else:
            QMessageBox.warning(self, "تنبيه", "لم يتم العثور على أعمدة متطابقة بالأسماء.")

    def go_next(self):
        if self.current_step == 0:
            if not self.file1_path:
                QMessageBox.warning(self, "خطأ", "يرجى اختيار الملف المرجعي أولاً.")
                return
        elif self.current_step == 1:
            if not self.file2_path:
                QMessageBox.warning(self, "خطأ", "يرجى اختيار الملف المقارن أولاً.")
                return
            # تحديث قوائم الخطوة 3 مباشرة من البيانات المحملة مسبقاً لتجنب إعادة القراءة ومسح الجدول
            self.cb_map1.clear()
            self.cb_map1.addItems([self.cb_key1.itemText(i) for i in range(self.cb_key1.count())])
            self.cb_map2.clear()
            self.cb_map2.addItems([self.cb_key2.itemText(i) for i in range(self.cb_key2.count())])
        elif self.current_step == 2:
            if self.table_maps.rowCount() == 0:
                QMessageBox.warning(self, "خطأ", "يرجى إضافة عمود واحد على الأقل للمقارنة.")
                return
            
            # الانتقال لصفحة النتائج أولاً ثم البدء
            self.current_step += 1
            self.stack.setCurrentIndex(self.current_step)
            self.update_navigation_ui()
            self.run_process()
            return

        self.current_step += 1
        self.stack.setCurrentIndex(self.current_step)
        self.update_navigation_ui()

    def go_back(self):
        self.current_step -= 1
        self.stack.setCurrentIndex(self.current_step)
        self.update_navigation_ui()

    def update_navigation_ui(self):
        # تحديث العنوان والتحكم في أزرار التنقل
        self.subtitle_label.setText(self.step_titles.get(self.current_step, ""))
        self.btn_back.setEnabled(self.current_step > 0)
        
        # في الصفحة الأخيرة، نريد زر "السابق" مفعلاً
        if self.current_step == 3:
            # إذا كانت العملية قيد التنفيذ، نعطل الزر
            if hasattr(self, 'worker') and self.worker.isRunning():
                self.btn_back.setEnabled(False)
            else:
                self.btn_back.setEnabled(True)

        # تحديث نص وأيقونة الزر في خطوة المطابقة
        if self.current_step == 2:
            self.btn_next.setText(" إجراء المقارنة")
            self.btn_next.setIcon(get_icon('fa5s.play', color='white'))
        else:
            self.btn_next.setText(" التالي")
            self.btn_next.setIcon(get_icon('fa5s.arrow-left', color='white'))

        # إخفاء زر "التالي" في الخطوة الأخيرة (خطوة المعالجة)
        is_last = (self.current_step == len(self.step_titles) - 1)
        self.btn_next.setVisible(not is_last)
        
        # إظهار زر "فتح التقرير" فقط في الخطوة الأخيرة إذا انتهت العملية
        self.btn_open_report.setVisible(is_last and hasattr(self, 'last_success') and self.last_success)

    def run_process(self):
        # منع تشغيل عمليتين في نفس الوقت
        if hasattr(self, 'worker') and self.worker.isRunning():
            QMessageBox.warning(self, "تنبيه", "هناك عملية مقارنة قيد التنفيذ حالياً. يرجى الانتظار.")
            return

        # التحقق من اختيار أعمدة المفاتيح
        if not self.cb_key1.currentText() or not self.cb_key2.currentText():
            QMessageBox.warning(self, "خطأ", "يرجى اختيار أعمدة المفتاح (Key Column) لكل من الملفين لتتم عملية المطابقة.")
            return

        default_dir = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        default_file = os.path.join(default_dir, 'تقرير_المقارنة.xlsx')
        output_path, _ = QFileDialog.getSaveFileName(self, "حفظ التقرير", default_file, "Excel Files (*.xlsx)")
        if not output_path: return

        self.output_file = output_path
        
        maps = []
        for r in range(self.table_maps.rowCount()):
            maps.append({
                'col1': self.table_maps.item(r, 0).text(),
                'col2': self.table_maps.item(r, 1).text()
            })

        params = {
            'file1': self.file1_path,
            'sheet1': self.cb_sheet1.currentText(),
            'header1': self.sb_header1.value() - 1,
            'key1': self.cb_key1.currentText(),
            'name1': self.le_name1.text(),
            
            'file2': self.file2_path,
            'sheet2': self.cb_sheet2.currentText(),
            'header2': self.sb_header2.value() - 1,
            'key2': self.cb_key2.currentText(),
            'name2': self.le_name2.text(),
            
            'mappings': maps,
            'output': output_path
        }

        self.btn_back.setEnabled(False)
        self.update_navigation_ui()
        
        self.worker = ComparisonWorker(params)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_finished)
        self.worker.start()

    def update_progress(self, val, msg):
        self.progress_bar.setValue(val)
        self.lbl_status.setText(msg)

    def on_finished(self, success, msg, stats):
        self.btn_back.setEnabled(True)
        self.update_navigation_ui()
        self.last_success = success
        
        if success:
            self.status_icon.setPixmap(get_icon('fa5s.check-circle', color='#1A7A3C').pixmap(64, 64))
            self.lbl_status.setText("تمت العملية بنجاح!")
            
            # ملء جدول الإحصائيات
            rows = [
                (f"إجمالي أسطر {self.le_name1.text()}", str(stats.get('common', 0) + stats.get('only1', 0)), "—"),
                (f"إجمالي أسطر {self.le_name2.text()}", str(stats.get('common', 0) + stats.get('only2', 0)), "—"),
                ("الأسطر المشتركة", str(stats.get('common', 0)), "—"),
                (f"فقط في {self.le_name1.text()}", str(stats.get('only1', 0)), f"{(stats.get('only1', 0)/max(stats.get('common', 1)+stats.get('only1', 0), 1))*100:.1f}%"),
                (f"فقط في {self.le_name2.text()}", str(stats.get('only2', 0)), f"{(stats.get('only2', 0)/max(stats.get('common', 1)+stats.get('only2', 0), 1))*100:.1f}%"),
                ("متطابقون تماماً", str(stats.get('match', 0)), f"{(stats.get('match', 0)/max(stats.get('common', 1), 1))*100:.1f}%"),
                ("يحتوون اختلافات", str(stats.get('diff', 0)), f"{(stats.get('diff', 0)/max(stats.get('common', 1), 1))*100:.1f}%"),
            ]
            
            for i, (label, val, pct) in enumerate(rows):
                self.table_stats.setItem(i, 0, QTableWidgetItem(label))
                self.table_stats.setItem(i, 1, QTableWidgetItem(val))
                self.table_stats.setItem(i, 2, QTableWidgetItem(pct))
                
                # تلوين الخلفية بناءً على الحالة
                if i == 5: # متطابقون
                    color = QColor("#D6F5D6")
                    for c in range(3): self.table_stats.item(i, c).setBackground(color)
                elif i == 6: # اختلافات
                    color = QColor("#FFD6D6")
                    for c in range(3): self.table_stats.item(i, c).setBackground(color)
                
                for c in range(3): self.table_stats.item(i, c).setTextAlignment(Qt.AlignCenter)

            self.table_stats.setVisible(True)
            self.btn_open_report.setVisible(True)
            QMessageBox.information(self, "نجاح", msg)
        else:
            self.status_icon.setPixmap(get_icon('fa5s.exclamation-triangle', color='#D32F2F').pixmap(64, 64))
            self.lbl_status.setText("حدث خطأ أثناء المعالجة")
            QMessageBox.critical(self, "خطأ", msg)

    def open_result_file(self):
        try:
            path = self.output_file
            if sys.platform == 'win32':
                os.startfile(path)
            elif sys.platform == 'darwin':
                subprocess.run(['open', path])
            else:
                subprocess.run(['xdg-open', path])
        except Exception as e:
            QMessageBox.warning(self, "تنبيه", f"تعذر فتح الملف تلقائياً: {str(e)}")

