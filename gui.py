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
from PySide6.QtCore import Qt, QThread, Signal, QSize, QStandardPaths, QTimer
from PySide6.QtGui import QIcon, QFont, QColor, QPalette, QScreen
from core import get_excel_info, get_sheet_columns, run_comparison, get_file_info

def create_font(size, bold=False):
    f = QFont()
    f.setFamilies(["Segoe UI", "Tahoma", "Arial", "Geeza Pro", "DejaVu Sans", "sans-serif"])
    f.setPointSizeF(size)
    if bold:
        f.setBold(True)
    return f

def get_icon(name, color='white'):
    import qtawesome as qta
    return qta.icon(name, color=color)


# ──────────────────────────────────────────────────────────────
# Worker threads
# ──────────────────────────────────────────────────────────────

class ComparisonWorker(QThread):
    progress = Signal(int, str)
    finished = Signal(bool, str, dict)

    def __init__(self, params):
        super().__init__()
        self.params = params

    def run(self):
        success, msg, stats = run_comparison(
            self.params['file1'], self.params['sheet1'], self.params['header1'],
            self.params['key1'],  self.params['name1'],
            self.params['file2'], self.params['sheet2'], self.params['header2'],
            self.params['key2'],  self.params['name2'],
            self.params['mappings'], self.params['output'], self.progress.emit
        )
        self.finished.emit(success, msg, stats)


class DeepComparisonWorker(QThread):
    progress = Signal(int, str)
    finished = Signal(bool, str, dict)

    def __init__(self, params):
        super().__init__()
        self.params = params

    def run(self):
        from core import run_deep_comparison
        p = self.params
        success, msg, stats = run_deep_comparison(
            p['file1'], p['sheet1'], p['header1'],
            p['match_col1'], p['anchor_col1'], p['name1'],
            p['file2'], p['sheet2'], p['header2'],
            p['match_col2'], p['anchor_col2'], p['name2'],
            p['mappings'], p['output'],
            p.get('min_similarity', 75),
            self.progress.emit
        )
        self.finished.emit(success, msg, stats)


# ──────────────────────────────────────────────────────────────
# النافذة الرئيسية
# ──────────────────────────────────────────────────────────────

class ExcelComparatorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("المطابق الذكي لملفات Excel Pro")
        self.setLayoutDirection(Qt.RightToLeft)

        # الحالة
        self.mode         = None       # 'fast' | 'deep'
        self.current_step = -1         # -1=الصفحة الرئيسية، 0-3=الخطوات
        self.file1_path   = ""
        self.file2_path   = ""
        self.file1_headers = {}
        self.file2_headers = {}
        self.cols1         = []
        self.cols2         = []

        self.step_titles = {
            -1: "اختر وضع المقارنة المناسب لبياناتك",
            0:  "الخطوة 1: إعداد الملف المرجعي",
            1:  "الخطوة 2: إعداد الملف المقارن",
            2:  "الخطوة 3: مطابقة الأعمدة",
            3:  "الخطوة 4: البدء والمخرجات",
        }

        self.setup_ui()
        self.apply_styles()
        self.setMinimumSize(800, 600)

    def showEvent(self, event):
        super().showEvent(event)
        if not hasattr(self, '_icons_loaded') or not self._icons_loaded:
            self._icons_loaded = True
            # تأجيل تحميل الأيقونات حتى بعد اكتمال دورة الرسم الأولى
            # — يمنع أخطاء QPainter::begin عند التحميل المباشر
            QTimer.singleShot(0, self._load_icons)

    def _load_icons(self):
        """تحميل أيقونات qtawesome — يُستدعى بعد عرض النافذة بإطار واحد."""
        self.setWindowIcon(get_icon('fa5s.file-excel', color='#1A7A3C'))
        self.icon_label.setPixmap(
            get_icon('fa5s.clipboard-check', color='#FFFFFF').pixmap(QSize(48, 48))
        )
        self.fast_card_icon.setPixmap(
            get_icon('fa5s.bolt',   color='#E67E00').pixmap(QSize(52, 52))
        )
        self.deep_card_icon.setPixmap(
            get_icon('fa5s.search', color='#7B1FA2').pixmap(QSize(52, 52))
        )
        self.btn_back.setIcon(get_icon('fa5s.arrow-right', color='#1E3A5F'))
        self.btn_next.setIcon(get_icon('fa5s.arrow-left',  color='white'))
        self.btn_open_report.setIcon(get_icon('fa5s.external-link-alt', color='white'))
        if hasattr(self, 'btn_file1'):
            self.btn_file1.setIcon(get_icon('fa5s.file-import'))
        if hasattr(self, 'btn_file2'):
            self.btn_file2.setIcon(get_icon('fa5s.file-import'))
        self.btn_auto_match.setIcon(get_icon('fa5s.magic',          color='white'))
        self.btn_add.setIcon(get_icon('fa5s.plus',                   color='white'))
        self.status_icon.setPixmap(
            get_icon('fa5s.hourglass-half', color='#2E5F8A').pixmap(64, 64)
        )

    # ──────────────────────────────────────────────────────────
    # بناء الواجهة
    # ──────────────────────────────────────────────────────────

    def setup_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        self.main_layout = QHBoxLayout(central) # استخدام تخطيط أفقي
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # ── الشريط الجانبي (Sidebar) ──────────────────────────
        self.sidebar_frame = QFrame()
        self.sidebar_frame.setObjectName("sidebarFrame")
        self.sidebar_frame.setFixedWidth(280)
        sidebar_layout = QVBoxLayout(self.sidebar_frame)
        sidebar_layout.setContentsMargins(20, 30, 20, 30)
        sidebar_layout.setSpacing(20)

        # ترويسة الشريط الجانبي
        header_container = QHBoxLayout()
        header_container.setSpacing(10)

        self.icon_label = QLabel()
        self.icon_label.setFixedSize(48, 48)
        self.icon_label.setAlignment(Qt.AlignCenter)
        header_container.addWidget(self.icon_label)

        title_container = QVBoxLayout()
        title_container.setSpacing(2)

        self.title_label = QLabel("المطابق الذكي")
        self.title_label.setObjectName("sidebarTitle")
        self.title_label.setFont(create_font(16, True))

        self.sidebar_subtitle = QLabel("لملفات Excel Pro")
        self.sidebar_subtitle.setObjectName("sidebarSubtitle")
        self.sidebar_subtitle.setFont(create_font(10))

        title_container.addWidget(self.title_label)
        title_container.addWidget(self.sidebar_subtitle)

        header_container.addLayout(title_container)
        sidebar_layout.addLayout(header_container)

        # خط فاصل
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setObjectName("sidebarLine")
        sidebar_layout.addWidget(line)

        # ويدجيت الخطوات
        self.steps_container = QFrame()
        self.steps_container.setObjectName("stepsContainer")
        steps_layout = QVBoxLayout(self.steps_container)
        steps_layout.setContentsMargins(0, 10, 0, 10)
        steps_layout.setSpacing(12)

        self.step_titles_sidebar = [
            ("الرئيسية", "fa5s.home"),
            ("الملف المرجعي", "fa5s.file-excel"),
            ("الملف المقارن", "fa5s.file-excel"),
            ("مطابقة الأعمدة", "fa5s.columns"),
            ("المعالجة والنتائج", "fa5s.chart-pie")
        ]

        self.step_widgets = []
        for idx, (title, icon_name) in enumerate(self.step_titles_sidebar):
            step_frame = QFrame()
            step_frame.setObjectName(f"stepFrame_{idx}")
            step_frame.setProperty("active", "false")
            step_frame.setCursor(Qt.PointingHandCursor)
            step_frame.mousePressEvent = lambda event, s_idx=idx-1: self.on_sidebar_clicked(s_idx)
            step_lay = QHBoxLayout(step_frame)
            step_lay.setContentsMargins(10, 8, 10, 8)
            step_lay.setSpacing(12)

            num_lbl = QLabel()
            num_lbl.setObjectName(f"stepNum_{idx}")
            num_lbl.setFixedSize(24, 24)
            num_lbl.setAlignment(Qt.AlignCenter)
            num_lbl.setFont(create_font(10, True))
            num_lbl.setStyleSheet("border-radius: 12px; background-color: #2E5F8A; color: white;")
            num_lbl.setText(str(idx + 1) if idx > 0 else "🏠")

            txt_lbl = QLabel(title)
            txt_lbl.setObjectName(f"stepTxt_{idx}")
            txt_lbl.setFont(create_font(10))

            step_lay.addWidget(num_lbl)
            step_lay.addWidget(txt_lbl, 1)

            steps_layout.addWidget(step_frame)
            self.step_widgets.append({
                'frame': step_frame,
                'num_lbl': num_lbl,
                'txt_lbl': txt_lbl,
                'icon_name': icon_name,
                'default_text': str(idx + 1) if idx > 0 else "🏠"
            })

        sidebar_layout.addWidget(self.steps_container)
        sidebar_layout.addStretch()

        self.main_layout.addWidget(self.sidebar_frame)

        # ── منطقة المحتوى الرئيسية ───────────────────────────
        self.content_widget = QWidget()
        self.content_widget.setObjectName("contentWidget")
        content_vlayout = QVBoxLayout(self.content_widget)
        content_vlayout.setContentsMargins(0, 0, 0, 0)
        content_vlayout.setSpacing(0)

        # شريط العنوان العلوي المبسط
        self.top_bar = QFrame()
        self.top_bar.setObjectName("topBar")
        self.top_bar.setFixedHeight(56)
        top_bar_layout = QHBoxLayout(self.top_bar)
        top_bar_layout.setContentsMargins(25, 0, 25, 0)

        self.subtitle_label = QLabel(self.step_titles[-1])
        self.subtitle_label.setObjectName("topBarSubtitle")
        self.subtitle_label.setFont(create_font(12, True))
        top_bar_layout.addWidget(self.subtitle_label)
        top_bar_layout.addStretch()
        content_vlayout.addWidget(self.top_bar)

        # الـ Stack الخاص بالصفحات
        self.stack = QStackedWidget()
        content_vlayout.addWidget(self.stack, 1)

        self.landing_widget = self.create_landing_page()
        self.step1_widget   = self.create_step1()
        self.step2_widget   = self.create_step2()
        self.step3_widget   = self.create_step3()
        self.step4_widget   = self.create_step4()

        self.stack.addWidget(self.landing_widget)                        # 0
        self.stack.addWidget(self.wrap_in_scroll_area(self.step1_widget))  # 1
        self.stack.addWidget(self.wrap_in_scroll_area(self.step2_widget))  # 2
        self.stack.addWidget(self.wrap_in_scroll_area(self.step3_widget))  # 3
        self.stack.addWidget(self.wrap_in_scroll_area(self.step4_widget))  # 4

        self.stack.setCurrentIndex(0)   # ابدأ بالصفحة الرئيسية

        # ── شريط التنقل السفلي ────────────────────────────────
        nav_frame = QFrame()
        nav_frame.setObjectName("navFrame")
        nav_layout = QHBoxLayout(nav_frame)
        nav_layout.setContentsMargins(20, 8, 20, 8)

        self.btn_back = QPushButton(" السابق")
        self.btn_back.clicked.connect(self.go_back)

        self.btn_next = QPushButton(" التالي")
        self.btn_next.setObjectName("nextBtn")
        self.btn_next.clicked.connect(self.go_next)

        self.btn_open_report = QPushButton(" فتح ملف التقرير")
        self.btn_open_report.setObjectName("runBtn")
        self.btn_open_report.setVisible(False)
        self.btn_open_report.clicked.connect(self.open_result_file)

        nav_layout.addWidget(self.btn_back)
        nav_layout.addStretch()
        nav_layout.addWidget(self.btn_open_report)
        nav_layout.addWidget(self.btn_next)

        content_vlayout.addWidget(nav_frame)
        self.main_layout.addWidget(self.content_widget, 1)
        self.update_navigation_ui()

    # ──────────────────────────────────────────────────────────
    # الصفحة الرئيسية: اختيار وضع المقارنة
    # ──────────────────────────────────────────────────────────

    def create_landing_page(self):
        page = QWidget()
        page.setObjectName("landingPage")
        outer = QVBoxLayout(page)
        outer.setContentsMargins(50, 35, 50, 30)
        outer.setSpacing(22)

        # ── العنوان ───────────────────────────────────────────
        title = QLabel("كيف تبدو بياناتك؟ اختر الوضع المناسب")
        title.setObjectName("landingMainTitle")
        title.setAlignment(Qt.AlignCenter)
        outer.addWidget(title)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setObjectName("landingSep")
        outer.addWidget(sep)

        # ── البطاقتان ─────────────────────────────────────────
        cards_row = QHBoxLayout()
        cards_row.setSpacing(28)

        fast_card = self._make_mode_card(
            title        = "⚡  المقارنة السريعة",
            tagline      = "يوجد رقم أو كود مشترك بين الجدولين",
            description  = (
                "مثل: رقم الموظف، رقم الحساب، رقم الهوية.\n"
                "النظام يطابق كل سجل برقمه مباشرةً — نتائج فورية ودقة 100%."
            ),
            bullets      = [
                "سريعة جداً حتى مع مئات الآلاف من السجلات",
                "تُظهر الاختلافات في البيانات بالتفصيل",
                "تُحدد السجلات الموجودة في جدول وغائبة عن الآخر",
            ],
            btn_text     = "ابدأ المقارنة السريعة",
            border_color = "#1E3A5F",
            btn_color    = "#1E3A5F",
            mode         = 'fast',
            is_fast      = True,
        )

        deep_card = self._make_mode_card(
            title        = "🔍  المقارنة العميقة",
            tagline      = "لا يوجد رقم مشترك — النظام يبحث بنفسه",
            description  = (
                "مثل: قائمتان بالأسماء واللقب فقط.\n"
                "النظام يقارن الكلمات ويتجاهل أخطاء الكتابة ويكتشف التطابق تلقائياً."
            ),
            bullets      = [
                "يتجاهل اختلافات الهمزة والحركات وترتيب الاسم واللقب",
                "يكتشف السجلات الزائدة في أي من الجدولين",
                "يمكن إضافة عمود ثانٍ (كالتاريخ) لتحسين الدقة",
            ],
            btn_text     = "ابدأ المقارنة العميقة",
            border_color = "#7B1FA2",
            btn_color    = "#7B1FA2",
            mode         = 'deep',
            is_fast      = False,
        )

        cards_row.addWidget(fast_card)
        cards_row.addWidget(deep_card)
        outer.addLayout(cards_row, 1)

        # ── صندوق النصيحة ─────────────────────────────────────
        tip = QLabel(
            "💡  نصيحة: ابحث عن عمود يحتوي على رقم أو كود فريد لكل صف في كلا الجدولين. "
            "إذا وجدته → اختر السريعة.    إذا لم تجده → اختر العميقة."
        )
        tip.setObjectName("landingTip")
        tip.setAlignment(Qt.AlignCenter)
        tip.setWordWrap(True)
        outer.addWidget(tip)

        return page

    def _make_mode_card(self, title, tagline, description, bullets,
                        btn_text, border_color, btn_color, mode, is_fast):
        card = QFrame()
        card.setObjectName("modeCardFast" if is_fast else "modeCardDeep")
        card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        v = QVBoxLayout(card)
        v.setContentsMargins(28, 28, 28, 24)
        v.setSpacing(12)

        # أيقونة
        icon_lbl = QLabel()
        icon_lbl.setAlignment(Qt.AlignCenter)
        icon_lbl.setFixedHeight(62)
        if is_fast:
            self.fast_card_icon = icon_lbl
        else:
            self.deep_card_icon = icon_lbl
        v.addWidget(icon_lbl)

        # عنوان البطاقة
        title_lbl = QLabel(title)
        title_lbl.setAlignment(Qt.AlignCenter)
        title_lbl.setObjectName("cardTitle")
        title_lbl.setStyleSheet(f"color: {border_color}; font-size: 17px; font-weight: bold;")
        v.addWidget(title_lbl)

        # الوسم التعريفي
        tagline_lbl = QLabel(tagline)
        tagline_lbl.setAlignment(Qt.AlignCenter)
        tagline_lbl.setWordWrap(True)
        tagline_lbl.setStyleSheet(
            f"color: {border_color}; font-size: 13px; font-weight: bold; "
            f"background: transparent; border: none;"
        )
        v.addWidget(tagline_lbl)

        # الوصف
        desc_lbl = QLabel(description)
        desc_lbl.setWordWrap(True)
        desc_lbl.setAlignment(Qt.AlignCenter)
        desc_lbl.setStyleSheet(
            "color: #555; font-size: 12px; font-weight: normal; "
            "background: transparent; border: none;"
        )
        v.addWidget(desc_lbl)

        # فاصل
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet(f"background-color: {border_color}; max-height: 1px; border: none;")
        v.addWidget(line)

        # النقاط
        for bullet in bullets:
            b = QLabel(f"✓   {bullet}")
            b.setWordWrap(True)
            b.setStyleSheet(
                "color: #333; font-size: 13px; font-weight: normal; "
                "background: transparent; border: none; padding: 1px 0;"
            )
            v.addWidget(b)

        v.addStretch()

        # زر البدء
        btn = QPushButton(btn_text)
        btn.setMinimumHeight(40)
        btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {btn_color};
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 14px;
                font-weight: bold;
                padding: 10px 20px;
            }}
            QPushButton:hover  {{ background-color: {btn_color}CC; }}
            QPushButton:pressed {{ background-color: {btn_color}99; }}
        """)
        btn.clicked.connect(lambda checked=False, m=mode: self.select_mode(m))
        v.addWidget(btn)

        return card

    # ──────────────────────────────────────────────────────────
    # منطق اختيار الوضع
    # ──────────────────────────────────────────────────────────

    def select_mode(self, mode: str):
        self.mode = mode

        # إعادة تعيين البيانات
        self.file1_path = ""
        self.file2_path = ""
        self.cols1 = []
        self.cols2 = []
        for le in [self.le_file1, self.le_file2]:
            le.setText("")
            le.setToolTip("")
        for cb in [self.cb_sheet1, self.cb_sheet2]:
            cb.blockSignals(True)
            cb.clear()
            cb.blockSignals(False)
        for cb in [self.cb_key1, self.cb_key2,
                   self.cb_match_col1, self.cb_match_col2,
                   self.cb_anchor_col1, self.cb_anchor_col2,
                   self.cb_map1, self.cb_map2]:
            cb.clear()
        self.table_maps.setRowCount(0)

        # إظهار/إخفاء الحقول المناسبة
        self._update_mode_widgets()

        self.current_step = 0
        self.stack.setCurrentIndex(1)
        self.update_navigation_ui()

    def _update_mode_widgets(self):
        """إظهار حقول العمود المفتاح (سريع) أو المطابقة (عميق) حسب الوضع."""
        is_fast = (self.mode == 'fast')
        self.key_widget1.setVisible(is_fast)
        self.key_widget2.setVisible(is_fast)
        self.deep_widget1.setVisible(not is_fast)
        self.deep_widget2.setVisible(not is_fast)
        if hasattr(self, 'thresh_group'):
            self.thresh_group.setVisible(not is_fast)

    def _on_threshold_changed(self, value: int):
        """تحديث التسمية والتلميح عند تغيير الشريط."""
        self.thresh_pct_label.setText(f"{value}%")

        if value >= 88:
            color = "#1A7A3C"
            hint  = (f"🟢  صارم جداً ({value}%) — يقبل فقط التطابقات شبه الحرفية. "
                     "مناسب للبيانات النظيفة المتسقة الكتابة.")
        elif value >= 75:
            color = "#1E3A5F"
            hint  = (f"🔵  متوازن ({value}%) — مناسب لمعظم الحالات. "
                     "يتجاهل أخطاء الهمزة والترتيب المعكوس للاسم واللقب.")
        elif value >= 62:
            color = "#E67E00"
            hint  = (f"🟠  متساهل ({value}%) — قد يُطابق أسماء مختلفة تشترك في كلمة. "
                     "ارفع النسبة إذا ظهرت مطابقات خاطئة في النتائج.")
        else:
            color = "#9B1C1C"
            hint  = (f"🔴  متساهل جداً ({value}%) — خطر مطابقات خاطئة مرتفع. "
                     "استخدم هذا فقط إذا كانت البيانات مكتوبة بأساليب مختلفة جداً.")

        self.thresh_pct_label.setStyleSheet(
            f"color: {color}; font-size: 18px; font-weight: bold;"
        )
        if hasattr(self, 'thresh_hint'):
            self.thresh_hint.setText(hint)
            self.thresh_hint.setStyleSheet(
                f"color: {color}; font-size: 12px; font-weight: normal; "
                "background: transparent; border: none; padding: 4px 0;"
            )

    # ──────────────────────────────────────────────────────────
    # مساعدات بناء الخطوات
    # ──────────────────────────────────────────────────────────

    def create_step_container(self, title):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(30, 16, 30, 16)
        layout.setSpacing(14)
        lbl = QLabel(title)
        lbl.setObjectName("stepTitle")
        lbl.setFont(create_font(14, True))
        layout.addWidget(lbl)
        return widget, layout

    def wrap_in_scroll_area(self, widget):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setObjectName("contentScrollArea")
        widget.setObjectName("stepContainerWidget")
        scroll.setWidget(widget)
        # منع scroll area من رفع الحد الأدنى لحجم النافذة
        # بدون هذا السطر، Qt يحسب الحد الأدنى من حجم المحتوى ويتجاوز الشاشة
        scroll.setMinimumSize(0, 0)
        scroll.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        return scroll

    def create_file_inputs(self, layout, idx):
        # اسم الملف في التقرير
        name_layout = QVBoxLayout()
        name_layout.addWidget(QLabel(
            f"اسم الملف في التقرير (مثلاً: {'البرنامج' if idx==1 else 'اليدوي'}):"
        ))
        le_name = QLineEdit("الملف الأول" if idx == 1 else "الملف الثاني")
        name_layout.addWidget(le_name)
        layout.addLayout(name_layout)

        # اختيار الملف
        file_layout = QVBoxLayout()
        file_layout.addWidget(QLabel("اختر ملف Excel:"))
        h_file = QHBoxLayout()
        le_file = QLineEdit()
        le_file.setPlaceholderText("لم يتم اختيار ملف...")
        le_file.setReadOnly(True)
        btn_file = QPushButton(" استيراد")
        btn_file.clicked.connect(lambda checked=False, i=idx: self.browse_file(i))
        if idx == 1:
            self.btn_file1 = btn_file
        else:
            self.btn_file2 = btn_file
        h_file.addWidget(le_file, 1)
        h_file.addWidget(btn_file)
        file_layout.addLayout(h_file)
        layout.addLayout(file_layout)

        # الورقة + صف العناوين
        opt_group = QGroupBox("إعدادات قراءة البيانات")
        opt_layout = QVBoxLayout(opt_group)
        h_opt = QHBoxLayout()

        sheet_vbox = QVBoxLayout()
        sheet_vbox.addWidget(QLabel("ورقة العمل (Sheet):"))
        cb_sheet = QComboBox()
        cb_sheet.setLayoutDirection(Qt.RightToLeft)
        cb_sheet.setMinimumHeight(36)
        cb_sheet.currentIndexChanged.connect(
            lambda _, i=idx: self.on_sheet_changed(i, auto_detect=True)
        )
        sheet_vbox.addWidget(cb_sheet)

        header_vbox = QVBoxLayout()
        header_vbox.addWidget(QLabel("صف العناوين (1 = الأول):"))
        sb_header = QSpinBox()
        sb_header.setMinimumHeight(36)
        sb_header.setRange(1, 500)
        sb_header.setValue(1)
        sb_header.valueChanged.connect(
            lambda _, i=idx: self.on_sheet_changed(i, auto_detect=False)
        )
        header_vbox.addWidget(sb_header)

        h_opt.addLayout(sheet_vbox, 2)
        h_opt.addLayout(header_vbox, 1)
        opt_layout.addLayout(h_opt)
        layout.addWidget(opt_group)

        # ── وضع سريع: العمود المفتاح ──────────────────────────
        key_widget = QWidget()
        kw_layout  = QVBoxLayout(key_widget)
        kw_layout.setContentsMargins(0, 0, 0, 0)
        kw_layout.setSpacing(6)
        kw_layout.addWidget(QLabel("العمود المفتاح (رقم أو كود فريد لكل سجل):"))
        cb_key = QComboBox()
        cb_key.setLayoutDirection(Qt.RightToLeft)
        cb_key.setMinimumHeight(36)
        kw_layout.addWidget(cb_key)
        layout.addWidget(key_widget)

        # ── وضع عميق: عمود المطابقة + عمود التدقيق ───────────
        deep_widget = QWidget()
        dw_layout   = QVBoxLayout(deep_widget)
        dw_layout.setContentsMargins(0, 0, 0, 0)
        dw_layout.setSpacing(10)

        dw_layout.addWidget(QLabel(
            "عمود المطابقة  (النص الذي سيستخدمه النظام للبحث عن التشابه):"
        ))
        cb_match = QComboBox()
        cb_match.setLayoutDirection(Qt.RightToLeft)
        cb_match.setMinimumHeight(36)
        dw_layout.addWidget(cb_match)

        anchor_lbl = QLabel(
            "عمود التدقيق — اختياري  "
            "(مثل التاريخ أو الرقم، يحسّن دقة المطابقة):"
        )
        dw_layout.addWidget(anchor_lbl)
        cb_anchor = QComboBox()
        cb_anchor.setLayoutDirection(Qt.RightToLeft)
        cb_anchor.setMinimumHeight(36)
        dw_layout.addWidget(cb_anchor)

        deep_widget.setVisible(False)   # مخفي حتى يُختار الوضع العميق
        layout.addWidget(deep_widget)

        # تخزين المراجع
        if idx == 1:
            (self.le_name1, self.le_file1, self.cb_sheet1,
             self.sb_header1, self.cb_key1,
             self.key_widget1, self.deep_widget1,
             self.cb_match_col1, self.cb_anchor_col1) = (
                le_name, le_file, cb_sheet, sb_header, cb_key,
                key_widget, deep_widget, cb_match, cb_anchor
            )
        else:
            (self.le_name2, self.le_file2, self.cb_sheet2,
             self.sb_header2, self.cb_key2,
             self.key_widget2, self.deep_widget2,
             self.cb_match_col2, self.cb_anchor_col2) = (
                le_name, le_file, cb_sheet, sb_header, cb_key,
                key_widget, deep_widget, cb_match, cb_anchor
            )

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

        self.btn_auto_match = QPushButton(" مطابقة تلقائية")
        self.btn_auto_match.setObjectName("autoMatchBtn")
        self.btn_auto_match.clicked.connect(self.auto_match_columns)
        layout.addWidget(self.btn_auto_match)

        # ── حساسية المطابقة (مرئية في الوضع العميق فقط) ──────
        self.thresh_group = QGroupBox("حساسية المطابقة  —  المقارنة العميقة")
        self.thresh_group.setObjectName("threshGroup")
        thresh_v = QVBoxLayout(self.thresh_group)
        thresh_v.setSpacing(10)

        # صف العنوان + القيمة الحالية
        thresh_top = QHBoxLayout()
        thresh_top.addWidget(QLabel("أدنى نسبة تشابه مقبولة بين السجلات:"))
        thresh_top.addStretch()
        self.thresh_pct_label = QLabel("75%")
        self.thresh_pct_label.setObjectName("threshPctLabel")
        thresh_top.addWidget(self.thresh_pct_label)
        thresh_v.addLayout(thresh_top)

        # الشريط المنزلق
        from PySide6.QtWidgets import QSlider
        self.threshold_slider = QSlider(Qt.Horizontal)
        self.threshold_slider.setRange(50, 100)
        self.threshold_slider.setValue(75)
        self.threshold_slider.setTickPosition(QSlider.TicksBelow)
        self.threshold_slider.setTickInterval(5)
        self.threshold_slider.setMinimumHeight(28)
        self.threshold_slider.valueChanged.connect(self._on_threshold_changed)
        thresh_v.addWidget(self.threshold_slider)

        # ملصقات الطرفين
        minmax_row = QHBoxLayout()
        lbl_low = QLabel("50%  —  متساهل")
        lbl_low.setStyleSheet("color:#888; font-weight:normal; font-size:12px;")
        lbl_high = QLabel("صارم  —  100%")
        lbl_high.setStyleSheet("color:#888; font-weight:normal; font-size:12px;")
        minmax_row.addWidget(lbl_low)
        minmax_row.addStretch()
        minmax_row.addWidget(lbl_high)
        thresh_v.addLayout(minmax_row)

        # شريط التلميح الديناميكي
        self.thresh_hint = QLabel()
        self.thresh_hint.setWordWrap(True)
        self.thresh_hint.setObjectName("threshHint")
        thresh_v.addWidget(self.thresh_hint)
        self._on_threshold_changed(75)   # تعيين النص الأولي

        self.thresh_group.setVisible(False)   # مخفي حتى يُختار الوضع العميق
        layout.addWidget(self.thresh_group)

        map_box = QGroupBox("إضافة مطابقة يدوية")
        map_h   = QHBoxLayout(map_box)

        v1 = QVBoxLayout()
        v1.addWidget(QLabel("عمود الملف 1:"))
        self.cb_map1 = QComboBox()
        self.cb_map1.setLayoutDirection(Qt.RightToLeft)
        self.cb_map1.setMinimumHeight(36)
        v1.addWidget(self.cb_map1)

        v2 = QVBoxLayout()
        v2.addWidget(QLabel("عمود الملف 2:"))
        self.cb_map2 = QComboBox()
        self.cb_map2.setLayoutDirection(Qt.RightToLeft)
        self.cb_map2.setMinimumHeight(36)
        v2.addWidget(self.cb_map2)

        btn_add = QPushButton(" إضافة")
        btn_add.setObjectName("addBtn")
        self.btn_add = btn_add
        btn_add.setFixedSize(110, 45)
        btn_add.clicked.connect(self.add_mapping)

        map_h.addLayout(v1, 1)
        map_h.addLayout(v2, 1)
        map_h.addWidget(btn_add, 0, Qt.AlignBottom)
        layout.addWidget(map_box)

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
        layout.addWidget(self.status_icon)

        self.lbl_status = QLabel("بانتظار بدء العملية...")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        self.lbl_status.setFont(create_font(12))
        layout.addWidget(self.lbl_status)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumHeight(25)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        self.table_stats = QTableWidget(9, 3)
        self.table_stats.setHorizontalHeaderLabels(["البيان", "القيمة", "النسبة %"])
        self.table_stats.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_stats.verticalHeader().setVisible(False)
        self.table_stats.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table_stats.setSelectionMode(QAbstractItemView.NoSelection)
        self.table_stats.setFixedHeight(280)
        self.table_stats.setVisible(False)
        self.table_stats.setStyleSheet("""
            QTableWidget { border: 1px solid #DEE2E6; border-radius: 4px; }
            QHeaderView::section { background-color: #2E5F8A; color: white; font-weight: bold; }
        """)
        layout.addWidget(self.table_stats)
        layout.addStretch()
        return widget

    # ──────────────────────────────────────────────────────────
    # معالجة الملفات
    # ──────────────────────────────────────────────────────────

    def browse_file(self, idx):
        default_dir = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        path, _ = QFileDialog.getOpenFileName(
            self, "اختر ملف Excel", default_dir, "Excel Files (*.xlsx *.xls)"
        )
        if not path:
            return
        le = self.le_file1 if idx == 1 else self.le_file2
        le.setText(os.path.basename(path))
        le.setToolTip(path)

        if idx == 1:
            self.file1_path = path
        else:
            self.file2_path = path

        success, sheets, headers = get_file_info(path)
        if success:
            if idx == 1:
                self.file1_headers = headers
            else:
                self.file2_headers = headers
            cb = self.cb_sheet1 if idx == 1 else self.cb_sheet2
            cb.blockSignals(True)
            cb.clear()
            cb.addItems(sheets)
            cb.blockSignals(False)
            self.on_sheet_changed(idx, auto_detect=True)
        else:
            QMessageBox.critical(self, "خطأ", f"فشل قراءة الملف: {sheets}")

    def on_sheet_changed(self, idx, auto_detect=False):
        path      = self.file1_path     if idx == 1 else self.file2_path
        cb_sheet  = self.cb_sheet1      if idx == 1 else self.cb_sheet2
        sb_header = self.sb_header1     if idx == 1 else self.sb_header2
        cb_key    = self.cb_key1        if idx == 1 else self.cb_key2
        cb_match  = self.cb_match_col1  if idx == 1 else self.cb_match_col2
        cb_anchor = self.cb_anchor_col1 if idx == 1 else self.cb_anchor_col2

        sheet = cb_sheet.currentText()
        if not path or not sheet:
            return

        if auto_detect:
            headers_dict = self.file1_headers if idx == 1 else self.file2_headers
            best = headers_dict.get(sheet, 0)
            sb_header.blockSignals(True)
            sb_header.setValue(best + 1)
            sb_header.blockSignals(False)

        success, cols = get_sheet_columns(path, sheet, sb_header.value() - 1)
        if success:
            if idx == 1:
                self.cols1 = list(cols)
            else:
                self.cols2 = list(cols)

            if self.current_step == 2 and self.table_maps.rowCount() > 0:
                self.table_maps.setRowCount(0)
                QMessageBox.information(
                    self, "تنبيه",
                    "تم مسح جدول المطابقات بسبب تغيير الورقة/العناوين. يرجى إعادة المطابقة."
                )

            # العمود المفتاح (سريع)
            cb_key.clear()
            cb_key.addItems(cols)

            # عمود المطابقة + عمود التدقيق (عميق)
            cb_match.clear()
            cb_match.addItems(cols)

            cb_anchor.clear()
            cb_anchor.addItem("— بدون عمود تدقيق —")
            cb_anchor.addItems(cols)

            # قوائم خطوة 3
            if idx == 1:
                self.cb_map1.clear()
                self.cb_map1.addItems(cols)
            else:
                self.cb_map2.clear()
                self.cb_map2.addItems(cols)
        else:
            cb_key.clear()
            QMessageBox.warning(self, "تنبيه", f"تعذر جلب الأعمدة: {cols}")

    # ──────────────────────────────────────────────────────────
    # خطوة 3: مطابقة الأعمدة
    # ──────────────────────────────────────────────────────────

    def add_mapping(self):
        c1 = self.cb_map1.currentText()
        c2 = self.cb_map2.currentText()
        if not c1 or not c2:
            return
        for r in range(self.table_maps.rowCount()):
            if (self.table_maps.item(r, 0).text() == c1 and
                    self.table_maps.item(r, 1).text() == c2):
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
        lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(btn_del)
        self.table_maps.setCellWidget(row, 2, container)

    def remove_map_row(self, checked=False):
        button = self.sender()
        if not button:
            return
        container = button.parentWidget()
        for r in range(self.table_maps.rowCount()):
            if self.table_maps.cellWidget(r, 2) == container:
                self.table_maps.removeRow(r)
                break

    def auto_match_columns(self):
        cols1 = [self.cb_map1.itemText(i) for i in range(self.cb_map1.count())]
        cols2 = [self.cb_map2.itemText(i) for i in range(self.cb_map2.count())]
        matched, unmatched = 0, []
        for c1 in cols1:
            found = False
            for c2 in cols2:
                if c1.strip().lower() == c2.strip().lower():
                    self.cb_map1.setCurrentText(c1)
                    self.cb_map2.setCurrentText(c2)
                    self.add_mapping()
                    matched += 1
                    found = True
                    break
            if not found:
                unmatched.append(c1)
        if matched > 0:
            msg = f"تمت مطابقة {matched} عمود تلقائياً."
            if unmatched:
                msg += "\n\n⚠️ لم يُوجد مثيل لـ:\n- " + "\n- ".join(unmatched[:10])
            QMessageBox.information(self, "نتائج المطابقة", msg)
        else:
            QMessageBox.warning(self, "تنبيه", "لم يتم العثور على أعمدة بنفس الاسم.")

    # ──────────────────────────────────────────────────────────
    # التنقل بين الخطوات
    # ──────────────────────────────────────────────────────────

    def go_next(self):
        if self.current_step == 0:
            if not self.file1_path:
                QMessageBox.warning(self, "خطأ", "يرجى اختيار الملف المرجعي أولاً.")
                return
        elif self.current_step == 1:
            if not self.file2_path:
                QMessageBox.warning(self, "خطأ", "يرجى اختيار الملف المقارن أولاً.")
                return
            
            self.cb_map1.clear()
            self.cb_map1.addItems(self.cols1)
            self.cb_map2.clear()
            self.cb_map2.addItems(self.cols2)
        elif self.current_step == 2:
            # مرحلة المقارنة: التحقق من الأعمدة ثم الإطلاق
            if self.mode == 'fast' and self.table_maps.rowCount() == 0:
                QMessageBox.warning(self, "خطأ", "يرجى إضافة عمود واحد على الأقل للمقارنة.")
                return
            self.current_step += 1
            self.stack.setCurrentIndex(self.current_step + 1)
            self.update_navigation_ui()
            self.run_process()
            return

        self.current_step += 1
        self.stack.setCurrentIndex(self.current_step + 1)
        self.update_navigation_ui()

    def go_back(self):
        self.current_step -= 1
        self.stack.setCurrentIndex(self.current_step + 1)
        self.update_navigation_ui()

    def on_sidebar_clicked(self, s_idx):
        if hasattr(self, 'worker') and self.worker.isRunning():
            return
        if s_idx == self.current_step:
            return
        if s_idx == -1:
            self.current_step = -1
            self.stack.setCurrentIndex(0)
            self.update_navigation_ui()
            return
        if self.mode is None:
            return

        if s_idx == 0:
            self.current_step = 0
            self.stack.setCurrentIndex(1)
            self.update_navigation_ui()
        elif s_idx == 1:
            if not self.file1_path:
                QMessageBox.warning(self, "تنبيه", "يرجى اختيار الملف المرجعي أولاً.")
                return
            self.current_step = 1
            self.stack.setCurrentIndex(2)
            self.update_navigation_ui()
        elif s_idx == 2:
            if not self.file1_path:
                QMessageBox.warning(self, "تنبيه", "يرجى اختيار الملف المرجعي أولاً.")
                return
            if not self.file2_path:
                QMessageBox.warning(self, "تنبيه", "يرجى اختيار الملف المقارن أولاً.")
                return
            self.current_step = 2
            self.stack.setCurrentIndex(3)
            self.update_navigation_ui()
        elif s_idx == 3:
            if self.current_step == 2:
                self.go_next()

    def update_sidebar_active_step(self):
        active_idx = self.current_step + 1
        for idx, item in enumerate(self.step_widgets):
            frame = item['frame']
            num_lbl = item['num_lbl']
            txt_lbl = item['txt_lbl']
            
            if idx == active_idx:
                frame.setProperty("active", "true")
                frame.setStyleSheet("""
                    QFrame {
                        background-color: #2E5F8A;
                        border-radius: 8px;
                    }
                    QFrame:hover {
                        background-color: #3A72A4;
                    }
                    QLabel {
                        color: white;
                        font-weight: bold;
                    }
                """)
                num_lbl.setStyleSheet("border-radius: 12px; background-color: #1A7A3C; color: white;")
            elif idx < active_idx:
                frame.setProperty("active", "completed")
                frame.setStyleSheet("""
                    QFrame {
                        background-color: transparent;
                    }
                    QFrame:hover {
                        background-color: rgba(255, 255, 255, 0.1);
                        border-radius: 8px;
                    }
                    QLabel {
                        color: #A0C4FF;
                        font-weight: normal;
                    }
                """)
                num_lbl.setText("✓")
                num_lbl.setStyleSheet("border-radius: 12px; background-color: #1A7A3C; color: white;")
            else:
                frame.setProperty("active", "false")
                frame.setStyleSheet("""
                    QFrame {
                        background-color: transparent;
                    }
                    QFrame:hover {
                        background-color: rgba(255, 255, 255, 0.05);
                        border-radius: 8px;
                    }
                    QLabel {
                        color: #B0BEC5;
                        font-weight: normal;
                    }
                """)
                num_lbl.setText(item['default_text'])
                num_lbl.setStyleSheet("border-radius: 12px; background-color: #37474F; color: #B0BEC5;")
            
            frame.style().unpolish(frame)
            frame.style().polish(frame)
            num_lbl.style().unpolish(num_lbl)
            num_lbl.style().polish(num_lbl)
            txt_lbl.style().unpolish(txt_lbl)
            txt_lbl.style().polish(txt_lbl)

    def update_navigation_ui(self):
        self.update_sidebar_active_step()
        is_landing = (self.current_step == -1)

        # إخفاء أزرار التنقل كلياً في الصفحة الرئيسية
        self.btn_back.setVisible(not is_landing)
        self.btn_next.setVisible(not is_landing)

        self.subtitle_label.setText(self.step_titles.get(self.current_step, ""))

        if is_landing:
            self.btn_open_report.setVisible(False)
            return

        self.btn_back.setEnabled(True)

        is_last = (self.current_step == 3)
        self.btn_next.setVisible(not is_last)

        if self.current_step == 2:
            self.btn_next.setText(" إجراء المقارنة")
            self.btn_next.setIcon(get_icon('fa5s.play', color='white'))
        else:
            self.btn_next.setText(" التالي")
            self.btn_next.setIcon(get_icon('fa5s.arrow-left', color='white'))

        if is_last and hasattr(self, 'worker') and self.worker.isRunning():
            self.btn_back.setEnabled(False)

        self.btn_open_report.setVisible(
            is_last and hasattr(self, 'last_success') and self.last_success
        )

    # ──────────────────────────────────────────────────────────
    # تشغيل المقارنة
    # ──────────────────────────────────────────────────────────

    def run_process(self):
        if hasattr(self, 'worker') and self.worker.isRunning():
            QMessageBox.warning(self, "تنبيه", "هناك عملية قيد التنفيذ. يرجى الانتظار.")
            return

        default_dir  = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        default_file = os.path.join(default_dir, 'تقرير_المقارنة.xlsx')
        output_path, _ = QFileDialog.getSaveFileName(
            self, "حفظ التقرير", default_file, "Excel Files (*.xlsx)"
        )
        if not output_path:
            self.go_back()
            return

        self.output_file = output_path
        maps = [
            {'col1': self.table_maps.item(r, 0).text(),
             'col2': self.table_maps.item(r, 1).text()}
            for r in range(self.table_maps.rowCount())
        ]
        self.btn_back.setEnabled(False)

        if self.mode == 'fast':
            if not self.cb_key1.currentText() or not self.cb_key2.currentText():
                QMessageBox.warning(self, "خطأ", "يرجى اختيار أعمدة المفتاح لكلا الملفين.")
                self.go_back()
                return
            params = {
                'file1': self.file1_path, 'sheet1': self.cb_sheet1.currentText(),
                'header1': self.sb_header1.value() - 1, 'key1': self.cb_key1.currentText(),
                'name1': self.le_name1.text(),
                'file2': self.file2_path, 'sheet2': self.cb_sheet2.currentText(),
                'header2': self.sb_header2.value() - 1, 'key2': self.cb_key2.currentText(),
                'name2': self.le_name2.text(),
                'mappings': maps, 'output': output_path,
            }
            self.worker = ComparisonWorker(params)

        else:  # deep
            if not self.cb_match_col1.currentText() or not self.cb_match_col2.currentText():
                QMessageBox.warning(self, "خطأ", "يرجى اختيار عمود المطابقة لكلا الملفين.")
                self.go_back()
                return
            anchor1_raw = self.cb_anchor_col1.currentText()
            anchor2_raw = self.cb_anchor_col2.currentText()
            anchor1 = None if anchor1_raw.startswith("—") else anchor1_raw
            anchor2 = None if anchor2_raw.startswith("—") else anchor2_raw
            min_sim = getattr(self, 'threshold_slider', None)
            min_sim = min_sim.value() if min_sim else 75
            params = {
                'file1': self.file1_path, 'sheet1': self.cb_sheet1.currentText(),
                'header1': self.sb_header1.value() - 1,
                'match_col1': self.cb_match_col1.currentText(),
                'anchor_col1': anchor1, 'name1': self.le_name1.text(),
                'file2': self.file2_path, 'sheet2': self.cb_sheet2.currentText(),
                'header2': self.sb_header2.value() - 1,
                'match_col2': self.cb_match_col2.currentText(),
                'anchor_col2': anchor2, 'name2': self.le_name2.text(),
                'mappings': maps, 'output': output_path,
                'min_similarity': min_sim,
            }
            self.worker = DeepComparisonWorker(params)

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
            self.status_icon.setPixmap(
                get_icon('fa5s.check-circle', color='#1A7A3C').pixmap(64, 64)
            )
            self.lbl_status.setText("تمت العملية بنجاح!")
            rows = [
                (f"إجمالي أسطر {self.le_name1.text()}", str(stats.get('total1', 0)), "—"),
                (f"إجمالي أسطر {self.le_name2.text()}", str(stats.get('total2', 0)), "—"),
                ("السجلات المشتركة/المتطابقة",           str(stats.get('common', 0)), "—"),
                (f"فقط في {self.le_name1.text()}",
                 str(stats.get('only1', 0)),
                 f"{stats.get('only1',0)/max(stats.get('common',1)+stats.get('only1',0),1)*100:.1f}%"),
                (f"فقط في {self.le_name2.text()}",
                 str(stats.get('only2', 0)),
                 f"{stats.get('only2',0)/max(stats.get('common',1)+stats.get('only2',0),1)*100:.1f}%"),
                ("متطابقون تماماً",   str(stats.get('match', 0)),
                 f"{stats.get('match',0)/max(stats.get('common',1),1)*100:.1f}%"),
                ("يحتوون اختلافات",   str(stats.get('diff',  0)),
                 f"{stats.get('diff', 0)/max(stats.get('common',1),1)*100:.1f}%"),
                (f"🔁 مكررات في {self.le_name1.text()}", str(stats.get('dup1', 0)), "—"),
                (f"🔁 مكررات في {self.le_name2.text()}", str(stats.get('dup2', 0)), "—"),
            ]
            for i, (label, val, pct) in enumerate(rows):
                self.table_stats.setItem(i, 0, QTableWidgetItem(label))
                self.table_stats.setItem(i, 1, QTableWidgetItem(val))
                self.table_stats.setItem(i, 2, QTableWidgetItem(pct))
                if i == 5:
                    color = QColor("#D6F5D6")
                    for c in range(3): self.table_stats.item(i, c).setBackground(color)
                elif i == 6:
                    color = QColor("#FFD6D6")
                    for c in range(3): self.table_stats.item(i, c).setBackground(color)
                elif i in (7, 8) and int(val) > 0:
                    color = QColor("#FFF3CD")
                    for c in range(3): self.table_stats.item(i, c).setBackground(color)
                for c in range(3):
                    self.table_stats.item(i, c).setTextAlignment(Qt.AlignCenter)
            self.table_stats.setVisible(True)
            self.btn_open_report.setVisible(True)
            QMessageBox.information(self, "نجاح", msg)
        else:
            self.status_icon.setPixmap(
                get_icon('fa5s.exclamation-triangle', color='#D32F2F').pixmap(64, 64)
            )
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

    # ──────────────────────────────────────────────────────────
    # التصميم
    # ──────────────────────────────────────────────────────────

    def apply_styles(self):
        base_path  = os.path.dirname(os.path.abspath(__file__))
        assets_path = os.path.join(base_path, "assets").replace("\\", "/")
        arrow_down  = f"{assets_path}/arrow_down.svg"
        arrow_up    = f"{assets_path}/arrow_up.svg"

        style = f"""
        QMainWindow {{ background-color: #F8F9FA; font-family: "Segoe UI", "Tahoma", "Arial", "Geeza Pro", "DejaVu Sans", sans-serif; font-size: 14px; }}
        #contentScrollArea {{ background-color: #F8F9FA; border: none; }}
        #stepContainerWidget {{ background-color: #F8F9FA; }}
        QStackedWidget {{ background-color: #F8F9FA; }}

        /* ── الشريط الجانبي (Sidebar) ── */
        #sidebarFrame {{
            background-color: #1E3A5F;
            border-left: 2px solid #1A7A3C;
        }}
        #sidebarTitle {{
            color: #FFFFFF;
            font-size: 18px;
            font-weight: bold;
        }}
        #sidebarSubtitle {{
            color: #A0C4FF;
            font-size: 11px;
        }}
        #sidebarLine {{
            background-color: #2E5F8A;
            max-height: 1px;
            border: none;
        }}
        
        #topBar {{
            background-color: #FFFFFF;
            border-bottom: 1px solid #DEE2E6;
        }}
        #topBarSubtitle {{
            color: #1E3A5F;
            font-size: 16px;
        }}

        /* ── شريط التنقل ── */
        #navFrame {{ background-color: #FFFFFF; border-top: 1px solid #DEE2E6; }}

        /* ── الصفحة الرئيسية ── */
        #landingPage {{ background-color: #F0F4FF; }}

        #landingMainTitle {{
            color: #1E3A5F;
            font-size: 20px;
            font-weight: bold;
            padding: 6px 0;
        }}
        #landingSep {{
            background-color: #C5D3E8;
            max-height: 1px;
            border: none;
        }}

        #modeCardFast {{
            background-color: #FFFFFF;
            border: 2px solid #1E3A5F;
            border-radius: 14px;
        }}
        #modeCardDeep {{
            background-color: #FFFFFF;
            border: 2px solid #7B1FA2;
            border-radius: 14px;
        }}

        #landingTip {{
            background-color: #FFF8E1;
            border: 1px solid #FFD54F;
            border-radius: 8px;
            padding: 10px 18px;
            color: #5D4037;
            font-size: 13px;
            font-weight: normal;
        }}

        /* ── خطوات المعالج ── */
        #stepTitle {{ color: #1E3A5F; margin-bottom: 10px; font-size: 18px; font-weight: bold; }}

        QLabel {{ color: #495057; font-size: 14px; font-weight: bold; }}

        QLineEdit {{
            padding: 10px; border: 1px solid #CED4DA; border-radius: 5px;
            background-color: #FFFFFF; color: #212529; font-size: 14px;
        }}
        QLineEdit:focus {{ border-color: #1E3A5F; }}

        QComboBox {{
            padding: 8px 10px 8px 35px; border: 1px solid #CED4DA; border-radius: 5px;
            background-color: #FFFFFF; color: #212529; font-size: 14px;
        }}
        QComboBox QAbstractItemView {{
            background-color: #FFFFFF; color: #212529;
            selection-background-color: #1E3A5F; selection-color: #FFFFFF;
        }}
        QComboBox::drop-down {{
            subcontrol-origin: padding; subcontrol-position: top right;
            width: 30px; border-right: 1px solid #CED4DA; border-left: none;
        }}
        QComboBox::down-arrow {{ image: url("{arrow_down}"); width: 14px; height: 14px; }}

        QSpinBox {{
            padding: 8px 25px 8px 12px; border: 1px solid #CED4DA; border-radius: 5px;
            background-color: #FFFFFF; color: #212529; font-size: 14px;
            qproperty-alignment: 'AlignRight | AlignVCenter';
        }}
        QSpinBox::up-button {{
            subcontrol-origin: border; subcontrol-position: top right;
            width: 25px; border-right: 1px solid #CED4DA;
            border-bottom: 1px solid #CED4DA; border-left: none;
        }}
        QSpinBox::down-button {{
            subcontrol-origin: border; subcontrol-position: bottom right;
            width: 25px; border-right: 1px solid #CED4DA; border-left: none;
        }}
        QSpinBox::up-arrow   {{ image: url("{arrow_up}");   width: 12px; height: 12px; }}
        QSpinBox::down-arrow {{ image: url("{arrow_down}"); width: 12px; height: 12px; }}

        QPushButton {{
            padding: 10px 20px; border-radius: 5px; font-weight: bold;
            background-color: #E9ECEF; border: 1px solid #ADB5BD;
            color: #1E3A5F; font-size: 14px;
        }}
        QPushButton:hover {{ background-color: #DEE2E6; }}

        #nextBtn     {{ background-color: #1E3A5F; color: white; border: none; }}
        #nextBtn:hover {{ background-color: #2E5F8A; }}
        #addBtn      {{ background-color: #2E5F8A; color: white; border: none; }}
        #runBtn      {{ background-color: #1A7A3C; color: white; border: none; font-size: 18px; }}
        #autoMatchBtn {{
            background-color: #17A2B8; color: white; border: none;
            font-weight: bold; padding: 12px 20px;
        }}
        #autoMatchBtn:hover {{ background-color: #138496; }}

        QGroupBox {{
            font-weight: bold; border: 1px solid #DEE2E6; border-radius: 8px;
            margin-top: 20px; padding-top: 25px; background-color: #FFFFFF;
            font-size: 15px; color: #1E3A5F;
        }}
        QGroupBox::title {{
            subcontrol-origin: margin; subcontrol-position: top center;
            padding: 0 10px; color: #1E3A5F;
        }}

        QTableWidget {{
            background-color: white; border: 1px solid #DEE2E6;
            border-radius: 5px; gridline-color: #F1F3F5;
            color: #212529; font-size: 14px;
        }}
        QTableWidget::item {{ color: #212529; }}
        QHeaderView::section {{
            background-color: #F8F9FA; padding: 10px; border: none;
            border-bottom: 2px solid #DEE2E6; font-weight: bold;
            color: #495057; font-size: 14px;
        }}

        QProgressBar {{
            border: 1px solid #DEE2E6; border-radius: 10px; text-align: center;
            background-color: #E9ECEF; color: #212529;
            font-size: 14px; font-weight: bold;
        }}
        QProgressBar::chunk {{ background-color: #1A7A3C; border-radius: 9px; }}

        QMessageBox {{ background-color: #FFFFFF; border: 1px solid #DEE2E6; }}
        QMessageBox QLabel {{ color: #212529; font-size: 14px; font-weight: normal; }}
        QMessageBox QPushButton {{
            background-color: #1E3A5F; color: white; border: 1px solid #1E3A5F;
            border-radius: 5px; padding: 8px 18px; font-weight: bold;
            min-width: 80px; font-size: 13px;
        }}
        QMessageBox QPushButton:hover {{ background-color: #2E5F8A; border-color: #2E5F8A; }}
        """
        self.setStyleSheet(style)
