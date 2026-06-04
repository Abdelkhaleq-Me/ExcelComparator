"""
core.py — منطق المعالجة والمقارنة (نسخة محسّنة للأداء)

الأوضاع:
  run_comparison()       ← المقارنة السريعة (بعمود مفتاح)
  run_deep_comparison()  ← المقارنة العميقة (بتشابه النصوص)

تحسينات الأداء:
  1. cdist (C + SIMD)  بدلاً من حلقة Python للمقارنة النصية   → ×16-50
  2. lru_cache على التطبيع النصي                               → ×10
  3. numpy .values بدلاً من .iloc لكل صف في التقرير           → ×17
  4. فصل الصفوف المختلفة مسبقاً عبر numpy mask               → ×24
  5. استراتيجية تكيّفية: Hungarian للمصفوفات الصغيرة،
     Top-K greedy للكبيرة (تجنّباً لاستهلاك الذاكرة)
"""

import os
import re
import warnings
import traceback
from collections import defaultdict
from functools import lru_cache
import openpyxl
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle
from openpyxl.utils import get_column_letter

try:
    from rapidfuzz import fuzz
    from rapidfuzz.process import cdist as rf_cdist
    from rapidfuzz.process import extract as rf_extract
    from scipy.optimize import linear_sum_assignment
    _HAS_DEEP_LIBS = True
except ImportError:
    _HAS_DEEP_LIBS = False

from dataclasses import dataclass

@dataclass
class LoadResultFast:
    success: bool
    error: str = ""
    df1: pd.DataFrame = None
    df2: pd.DataFrame = None
    dup1_df: pd.DataFrame = None
    dup2_df: pd.DataFrame = None
    common: list = None
    only1: list = None
    only2: list = None
    total1: int = 0
    total2: int = 0

@dataclass
class LoadResultDeep:
    success: bool
    error: str = ""
    df1: pd.DataFrame = None
    df2: pd.DataFrame = None
    n1: int = 0
    n2: int = 0
    tokens1: list = None
    tokens2: list = None
    anchors1: list = None
    anchors2: list = None
    dup1_df: pd.DataFrame = None
    dup2_df: pd.DataFrame = None


# ─────────────────────────────────────────────────────────────
# ثوابت الأداء
# ─────────────────────────────────────────────────────────────

# الحدّ الأقصى لحجم مصفوفة الدرجات قبل التبديل إلى Top-K
# 4M خلية × 4 بايت (float32) = 16MB — آمن حتى على أجهزة 2GB RAM
_MAX_MATRIX_CELLS = 4_000_000

# عدد المرشّحين في وضع Top-K Greedy
_TOP_K = 10


# ─────────────────────────────────────────────────────────────
# دوال تنظيف البيانات — وضع سريع
# ─────────────────────────────────────────────────────────────

def clean_id(s):
    if pd.isna(s):
        return ''
    s_str = str(s).strip()
    result = s_str.lstrip('0')
    return result if result else s_str


@lru_cache(maxsize=4096)
def _norm_cached(s: str) -> str:
    if not s or s in ('nan', 'NaT', 'None'):
        return ''
    s = re.sub(r'[\u064b-\u0652\u0670\u0640]', '', s)
    try:
        f = float(s)
        return str(int(f)) if f == int(f) else f'{f:.10g}'
    except (ValueError, TypeError, OverflowError):
        pass
    try:
        return pd.to_datetime(s, dayfirst=True).strftime('%Y-%m-%d')
    except (ValueError, TypeError):
        pass
    return s


def norm(v) -> str:
    if pd.isna(v):
        return ''
    return _norm_cached(str(v).strip())


# ─────────────────────────────────────────────────────────────
# دوال تطبيع النصوص — وضع عميق  (مُحسَّنة بـ lru_cache)
# ─────────────────────────────────────────────────────────────

# الدالة الأساسية: تُطبَّق على سلاسل نصية نظيفة فقط،
# ممّا يتيح lru_cache الاستفادة الكاملة من النتائج المحفوظة.
@lru_cache(maxsize=131072)
def _sort_tokens_str(text: str) -> str:
    """
    تطبيع عميق + فرز الكلمات أبجدياً (مخزَّن مؤقتاً).

    التطبيعات المُطبَّقة بالترتيب:
      1. حذف التشكيل والحركات
      2. توحيد الألف وأشكالها  (أ إ آ ٱ  →  ا)
      3. توحيد الياء            (ى  →  ي)
      4. توحيد التاء المربوطة   (ة  →  ه)
      5. حذف ألف الوصل من أول كل كلمة إذا لم تكن "ال"
         مثال: "امبارك" → "مبارك" | "ابن" → "بن"
      6. دمج "عبد" مع الكلمة التالية
         مثال: "عبد الخالق" → "عبدالخالق"
      7. فرز الكلمات أبجدياً   (يعالج عكس الاسم واللقب)
    """
    # 1-4: تطبيعات الأحرف الأساسية
    text = re.sub(r'[\u064B-\u065F\u0670\u0640]', '', text)
    text = re.sub(r'[أإآٱ]', 'ا', text)
    text = re.sub(r'ى',      'ي', text)
    text = re.sub(r'ة',      'ه', text)
    text = re.sub(r'\s+',    ' ', text).strip()

    # 7: حذف ألف الوصل من بداية كل كلمة
    # القاعدة: اِحذف الألف إذا:
    #   • ما يليها ليس "ل" (لتجنّب تفكيك "ال" التعريف)
    #   • ما تبقى بعد الحذف >= حرفان (لتجنّب كلمات مفردة الحرف)
    tokens = []
    for tok in text.split():
        if len(tok) >= 3 and tok[0] == 'ا' and tok[1] != 'ل':
            tok = tok[1:]
        tokens.append(tok)

    # 8: دمج "عبد" مع الكلمة التالية
    # يعالج: "عبد الخالق" ↔ "عبدالخالق" | "عبد الرحمن" ↔ "عبدالرحمن"
    # "عبد" وحدها (آخر كلمة أو الاسم كله) تبقى كما هي
    merged = []
    i = 0
    while i < len(tokens):
        if tokens[i] == 'عبد' and i + 1 < len(tokens):
            merged.append('عبد' + tokens[i + 1])
            i += 2
        else:
            merged.append(tokens[i])
            i += 1
    tokens = merged

    # 9: الفرز الأبجدي
    return ' '.join(sorted(tokens))


def _sort_tokens(v) -> str:
    """غلاف يعالج القيم الخاصة (NaN, None) قبل الإرسال للدالة المخزَّنة."""
    if pd.isna(v):
        return ''
    return _sort_tokens_str(str(v).strip())


@lru_cache(maxsize=32768)
def _norm_anchor_str(s: str) -> str:
    """تطبيع عمود التدقيق (تاريخ أو نص) — مخزَّن مؤقتاً."""
    try:
        return pd.to_datetime(s).strftime('%Y-%m-%d')
    except Exception:
        return _sort_tokens_str(s)


def _norm_anchor(v) -> str:
    if pd.isna(v):
        return ''
    return _norm_anchor_str(str(v).strip())


# ─────────────────────────────────────────────────────────────
# دوال قراءة معلومات الملف
# ─────────────────────────────────────────────────────────────

def _detect_header_in_ws(ws, max_rows: int = 25) -> int:
    best_row, best_score = 0, -1
    for r_idx, row in enumerate(ws.iter_rows(max_row=max_rows), 0):
        cells = [c for c in row if c.value is not None and str(c.value).strip()]
        total = len(cells)
        if total == 0:
            continue
        text_ratio = sum(1 for c in cells if c.data_type == 's') / total
        score = total * (0.5 + text_ratio)
        if score > best_score:
            best_score, best_row = score, r_idx
    return best_row


def get_file_info(file_path: str):
    try:
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning)
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        headers = {s: _detect_header_in_ws(wb[s]) for s in sheet_names}
        wb.close()
        return True, sheet_names, headers
    except Exception as e:
        return False, str(e), {}


def get_excel_info(file_path: str):
    try:
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning)
            wb = openpyxl.load_workbook(file_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return True, names
    except Exception as e:
        return False, str(e)


def get_sheet_columns(file_path: str, sheet_name: str, header_row: int = 0):
    try:
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning)
            df = pd.read_excel(file_path, sheet_name=sheet_name,
                               header=header_row, nrows=0, engine='openpyxl')
        cols = [str(c).strip().replace('\n', ' ').replace('  ', ' ')
                for c in df.columns]
        return True, cols
    except ValueError as e:
        if 'header' in str(e).lower():
            return False, f"خطأ: صف العناوين ({header_row + 1}) يتجاوز عدد الأسطر."
        return False, str(e)
    except Exception as e:
        return False, str(e)


def get_sheet_preview(file_path: str, sheet_name: str, nrows: int = 12):
    try:
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning)
            df = pd.read_excel(file_path, sheet_name=sheet_name,
                               header=None, nrows=nrows, engine='openpyxl')
        return True, df.fillna('').values.tolist()
    except Exception as e:
        return False, str(e)





# ─────────────────────────────────────────────────────────────
# NamedStyles للتقرير
# ─────────────────────────────────────────────────────────────

def _build_named_styles():
    thin    = Side(style='thin',   color='FFB0BEC5')
    med     = Side(style='medium', color='FF78909C')
    brd     = Border(top=thin, bottom=thin, left=thin, right=thin)
    brd_med = Border(top=med,  bottom=med,  left=med,  right=med)
    align_c = Alignment(horizontal='center', vertical='center', wrap_text=True)

    C = dict(
        hdr='FF1E3A5F', sub='FF2E5F8A', white='FFFFFFFF', stripe='FFF0F4F8',
        match='FFD6F5D6', diff_row='FFFFD6D6', only='FFFFFF99',
        green_h='FF1A7A3C', red_h='FF9B1C1C', gold_bg='FFFFFCE8',
        diff_cell='FFFEE2E2', slate='FF455A64', dark_grn='FF1A5C2A',
        brown2='FF9E6A00', purple2='FF7B1FA2', deep_hdr='FF4A148C',
    )

    def _ns(name, fg, fill, bdr, bold=False, sz=14):
        ns           = NamedStyle(name=name)
        ns.font      = Font(name='Calibri', size=sz, bold=bold, color=fg)
        ns.fill      = PatternFill('solid', fgColor=fill)
        ns.alignment = align_c
        ns.border    = bdr
        return ns

    W = 'FFFFFFFF'
    D = 'FF1E1E1E'
    return [
        _ns('ns_hdr',        W, C['hdr'],       brd_med, bold=True, sz=16),
        _ns('ns_sub',        W, C['sub'],        brd_med, bold=True, sz=14),
        _ns('ns_slate',      W, C['slate'],      brd_med, bold=True, sz=14),
        _ns('ns_dark_grn',   W, C['dark_grn'],   brd_med, bold=True, sz=14),
        _ns('ns_brown2',     W, C['brown2'],     brd_med, bold=True, sz=14),
        _ns('ns_purple2',    W, C['purple2'],    brd_med, bold=True, sz=14),
        _ns('ns_deep_hdr',   W, C['deep_hdr'],   brd_med, bold=True, sz=16),
        _ns('ns_dat_white',  D, C['white'],      brd, sz=14),
        _ns('ns_dat_stripe', D, C['stripe'],     brd, sz=14),
        _ns('ns_dat_match',  C['green_h'], C['match'],     brd, bold=True, sz=14),
        _ns('ns_dat_diff',   C['red_h'],   C['diff_row'],  brd, bold=True, sz=14),
        _ns('ns_diff_cell',  C['red_h'],   C['diff_cell'], brd, bold=True, sz=14),
        _ns('ns_dat_gold',   D, C['gold_bg'],    brd, sz=14),
        _ns('ns_dat_only',   D, C['only'],       brd, sz=14),
    ], C

_GLOBAL_STYLES, _GLOBAL_COLORS = _build_named_styles()


# ─────────────────────────────────────────────────────────────
# دوال مساعدة لكتابة التقرير
# ─────────────────────────────────────────────────────────────

def _wc(cell, style, value=None):
    cell.style = style
    if value is not None:
        cell.value = value

def _write_only_sheet(ws, title, ids, df, nc, s_hdr):
    """
    كتابة ورقة السجلات الموجودة في جدول واحد فقط.
    مُحسَّنة: تستخدم .values بدلاً من .loc لكل صف → ×17 أسرع.
    """
    ws.sheet_view.rightToLeft = True
    ws.row_dimensions[1].height = 36
    ws.merge_cells('A1:C1')
    _wc(ws['A1'], s_hdr, f'{title} ({len(ids)})')
    for ci, (h, w) in enumerate(zip(['#', 'المعرف', 'الاسم'], [5, 18, 28]), 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
        _wc(ws.cell(2, ci), s_hdr, h)
    ws.row_dimensions[2].height = 26

    if nc:
        names_arr = df.loc[ids, nc].fillna('').astype(str).values
    else:
        names_arr = np.full(len(ids), '', dtype=object)

    for si, (eid, name) in enumerate(zip(ids, names_arr), 1):
        sty = 'ns_dat_only' if si % 2 == 0 else 'ns_dat_gold'
        _wc(ws.cell(si + 2, 1), sty, si)
        _wc(ws.cell(si + 2, 2), sty, eid)
        _wc(ws.cell(si + 2, 3), sty, name)


def _write_duplicates_sheet(ws, title, dup_df, sort_col, display_cols, s_hdr):
    """
    كتابة ورقة المكررات.
    المكررات بنفس المفتاح تُجمع معاً وتُلوَّن بألوان متناوبة بين المجموعات.

    sort_col     : العمود المستخدم للترتيب والتجميع
    display_cols : الأعمدة المعروضة (بدون الأعمدة الداخلية)
    """
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = 'A3'
    ws.row_dimensions[1].height = 42
    ncols = len(display_cols) + 1          # +1 لعمود الترقيم
    ws.merge_cells(f'A1:{get_column_letter(ncols)}1')
    _wc(ws['A1'], s_hdr, f'{title}  —  {len(dup_df)} سجل مكرر')
    ws.row_dimensions[2].height = 28
    headers = ['#'] + display_cols
    widths  = [5] + [max(12, min(len(str(c)) * 2, 35)) for c in display_cols]
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
        _wc(ws.cell(2, ci), s_hdr, h)

    # ترتيب حسب المفتاح لتجميع النسخ المكررة معاً
    sorted_df = dup_df.sort_values(sort_col, na_position='last').reset_index(drop=True)
    vals      = sorted_df[display_cols].fillna('').values
    keys      = sorted_df[sort_col].fillna('').astype(str).values

    # لون متناوب بين كل مجموعة (مجموعة = سجلات بنفس المفتاح)
    grp, prev = 0, None
    for si, (row_vals, k) in enumerate(zip(vals, keys), 1):
        if k != prev:
            grp += 1
            prev = k
        sty = 'ns_dat_gold' if grp % 2 == 1 else 'ns_dat_diff'
        _wc(ws.cell(si + 2, 1), sty, si)
        for ci, val in enumerate(row_vals, 2):
            _wc(ws.cell(si + 2, ci), sty, val)


def clean_cols(df):
    df.columns = [str(c).strip().replace('\n', ' ').replace('  ', ' ')
                  for c in df.columns]
    return df


def _write_extras_sheet(ws, title, indices, src_df, all_cols, col_ws, s_hdr):
    """
    تحسين: استخراج كل البيانات دفعةً واحدة بـ .values
    بدلاً من .iloc لكل صف → ×17 أسرع.
    """
    ws.sheet_view.rightToLeft = True
    ws.freeze_panes = 'A3'
    ws.row_dimensions[1].height = 42
    ncols = len(all_cols)
    if ncols > 1:
        ws.merge_cells(f'A1:{get_column_letter(ncols)}1')
    _wc(ws['A1'], s_hdr, f'{title}  —  عدد السجلات: {len(indices)}')
    ws.row_dimensions[2].height = 28
    for ci, (col, w) in enumerate(zip(all_cols, col_ws), 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
        _wc(ws.cell(2, ci), s_hdr, col)

    # ── استخراج كل القيم دفعةً واحدة ──
    subset_vals = (src_df.iloc[indices][all_cols]
                   .fillna('').values)  # numpy array (M, ncols)

    for si, row_vals in enumerate(subset_vals, 1):
        sty = 'ns_dat_only' if si % 2 == 0 else 'ns_dat_gold'
        for ci, val in enumerate(row_vals, 1):
            _wc(ws.cell(si + 2, ci), sty, val)


# ─────────────────────────────────────────────────────────────
# المقارنة السريعة
# ─────────────────────────────────────────────────────────────

def _load_and_validate_fast(file1, sheet1, header1, key1, name1,
                            file2, sheet2, header2, key2, name2,
                            mappings, prog):
    try:
        prog(10, f"جاري قراءة {name1}...")
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning)
            df1 = pd.read_excel(file1, sheet_name=sheet1, header=header1,
                                 dtype={key1: str}, engine='openpyxl')

        prog(25, f"جاري قراءة {name2}...")
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning)
            df2 = pd.read_excel(file2, sheet_name=sheet2, header=header2,
                                 dtype={key2: str}, engine='openpyxl')

        df1, df2 = clean_cols(df1), clean_cols(df2)

        for col, df, lbl in [(key1, df1, name1), (key2, df2, name2)]:
            if col not in df.columns:
                return LoadResultFast(success=False, error=f"العمود المفتاح '{col}' غير موجود في {lbl}.")
        for m in mappings:
            if m['col1'] not in df1.columns:
                return LoadResultFast(success=False, error=f"الحقل '{m['col1']}' غير موجود في {name1}.")
            if m['col2'] not in df2.columns:
                return LoadResultFast(success=False, error=f"الحقل '{m['col2']}' غير موجود في {name2}.")

        total1 = len(df1)
        total2 = len(df2)

        prog(40, "جاري معالجة البيانات...")
        df1['_id'] = df1[key1].apply(clean_id)
        df2['_id'] = df2[key2].apply(clean_id)

        # ── اكتشاف المكررات قبل حذفها ────────────────────────
        _dup_mask1 = df1['_id'].duplicated(keep=False) & (df1['_id'] != '')
        _dup_mask2 = df2['_id'].duplicated(keep=False) & (df2['_id'] != '')
        dup1_df    = df1[_dup_mask1].copy()
        dup2_df    = df2[_dup_mask2].copy()

        df1 = df1[df1['_id'] != ''].drop_duplicates('_id').set_index('_id')
        df2 = df2[df2['_id'] != ''].drop_duplicates('_id').set_index('_id')

        ids1   = set(df1.index)
        ids2   = set(df2.index)
        common = sorted(ids1 & ids2)
        only1  = sorted(ids1 - ids2)
        only2  = sorted(ids2 - ids1)

        return LoadResultFast(
            success=True, df1=df1, df2=df2, dup1_df=dup1_df, dup2_df=dup2_df,
            common=common, only1=only1, only2=only2, total1=total1, total2=total2
        )
    except Exception as e:
        return LoadResultFast(
            success=False, error=f"حدث خطأ أثناء تحميل البيانات أو التحقق منها:\n{str(e)}"
        )


def _compare_mapped_columns_fast(df1, df2, common, mappings, name_col1, prog):
    prog(52, "جاري المقارنة...")
    cols1 = [m['col1'] for m in mappings]
    cols2 = [m['col2'] for m in mappings]

    # ── تحسين 1: map مع lru_cache يُجنّب إعادة الحساب ──
    sub1 = df1.loc[common, cols1].map(norm)
    sub2 = df2.loc[common, cols2].map(norm)
    sub2.columns = cols1

    # ── تحسين 2: diff_mask بشكل متجهي ───────────────────
    diff_mask     = (sub1 != sub2)
    col_diff_count = {c: int(diff_mask[c].sum()) for c in cols1}

    # ── تحسين 3: الفصل المسبق (mask numpy) ──────────────
    diff_mask_arr = diff_mask.values              # numpy bool (N, ncols)
    has_diff_bool = diff_mask_arr.any(axis=1)     # numpy bool (N,)
    common_arr    = np.array(common)

    diff_eids  = common_arr[has_diff_bool].tolist()
    match_eids = common_arr[~has_diff_bool].tolist()

    # ── تحسين 4: بناء سجلات الاختلاف فقط للصفوف المختلفة ─
    diff_indices    = np.where(has_diff_bool)[0]
    sub1_vals       = sub1.values        # numpy str (N, ncols)
    sub2_vals       = sub2.values
    diff_mask_vals  = diff_mask_arr

    names_s = (df1[name_col1].fillna('').astype(str)
               if name_col1 else pd.Series('', index=df1.index))

    diff_records = []
    for local_i in diff_indices:
        eid  = common_arr[local_i]
        mask = diff_mask_vals[local_i]
        diffs = [
            (cols1[ci], cols2[ci], sub1_vals[local_i, ci], sub2_vals[local_i, ci])
            for ci in range(len(cols1)) if mask[ci]
        ]
        diff_records.append({
            'id': eid, 'name': names_s.get(eid, ''), 'diffs': diffs
        })

    match_count = len(match_eids)
    diff_count  = len(diff_eids)

    return match_count, diff_count, diff_records, col_diff_count, match_eids


def _write_fast_report(output_path, df1, df2, name1, name2, common, only1, only2,
                       dup1_df, dup2_df, match_count, diff_count, diff_records, col_diff_count,
                       match_eids, name_col1, name_col2, N, total1, total2, prog):
    try:
        prog(68, "جاري إنشاء التقرير...")
        import copy
        wb = Workbook()
        for ns in _GLOBAL_STYLES:
            try:
                wb.add_named_style(copy.copy(ns))
            except Exception:
                pass

        # ── ورقة 1: الملخص ──────────────────────────────────
        ws1 = wb.active
        ws1.title = 'الملخص التنفيذي'
        ws1.sheet_view.rightToLeft = True
        for col, w in [('A', 38), ('B', 18), ('C', 18)]:
            ws1.column_dimensions[col].width = w
        ws1.row_dimensions[1].height = 50
        ws1.merge_cells('A1:C1')
        _wc(ws1['A1'], 'ns_hdr', 'تقرير مقارنة البيانات الذكي')

        rows_s = [
            ('', 'القيمة', 'النسبة %'),
            (f'إجمالي أسطر {name1}',  total1,       '—'),
            (f'إجمالي أسطر {name2}',  total2,       '—'),
            ('الأسطر المشتركة',         N,            '—'),
            (f'فقط في {name1}',         len(only1),   f'{len(only1)/max(len(df1.index),1)*100:.1f}%'),
            (f'فقط في {name2}',         len(only2),   f'{len(only2)/max(len(df2.index),1)*100:.1f}%'),
            ('✅  متطابقون تماماً',      match_count,  f'{match_count/max(N,1)*100:.1f}%'),
            ('⚠️  يحتوون اختلافات',     diff_count,   f'{diff_count/max(N,1)*100:.1f}%'),
            (f'🔁  مكررات في {name1}',  len(dup1_df), '—'),
            (f'🔁  مكررات في {name2}',  len(dup2_df), '—'),
        ]
        start_row = 3
        for idx, (lbl, val, pct) in enumerate(rows_s):
            i = start_row + idx
            ws1.row_dimensions[i].height = 26
            if i == 3:
                for ci, v in enumerate([lbl, val, pct], 1):
                    _wc(ws1.cell(i, ci), 'ns_sub', v)
            elif i == 9:
                for ci, v in enumerate([lbl, val, pct], 1):
                    _wc(ws1.cell(i, ci), 'ns_dat_match', v)
            elif i == 10:
                for ci, v in enumerate([lbl, val, pct], 1):
                    _wc(ws1.cell(i, ci), 'ns_dat_diff', v)
            elif i in (11, 12):
                for ci, v in enumerate([lbl, val, pct], 1):
                    _wc(ws1.cell(i, ci),
                        'ns_dat_stripe' if val == 0 else 'ns_dat_diff', v)
            else:
                sty = 'ns_dat_stripe' if i % 2 == 0 else 'ns_dat_white'
                for ci, v in enumerate([lbl, val, pct], 1):
                    _wc(ws1.cell(i, ci), sty, v)

        next_row = start_row + len(rows_s)
        header_row = next_row + 1
        col_headers_row = next_row + 2

        ws1.row_dimensions[header_row].height = 30
        ws1.merge_cells(f'A{header_row}:C{header_row}')
        _wc(ws1[f'A{header_row}'], 'ns_sub', 'تفصيل الاختلافات حسب الأعمدة')
        ws1.row_dimensions[col_headers_row].height = 26
        for ci, lbl in enumerate([f'العمود ({name1})', 'عدد الاختلافات', 'النسبة %'], 1):
            _wc(ws1.cell(col_headers_row, ci), 'ns_slate', lbl)
        r = col_headers_row + 1
        for i, (col, cnt) in enumerate(sorted(col_diff_count.items(), key=lambda x: -x[1])):
            ws1.row_dimensions[r].height = 22
            sty_n = 'ns_dat_diff' if cnt > 0 else 'ns_dat_match'
            sty_r = 'ns_dat_stripe' if i % 2 == 0 else 'ns_dat_white'
            _wc(ws1.cell(r, 1, col),                  sty_r)
            _wc(ws1.cell(r, 2, cnt),                  sty_n)
            _wc(ws1.cell(r, 3, f'{cnt/max(N,1)*100:.1f}%'), sty_n)
            r += 1

        prog(80, "جاري كتابة التفاصيل...")

        # ── ورقة 2: تفاصيل الاختلافات ───────────────────────
        ws2 = wb.create_sheet('تفاصيل الاختلافات')
        ws2.sheet_view.rightToLeft = True
        ws2.freeze_panes = 'A3'
        ws2.row_dimensions[1].height = 40
        ws2.merge_cells('A1:G1')
        _wc(ws2['A1'], 'ns_hdr', '⚠  تفاصيل الاختلافات')
        h2 = ['#', 'المعرف', 'الاسم',
               f'عمود ({name1})', f'عمود ({name2})',
               f'قيمة ({name1})', f'قيمة ({name2})']
        w2 = [5, 18, 25, 22, 22, 20, 20]
        for ci, (h, w) in enumerate(zip(h2, w2), 1):
            ws2.column_dimensions[get_column_letter(ci)].width = w
            _wc(ws2.cell(2, ci), 'ns_sub', h)
        ws2.row_dimensions[2].height = 28
        row2, seq = 3, 0
        for rec in diff_records:
            for j, (c1, c2, v1, v2) in enumerate(rec['diffs']):
                seq += 1
                sty = 'ns_dat_stripe' if seq % 2 == 0 else 'ns_dat_gold'
                vals = [seq, rec['id'], rec['name'] if j == 0 else '',
                        c1, c2, v1, v2]
                for ci, val in enumerate(vals, 1):
                    cell = ws2.cell(row2, ci)
                    cell.value = val
                    cell.style = 'ns_diff_cell' if ci in (6, 7) else sty
                row2 += 1

        # ── ورقة 3: المتطابقون ────────────────────────────────
        ws3 = wb.create_sheet('المتطابقون تماماً')
        ws3.sheet_view.rightToLeft = True
        ws3.freeze_panes = 'A3'
        ws3.row_dimensions[1].height = 40
        ws3.merge_cells('A1:C1')
        _wc(ws3['A1'], 'ns_dark_grn', f'✅  الأسطر المتطابقة ({match_count})')
        for ci, (h, w) in enumerate(zip(['#', 'المعرف', 'الاسم'], [5, 18, 28]), 1):
            ws3.column_dimensions[get_column_letter(ci)].width = w
            _wc(ws3.cell(2, ci), 'ns_dark_grn', h)
        ws3.row_dimensions[2].height = 26

        if match_eids and name_col1:
            match_names = df1.loc[match_eids, name_col1].fillna('').astype(str).values
        else:
            match_names = np.full(len(match_eids), '', dtype=object)

        for si, (eid, name) in enumerate(zip(match_eids, match_names), 1):
            sty = 'ns_dat_match' if si % 2 == 0 else 'ns_dat_white'
            _wc(ws3.cell(si + 2, 1), sty, si)
            _wc(ws3.cell(si + 2, 2), sty, eid)
            _wc(ws3.cell(si + 2, 3), sty, name)

        if only1:
            ws4 = wb.create_sheet(f'في {name1[:23]} فقط')
            _write_only_sheet(ws4, f'فقط في {name1}', only1, df1, name_col1, 'ns_brown2')
        if only2:
            ws5 = wb.create_sheet(f'في {name2[:23]} فقط')
            _write_only_sheet(ws5, f'فقط في {name2}', only2, df2, name_col2, 'ns_purple2')

        # ── ورقات المكررات ────────────────────────────────────
        _all_cols1 = [c for c in dup1_df.columns if not c.startswith('_')]
        _all_cols2 = [c for c in dup2_df.columns if not c.startswith('_')]
        if not dup1_df.empty:
            ws_d1 = wb.create_sheet(f'مكررات {name1[:20]}')
            _write_duplicates_sheet(
                ws_d1, f'🔁  سجلات مكررة في {name1}',
                dup1_df, '_id', _all_cols1, 'ns_brown2'
            )
        if not dup2_df.empty:
            ws_d2 = wb.create_sheet(f'مكررات {name2[:20]}')
            _write_duplicates_sheet(
                ws_d2, f'🔁  سجلات مكررة في {name2}',
                dup2_df, '_id', _all_cols2, 'ns_purple2'
            )

        prog(96, "جاري حفظ التقرير...")
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        prog(100, "تمت العملية بنجاح!")
        return True, "تمت المطابقة وتوليد التقرير بنجاح."

    except Exception as e:
        return False, f"حدث خطأ أثناء كتابة التقرير وتنسيقه:\n{str(e)}"


def run_comparison(file1, sheet1, header1, key1, name1,
                   file2, sheet2, header2, key2, name2,
                   mappings, output_path,
                   progress_callback=None):
    try:
        def prog(p, msg=''):
            if progress_callback:
                progress_callback(p, msg)

        # 1. قراءة البيانات والتحقق من صحتها
        load_res = _load_and_validate_fast(
            file1, sheet1, header1, key1, name1,
            file2, sheet2, header2, key2, name2,
            mappings, prog
        )
        if not load_res.success:
            return False, load_res.error, {}
        df1 = load_res.df1
        df2 = load_res.df2
        dup1_df = load_res.dup1_df
        dup2_df = load_res.dup2_df
        common = load_res.common
        only1 = load_res.only1
        only2 = load_res.only2
        total1 = load_res.total1
        total2 = load_res.total2

        N = len(common)
        if N == 0:
            return False, "لا توجد قيود مشتركة بين الملفين بناءً على العمود المفتاحي.", {}

        name_col1 = next((c for c in df1.columns
                          if any(k in c for k in ['اسم', 'لقب', 'name'])), None)
        name_col2 = next((c for c in df2.columns
                          if any(k in c for k in ['اسم', 'لقب', 'name'])), None)

        # 2. مقارنة قيم الأعمدة المحددة
        match_count, diff_count, diff_records, col_diff_count, match_eids = _compare_mapped_columns_fast(
            df1, df2, common, mappings, name_col1, prog
        )

        # 3. توليد وكتابة التقرير النهائي
        success_write, msg_write = _write_fast_report(
            output_path, df1, df2, name1, name2, common, only1, only2,
            dup1_df, dup2_df, match_count, diff_count, diff_records, col_diff_count,
            match_eids, name_col1, name_col2, N, total1, total2, prog
        )
        if not success_write:
            return False, msg_write, {}

        stats = {
            'common': N, 'match': match_count, 'diff': diff_count,
            'only1': len(only1), 'only2': len(only2),
            'total1': total1,    'total2': total2,
            'dup1': len(dup1_df), 'dup2': len(dup2_df),
        }
        return True, "تمت المطابقة وتوليد التقرير بنجاح.", stats

    except Exception as e:
        traceback.print_exc()
        return False, f"حدث خطأ في المقارنة السريعة:\n{str(e)}", {}


# ─────────────────────────────────────────────────────────────
# المقارنة العميقة — مُحسَّنة (مُقسَّمة إلى دوال فرعية)
# ─────────────────────────────────────────────────────────────

def _load_and_validate(file1, sheet1, header1, match_col1, anchor_col1, label1,
                      file2, sheet2, header2, match_col2, anchor_col2, label2,
                      mappings, prog):
    try:
        prog(8, f"جاري قراءة {label1}...")
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning)
            df1 = pd.read_excel(file1, sheet_name=sheet1, header=header1, engine='openpyxl')

        prog(20, f"جاري قراءة {label2}...")
        with warnings.catch_warnings():
            warnings.filterwarnings('ignore', category=UserWarning)
            df2 = pd.read_excel(file2, sheet_name=sheet2, header=header2, engine='openpyxl')

        df1 = clean_cols(df1)
        df2 = clean_cols(df2)

        # ── التحقق من الأعمدة ─────────────────────────────────
        for col, df, lbl in [(match_col1, df1, label1), (match_col2, df2, label2)]:
            if col not in df.columns:
                return LoadResultDeep(success=False, error=f"عمود المطابقة '{col}' غير موجود في {lbl}.")
        if anchor_col1 and anchor_col1 not in df1.columns:
            return LoadResultDeep(success=False, error=f"عمود التدقيق '{anchor_col1}' غير موجود في {label1}.")
        if anchor_col2 and anchor_col2 not in df2.columns:
            return LoadResultDeep(success=False, error=f"عمود التدقيق '{anchor_col2}' غير موجود في {label2}.")
        for m in mappings:
            if m['col1'] not in df1.columns:
                return LoadResultDeep(success=False, error=f"الحقل '{m['col1']}' غير موجود في {label1}.")
            if m['col2'] not in df2.columns:
                return LoadResultDeep(success=False, error=f"الحقل '{m['col2']}' غير موجود في {label2}.")

        n1, n2 = len(df1), len(df2)

        # ── تحسين: التطبيع مع lru_cache ─────────────────────
        prog(30, "جاري تطبيع النصوص (lru_cache مُفعَّل)...")

        tokens1  = df1[match_col1].apply(_sort_tokens).tolist()
        tokens2  = df2[match_col2].apply(_sort_tokens).tolist()
        anchors1 = (df1[anchor_col1].apply(_norm_anchor).tolist()
                    if anchor_col1 else [''] * n1)
        anchors2 = (df2[anchor_col2].apply(_norm_anchor).tolist()
                    if anchor_col2 else [''] * n2)

        # ── اكتشاف المكررات داخل كل جدول ────────────────────
        _dk1 = pd.Series([f"{tokens1[i]}|{anchors1[i]}" for i in range(n1)])
        _dk2 = pd.Series([f"{tokens2[i]}|{anchors2[i]}" for i in range(n2)])
        _dup1_mask = _dk1.duplicated(keep=False).values
        _dup2_mask = _dk2.duplicated(keep=False).values
        dup1_df = df1.iloc[np.where(_dup1_mask)[0]].copy()
        dup2_df = df2.iloc[np.where(_dup2_mask)[0]].copy()
        dup1_df['__dup_key__'] = _dk1[_dup1_mask].values
        dup2_df['__dup_key__'] = _dk2[_dup2_mask].values

        return LoadResultDeep(
            success=True, df1=df1, df2=df2, n1=n1, n2=n2,
            tokens1=tokens1, tokens2=tokens2, anchors1=anchors1, anchors2=anchors2,
            dup1_df=dup1_df, dup2_df=dup2_df
        )
    except Exception as e:
        return LoadResultDeep(
            success=False, error=f"حدث خطأ أثناء تحميل البيانات أو التحقق منها:\n{str(e)}"
        )


def _phase1_high_confidence(tokens1, tokens2, anchors1, anchors2, n1, n2, phase1_thresh):
    anchor_to_j = defaultdict(list)
    for j, anc in enumerate(anchors1):
        if anc:
            anchor_to_j[anc].append(j)

    used1 = [False] * n1
    used2 = [False] * n2
    matched_pairs = []

    high_conf = []
    for i2, (tok2, anc2) in enumerate(zip(tokens2, anchors2)):
        if not anc2:
            continue
        for j1 in anchor_to_j.get(anc2, []):
            score = fuzz.token_sort_ratio(tok2, tokens1[j1])
            if score >= phase1_thresh:
                high_conf.append((score, i2, j1))

    high_conf.sort(reverse=True)
    for score, i2, j1 in high_conf:
        if used2[i2] or used1[j1]:
            continue
        used2[i2] = True
        used1[j1] = True
        matched_pairs.append((i2, j1))

    return used1, used2, matched_pairs


def _phase2_hungarian_or_topk(tokens1, tokens2, anchors1, anchors2, used1, used2, matched_pairs,
                              phase2_with_anchor, phase2_no_anchor, prog):

    n1, n2 = len(tokens1), len(tokens2)
    remaining2 = [i for i in range(n2) if not used2[i]]
    remaining1 = [j for j in range(n1) if not used1[j]]
    nr2, nr1   = len(remaining2), len(remaining1)

    if not remaining2 or not remaining1:
        return used1, used2, matched_pairs

    matrix_cells = nr2 * nr1

    if matrix_cells <= _MAX_MATRIX_CELLS:
        prog(55, f"المرحلة 2 (cdist + Hungarian) — {nr2}×{nr1} = {matrix_cells:,} خلية...")

        toks_r2 = [tokens2[i] for i in remaining2]
        toks_r1 = [tokens1[j] for j in remaining1]

        name_scores = rf_cdist(
            toks_r2, toks_r1,
            scorer=fuzz.token_sort_ratio,
            dtype=np.float32,
            workers=-1,
        )

        anc_r2 = np.array([anchors2[i] for i in remaining2], dtype=object)
        anc_r1 = np.array([anchors1[j] for j in remaining1], dtype=object)
        has_anc     = (anc_r2 != '').reshape(-1, 1)
        same_anchor = (anc_r2.reshape(-1, 1) == anc_r1.reshape(1, -1))
        score_mat   = name_scores + (same_anchor & has_anc).astype(np.float32) * 20

        row_ind, col_ind = linear_sum_assignment(-score_mat)

        for ii, jj in zip(row_ind, col_ind):
            i2 = remaining2[ii]
            j1 = remaining1[jj]
            name_score      = float(name_scores[ii, jj])
            has_anc_match   = bool(anchors2[i2]) and (anchors2[i2] == anchors1[j1])
            accept_thresh   = phase2_with_anchor if has_anc_match else phase2_no_anchor

            if name_score >= accept_thresh:
                if not used2[i2] and not used1[j1]:
                    used2[i2] = True
                    used1[j1] = True
                    matched_pairs.append((i2, j1))

    else:
        prog(55, f"المرحلة 2 (Top-{_TOP_K} Greedy) — {nr2}×{nr1} = {matrix_cells:,} خلية...")

        toks_r1   = [tokens1[j] for j in remaining1]
        used1_loc = [False] * nr1
        candidates = []

        for ii, i2 in enumerate(remaining2):
            results = rf_extract(
                tokens2[i2], toks_r1,
                scorer=fuzz.token_sort_ratio,
                limit=_TOP_K,
            )
            for _tok, ns, jj in results:
                j1 = remaining1[jj]
                has_anc = (anchors2[i2] and anchors2[i2] == anchors1[j1])
                candidates.append((ns + (20 if has_anc else 0), ns, ii, jj))

        candidates.sort(reverse=True)

        for total_score, ns, ii, jj in candidates:
            i2 = remaining2[ii]
            j1 = remaining1[jj]
            if used2[i2] or used1[j1] or used1_loc[jj]:
                continue
            has_anc    = (anchors2[i2] and anchors2[i2] == anchors1[j1])
            accept_thr = phase2_with_anchor if has_anc else phase2_no_anchor
            if ns >= accept_thr:
                used2[i2]     = True
                used1[j1]     = True
                used1_loc[jj] = True
                matched_pairs.append((i2, j1))

    return used1, used2, matched_pairs


def _compute_matches(tokens1, tokens2, anchors1, anchors2, n1, n2,
                     phase1_thresh, phase2_with_anchor, phase2_no_anchor, prog):
    prog(40, "المرحلة 1: المطابقة العالية الثقة...")
    used1, used2, matched_pairs = _phase1_high_confidence(
        tokens1, tokens2, anchors1, anchors2, n1, n2, phase1_thresh
    )

    used1, used2, matched_pairs = _phase2_hungarian_or_topk(
        tokens1, tokens2, anchors1, anchors2, used1, used2, matched_pairs,
        phase2_with_anchor, phase2_no_anchor, prog
    )

    only2_idx = [i for i in range(n2) if not used2[i]]
    only1_idx = [j for j in range(n1) if not used1[j]]

    return matched_pairs, only1_idx, only2_idx


def _compare_mapped_columns(df1, df2, matched_pairs, mappings, match_col2, anchor_col2, prog):
    prog(68, "جاري مقارنة الحقول...")

    cols1 = [m['col1'] for m in mappings]
    cols2 = [m['col2'] for m in mappings]
    col_diff_count = {m['col1']: 0 for m in mappings}

    # ── استخراج القيم مرة واحدة بـ numpy ─────────
    if matched_pairs and mappings:
        i2_list = [p[0] for p in matched_pairs]
        j1_list = [p[1] for p in matched_pairs]

        sub2_m = df2.iloc[i2_list][cols2].map(norm).values
        sub1_m = df1.iloc[j1_list][cols1].map(norm).values
        diff_m = (sub1_m != sub2_m)
    else:
        diff_m = np.empty((0, len(cols1)))
        sub2_m = sub1_m = np.empty((0, len(cols1)))

    match_count  = 0
    diff_count   = 0
    diff_records = []

    if matched_pairs:
        i2_list = [p[0] for p in matched_pairs]
        display_names = df2.iloc[i2_list][match_col2].fillna('').astype(str).tolist()
        anchor_disps  = (df2.iloc[i2_list][anchor_col2].fillna('').astype(str).tolist()
                         if anchor_col2 else [''] * len(matched_pairs))
    else:
        display_names = []
        anchor_disps = []

    for k, (i2, j1) in enumerate(matched_pairs):
        display_name = display_names[k]
        anchor_disp  = anchor_disps[k]

        if mappings:
            mask_row = diff_m[k]
            diffs = [
                (cols1[ci], cols2[ci], sub1_m[k, ci], sub2_m[k, ci])
                for ci in range(len(cols1)) if mask_row[ci]
            ]
            for ci in range(len(cols1)):
                if mask_row[ci]:
                    col_diff_count[cols1[ci]] += 1
        else:
            diffs = []

        if diffs:
            diff_count += 1
        else:
            match_count += 1

        diff_records.append({
            'name': display_name, 'anchor': anchor_disp, 'diffs': diffs,
        })

    return match_count, diff_count, diff_records, col_diff_count


def _write_deep_report(output_path, df1, df2, n1, n2, label1, label2, mappings,
                      min_similarity, phase2_no_anchor, phase2_with_anchor,
                      anchor_col1, anchor_col2, only1_idx, only2_idx,
                      match_count, diff_count, diff_records, col_diff_count,
                      dup1_df, dup2_df, N, prog):
    try:
        import copy
        wb = Workbook()
        for ns in _GLOBAL_STYLES:
            try:
                wb.add_named_style(copy.copy(ns))
            except Exception:
                pass

        anchor_lbl = (f'عمود التدقيق ({anchor_col1 or anchor_col2 or "—"})')

        # ── ورقة 1: الملخص ───────────────────────────────────
        ws1 = wb.active
        ws1.title = 'الملخص التنفيذي'
        ws1.sheet_view.rightToLeft = True
        for col, w in [('A', 42), ('B', 18), ('C', 18)]:
            ws1.column_dimensions[col].width = w
        ws1.row_dimensions[1].height = 55
        ws1.merge_cells('A1:C1')
        _wc(ws1['A1'], 'ns_deep_hdr', '🔍  تقرير المقارنة العميقة الذكية')

        ws1.row_dimensions[2].height = 22
        ws1.merge_cells('A2:C2')
        anc_note = f" | عمود التدقيق: موجود ← عتبة أخف ({phase2_with_anchor}%)" if (anchor_col1 or anchor_col2) else ""
        _wc(ws1['A2'], 'ns_dat_stripe',
            f"⚙️  إعدادات المطابقة — أدنى تشابه: {min_similarity}%  (بدون تدقيق: {phase2_no_anchor}%{anc_note})")

        rows_s = [
            ('',                              'القيمة',       'النسبة %'),
            (f'إجمالي سجلات {label1}',        n1,             '—'),
            (f'إجمالي سجلات {label2}',        n2,             '—'),
            ('السجلات المتطابقة (مُقارَنة)',   N,              '—'),
            (f'زائد في {label1} فقط',          len(only1_idx), f'{len(only1_idx)/max(n1,1)*100:.1f}%'),
            (f'زائد في {label2} فقط',          len(only2_idx), f'{len(only2_idx)/max(n2,1)*100:.1f}%'),
            ('✅  متطابقون في جميع الحقول',    match_count,    f'{match_count/max(N,1)*100:.1f}%'),
            ('⚠️  بهم اختلافات في البيانات',   diff_count,     f'{diff_count/max(N,1)*100:.1f}%'),
            (f'🔁  مكررات في {label1}',         len(dup1_df),   '—'),
            (f'🔁  مكررات في {label2}',         len(dup2_df),   '—'),
        ]
        start_row = 4
        for idx, (lbl, val, pct) in enumerate(rows_s):
            i = start_row + idx
            ws1.row_dimensions[i].height = 26
            if i == 4:
                for ci, v in enumerate([lbl, val, pct], 1):
                    _wc(ws1.cell(i, ci), 'ns_sub', v)
            elif i == 10:
                for ci, v in enumerate([lbl, val, pct], 1):
                    _wc(ws1.cell(i, ci), 'ns_dat_match', v)
            elif i == 11:
                for ci, v in enumerate([lbl, val, pct], 1):
                    _wc(ws1.cell(i, ci), 'ns_dat_diff', v)
            elif i in (12, 13):
                for ci, v in enumerate([lbl, val, pct], 1):
                    _wc(ws1.cell(i, ci),
                        'ns_dat_stripe' if val == 0 else 'ns_dat_diff', v)
            else:
                sty = 'ns_dat_stripe' if i % 2 == 0 else 'ns_dat_white'
                for ci, v in enumerate([lbl, val, pct], 1):
                    _wc(ws1.cell(i, ci), sty, v)

        if mappings and N > 0:
            next_row = start_row + len(rows_s)
            header_row = next_row + 1
            col_headers_row = next_row + 2

            ws1.row_dimensions[header_row].height = 30
            ws1.merge_cells(f'A{header_row}:C{header_row}')
            _wc(ws1[f'A{header_row}'], 'ns_sub', 'تفصيل الاختلافات حسب الأعمدة')
            ws1.row_dimensions[col_headers_row].height = 26
            for ci, lbl in enumerate([f'العمود ({label1})', 'عدد الاختلافات', 'النسبة %'], 1):
                _wc(ws1.cell(col_headers_row, ci), 'ns_slate', lbl)
            r = col_headers_row + 1
            for i, (col, cnt) in enumerate(
                    sorted(col_diff_count.items(), key=lambda x: -x[1])):
                ws1.row_dimensions[r].height = 22
                sty_n = 'ns_dat_diff' if cnt > 0 else 'ns_dat_match'
                sty_r = 'ns_dat_stripe' if i % 2 == 0 else 'ns_dat_white'
                _wc(ws1.cell(r, 1, col),                  sty_r)
                _wc(ws1.cell(r, 2, cnt),                  sty_n)
                _wc(ws1.cell(r, 3, f'{cnt/N*100:.1f}%'),  sty_n)
                r += 1

        # ── ورقة 2 & 3: السجلات الزائدة ─────────────────────
        prog(85, "جاري كتابة السجلات الزائدة...")

        all_cols2 = [c for c in df2.columns if not c.startswith('_')]
        all_cols1 = [c for c in df1.columns if not c.startswith('_')]
        cw2 = [max(12, min(len(str(c)) * 2, 35)) for c in all_cols2]
        cw1 = [max(12, min(len(str(c)) * 2, 35)) for c in all_cols1]

        if only2_idx:
            ws_e2 = wb.create_sheet(f'زائد في {label2[:22]}')
            _write_extras_sheet(
                ws_e2, f'سجلات زائدة في {label2}',
                only2_idx, df2, all_cols2, cw2, 'ns_purple2'
            )
        if only1_idx:
            ws_e1 = wb.create_sheet(f'زائد في {label1[:22]}')
            _write_extras_sheet(
                ws_e1, f'سجلات زائدة في {label1}',
                only1_idx, df1, all_cols1, cw1, 'ns_brown2'
            )

        # ── ورقة 4: تفاصيل الاختلافات ────────────────────────
        if mappings and diff_count > 0:
            ws_diff = wb.create_sheet('تفاصيل الاختلافات')
            ws_diff.sheet_view.rightToLeft = True
            ws_diff.freeze_panes = 'A3'
            ws_diff.row_dimensions[1].height = 40
            ws_diff.merge_cells('A1:G1')
            _wc(ws_diff['A1'], 'ns_hdr', '⚠  تفاصيل اختلافات السجلات المتطابقة')
            hdrs = ['#', f'نص المطابقة ({label2})', anchor_lbl,
                    f'عمود ({label1})', f'عمود ({label2})',
                    f'قيمة ({label1})', f'قيمة ({label2})']
            wds  = [5, 30, 18, 22, 22, 20, 20]
            for ci, (h, w) in enumerate(zip(hdrs, wds), 1):
                ws_diff.column_dimensions[get_column_letter(ci)].width = w
                _wc(ws_diff.cell(2, ci), 'ns_sub', h)
            ws_diff.row_dimensions[2].height = 28
            rw, seq = 3, 0
            for rec in diff_records:
                if not rec['diffs']:
                    continue
                for j, (c1, c2, v1, v2) in enumerate(rec['diffs']):
                    seq += 1
                    sty = 'ns_dat_stripe' if seq % 2 == 0 else 'ns_dat_gold'
                    vals = [seq,
                            rec['name']   if j == 0 else '',
                            rec['anchor'] if j == 0 else '',
                            c1, c2, v1, v2]
                    for ci, val in enumerate(vals, 1):
                        cell = ws_diff.cell(rw, ci)
                        cell.value = val
                        cell.style = 'ns_diff_cell' if ci in (6, 7) else sty
                    rw += 1

        # ── ورقة 5: المتطابقون تماماً ─────────────────────────
        if match_count > 0:
            ws_ok = wb.create_sheet('المتطابقون تماماً')
            ws_ok.sheet_view.rightToLeft = True
            ws_ok.freeze_panes = 'A3'
            ws_ok.row_dimensions[1].height = 40
            ws_ok.merge_cells('A1:C1')
            _wc(ws_ok['A1'], 'ns_dark_grn',
                f'✅  سجلات متطابقة في جميع الحقول ({match_count})')
            for ci, (h, w) in enumerate(
                    zip(['#', f'نص المطابقة ({label2})', anchor_lbl], [5, 35, 20]), 1):
                ws_ok.column_dimensions[get_column_letter(ci)].width = w
                _wc(ws_ok.cell(2, ci), 'ns_dark_grn', h)
            ws_ok.row_dimensions[2].height = 26
            seq_ok = 0
            for rec in diff_records:
                if rec['diffs']:
                    continue
                seq_ok += 1
                sty = 'ns_dat_match' if seq_ok % 2 == 0 else 'ns_dat_white'
                _wc(ws_ok.cell(seq_ok + 2, 1), sty, seq_ok)
                _wc(ws_ok.cell(seq_ok + 2, 2), sty, rec['name'])
                _wc(ws_ok.cell(seq_ok + 2, 3), sty, rec['anchor'])

        # ── ورقات المكررات ────────────────────────────────────
        _dcols1 = [c for c in dup1_df.columns if not c.startswith('__')]
        _dcols2 = [c for c in dup2_df.columns if not c.startswith('__')]
        if not dup1_df.empty:
            ws_dup1 = wb.create_sheet(f'مكررات {label1[:20]}')
            _write_duplicates_sheet(
                ws_dup1, f'🔁  سجلات مكررة في {label1}',
                dup1_df, '__dup_key__', _dcols1, 'ns_brown2'
            )
        if not dup2_df.empty:
            ws_dup2 = wb.create_sheet(f'مكررات {label2[:20]}')
            _write_duplicates_sheet(
                ws_dup2, f'🔁  سجلات مكررة في {label2}',
                dup2_df, '__dup_key__', _dcols2, 'ns_purple2'
            )

        # ── حفظ ──────────────────────────────────────────────
        prog(96, "جاري حفظ التقرير...")
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        wb.save(output_path)
        prog(100, "تمت العملية بنجاح!")
        return True, "تم حفظ التقرير بنجاح."

    except Exception as e:
        return False, f"حدث خطأ أثناء كتابة التقرير وتنسيقه:\n{str(e)}"


def run_deep_comparison(file1, sheet1, header1, match_col1, anchor_col1, label1,
                        file2, sheet2, header2, match_col2, anchor_col2, label2,
                        mappings, output_path,
                        min_similarity: int = 75,
                        progress_callback=None):
    """
    المقارنة العميقة — تجد التطابق بين أي جدولين بتشابه النصوص.

    min_similarity : أدنى نسبة تشابه مقبولة للمطابقة (50-100، الافتراضي 75).
                     - بدون عمود تدقيق : يجب أن يبلغ تشابه النص min_similarity على الأقل.
                     - مع تطابق عمود التدقيق : يُخفَّف الشرط بـ 15 نقطة
                       (أي min_similarity - 15)، لأن التدقيق يعزّز الثقة.
    """
    if not _HAS_DEEP_LIBS:
        return False, (
            "مكتبة مطلوبة غير مثبتة: rapidfuzz أو scipy.\n"
            "قم بتشغيل:  pip install rapidfuzz scipy"
        ), {}
    try:
        def prog(p, msg=''):
            if progress_callback:
                progress_callback(p, msg)

        # 1. قراءة البيانات والتحقق من صحتها
        load_res = _load_and_validate(
            file1, sheet1, header1, match_col1, anchor_col1, label1,
            file2, sheet2, header2, match_col2, anchor_col2, label2,
            mappings, prog
        )
        if not load_res.success:
            return False, load_res.error, {}
        df1 = load_res.df1
        df2 = load_res.df2
        n1 = load_res.n1
        n2 = load_res.n2
        tokens1 = load_res.tokens1
        tokens2 = load_res.tokens2
        anchors1 = load_res.anchors1
        anchors2 = load_res.anchors2
        dup1_df = load_res.dup1_df
        dup2_df = load_res.dup2_df

        # حساب العتبات المخصصة
        min_similarity = max(50, min(100, int(min_similarity)))
        phase1_thresh = min(100, max(min_similarity + 7, 82))
        phase2_with_anchor = max(min_similarity - 15, 50)
        phase2_no_anchor = min_similarity

        # 2. حساب المطابقة بين السجلات
        matched_pairs, only1_idx, only2_idx = _compute_matches(
            tokens1, tokens2, anchors1, anchors2, n1, n2,
            phase1_thresh, phase2_with_anchor, phase2_no_anchor, prog
        )
        N = len(matched_pairs)

        # 3. مقارنة قيم الأعمدة المحددة
        match_count, diff_count, diff_records, col_diff_count = _compare_mapped_columns(
            df1, df2, matched_pairs, mappings, match_col2, anchor_col2, prog
        )

        # 4. توليد وكتابة التقرير النهائي
        success_write = _write_deep_report(
            output_path, df1, df2, n1, n2, label1, label2, mappings,
            min_similarity, phase2_no_anchor, phase2_with_anchor,
            anchor_col1, anchor_col2, only1_idx, only2_idx,
            match_count, diff_count, diff_records, col_diff_count,
            dup1_df, dup2_df, N, prog
        )
        if not success_write[0]:
            return False, success_write[1], {}

        stats = {
            'common': N, 'match': match_count, 'diff': diff_count,
            'only1': len(only1_idx), 'only2': len(only2_idx),
            'total1': n1, 'total2': n2,
            'dup1': len(dup1_df), 'dup2': len(dup2_df),
        }
        return True, "تمت المقارنة العميقة وتوليد التقرير بنجاح.", stats

    except Exception as e:
        traceback.print_exc()
        return False, f"حدث خطأ في المقارنة العميقة:\n{str(e)}", {}
