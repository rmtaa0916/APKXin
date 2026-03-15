# =========================
# PART 1 / 4
# Imports + Compatibility + Helpers + Engine Base
# =========================

import os
import io
import json
import math
import base64
import zipfile
import traceback
from functools import partial

# -------------------------------------------------------------------
# Python 3.10+ compatibility patch for older reportlab builds
# Fixes:
# ImportError: cannot import name 'decodestring' from 'base64'
# -------------------------------------------------------------------
if not hasattr(base64, "decodestring"):
    base64.decodestring = base64.decodebytes
if not hasattr(base64, "encodestring"):
    base64.encodestring = base64.encodebytes

import cv2
import fitz
import numpy as np
import pandas as pd

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas

from kivy.app import App
from kivy.clock import Clock
from kivy.core.image import Image as CoreImage
from kivy.graphics import Color, Line, Rectangle
from kivy.metrics import dp
from kivy.properties import StringProperty, NumericProperty, BooleanProperty, ListProperty, ObjectProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.image import Image
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.slider import Slider
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.utils import platform


# ============================================================
# Defaults
# ============================================================
APP_TITLE = "MediMap Pro: Intelligent Form Automator"
CONFIG_FILENAME = "medimap_config.json"
ZOOM = 4.0
PREVIEW_SCALE = 1.25

DEFAULTS = {
    "F_Area": 500,
    "F_MinW": 15,
    "F_MinH": 35,
    "F_Close": 1,
    "Line_MinW": 100,
    "Line_MaxW": 1680,
    "C_Strict": 40,
    "C_Size": (14, 65),
    "C_Border": 0.10,
    "C_Inner": 0.40,
    "ROI_Max": 200,
    "C_Open": 1,
    "C_Close": 0,
    "C_BandPct": 0.18,
    "C_AspectTol": 0.10,
    "Ext_Low": 0.04,
    "Ext_High": 0.60,
    "C_FillMin": 0.45,
    "C_Eps": 0.04,
    "Use_Extent": False,
    "Is_Grid": False,
    "Grid_N": 1,
}


# ============================================================
# Small helpers
# ============================================================
def safe_name(text):
    text = str(text or "").strip()
    text = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in text)
    return text.strip() or "Unknown"

def _norm_str(v):
    return str(v or "").strip()

def _rect_area(r):
    return max(0.0, r.width) * max(0.0, r.height)

def _rect_intersection_area(a, b):
    ix0 = max(a.x0, b.x0)
    iy0 = max(a.y0, b.y0)
    ix1 = min(a.x1, b.x1)
    iy1 = min(a.y1, b.y1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    return (ix1 - ix0) * (iy1 - iy0)

def _rect_iou(a, b):
    inter = _rect_intersection_area(a, b)
    if inter <= 0:
        return 0.0
    union = (a.width * a.height) + (b.width * b.height) - inter
    return inter / union if union > 0 else 0.0

def _x_intersection(a, b):
    return max(0.0, min(a.x1, b.x1) - max(a.x0, b.x0))

def _x_overlap_ratio(a, b):
    inter_x = _x_intersection(a, b)
    return inter_x / max(min(a.width, b.width), 1e-6)

def _rect_union(rects):
    return fitz.Rect(
        min(r.x0 for r in rects),
        min(r.y0 for r in rects),
        max(r.x1 for r in rects),
        max(r.y1 for r in rects)
    )

def _rect_close(a, b, tol=0.20):
    return (
        abs(a.x0 - b.x0) <= tol and
        abs(a.y0 - b.y0) <= tol and
        abs(a.x1 - b.x1) <= tol and
        abs(a.y1 - b.y1) <= tol
    )

def _rect_list_close(rects_a, rects_b, tol=0.20):
    if len(rects_a) != len(rects_b):
        return False
    aa = sorted(rects_a, key=lambda r: (round(r.y0, 3), r.x0, r.x1, r.y1))
    bb = sorted(rects_b, key=lambda r: (round(r.y0, 3), r.x0, r.x1, r.y1))
    return all(_rect_close(a, b, tol=tol) for a, b in zip(aa, bb))

def find_col_safe(df, *keywords):
    keys = [k.lower() for k in keywords]
    for c in df.columns:
        if all(k in str(c).lower() for k in keys):
            return c
    return None


# ============================================================
# Detection / mapping engine
# ============================================================
class MediMapEngine:
    def __init__(self):
        self.pdf_path = ""
        self.df = pd.DataFrame()
        self.patient_names = []

        self.first_col = None
        self.last_col = None
        self.mid_col = None
        self.suf_col = None
        self.dob_col = None
        self.phil_col = None

        self.geom = {"names": [], "dob": [], "phil": []}
        self.all_boxes = []
        self.box_types = []
        self.custom_mappings = {}
        self.selected_box_ids = []

        self.settings = dict(DEFAULTS)
        self.settings["C_Size"] = list(DEFAULTS["C_Size"])

    # --------------------------------------------------------
    # Data loading
    # --------------------------------------------------------
    def load_dataframe(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            df = pd.read_csv(path, dtype=str).fillna("")
        else:
            df = pd.read_excel(path, dtype=str).fillna("")

        self.df = df
        self.first_col = find_col_safe(df, "first", "name")
        self.last_col = find_col_safe(df, "surname") or find_col_safe(df, "last", "name")
        self.mid_col = find_col_safe(df, "middle", "name")
        self.suf_col = find_col_safe(df, "suffix")
        self.dob_col = find_col_safe(df, "birth", "date") or find_col_safe(df, "birthdate")
        self.phil_col = find_col_safe(df, "philhealth")

        f_col_safe = self.first_col if self.first_col else df.columns[0]
        l_col_safe = self.last_col if self.last_col else df.columns[0]

        self.df["_DISPLAY_NAME"] = (
            df[f_col_safe].fillna("").astype(str) + " " +
            df[l_col_safe].fillna("").astype(str)
        ).str.strip()

        self.patient_names = sorted([
            str(x).strip()
            for x in self.df["_DISPLAY_NAME"].dropna().tolist()
            if str(x).strip()
        ])

    def total_pages(self):
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            return 1
        try:
            with fitz.open(self.pdf_path) as doc:
                return max(len(doc), 1)
        except Exception:
            return 1

    # --------------------------------------------------------
    # Box append helpers
    # --------------------------------------------------------
    def _append_box_unique(self, rect, box_type="check", iou_thresh=0.55):
        for j, existing in enumerate(self.all_boxes):
            if self.box_types[j] != box_type:
                continue
            if _rect_iou(rect, existing) >= iou_thresh:
                return
        self.all_boxes.append(rect)
        self.box_types.append(box_type)

    def _append_geom_fields(self):
        for key in ["names", "dob", "phil"]:
            for r in self.geom.get(key, []):
                self._append_box_unique(r, "field", iou_thresh=0.60)

    # --------------------------------------------------------
    # Cleanup helpers
    # --------------------------------------------------------
    def _cleanup_field_fragments(self):
        if not self.all_boxes:
            return

        keep = [True] * len(self.all_boxes)
        field_idxs = [i for i, t in enumerate(self.box_types) if t == "field"]

        for si in field_idxs:
            if not keep[si]:
                continue

            sr = self.all_boxes[si]
            s_area = _rect_area(sr)
            if s_area <= 0:
                keep[si] = False
                continue

            for bi in field_idxs:
                if si == bi or not keep[bi]:
                    continue

                br = self.all_boxes[bi]
                if br.height < sr.height or _rect_area(br) <= s_area:
                    continue

                x_overlap = _x_overlap_ratio(sr, br)
                gap = br.y0 - sr.y1

                if (
                    x_overlap >= 0.80 and
                    0 <= gap <= 6.0 and
                    sr.height <= max(4.0, br.height * 0.45)
                ):
                    keep[si] = False
                    break

                inter = _rect_intersection_area(sr, br)
                if inter > 0:
                    cover_small = inter / max(s_area, 1e-6)
                    if cover_small >= 0.75 and sr.y1 <= (br.y0 + br.height * 0.45):
                        keep[si] = False
                        break

        self.all_boxes = [r for k, r in zip(keep, self.all_boxes) if k]
        self.box_types = [t for k, t in zip(keep, self.box_types) if k]

    def _cleanup_line_field_conflicts(self):
        if not self.all_boxes:
            return

        keep = [True] * len(self.all_boxes)
        field_idxs = [i for i, t in enumerate(self.box_types) if t == "field"]
        line_idxs = [i for i, t in enumerate(self.box_types) if t == "line"]

        for li in line_idxs:
            if not keep[li]:
                continue

            lr = self.all_boxes[li]
            la = _rect_area(lr)
            if la <= 0:
                keep[li] = False
                continue

            lmid_y = (lr.y0 + lr.y1) / 2.0

            for fi in field_idxs:
                fr = self.all_boxes[fi]

                inter = _rect_intersection_area(lr, fr)
                overlap_ratio = inter / max(la, 1e-6)
                x_overlap = _x_intersection(lr, fr) / max(lr.width, 1e-6)

                gap_above = fr.y0 - lr.y1
                gap_below = lr.y0 - fr.y1

                if overlap_ratio >= 0.35:
                    keep[li] = False
                    break

                if x_overlap >= 0.70 and (fr.y0 - 1.5) <= lmid_y <= (fr.y0 + fr.height * 0.60):
                    keep[li] = False
                    break

                if x_overlap >= 0.70 and (0 <= gap_above <= 4.0 or 0 <= gap_below <= 4.0):
                    keep[li] = False
                    break

                line_mid_x = (lr.x0 + lr.x1) / 2.0
                near_top_band = fr.y0 <= lr.y1 <= (fr.y0 + fr.height * 0.55)

                if (fr.x0 <= line_mid_x <= fr.x1) and near_top_band:
                    keep[li] = False
                    break

        self.all_boxes = [r for k, r in zip(keep, self.all_boxes) if k]
        self.box_types = [t for k, t in zip(keep, self.box_types) if k]

    # --------------------------------------------------------
    # Field refinement
    # --------------------------------------------------------
    def _refine_field_rect_from_mask(
        self,
        binv, x, y, w, h,
        zoom_factor=1.0,
        row_frac_thresh=0.45,
        col_frac_thresh=0.18,
        min_inner_h=18,
        pad_px=2
    ):
        roi = binv[y:y+h, x:x+w]
        if roi.size == 0:
            return fitz.Rect(x/zoom_factor, y/zoom_factor, (x+w)/zoom_factor, (y+h)/zoom_factor)

        row_frac = (roi > 0).sum(axis=1) / float(max(w, 1))
        col_frac = (roi > 0).sum(axis=0) / float(max(h, 1))

        strong_rows = np.where(row_frac >= row_frac_thresh)[0]
        strong_cols = np.where(col_frac >= col_frac_thresh)[0]

        if len(strong_rows) >= 2 and len(strong_cols) >= 2:
            ry0 = max(0, int(strong_rows[0]) - pad_px)
            ry1 = min(h, int(strong_rows[-1]) + pad_px + 1)
            rx0 = max(0, int(strong_cols[0]) - pad_px)
            rx1 = min(w, int(strong_cols[-1]) + pad_px + 1)

            if (ry1 - ry0) >= min_inner_h and (rx1 - rx0) >= max(20, int(w * 0.40)):
                return fitz.Rect(
                    (x + rx0) / zoom_factor,
                    (y + ry0) / zoom_factor,
                    (x + rx1) / zoom_factor,
                    (y + ry1) / zoom_factor
                )

        return fitz.Rect(x/zoom_factor, y/zoom_factor, (x+w)/zoom_factor, (y+h)/zoom_factor)

    # --------------------------------------------------------
    # Checkbox logic
    # --------------------------------------------------------
    def looks_like_checkbox(self, binv, x, y, w, h, cc_area=None):
        strictness = self.settings["C_Strict"]
        border_override = self.settings["C_Border"]
        inner_override = self.settings["C_Inner"]
        band_pct = self.settings["C_BandPct"]
        aspect_tol = self.settings["C_AspectTol"]
        extent_low = self.settings["Ext_Low"]
        extent_high = self.settings["Ext_High"]
        fill_min = self.settings["C_FillMin"]
        eps = self.settings["C_Eps"]
        use_extent = self.settings["Use_Extent"]

        roi = binv[y:y+h, x:x+w]
        if roi.size == 0:
            return False

        s = max(10.0, min(float(strictness), 100.0))
        delta = (s - 40.0) / 60.0

        eff_border = border_override + (delta * 0.05)
        eff_inner  = inner_override - (delta * 0.06)
        eff_aspect = aspect_tol - (delta * 0.05)
        eff_fill   = fill_min + (delta * 0.05)

        eff_border = max(0.03, min(0.95, eff_border))
        eff_inner  = max(0.05, min(0.95, eff_inner))
        eff_aspect = max(0.08, min(1.00, eff_aspect))
        eff_fill   = max(0.18, min(0.95, eff_fill))

        aspect = w / float(h)
        if not (1 - eff_aspect <= aspect <= 1 + eff_aspect):
            return False

        if use_extent and cc_area is not None:
            extent = cc_area / float(w * h)
            if extent < extent_low or extent > extent_high:
                return False

        t = max(1, int(min(w, h) * band_pct))
        if w <= 2 * t or h <= 2 * t:
            return False

        top, bottom, left, right = roi[:t, :], roi[-t:, :], roi[:, :t], roi[:, -t:]
        inner = roi[t:-t, t:-t]

        def frac(a):
            return cv2.countNonZero(a) / float(a.size) if a.size > 0 else 0

        if min(frac(top), frac(bottom), frac(left), frac(right)) < eff_border:
            return False
        if frac(inner) > eff_inner:
            return False

        cnts, _ = cv2.findContours(roi, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not cnts:
            return False

        c = max(cnts, key=cv2.contourArea)
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, eps * peri, True)

        if len(approx) != 4 or not cv2.isContourConvex(approx):
            return False

        bx, by, bw, bh = cv2.boundingRect(c)
        c_area = cv2.contourArea(c)
        rect_fill = c_area / float(max(bw * bh, 1))
        if rect_fill < 0.70:
            return False

        pad = max(2, int(min(w, h) * 0.35))
        y0 = max(0, y - pad)
        y1 = min(binv.shape[0], y + h + pad)
        x0 = max(0, x - pad)
        x1 = min(binv.shape[1], x + w + pad)
        outer = binv[y0:y1, x0:x1].copy()

        ix0 = x - x0
        iy0 = y - y0
        ix1 = ix0 + w
        iy1 = iy0 + h
        outer[iy0:iy1, ix0:ix1] = 0

        outer_frac = cv2.countNonZero(outer) / float(max(outer.size, 1))
        if outer_frac > 0.10:
            return False

        if (c_area / float(w * h)) < eff_fill:
            return False

        return True

# =========================
# PART 1 / 4
# Imports + Compatibility + Helpers + Engine Base
# =========================

import os
import io
import json
import math
import base64
import zipfile
import traceback
from functools import partial

# -------------------------------------------------------------------
# Python 3.10+ compatibility patch for older reportlab builds
# Fixes:
# ImportError: cannot import name 'decodestring' from 'base64'
# -------------------------------------------------------------------
if not hasattr(base64, "decodestring"):
    base64.decodestring = base64.decodebytes
if not hasattr(base64, "encodestring"):
    base64.encodestring = base64.encodebytes

import cv2
import fitz
import numpy as np
import pandas as pd

from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas as rl_canvas

from kivy.app import App
from kivy.clock import Clock
from kivy.core.image import Image as CoreImage
from kivy.graphics import Color, Line, Rectangle
from kivy.metrics import dp
from kivy.properties import StringProperty, NumericProperty, BooleanProperty, ListProperty, ObjectProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.gridlayout import GridLayout
from kivy.uix.image import Image
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.slider import Slider
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.utils import platform


# ============================================================
# Defaults
# ============================================================
APP_TITLE = "MediMap Pro: Intelligent Form Automator"
CONFIG_FILENAME = "medimap_config.json"
ZOOM = 4.0
PREVIEW_SCALE = 1.25

DEFAULTS = {
    "F_Area": 500,
    "F_MinW": 15,
    "F_MinH": 35,
    "F_Close": 1,
    "Line_MinW": 100,
    "Line_MaxW": 1680,
    "C_Strict": 40,
    "C_Size": (14, 65),
    "C_Border": 0.10,
    "C_Inner": 0.40,
    "ROI_Max": 200,
    "C_Open": 1,
    "C_Close": 0,
    "C_BandPct": 0.18,
    "C_AspectTol": 0.10,
    "Ext_Low": 0.04,
    "Ext_High": 0.60,
    "C_FillMin": 0.45,
    "C_Eps": 0.04,
    "Use_Extent": False,
    "Is_Grid": False,
    "Grid_N": 1,
}


# ============================================================
# Small helpers
# ============================================================
def safe_name(text):
    text = str(text or "").strip()
    text = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in text)
    return text.strip() or "Unknown"

def _norm_str(v):
    return str(v or "").strip()

def _rect_area(r):
    return max(0.0, r.width) * max(0.0, r.height)

def _rect_intersection_area(a, b):
    ix0 = max(a.x0, b.x0)
    iy0 = max(a.y0, b.y0)
    ix1 = min(a.x1, b.x1)
    iy1 = min(a.y1, b.y1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    return (ix1 - ix0) * (iy1 - iy0)

def _rect_iou(a, b):
    inter = _rect_intersection_area(a, b)
    if inter <= 0:
        return 0.0
    union = (a.width * a.height) + (b.width * b.height) - inter
    return inter / union if union > 0 else 0.0

def _x_intersection(a, b):
    return max(0.0, min(a.x1, b.x1) - max(a.x0, b.x0))

def _x_overlap_ratio(a, b):
    inter_x = _x_intersection(a, b)
    return inter_x / max(min(a.width, b.width), 1e-6)

def _rect_union(rects):
    return fitz.Rect(
        min(r.x0 for r in rects),
        min(r.y0 for r in rects),
        max(r.x1 for r in rects),
        max(r.y1 for r in rects)
    )

def _rect_close(a, b, tol=0.20):
    return (
        abs(a.x0 - b.x0) <= tol and
        abs(a.y0 - b.y0) <= tol and
        abs(a.x1 - b.x1) <= tol and
        abs(a.y1 - b.y1) <= tol
    )

def _rect_list_close(rects_a, rects_b, tol=0.20):
    if len(rects_a) != len(rects_b):
        return False
    aa = sorted(rects_a, key=lambda r: (round(r.y0, 3), r.x0, r.x1, r.y1))
    bb = sorted(rects_b, key=lambda r: (round(r.y0, 3), r.x0, r.x1, r.y1))
    return all(_rect_close(a, b, tol=tol) for a, b in zip(aa, bb))

def find_col_safe(df, *keywords):
    keys = [k.lower() for k in keywords]
    for c in df.columns:
        if all(k in str(c).lower() for k in keys):
            return c
    return None


# ============================================================
# Detection / mapping engine
# ============================================================
class MediMapEngine:
    def __init__(self):
        self.pdf_path = ""
        self.df = pd.DataFrame()
        self.patient_names = []

        self.first_col = None
        self.last_col = None
        self.mid_col = None
        self.suf_col = None
        self.dob_col = None
        self.phil_col = None

        self.geom = {"names": [], "dob": [], "phil": []}
        self.all_boxes = []
        self.box_types = []
        self.custom_mappings = {}
        self.selected_box_ids = []

        self.settings = dict(DEFAULTS)
        self.settings["C_Size"] = list(DEFAULTS["C_Size"])

    # --------------------------------------------------------
    # Data loading
    # --------------------------------------------------------
    def load_dataframe(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".csv":
            df = pd.read_csv(path, dtype=str).fillna("")
        else:
            df = pd.read_excel(path, dtype=str).fillna("")

        self.df = df
        self.first_col = find_col_safe(df, "first", "name")
        self.last_col = find_col_safe(df, "surname") or find_col_safe(df, "last", "name")
        self.mid_col = find_col_safe(df, "middle", "name")
        self.suf_col = find_col_safe(df, "suffix")
        self.dob_col = find_col_safe(df, "birth", "date") or find_col_safe(df, "birthdate")
        self.phil_col = find_col_safe(df, "philhealth")

        f_col_safe = self.first_col if self.first_col else df.columns[0]
        l_col_safe = self.last_col if self.last_col else df.columns[0]

        self.df["_DISPLAY_NAME"] = (
            df[f_col_safe].fillna("").astype(str) + " " +
            df[l_col_safe].fillna("").astype(str)
        ).str.strip()

        self.patient_names = sorted([
            str(x).strip()
            for x in self.df["_DISPLAY_NAME"].dropna().tolist()
            if str(x).strip()
        ])

    def total_pages(self):
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            return 1
        try:
            with fitz.open(self.pdf_path) as doc:
                return max(len(doc), 1)
        except Exception:
            return 1

    # --------------------------------------------------------
    # Box append helpers
    # --------------------------------------------------------
    def _append_box_unique(self, rect, box_type="check", iou_thresh=0.55):
        for j, existing in enumerate(self.all_boxes):
            if self.box_types[j] != box_type:
                continue
            if _rect_iou(rect, existing) >= iou_thresh:
                return
        self.all_boxes.append(rect)
        self.box_types.append(box_type)

    def _append_geom_fields(self):
        for key in ["names", "dob", "phil"]:
            for r in self.geom.get(key, []):
                self._append_box_unique(r, "field", iou_thresh=0.60)

    # --------------------------------------------------------
    # Cleanup helpers
    # --------------------------------------------------------
    def _cleanup_field_fragments(self):
        if not self.all_boxes:
            return

        keep = [True] * len(self.all_boxes)
        field_idxs = [i for i, t in enumerate(self.box_types) if t == "field"]

        for si in field_idxs:
            if not keep[si]:
                continue

            sr = self.all_boxes[si]
            s_area = _rect_area(sr)
            if s_area <= 0:
                keep[si] = False
                continue

            for bi in field_idxs:
                if si == bi or not keep[bi]:
                    continue

                br = self.all_boxes[bi]
                if br.height < sr.height or _rect_area(br) <= s_area:
                    continue

                x_overlap = _x_overlap_ratio(sr, br)
                gap = br.y0 - sr.y1

                if (
                    x_overlap >= 0.80 and
                    0 <= gap <= 6.0 and
                    sr.height <= max(4.0, br.height * 0.45)
                ):
                    keep[si] = False
                    break

                inter = _rect_intersection_area(sr, br)
                if inter > 0:
                    cover_small = inter / max(s_area, 1e-6)
                    if cover_small >= 0.75 and sr.y1 <= (br.y0 + br.height * 0.45):
                        keep[si] = False
                        break

        self.all_boxes = [r for k, r in zip(keep, self.all_boxes) if k]
        self.box_types = [t for k, t in zip(keep, self.box_types) if k]

    def _cleanup_line_field_conflicts(self):
        if not self.all_boxes:
            return

        keep = [True] * len(self.all_boxes)
        field_idxs = [i for i, t in enumerate(self.box_types) if t == "field"]
        line_idxs = [i for i, t in enumerate(self.box_types) if t == "line"]

        for li in line_idxs:
            if not keep[li]:
                continue

            lr = self.all_boxes[li]
            la = _rect_area(lr)
            if la <= 0:
                keep[li] = False
                continue

            lmid_y = (lr.y0 + lr.y1) / 2.0

            for fi in field_idxs:
                fr = self.all_boxes[fi]

                inter = _rect_intersection_area(lr, fr)
                overlap_ratio = inter / max(la, 1e-6)
                x_overlap = _x_intersection(lr, fr) / max(lr.width, 1e-6)

                gap_above = fr.y0 - lr.y1
                gap_below = lr.y0 - fr.y1

                if overlap_ratio >= 0.35:
                    keep[li] = False
                    break

                if x_overlap >= 0.70 and (fr.y0 - 1.5) <= lmid_y <= (fr.y0 + fr.height * 0.60):
                    keep[li] = False
                    break

                if x_overlap >= 0.70 and (0 <= gap_above <= 4.0 or 0 <= gap_below <= 4.0):
                    keep[li] = False
                    break

                line_mid_x = (lr.x0 + lr.x1) / 2.0
                near_top_band = fr.y0 <= lr.y1 <= (fr.y0 + fr.height * 0.55)

                if (fr.x0 <= line_mid_x <= fr.x1) and near_top_band:
                    keep[li] = False
                    break

        self.all_boxes = [r for k, r in zip(keep, self.all_boxes) if k]
        self.box_types = [t for k, t in zip(keep, self.box_types) if k]

    # --------------------------------------------------------
    # Field refinement
    # --------------------------------------------------------
    def _refine_field_rect_from_mask(
        self,
        binv, x, y, w, h,
        zoom_factor=1.0,
        row_frac_thresh=0.45,
        col_frac_thresh=0.18,
        min_inner_h=18,
        pad_px=2
    ):
        roi = binv[y:y+h, x:x+w]
        if roi.size == 0:
            return fitz.Rect(x/zoom_factor, y/zoom_factor, (x+w)/zoom_factor, (y+h)/zoom_factor)

        row_frac = (roi > 0).sum(axis=1) / float(max(w, 1))
        col_frac = (roi > 0).sum(axis=0) / float(max(h, 1))

        strong_rows = np.where(row_frac >= row_frac_thresh)[0]
        strong_cols = np.where(col_frac >= col_frac_thresh)[0]

        if len(strong_rows) >= 2 and len(strong_cols) >= 2:
            ry0 = max(0, int(strong_rows[0]) - pad_px)
            ry1 = min(h, int(strong_rows[-1]) + pad_px + 1)
            rx0 = max(0, int(strong_cols[0]) - pad_px)
            rx1 = min(w, int(strong_cols[-1]) + pad_px + 1)

            if (ry1 - ry0) >= min_inner_h and (rx1 - rx0) >= max(20, int(w * 0.40)):
                return fitz.Rect(
                    (x + rx0) / zoom_factor,
                    (y + ry0) / zoom_factor,
                    (x + rx1) / zoom_factor,
                    (y + ry1) / zoom_factor
                )

        return fitz.Rect(x/zoom_factor, y/zoom_factor, (x+w)/zoom_factor, (y+h)/zoom_factor)

    # --------------------------------------------------------
    # Checkbox logic
    # --------------------------------------------------------
    def looks_like_checkbox(self, binv, x, y, w, h, cc_area=None):
        strictness = self.settings["C_Strict"]
        border_override = self.settings["C_Border"]
        inner_override = self.settings["C_Inner"]
        band_pct = self.settings["C_BandPct"]
        aspect_tol = self.settings["C_AspectTol"]
        extent_low = self.settings["Ext_Low"]
        extent_high = self.settings["Ext_High"]
        fill_min = self.settings["C_FillMin"]
        eps = self.settings["C_Eps"]
        use_extent = self.settings["Use_Extent"]

        roi = binv[y:y+h, x:x+w]
        if roi.size == 0:
            return False

        s = max(10.0, min(float(strictness), 100.0))
        delta = (s - 40.0) / 60.0

        eff_border = border_override + (delta * 0.05)
        eff_inner  = inner_override - (delta * 0.06)
        eff_aspect = aspect_tol - (delta * 0.05)
        eff_fill   = fill_min + (delta * 0.05)

        eff_border = max(0.03, min(0.95, eff_border))
        eff_inner  = max(0.05, min(0.95, eff_inner))
        eff_aspect = max(0.08, min(1.00, eff_aspect))
        eff_fill   = max(0.18, min(0.95, eff_fill))

        aspect = w / float(h)
        if not (1 - eff_aspect <= aspect <= 1 + eff_aspect):
            return False

        if use_extent and cc_area is not None:
            extent = cc_area / float(w * h)
            if extent < extent_low or extent > extent_high:
                return False

        t = max(1, int(min(w, h) * band_pct))
        if w <= 2 * t or h <= 2 * t:
            return False

        top, bottom, left, right = roi[:t, :], roi[-t:, :], roi[:, :t], roi[:, -t:]
        inner = roi[t:-t, t:-t]

        def frac(a):
            return cv2.countNonZero(a) / float(a.size) if a.size > 0 else 0

        if min(frac(top), frac(bottom), frac(left), frac(right)) < eff_border:
            return False
        if frac(inner) > eff_inner:
            return False

        cnts, _ = cv2.findContours(roi, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        if not cnts:
            return False

        c = max(cnts, key=cv2.contourArea)
        peri = cv2.arcLength(c, True)
        approx = cv2.approxPolyDP(c, eps * peri, True)

        if len(approx) != 4 or not cv2.isContourConvex(approx):
            return False

        bx, by, bw, bh = cv2.boundingRect(c)
        c_area = cv2.contourArea(c)
        rect_fill = c_area / float(max(bw * bh, 1))
        if rect_fill < 0.70:
            return False

        pad = max(2, int(min(w, h) * 0.35))
        y0 = max(0, y - pad)
        y1 = min(binv.shape[0], y + h + pad)
        x0 = max(0, x - pad)
        x1 = min(binv.shape[1], x + w + pad)
        outer = binv[y0:y1, x0:x1].copy()

        ix0 = x - x0
        iy0 = y - y0
        ix1 = ix0 + w
        iy1 = iy0 + h
        outer[iy0:iy1, ix0:ix1] = 0

        outer_frac = cv2.countNonZero(outer) / float(max(outer.size, 1))
        if outer_frac > 0.10:
            return False

        if (c_area / float(w * h)) < eff_fill:
            return False

        return True

# =========================
# PART 3 / 4
# Fill + Config Engine
# Add this BELOW Part 2, inside the MediMapEngine class
# =========================

    # --------------------------------------------------------
    # Drawing helpers
    # --------------------------------------------------------
    def _rect_union(self, rects):
        return fitz.Rect(
            min(r.x0 for r in rects),
            min(r.y0 for r in rects),
            max(r.x1 for r in rects),
            max(r.y1 for r in rects)
        )

    def _allocate_cells_by_width(self, rects, total_cells):
        total_cells = max(int(total_cells), 1)
        widths = [max(float(r.width), 1e-6) for r in rects]
        total_w = sum(widths)
        raw = [(w / total_w) * total_cells for w in widths]
        base = [int(np.floor(v)) for v in raw]
        remain = total_cells - sum(base)

        order = sorted(
            range(len(raw)),
            key=lambda i: (raw[i] - base[i]),
            reverse=True
        )
        for i in order[:remain]:
            base[i] += 1

        if sum(base) == 0 and base:
            base[-1] = total_cells

        return base

    def _draw_single_rect_text(self, page, val, rect, ox=0, oy=0, fs_scale=0.60):
        fs = rect.height * fs_scale
        est_w = len(val) * fs * 0.6
        p = fitz.Point(
            rect.x0 + max((rect.width - est_w) / 2, 1) + ox,
            rect.y1 - (rect.height * 0.2) + oy
        )
        page.insert_text(p, val, fontsize=fs, fontname="helv")

    def _draw_grid_rect_text(self, page, val, rect, grid_n=1, ox=0, oy=0, fs_scale=0.60):
        grid_n = max(int(grid_n), 1)
        fs = rect.height * fs_scale
        cell_w = rect.width / grid_n

        for i, ch in enumerate(val[:grid_n]):
            p = fitz.Point(
                rect.x0 + i * cell_w + (cell_w * 0.25) + ox,
                rect.y1 - (rect.height * 0.2) + oy
            )
            page.insert_text(p, ch, fontsize=fs, fontname="helv")

    def draw_logic(self, page, text, rect_or_rects, is_grid=False, grid_n=1, ox=0, oy=0, fs_scale=0.65):
        val = str(text).strip()
        if not val or val.lower() in ["nan", "none"]:
            return

        if val.endswith(".0"):
            val = val[:-2]
        val = val.upper()

        rects = rect_or_rects if isinstance(rect_or_rects, list) else [rect_or_rects]
        rects = [r if isinstance(r, fitz.Rect) else fitz.Rect(*r) for r in rects]
        rects = sorted(rects, key=lambda rr: (round(rr.y0, 3), rr.x0))
        if not rects:
            return

        if is_grid and len(rects) > 1:
            counts = self._allocate_cells_by_width(rects, grid_n)
            pos = 0
            for rect, n_cells in zip(rects, counts):
                if n_cells <= 0:
                    continue
                seg = val[pos:pos + n_cells]
                if not seg:
                    break
                self._draw_grid_rect_text(page, seg, rect, grid_n=n_cells, ox=ox, oy=oy, fs_scale=fs_scale)
                pos += n_cells
            return

        if is_grid:
            target = rects[0] if len(rects) == 1 else self._rect_union(rects)
            self._draw_grid_rect_text(page, val, target, grid_n=grid_n, ox=ox, oy=oy, fs_scale=fs_scale)
        else:
            target = rects[0] if len(rects) == 1 else self._rect_union(rects)
            self._draw_single_rect_text(page, val, target, ox=ox, oy=oy, fs_scale=fs_scale)

    # --------------------------------------------------------
    # Document processing
    # --------------------------------------------------------
    def process_doc(self, patient_name, page_idx=0):
        if self.df is None or self.df.empty:
            raise ValueError("DataFrame is empty.")

        matches = self.df[self.df["_DISPLAY_NAME"] == patient_name]
        if matches.empty:
            raise ValueError(f"Patient not found: {patient_name}")

        row = matches.iloc[0]
        doc = fitz.open(self.pdf_path)
        page = doc[page_idx]

        if page_idx == 0:
            name_values = [
                row.get(self.first_col, ""),
                row.get(self.mid_col, ""),
                row.get(self.last_col, ""),
                row.get(self.suf_col, "")
            ]

            for val, rect in zip(name_values, self.geom.get("names", [])):
                self.draw_logic(page, val, rect)

            dt = pd.to_datetime(row.get(self.dob_col, ""), errors="coerce")
            if not pd.isna(dt) and self.geom.get("dob"):
                dob_values = [dt.strftime("%m"), dt.strftime("%d"), dt.strftime("%Y")]
                dob_grid_n = [2, 2, 4]

                for val, n, rect in zip(dob_values, dob_grid_n, self.geom["dob"][:3]):
                    self.draw_logic(page, val, rect, is_grid=True, grid_n=n)

        for configs in self.custom_mappings.values():
            for c in configs:
                if c.get("page", 0) != page_idx:
                    continue

                target_rects = self._mapping_rect_list(c)
                if not target_rects:
                    continue

                csv_val = str(row.get(c["column"], "")).strip().upper()
                trigger = str(c.get("trigger", "")).strip().upper()

                if trigger and csv_val == trigger:
                    target = self._rect_union(target_rects)
                    fs = target.height * 0.95
                    p = fitz.Point(
                        target.x0 + (target.width * 0.15),
                        target.y1 - (target.height * 0.15)
                    )
                    page.insert_text(p, "X", fontsize=fs, fontname="Helvetica-Bold")

                elif not trigger:
                    self.draw_logic(
                        page,
                        row.get(c["column"], ""),
                        target_rects,
                        is_grid=c.get("g", False),
                        grid_n=c.get("n", 1),
                        ox=0,
                        oy=0,
                        fs_scale=0.65
                    )

        return doc

    # --------------------------------------------------------
    # Config helpers
    # --------------------------------------------------------
    def _norm_str(self, v):
        return str(v or "").strip()

    def _rect_close(self, a, b, tol=0.20):
        return (
            abs(a.x0 - b.x0) <= tol and
            abs(a.y0 - b.y0) <= tol and
            abs(a.x1 - b.x1) <= tol and
            abs(a.y1 - b.y1) <= tol
        )

    def _rect_list_close(self, rects_a, rects_b, tol=0.20):
        if len(rects_a) != len(rects_b):
            return False

        aa = sorted(rects_a, key=lambda r: (round(r.y0, 3), r.x0, r.x1, r.y1))
        bb = sorted(rects_b, key=lambda r: (round(r.y0, 3), r.x0, r.x1, r.y1))
        return all(self._rect_close(a, b, tol=tol) for a, b in zip(aa, bb))

    def _mapping_entry_to_current(self, item, fallback_key=""):
        column = self._norm_str(item.get("column", ""))
        trigger = self._norm_str(item.get("trigger", ""))

        if not column and fallback_key:
            if "_" in fallback_key:
                parts = fallback_key.rsplit("_", 1)
                if len(parts) == 2:
                    column = self._norm_str(parts[0])
                    if not trigger:
                        trigger = self._norm_str(parts[1])
            else:
                column = self._norm_str(fallback_key)

        rects = []
        if item.get("rects"):
            for r in item.get("rects", []):
                if isinstance(r, fitz.Rect):
                    rects.append(r)
                elif isinstance(r, (list, tuple)) and len(r) == 4:
                    rects.append(fitz.Rect(*r))
        elif item.get("rect") is not None:
            r = item.get("rect")
            if isinstance(r, fitz.Rect):
                rects = [r]
            elif isinstance(r, (list, tuple)) and len(r) == 4:
                rects = [fitz.Rect(*r)]

        rects = sorted(rects, key=lambda rr: (round(rr.y0, 3), rr.x0, rr.x1, rr.y1))

        return {
            "column": column,
            "trigger": trigger,
            "rects": rects,
            "page": int(item.get("page", 0)),
            "g": bool(item.get("g", item.get("is_grid", False))),
            "n": int(item.get("n", item.get("grid_n", 1))),
        }

    def _extract_current_mapping_entries(self, mapping_dict):
        out = []
        for map_key, lst in mapping_dict.items():
            for item in lst:
                out.append(self._mapping_entry_to_current(item, fallback_key=map_key))
        return out

    def _extract_cfg_mapping_entries(self, cfg):
        out = []
        loaded = cfg.get("custom_mappings", {}) or {}
        for map_key, lst in loaded.items():
            if isinstance(lst, dict):
                lst = [lst]
            for item in lst:
                out.append(self._mapping_entry_to_current(item, fallback_key=map_key))
        return out

    def _same_mapping_identity(self, a, b, tol=0.20):
        return (
            int(a.get("page", 0)) == int(b.get("page", 0)) and
            self._norm_str(a.get("column", "")).upper() == self._norm_str(b.get("column", "")).upper() and
            self._norm_str(a.get("trigger", "")).upper() == self._norm_str(b.get("trigger", "")).upper() and
            self._rect_list_close(a.get("rects", []), b.get("rects", []), tol=tol)
        )

    def _same_target_rects(self, a, b, tol=0.20):
        return (
            int(a.get("page", 0)) == int(b.get("page", 0)) and
            self._rect_list_close(a.get("rects", []), b.get("rects", []), tol=tol)
        )

    def _entries_to_mapping_dict(self, entries):
        out = {}
        for e in entries:
            map_key = f"{self._norm_str(e.get('column', ''))}_{self._norm_str(e.get('trigger', ''))}"
            if map_key not in out:
                out[map_key] = []

            out[map_key].append({
                "column": self._norm_str(e.get("column", "")),
                "trigger": self._norm_str(e.get("trigger", "")),
                "rects": [r if isinstance(r, fitz.Rect) else fitz.Rect(*r) for r in e.get("rects", [])],
                "page": int(e.get("page", 0)),
                "g": bool(e.get("g", False)),
                "n": int(e.get("n", 1)),
            })
        return out

    def collect_config(self):
        mappings_serial = {}

        for k, lst in self.custom_mappings.items():
            mappings_serial[k] = []
            for item in lst:
                rects = self._mapping_rect_list(item)
                mappings_serial[k].append({
                    "column": item.get("column", ""),
                    "trigger": item.get("trigger", ""),
                    "rects": [[r.x0, r.y0, r.x1, r.y1] for r in rects],
                    "page": int(item.get("page", 0)),
                    "g": bool(item.get("g", False)),
                    "n": int(item.get("n", 1)),
                })

        return {
            "version": 3,
            "pdf_path": self.pdf_path,
            "zoom": float(ZOOM),

            "F_Area": int(self.settings["F_Area"]),
            "F_MinW": int(self.settings["F_MinW"]),
            "F_MinH": int(self.settings["F_MinH"]),
            "F_Close": int(self.settings["F_Close"]),

            "Line_MinW": int(self.settings["Line_MinW"]),
            "Line_MaxW": int(self.settings["Line_MaxW"]),

            "C_Strict": int(self.settings["C_Strict"]),
            "C_Size": [int(self.settings["C_Size"][0]), int(self.settings["C_Size"][1])],

            "C_Border": float(self.settings["C_Border"]),
            "C_Inner": float(self.settings["C_Inner"]),
            "ROI_Max": int(self.settings["ROI_Max"]),

            "C_Open": int(self.settings["C_Open"]),
            "C_Close": int(self.settings["C_Close"]),

            "C_BandPct": float(self.settings["C_BandPct"]),
            "C_AspectTol": float(self.settings["C_AspectTol"]),

            "Ext_Low": float(self.settings["Ext_Low"]),
            "Ext_High": float(self.settings["Ext_High"]),

            "C_FillMin": float(self.settings["C_FillMin"]),
            "C_Eps": float(self.settings["C_Eps"]),
            "Use_Extent": bool(self.settings["Use_Extent"]),

            "Is_Grid": bool(self.settings["Is_Grid"]),
            "Grid_N": int(self.settings["Grid_N"]),

            "custom_mappings": mappings_serial,
        }

    def apply_config(self, cfg):
        self.settings["F_Area"] = int(cfg.get("F_Area", self.settings["F_Area"]))
        self.settings["F_MinW"] = int(cfg.get("F_MinW", self.settings["F_MinW"]))
        self.settings["F_MinH"] = int(cfg.get("F_MinH", self.settings["F_MinH"]))
        self.settings["F_Close"] = int(cfg.get("F_Close", self.settings["F_Close"]))

        self.settings["Line_MinW"] = int(cfg.get("Line_MinW", self.settings["Line_MinW"]))
        self.settings["Line_MaxW"] = int(cfg.get("Line_MaxW", self.settings["Line_MaxW"]))

        self.settings["C_Strict"] = int(cfg.get("C_Strict", self.settings["C_Strict"]))

        c_size = cfg.get("C_Size", self.settings["C_Size"])
        if isinstance(c_size, (list, tuple)) and len(c_size) == 2:
            self.settings["C_Size"] = (int(c_size[0]), int(c_size[1]))

        self.settings["C_Border"] = float(cfg.get("C_Border", self.settings["C_Border"]))
        self.settings["C_Inner"] = float(cfg.get("C_Inner", self.settings["C_Inner"]))
        self.settings["ROI_Max"] = int(cfg.get("ROI_Max", self.settings["ROI_Max"]))
        self.settings["C_Open"] = int(cfg.get("C_Open", self.settings["C_Open"]))
        self.settings["C_Close"] = int(cfg.get("C_Close", self.settings["C_Close"]))
        self.settings["C_BandPct"] = float(cfg.get("C_BandPct", self.settings["C_BandPct"]))
        self.settings["C_AspectTol"] = float(cfg.get("C_AspectTol", self.settings["C_AspectTol"]))
        self.settings["Ext_Low"] = float(cfg.get("Ext_Low", self.settings["Ext_Low"]))
        self.settings["Ext_High"] = float(cfg.get("Ext_High", self.settings["Ext_High"]))
        self.settings["C_FillMin"] = float(cfg.get("C_FillMin", self.settings["C_FillMin"]))
        self.settings["C_Eps"] = float(cfg.get("C_Eps", self.settings["C_Eps"]))
        self.settings["Use_Extent"] = bool(cfg.get("Use_Extent", self.settings["Use_Extent"]))
        self.settings["Is_Grid"] = bool(cfg.get("Is_Grid", self.settings["Is_Grid"]))
        self.settings["Grid_N"] = int(cfg.get("Grid_N", self.settings["Grid_N"]))

        self.custom_mappings.clear()
        loaded_mappings = cfg.get("custom_mappings", {}) or {}
        for k, lst in loaded_mappings.items():
            if isinstance(lst, dict):
                lst = [lst]
            elif not isinstance(lst, list):
                continue

            self.custom_mappings[k] = []
            for item in lst:
                if not isinstance(item, dict):
                    continue

                rects = []
                if item.get("rects"):
                    for coords in item.get("rects", []):
                        if isinstance(coords, (list, tuple)) and len(coords) == 4:
                            rects.append(fitz.Rect(*coords))
                else:
                    rect_coords = item.get("rect", [0, 0, 0, 0])
                    if isinstance(rect_coords, (list, tuple)) and len(rect_coords) == 4:
                        rects = [fitz.Rect(*rect_coords)]

                self.custom_mappings[k].append({
                    "column": str(item.get("column", "")),
                    "trigger": str(item.get("trigger", "")),
                    "rects": rects,
                    "page": int(item.get("page", 0)),
                    "g": bool(item.get("g", item.get("is_grid", False))),
                    "n": int(item.get("n", item.get("grid_n", 1))),
                })

    def merge_config_into_current(self, incoming_cfg, keep_current_detection=True, prefer="incoming"):
        current_cfg = self.collect_config()

        if keep_current_detection:
            merged_cfg = dict(current_cfg)
        else:
            merged_cfg = dict(current_cfg)
            for k, v in incoming_cfg.items():
                if k != "custom_mappings":
                    merged_cfg[k] = v

        current_entries = self._extract_current_mapping_entries(self.custom_mappings)
        incoming_entries = self._extract_cfg_mapping_entries(incoming_cfg)

        merged_entries = list(current_entries)

        for inc in incoming_entries:
            if not inc.get("column") and not inc.get("rects"):
                continue

            if any(self._same_mapping_identity(old, inc) for old in merged_entries):
                continue

            conflict_idx = None
            for i, old in enumerate(merged_entries):
                if self._same_target_rects(old, inc):
                    conflict_idx = i
                    break

            if conflict_idx is not None:
                if prefer == "incoming":
                    merged_entries[conflict_idx] = inc
            else:
                merged_entries.append(inc)

        merged_cfg["custom_mappings"] = {}
        merged_mapping_dict = self._entries_to_mapping_dict(merged_entries)

        for k, lst in merged_mapping_dict.items():
            merged_cfg["custom_mappings"][k] = []
            for item in lst:
                merged_cfg["custom_mappings"][k].append({
                    "column": item.get("column", ""),
                    "trigger": item.get("trigger", ""),
                    "rects": [[r.x0, r.y0, r.x1, r.y1] for r in item.get("rects", [])],
                    "page": int(item.get("page", 0)),
                    "g": bool(item.get("g", False)),
                    "n": int(item.get("n", 1)),
                })

        self.apply_config(merged_cfg)
        return merged_cfg

    def save_config(self, path):
        cfg = self.collect_config()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)

    def load_config(self, path):
        with open(path, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        self.apply_config(cfg)
        return cfg

    def get_preview_pixmap_with_boxes(self, patient_name, page_idx=0, preview_zoom=1.5):
        if not self.all_boxes:
            self.run_detection(page_idx=page_idx)

        doc = self.process_doc(patient_name, page_idx=page_idx)
        page = doc[page_idx]
        pix = page.get_pixmap(matrix=fitz.Matrix(preview_zoom, preview_zoom))
        img = cv2.imdecode(np.frombuffer(pix.tobytes("png"), np.uint8), cv2.IMREAD_COLOR)

        if img is None:
            doc.close()
            raise ValueError("Failed to build preview image.")

        for i, r in enumerate(self.all_boxes):
            x0 = int(r.x0 * preview_zoom)
            y0 = int(r.y0 * preview_zoom)
            x1 = int(r.x1 * preview_zoom)
            y1 = int(r.y1 * preview_zoom)

            box_type = self.box_types[i]
            if box_type == "check":
                color = (0, 0, 255)      # red
            elif box_type == "line":
                color = (0, 215, 255)    # yellow-ish
            else:
                color = (0, 255, 0)      # green

            cv2.rectangle(img, (x0, y0), (x1, y1), color, 2)

            label = str(i)
            cv2.rectangle(
                img,
                (x0, max(0, y0 - 14)),
                (x0 + 20, y0),
                (0, 0, 0),
                -1
            )
            cv2.putText(
                img,
                label,
                (x0 + 2, y0 - 3),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.35,
                (255, 255, 255),
                1,
                cv2.LINE_AA
            )

        doc.close()
        return img

# =========================
# PART 4 / 4
# Kivy App UI + Wiring
# Paste this BELOW Part 3
# =========================

class MediMapProLayout(BoxLayout):
    pass


class MediMapProApp(App):
    def build(self):
        self.title = "MediMap Pro"
        self.engine = MediMapEngine()

        root = BoxLayout(orientation="vertical", spacing=8, padding=8)

        # -----------------------------
        # Header / status
        # -----------------------------
        self.status_lbl = Label(
            text="MediMap Pro\nLoad CSV + PDF to begin",
            size_hint_y=None,
            height=70,
            halign="left",
            valign="middle"
        )
        self.status_lbl.bind(size=self._sync_label_text_size)
        root.add_widget(self.status_lbl)

        # -----------------------------
        # File chooser
        # -----------------------------
        chooser_wrap = BoxLayout(orientation="vertical", size_hint_y=0.38, spacing=6)

        self.file_chooser = FileChooserListView(
            path="/sdcard/Download" if platform == "android" else os.path.expanduser("~")
        )
        chooser_wrap.add_widget(self.file_chooser)

        file_btn_row = BoxLayout(size_hint_y=None, height=48, spacing=6)

        self.btn_load_pdf = Button(text="Load PDF")
        self.btn_load_pdf.bind(on_release=self.on_load_pdf)
        file_btn_row.add_widget(self.btn_load_pdf)

        self.btn_load_csv = Button(text="Load CSV/XLSX")
        self.btn_load_csv.bind(on_release=self.on_load_csv)
        file_btn_row.add_widget(self.btn_load_csv)

        self.btn_load_cfg = Button(text="Load Config")
        self.btn_load_cfg.bind(on_release=self.on_load_config)
        file_btn_row.add_widget(self.btn_load_cfg)

        self.btn_merge_cfg = Button(text="Merge Config")
        self.btn_merge_cfg.bind(on_release=self.on_merge_config)
        file_btn_row.add_widget(self.btn_merge_cfg)

        self.btn_save_cfg = Button(text="Save Config")
        self.btn_save_cfg.bind(on_release=self.on_save_config)
        file_btn_row.add_widget(self.btn_save_cfg)

        chooser_wrap.add_widget(file_btn_row)
        root.add_widget(chooser_wrap)

        # -----------------------------
        # Main controls
        # -----------------------------
        ctl1 = BoxLayout(size_hint_y=None, height=46, spacing=6)

        self.patient_spinner = Spinner(
            text="Select Patient",
            values=[],
            size_hint_x=0.42
        )
        self.patient_spinner.bind(text=self.on_patient_change)
        ctl1.add_widget(self.patient_spinner)

        self.page_input = TextInput(
            text="0",
            multiline=False,
            hint_text="Page Index",
            size_hint_x=0.12
        )
        ctl1.add_widget(self.page_input)

        self.btn_prev = Button(text="Prev", size_hint_x=0.10)
        self.btn_prev.bind(on_release=self.on_prev_page)
        ctl1.add_widget(self.btn_prev)

        self.btn_next = Button(text="Next", size_hint_x=0.10)
        self.btn_next.bind(on_release=self.on_next_page)
        ctl1.add_widget(self.btn_next)

        self.btn_detect = Button(text="Run Detect", size_hint_x=0.13)
        self.btn_detect.bind(on_release=self.on_run_detect)
        ctl1.add_widget(self.btn_detect)

        self.btn_preview = Button(text="Preview", size_hint_x=0.13)
        self.btn_preview.bind(on_release=self.on_preview)
        ctl1.add_widget(self.btn_preview)

        root.add_widget(ctl1)

        # -----------------------------
        # Settings row A
        # -----------------------------
        ctl2 = BoxLayout(size_hint_y=None, height=42, spacing=4)

        self.f_area = TextInput(text=str(DEFAULTS["F_Area"]), multiline=False, hint_text="F_Area")
        self.f_minw = TextInput(text=str(DEFAULTS["F_MinW"]), multiline=False, hint_text="F_MinW")
        self.f_minh = TextInput(text=str(DEFAULTS["F_MinH"]), multiline=False, hint_text="F_MinH")
        self.f_close = TextInput(text=str(DEFAULTS["F_Close"]), multiline=False, hint_text="F_Close")
        self.line_minw = TextInput(text=str(DEFAULTS["Line_MinW"]), multiline=False, hint_text="Line_MinW")
        self.line_maxw = TextInput(text=str(DEFAULTS["Line_MaxW"]), multiline=False, hint_text="Line_MaxW")

        for w in [self.f_area, self.f_minw, self.f_minh, self.f_close, self.line_minw, self.line_maxw]:
            ctl2.add_widget(w)

        root.add_widget(ctl2)

        # -----------------------------
        # Settings row B
        # -----------------------------
        ctl3 = BoxLayout(size_hint_y=None, height=42, spacing=4)

        self.c_strict = TextInput(text=str(DEFAULTS["C_Strict"]), multiline=False, hint_text="C_Strict")
        self.c_size_min = TextInput(text=str(DEFAULTS["C_Size"][0]), multiline=False, hint_text="C_Size_Min")
        self.c_size_max = TextInput(text=str(DEFAULTS["C_Size"][1]), multiline=False, hint_text="C_Size_Max")
        self.c_border = TextInput(text=str(DEFAULTS["C_Border"]), multiline=False, hint_text="C_Border")
        self.c_inner = TextInput(text=str(DEFAULTS["C_Inner"]), multiline=False, hint_text="C_Inner")
        self.roi_max = TextInput(text=str(DEFAULTS["ROI_Max"]), multiline=False, hint_text="ROI_Max")

        for w in [self.c_strict, self.c_size_min, self.c_size_max, self.c_border, self.c_inner, self.roi_max]:
            ctl3.add_widget(w)

        root.add_widget(ctl3)

        # -----------------------------
        # Settings row C
        # -----------------------------
        ctl4 = BoxLayout(size_hint_y=None, height=42, spacing=4)

        self.c_open = TextInput(text=str(DEFAULTS["C_Open"]), multiline=False, hint_text="C_Open")
        self.c_close = TextInput(text=str(DEFAULTS["C_Close"]), multiline=False, hint_text="C_Close")
        self.c_band = TextInput(text=str(DEFAULTS["C_BandPct"]), multiline=False, hint_text="C_BandPct")
        self.c_aspect = TextInput(text=str(DEFAULTS["C_AspectTol"]), multiline=False, hint_text="C_AspectTol")
        self.ext_low = TextInput(text=str(DEFAULTS["Ext_Low"]), multiline=False, hint_text="Ext_Low")
        self.ext_high = TextInput(text=str(DEFAULTS["Ext_High"]), multiline=False, hint_text="Ext_High")
        self.c_fill = TextInput(text=str(DEFAULTS["C_FillMin"]), multiline=False, hint_text="C_FillMin")
        self.c_eps = TextInput(text=str(DEFAULTS["C_Eps"]), multiline=False, hint_text="C_Eps")
        self.use_extent = TextInput(text="0", multiline=False, hint_text="Use_Extent 0/1")

        for w in [self.c_open, self.c_close, self.c_band, self.c_aspect, self.ext_low, self.ext_high, self.c_fill, self.c_eps, self.use_extent]:
            ctl4.add_widget(w)

        root.add_widget(ctl4)

        # -----------------------------
        # Mapping controls
        # -----------------------------
        map_row = BoxLayout(size_hint_y=None, height=42, spacing=6)

        self.box_ids_input = TextInput(multiline=False, hint_text="Box IDs e.g. 15 or 15,16", size_hint_x=0.23)
        self.column_spinner = Spinner(text="Select Column", values=[], size_hint_x=0.26)
        self.trigger_input = TextInput(multiline=False, hint_text="Trigger", size_hint_x=0.17)
        self.grid_flag_input = TextInput(text="0", multiline=False, hint_text="Grid 0/1", size_hint_x=0.10)
        self.grid_n_input = TextInput(text="1", multiline=False, hint_text="Grid N", size_hint_x=0.10)

        self.btn_assign = Button(text="Assign Mapping", size_hint_x=0.14)
        self.btn_assign.bind(on_release=self.on_assign_mapping)

        map_row.add_widget(self.box_ids_input)
        map_row.add_widget(self.column_spinner)
        map_row.add_widget(self.trigger_input)
        map_row.add_widget(self.grid_flag_input)
        map_row.add_widget(self.grid_n_input)
        map_row.add_widget(self.btn_assign)

        root.add_widget(map_row)

        # -----------------------------
        # Output controls
        # -----------------------------
        out_row = BoxLayout(size_hint_y=None, height=48, spacing=6)

        self.btn_generate_one = Button(text="Generate Single PDF")
        self.btn_generate_one.bind(on_release=self.on_generate_single)
        out_row.add_widget(self.btn_generate_one)

        self.btn_generate_batch = Button(text="Generate Batch PDFs")
        self.btn_generate_batch.bind(on_release=self.on_generate_batch)
        out_row.add_widget(self.btn_generate_batch)

        root.add_widget(out_row)

        # -----------------------------
        # Preview image
        # -----------------------------
        self.preview = Image(
            allow_stretch=True,
            keep_ratio=True
        )
        root.add_widget(self.preview)

        return root

    # --------------------------------------------------------
    # UI helpers
    # --------------------------------------------------------
    def _sync_label_text_size(self, instance, value):
        instance.text_size = value

    def set_status(self, text):
        self.status_lbl.text = text

    def get_selected_file(self):
        if not self.file_chooser.selection:
            return None
        return self.file_chooser.selection[0]

    def current_page_idx(self):
        try:
            return max(0, int(self.page_input.text.strip()))
        except Exception:
            return 0

    def apply_ui_settings_to_engine(self):
        try:
            self.engine.settings["F_Area"] = int(self.f_area.text)
            self.engine.settings["F_MinW"] = int(self.f_minw.text)
            self.engine.settings["F_MinH"] = int(self.f_minh.text)
            self.engine.settings["F_Close"] = int(self.f_close.text)

            self.engine.settings["Line_MinW"] = int(self.line_minw.text)
            self.engine.settings["Line_MaxW"] = int(self.line_maxw.text)

            self.engine.settings["C_Strict"] = int(self.c_strict.text)
            self.engine.settings["C_Size"] = (
                int(self.c_size_min.text),
                int(self.c_size_max.text)
            )

            self.engine.settings["C_Border"] = float(self.c_border.text)
            self.engine.settings["C_Inner"] = float(self.c_inner.text)
            self.engine.settings["ROI_Max"] = int(self.roi_max.text)

            self.engine.settings["C_Open"] = int(self.c_open.text)
            self.engine.settings["C_Close"] = int(self.c_close.text)
            self.engine.settings["C_BandPct"] = float(self.c_band.text)
            self.engine.settings["C_AspectTol"] = float(self.c_aspect.text)
            self.engine.settings["Ext_Low"] = float(self.ext_low.text)
            self.engine.settings["Ext_High"] = float(self.ext_high.text)
            self.engine.settings["C_FillMin"] = float(self.c_fill.text)
            self.engine.settings["C_Eps"] = float(self.c_eps.text)
            self.engine.settings["Use_Extent"] = bool(int(self.use_extent.text.strip() or "0"))
        except Exception as e:
            raise ValueError(f"Invalid settings input: {e}")

    def push_engine_settings_to_ui(self):
        s = self.engine.settings
        self.f_area.text = str(s["F_Area"])
        self.f_minw.text = str(s["F_MinW"])
        self.f_minh.text = str(s["F_MinH"])
        self.f_close.text = str(s["F_Close"])
        self.line_minw.text = str(s["Line_MinW"])
        self.line_maxw.text = str(s["Line_MaxW"])

        self.c_strict.text = str(s["C_Strict"])
        self.c_size_min.text = str(s["C_Size"][0])
        self.c_size_max.text = str(s["C_Size"][1])
        self.c_border.text = str(s["C_Border"])
        self.c_inner.text = str(s["C_Inner"])
        self.roi_max.text = str(s["ROI_Max"])

        self.c_open.text = str(s["C_Open"])
        self.c_close.text = str(s["C_Close"])
        self.c_band.text = str(s["C_BandPct"])
        self.c_aspect.text = str(s["C_AspectTol"])
        self.ext_low.text = str(s["Ext_Low"])
        self.ext_high.text = str(s["Ext_High"])
        self.c_fill.text = str(s["C_FillMin"])
        self.c_eps.text = str(s["C_Eps"])
        self.use_extent.text = "1" if s["Use_Extent"] else "0"

    def refresh_patient_and_column_lists(self):
        if self.engine.df is None or self.engine.df.empty:
            self.patient_spinner.values = []
            self.column_spinner.values = []
            self.patient_spinner.text = "Select Patient"
            self.column_spinner.text = "Select Column"
            return

        patient_names = sorted([
            str(x).strip()
            for x in self.engine.df["_DISPLAY_NAME"].dropna().tolist()
            if str(x).strip()
        ])
        self.patient_spinner.values = patient_names
        self.patient_spinner.text = patient_names[0] if patient_names else "Select Patient"

        cols = sorted([str(c) for c in self.engine.df.columns.tolist()])
        self.column_spinner.values = cols
        self.column_spinner.text = cols[0] if cols else "Select Column"

    def selected_patient(self):
        val = self.patient_spinner.text.strip()
        if not val or val == "Select Patient":
            return None
        return val

    def render_preview_image(self, img_bgr):
        if img_bgr is None:
            return

        rgba = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGBA)
        buf = rgba.tobytes()
        texture = Texture.create(size=(rgba.shape[1], rgba.shape[0]), colorfmt="rgba")
        texture.blit_buffer(buf, colorfmt="rgba", bufferfmt="ubyte")
        texture.flip_vertical()
        self.preview.texture = texture

    # --------------------------------------------------------
    # File loading
    # --------------------------------------------------------
    def on_load_pdf(self, instance):
        try:
            path = self.get_selected_file()
            if not path or not path.lower().endswith(".pdf"):
                self.set_status("Please select a PDF file.")
                return

            self.engine.pdf_path = path

            with fitz.open(path) as doc:
                total_pages = len(doc)

            self.page_input.text = "0"
            self.set_status(f"PDF loaded:\n{os.path.basename(path)}\nPages: {total_pages}")
        except Exception as e:
            self.set_status(f"Load PDF error:\n{e}")

    def on_load_csv(self, instance):
        try:
            path = self.get_selected_file()
            if not path:
                self.set_status("Please select a CSV/XLSX file.")
                return

            self.engine.load_dataframe(path)
            self.refresh_patient_and_column_lists()
            self.set_status(f"Data loaded:\n{os.path.basename(path)}\nRows: {len(self.engine.df)}")
        except Exception as e:
            self.set_status(f"Load data error:\n{e}")

    def on_load_config(self, instance):
        try:
            path = self.get_selected_file()
            if not path or not path.lower().endswith(".json"):
                self.set_status("Please select a JSON config file.")
                return

            self.engine.load_config(path)
            self.push_engine_settings_to_ui()
            self.set_status(f"Config loaded:\n{os.path.basename(path)}")
        except Exception as e:
            self.set_status(f"Load config error:\n{e}")

    def on_merge_config(self, instance):
        try:
            path = self.get_selected_file()
            if not path or not path.lower().endswith(".json"):
                self.set_status("Please select a JSON config file.")
                return

            with open(path, "r", encoding="utf-8") as f:
                incoming = json.load(f)

            self.apply_ui_settings_to_engine()
            self.engine.merge_config_into_current(
                incoming_cfg=incoming,
                keep_current_detection=True,
                prefer="incoming"
            )
            self.push_engine_settings_to_ui()
            self.set_status(f"Config merged:\n{os.path.basename(path)}")
        except Exception as e:
            self.set_status(f"Merge config error:\n{e}")

    def on_save_config(self, instance):
        try:
            if not self.engine.pdf_path:
                self.set_status("Load a PDF first before saving config.")
                return

            self.apply_ui_settings_to_engine()
            out_path = os.path.join(
                os.path.dirname(self.engine.pdf_path),
                "medimap_config.json"
            )
            self.engine.save_config(out_path)
            self.set_status(f"Config saved:\n{out_path}")
        except Exception as e:
            self.set_status(f"Save config error:\n{e}")

    # --------------------------------------------------------
    # Navigation
    # --------------------------------------------------------
    def on_prev_page(self, instance):
        try:
            idx = max(0, self.current_page_idx() - 1)
            self.page_input.text = str(idx)
            self.on_preview(None)
        except Exception as e:
            self.set_status(f"Prev page error:\n{e}")

    def on_next_page(self, instance):
        try:
            idx = self.current_page_idx() + 1
            self.page_input.text = str(idx)
            self.on_preview(None)
        except Exception as e:
            self.set_status(f"Next page error:\n{e}")

    def on_patient_change(self, spinner, text):
        if text and text != "Select Patient":
            Clock.schedule_once(lambda dt: self.on_preview(None), 0.1)

    # --------------------------------------------------------
    # Detection / preview
    # --------------------------------------------------------
    def on_run_detect(self, instance):
        try:
            if not self.engine.pdf_path:
                self.set_status("Load PDF first.")
                return

            self.apply_ui_settings_to_engine()
            page_idx = self.current_page_idx()
            self.engine.run_detection(page_idx=page_idx)
            self.set_status(
                f"Detection done.\n"
                f"Page: {page_idx}\n"
                f"Boxes: {len(self.engine.all_boxes)}"
            )
            self.on_preview(None)
        except Exception as e:
            traceback.print_exc()
            self.set_status(f"Detect error:\n{e}")

    def on_preview(self, instance):
        try:
            if not self.engine.pdf_path:
                self.set_status("Load PDF first.")
                return

            patient = self.selected_patient()
            if not patient:
                self.set_status("Load data and select a patient first.")
                return

            self.apply_ui_settings_to_engine()
            page_idx = self.current_page_idx()

            img = self.engine.get_preview_pixmap_with_boxes(
                patient_name=patient,
                page_idx=page_idx,
                preview_zoom=1.5
            )
            self.render_preview_image(img)

            self.set_status(
                f"Preview rendered.\n"
                f"Patient: {patient}\n"
                f"Page: {page_idx}\n"
                f"Boxes: {len(self.engine.all_boxes)}"
            )
        except Exception as e:
            traceback.print_exc()
            self.set_status(f"Preview error:\n{e}")

    # --------------------------------------------------------
    # Mapping
    # --------------------------------------------------------
    def on_assign_mapping(self, instance):
        try:
            column = self.column_spinner.text.strip()
            if not column or column == "Select Column":
                self.set_status("Please select a column.")
                return

            raw_ids = [x.strip() for x in self.box_ids_input.text.split(",") if x.strip()]
            box_ids = []
            for x in raw_ids:
                if not x.isdigit():
                    raise ValueError(f"Invalid Box ID: {x}")
                box_ids.append(int(x))

            if not box_ids:
                raise ValueError("Please enter at least one Box ID.")

            selected_rects = []
            for b_id in box_ids:
                if b_id < 0 or b_id >= len(self.engine.all_boxes):
                    raise ValueError(f"Box ID {b_id} is out of range.")
                selected_rects.append(self.engine.all_boxes[b_id])

            selected_rects = sorted(selected_rects, key=lambda rr: (round(rr.y0, 3), rr.x0))
            trigger = self.trigger_input.text.strip()
            is_grid = bool(int(self.grid_flag_input.text.strip() or "0"))
            grid_n = int(self.grid_n_input.text.strip() or "1")
            page_idx = self.current_page_idx()

            map_key = f"{column}_{trigger}"

            def mapping_conflicts_with_selected(mapping_item, selected_rects):
                existing_rects = self.engine._mapping_rect_list(mapping_item)
                if not existing_rects:
                    return False

                for er in existing_rects:
                    for sr in selected_rects:
                        if self.engine._rects_refer_to_same_target(er, sr, tol=0.20):
                            return True
                return False

            for k in list(self.engine.custom_mappings.keys()):
                kept = []
                for m in self.engine.custom_mappings[k]:
                    if m.get("page", 0) != page_idx:
                        kept.append(m)
                        continue

                    if mapping_conflicts_with_selected(m, selected_rects):
                        continue

                    kept.append(m)

                if kept:
                    self.engine.custom_mappings[k] = kept
                else:
                    del self.engine.custom_mappings[k]

            if map_key not in self.engine.custom_mappings:
                self.engine.custom_mappings[map_key] = []

            self.engine.custom_mappings[map_key].append({
                "column": column,
                "trigger": trigger,
                "rects": selected_rects,
                "page": page_idx,
                "g": is_grid,
                "n": grid_n
            })

            self.set_status(
                f"Mapping saved.\n"
                f"Page: {page_idx}\n"
                f"Column: {column}\n"
                f"Boxes: {box_ids}"
            )
            self.on_preview(None)
        except Exception as e:
            self.set_status(f"Assign mapping error:\n{e}")

    # --------------------------------------------------------
    # Output generation
    # --------------------------------------------------------
    def on_generate_single(self, instance):
        try:
            if not self.engine.pdf_path:
                self.set_status("Load PDF first.")
                return

            patient = self.selected_patient()
            if not patient:
                self.set_status("Select a patient first.")
                return

            page_idx = self.current_page_idx()
            doc = self.engine.process_doc(patient, page_idx=page_idx)

            safe_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in patient).strip()
            if not safe_name:
                safe_name = "Unknown"

            out_path = os.path.join(
                os.path.dirname(self.engine.pdf_path),
                f"Filled_{safe_name}_page_{page_idx + 1}.pdf"
            )

            doc.save(out_path)
            doc.close()

            self.set_status(f"Generated:\n{out_path}")
        except Exception as e:
            traceback.print_exc()
            self.set_status(f"Generate single error:\n{e}")

    def on_generate_batch(self, instance):
        try:
            if self.engine.df is None or self.engine.df.empty:
                self.set_status("Load CSV/XLSX first.")
                return

            if not self.engine.pdf_path:
                self.set_status("Load PDF first.")
                return

            names = sorted(self.engine.df["_DISPLAY_NAME"].dropna().astype(str).unique())
            out_dir = os.path.join(os.path.dirname(self.engine.pdf_path), "batch_output")
            os.makedirs(out_dir, exist_ok=True)

            success = 0
            skipped = 0

            for patient_name in names:
                try:
                    doc = self.engine.process_doc(patient_name, page_idx=self.current_page_idx())

                    safe_name = "".join(
                        c if c.isalnum() or c in (" ", "-", "_") else "_"
                        for c in patient_name
                    ).strip()
                    if not safe_name:
                        safe_name = "Unknown"

                    out_path = os.path.join(out_dir, f"Filled_{safe_name}.pdf")
                    doc.save(out_path)
                    doc.close()
                    success += 1
                except Exception:
                    skipped += 1

            self.set_status(
                f"Batch done.\n"
                f"Folder: {out_dir}\n"
                f"Success: {success}\n"
                f"Skipped: {skipped}"
            )
        except Exception as e:
            traceback.print_exc()
            self.set_status(f"Generate batch error:\n{e}")


if __name__ == "__main__":
    MediMapProApp().run()
