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
import re
from urllib.parse import urlparse, parse_qs
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

from kivy.app import App
from kivy.clock import Clock
from kivy.core.image import Image as CoreImage
from kivy.graphics import Color, Line, Rectangle
from kivy.graphics.texture import Texture
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
PREVIEW_SCALE = 2.2

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

def extract_gsheet_id(url):
    """
    Extract Google Sheet ID from common Google Sheets URLs.
    """
    if not url:
        return None

    url = str(url).strip()

    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if m:
        return m.group(1)

    parsed = urlparse(url)
    qs = parse_qs(parsed.query)
    if "id" in qs and qs["id"]:
        return qs["id"][0]

    return None


def extract_gsheet_gid(url):
    """
    Extract gid from Google Sheet URL. Defaults to 0 if missing.
    """
    if not url:
        return "0"

    parsed = urlparse(str(url).strip())
    qs = parse_qs(parsed.query)

    if "gid" in qs and qs["gid"]:
        return str(qs["gid"][0])

    if parsed.fragment:
        frag_qs = parse_qs(parsed.fragment)
        if "gid" in frag_qs and frag_qs["gid"]:
            return str(frag_qs["gid"][0])

        m = re.search(r"gid=([0-9]+)", parsed.fragment)
        if m:
            return m.group(1)

    return "0"


def gsheet_url_to_csv_export(url):
    """
    Convert a Google Sheet URL into a direct CSV export URL.
    Works for sheets that are accessible without OAuth
    (public/shared appropriately).
    """
    sheet_id = extract_gsheet_id(url)
    if not sheet_id:
        raise ValueError("Could not extract Google Sheet ID from URL.")

    gid = extract_gsheet_gid(url)
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"

# ============================================================
# Detection / mapping engine
# ============================================================
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
    def load_dataframe(self, path_or_url):
        src = str(path_or_url).strip()
    
        if not src:
            raise ValueError("No data source provided.")
    
        # ---------------------------------------------
        # Google Sheet URL support (CSV export endpoint)
        # ---------------------------------------------
        if src.startswith("http://") or src.startswith("https://"):
            if "docs.google.com/spreadsheets" in src:
                csv_url = gsheet_url_to_csv_export(src)
                try:
                    df = pd.read_csv(csv_url, dtype=str).fillna("")
                except Exception as e:
                    raise ValueError(
                        "Failed to load Google Sheet from URL. "
                        "Make sure the sheet/tab is accessible. "
                        f"Details: {e}"
                    )
            else:
                raise ValueError("Only Google Sheets URLs are supported for URL loading.")
    
        # ---------------------------------------------
        # Local file support
        # ---------------------------------------------
        else:
            ext = os.path.splitext(src)[1].lower()
            if ext == ".csv":
                df = pd.read_csv(src, dtype=str).fillna("")
            elif ext in [".xlsx", ".xls"]:
                df = pd.read_excel(src, dtype=str).fillna("")
            else:
                raise ValueError("Unsupported file type. Use CSV, XLSX, XLS, or a Google Sheet URL.")
    
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

    def load_pdf(self, path):
        """Set PDF path, validate file, and return page count."""
        if not path:
            raise ValueError("No PDF path provided.")
        if not os.path.exists(path):
            raise FileNotFoundError(f"PDF file not found: {path}")

        try:
            with fitz.open(path) as doc:
                total = len(doc)
                if total <= 0:
                    raise ValueError("PDF has no pages.")
        except Exception as e:
            raise ValueError(f"Failed to open PDF: {e}")

        self.pdf_path = path
        self.all_boxes = []
        self.box_types = []
        self.geom = {"names": [], "dob": [], "phil": []}
        return total

    def total_pages(self):
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            return 1
        try:
            with fitz.open(self.pdf_path) as doc:
                return max(len(doc), 1)
        except Exception:
            return 1

    def get_raw_preview_pixmap(self, page_idx=0, preview_zoom=1.5):
        """Render the raw loaded PDF page without filling or boxes."""
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            raise FileNotFoundError("PDF path is missing or invalid.")

        with fitz.open(self.pdf_path) as doc:
            total = len(doc)
            if total <= 0:
                raise ValueError("PDF has no pages.")
            page_idx = max(0, min(int(page_idx), total - 1))
            page = doc[page_idx]
            pix = page.get_pixmap(matrix=fitz.Matrix(preview_zoom, preview_zoom))
            img = cv2.imdecode(np.frombuffer(pix.tobytes("png"), np.uint8), cv2.IMREAD_COLOR)

        if img is None:
            raise ValueError("Failed to render raw PDF preview image.")

        return img



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
# PART 2 / 4
# Detection Engine Continuation
# Add this BELOW Part 1, inside the MediMapEngine class
# =========================

    # --------------------------------------------------------
    # Line detection
    # --------------------------------------------------------
    def find_answer_lines(self, img, zoom_factor=3.0, min_line_w=100, max_line_w=900):
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        gray = cv2.GaussianBlur(gray, (3, 3), 0)
        _, binv = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (max(40, int(min_line_w) // 4), 2))
        detected_lines = cv2.morphologyEx(binv, cv2.MORPH_OPEN, kernel, iterations=1)
        cnts, _ = cv2.findContours(detected_lines, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        line_rects = []
        img_h, img_w = img.shape[:2]

        for c in cnts:
            x, y, w, h = cv2.boundingRect(c)

            effective_max_line_w = int(max_line_w)

            if w > min_line_w and w < effective_max_line_w and 1 <= h <= 12:
                roi_tl = binv[max(0, y - 5):y, x:x + 4]
                roi_tr = binv[max(0, y - 5):y, x + w - 4:x + w]
                if cv2.countNonZero(roi_tl) > 4 and cv2.countNonZero(roi_tr) > 4:
                    continue

                roi_left_edge = binv[y:y + h, max(0, x - 5):x]
                roi_right_edge = binv[y:y + h, x + w:min(img_w, x + w + 5)]
                if cv2.countNonZero(roi_left_edge) > 8 and cv2.countNonZero(roi_right_edge) > 8:
                    continue

                roi_above = binv[max(0, y - 12):y, x:x + w]
                if roi_above.size > 0 and (cv2.countNonZero(roi_above) / float(roi_above.size)) > 0.35:
                    continue

                if y < (img_h * 0.05) or y > (img_h * 0.95):
                    continue

                line_rects.append(
                    fitz.Rect(
                        x / zoom_factor,
                        (y - 28) / zoom_factor,
                        (x + w) / zoom_factor,
                        y / zoom_factor
                    )
                )

        return line_rects

    # --------------------------------------------------------
    # Checkbox ROI scan
    # --------------------------------------------------------
    def find_checkbox_rects_in_roi(self, binv, x, y, w, h):
        roi = binv[y:y + h, x:x + w]
        if roi.size == 0:
            return []

        min_sz, max_sz = self.settings["C_Size"]
        cnts, _ = cv2.findContours(roi, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        found = []

        for c in cnts:
            rx, ry, rw, rh = cv2.boundingRect(c)
            if not (min_sz <= rw <= max_sz and min_sz <= rh <= max_sz):
                continue

            cc_area = cv2.contourArea(c)
            if self.looks_like_checkbox(binv, x + rx, y + ry, rw, rh, cc_area=cc_area):
                found.append((x + rx, y + ry, rw, rh))

        return found

    # --------------------------------------------------------
    # Red checkbox color-assisted detection
    # --------------------------------------------------------
    def find_red_checkbox_candidates(
        self,
        img_bgr,
        zoom_factor=3.0,
        min_sz=8,
        max_sz=40,
        border_min=0.06,
        inner_max=0.65,
        band_pct=0.12,
        aspect_tol=0.45,
        fill_min=0.18,
        eps=0.06
    ):
        hsv = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2HSV)

        lower1 = np.array([0, 70, 60], dtype=np.uint8)
        upper1 = np.array([12, 255, 255], dtype=np.uint8)
        lower2 = np.array([168, 70, 60], dtype=np.uint8)
        upper2 = np.array([180, 255, 255], dtype=np.uint8)

        mask1 = cv2.inRange(hsv, lower1, upper1)
        mask2 = cv2.inRange(hsv, lower2, upper2)
        red_mask = cv2.bitwise_or(mask1, mask2)

        red_mask = cv2.morphologyEx(red_mask, cv2.MORPH_CLOSE, np.ones((2, 2), np.uint8))
        red_mask = cv2.morphologyEx(red_mask, cv2.MORPH_OPEN, np.ones((1, 1), np.uint8))

        cnts, _ = cv2.findContours(red_mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        out = []

        for c in cnts:
            x, y, w, h = cv2.boundingRect(c)

            if not (min_sz <= w <= max_sz and min_sz <= h <= max_sz):
                continue

            aspect = w / float(h) if h > 0 else 999
            if not (1 - aspect_tol <= aspect <= 1 + aspect_tol):
                continue

            roi = red_mask[y:y + h, x:x + w]
            if roi.size == 0:
                continue

            t = max(1, int(min(w, h) * band_pct))
            if w <= 2 * t or h <= 2 * t:
                continue

            top, bottom = roi[:t, :], roi[-t:, :]
            left, right = roi[:, :t], roi[:, -t:]
            inner = roi[t:-t, t:-t]

            def frac(a):
                return cv2.countNonZero(a) / float(a.size) if a.size > 0 else 0.0

            if min(frac(top), frac(bottom), frac(left), frac(right)) < border_min:
                continue
            if frac(inner) > inner_max:
                continue

            peri = cv2.arcLength(c, True)
            approx = cv2.approxPolyDP(c, eps * peri, True)
            if len(approx) != 4:
                continue

            area = cv2.contourArea(c)
            if (area / float(w * h)) < fill_min:
                continue

            out.append(
                fitz.Rect(
                    x / zoom_factor,
                    y / zoom_factor,
                    (x + w) / zoom_factor,
                    (y + h) / zoom_factor
                )
            )

        return out

    # --------------------------------------------------------
    # Main detection entry
    # --------------------------------------------------------
    def run_detection(self, page_idx=0):
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            raise FileNotFoundError("PDF path is missing or invalid.")

        doc = fitz.open(self.pdf_path)
        page_obj = doc[page_idx]
        pix = page_obj.get_pixmap(matrix=fitz.Matrix(ZOOM, ZOOM))
        img = cv2.imdecode(np.frombuffer(pix.tobytes(), np.uint8), cv2.IMREAD_COLOR)

        if img is None:
            doc.close()
            raise ValueError("Failed to render PDF page into image.")

        h_img, w_img = img.shape[:2]

        field_area_thresh = int(self.settings["F_Area"])
        field_min_w_val = int(self.settings["F_MinW"])
        field_min_h_val = int(self.settings["F_MinH"])
        field_close_k = int(self.settings["F_Close"])
        line_min_w_val = int(self.settings["Line_MinW"])
        line_max_w_val = int(self.settings["Line_MaxW"])
        checkbox_min_sz = int(self.settings["C_Size"][0])
        checkbox_max_sz = int(self.settings["C_Size"][1])

        red_border_min = float(self.settings["C_Border"])
        red_inner_max = float(self.settings["C_Inner"])
        red_roi_max_val = int(self.settings["ROI_Max"])
        red_open_k = int(self.settings["C_Open"])
        red_close_k = int(self.settings["C_Close"])
        red_band_pct = float(self.settings["C_BandPct"])
        red_aspect_tol = float(self.settings["C_AspectTol"])
        red_extent_low_val = float(self.settings["Ext_Low"])
        red_extent_high_val = float(self.settings["Ext_High"])
        red_fill_min = float(self.settings["C_FillMin"])
        red_eps_val = float(self.settings["C_Eps"])
        red_use_extent_val = bool(self.settings["Use_Extent"])

        # --- Page 0 DEM ROI band detection
        self.geom = {"names": [], "dob": [], "phil": []}
        if page_idx == 0:
            y0_roi, y1_roi = int(0.22 * h_img), int(0.82 * h_img)
            gray_roi = cv2.cvtColor(img[y0_roi:y1_roi, :], cv2.COLOR_BGR2GRAY)
            _, bin_roi = cv2.threshold(gray_roi, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
            num, _, stats, _ = cv2.connectedComponentsWithStats(bin_roi)

            cands = [
                (x, y, ww, hh)
                for i, (x, y, ww, hh, area) in enumerate(stats[1:], 1)
                if hh >= 20 and area > field_area_thresh
            ]

            def get_band(l, h):
                b = [c for c in cands if l <= c[1] / gray_roi.shape[0] <= h]
                return [
                    fitz.Rect(
                        x / ZOOM,
                        (y + y0_roi) / ZOOM,
                        (x + ww) / ZOOM,
                        (y + hh + y0_roi) / ZOOM
                    )
                    for x, y, ww, hh in sorted(b, key=lambda t: t[0])
                ]

            self.geom = {
                "names": get_band(0.46, 0.50),
                "dob": get_band(0.56, 0.66),
                "phil": get_band(0.36, 0.40),
            }

        # --- General detection
        full_gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        base_bin = cv2.threshold(full_gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

        bin_checks = base_bin.copy()
        if red_close_k > 0:
            bin_checks = cv2.morphologyEx(
                bin_checks,
                cv2.MORPH_CLOSE,
                np.ones((red_close_k, red_close_k), np.uint8)
            )
        if red_open_k > 0:
            bin_checks = cv2.morphologyEx(
                bin_checks,
                cv2.MORPH_OPEN,
                np.ones((red_open_k, red_open_k), np.uint8)
            )

        if field_close_k > 0:
            bin_fields = cv2.morphologyEx(
                base_bin,
                cv2.MORPH_CLOSE,
                np.ones((field_close_k, field_close_k), np.uint8)
            )
        else:
            bin_fields = base_bin

        self.all_boxes = []
        self.box_types = []

        if page_idx == 0:
            self._append_geom_fields()

        # --- answer lines
        line_rects = self.find_answer_lines(
            img,
            ZOOM,
            min_line_w=line_min_w_val,
            max_line_w=line_max_w_val
        )
        for lr in line_rects:
            self.all_boxes.append(lr)
            self.box_types.append("line")

        # --- color-assisted red checkbox detection
        red_color_boxes = self.find_red_checkbox_candidates(
            img,
            zoom_factor=ZOOM,
            min_sz=max(7, checkbox_min_sz - 5),
            max_sz=min(int(checkbox_max_sz), 42),
            border_min=max(0.05, float(red_border_min) * 0.7),
            inner_max=min(0.75, float(red_inner_max) + 0.20),
            band_pct=max(0.10, min(float(red_band_pct), 0.16)),
            aspect_tol=max(0.30, float(red_aspect_tol)),
            fill_min=max(0.12, float(red_fill_min) * 0.45),
            eps=max(0.04, float(red_eps_val))
        )
        for rr in red_color_boxes:
            self._append_box_unique(rr, "check", iou_thresh=0.45)

        # --- grayscale checkbox detection
        nf, _, sf, _ = cv2.connectedComponentsWithStats(bin_checks)
        small_min = max(7, int(checkbox_min_sz) - 5)

        for i in range(1, nf):
            x, y, w, h, area = sf[i]
            if w >= w_img * 0.9:
                continue

            aspect = w / float(h) if h > 0 else 999

            if (
                small_min <= w <= checkbox_max_sz and
                small_min <= h <= checkbox_max_sz and
                0.68 <= aspect <= 1.45
            ):
                if self.looks_like_checkbox(bin_checks, x, y, w, h, area):
                    self._append_box_unique(
                        fitz.Rect(x / ZOOM, y / ZOOM, (x + w) / ZOOM, (y + h) / ZOOM),
                        "check",
                        iou_thresh=0.45
                    )

            elif (
                max(w, h) <= int(red_roi_max_val) and
                min(w, h) >= small_min and
                0.55 <= aspect <= 1.75
            ):
                for fx, fy, fw, fh in self.find_checkbox_rects_in_roi(bin_checks, x, y, w, h):
                    self._append_box_unique(
                        fitz.Rect(fx / ZOOM, fy / ZOOM, (fx + fw) / ZOOM, (fy + fh) / ZOOM),
                        "check",
                        iou_thresh=0.45
                    )

            elif (
                max(w, h) <= int(red_roi_max_val) and
                min(w, h) >= small_min and
                0.60 <= aspect <= 1.60
            ):
                for fx, fy, fw, fh in self.find_checkbox_rects_in_roi(bin_checks, x, y, w, h):
                    self.all_boxes.append(
                        fitz.Rect(fx / ZOOM, fy / ZOOM, (fx + fw) / ZOOM, (fy + fh) / ZOOM)
                    )
                    self.box_types.append("check")

            elif max(w, h) <= int(red_roi_max_val) and min(w, h) >= checkbox_min_sz:
                for fx, fy, fw, fh in self.find_checkbox_rects_in_roi(bin_checks, x, y, w, h):
                    self.all_boxes.append(
                        fitz.Rect(fx / ZOOM, fy / ZOOM, (fx + fw) / ZOOM, (fy + fh) / ZOOM)
                    )
                    self.box_types.append("check")

        # --- field detection
        nf2, _, sf2, _ = cv2.connectedComponentsWithStats(bin_fields)
        for i in range(1, nf2):
            x, y, w, h, area = sf2[i]

            if w >= w_img * 0.90 or h >= h_img * 0.12 or area >= (w_img * h_img * 0.03):
                continue

            if (w > 1200 and h > 120) or (w > 1600 and h > 80):
                continue

            if h < field_min_h_val or w < field_min_w_val or area < field_area_thresh:
                continue

            refined_rect = self._refine_field_rect_from_mask(
                bin_fields,
                x, y, w, h,
                zoom_factor=ZOOM,
                row_frac_thresh=0.45,
                col_frac_thresh=0.18,
                min_inner_h=max(18, int(field_min_h_val * 0.45)),
                pad_px=2
            )

            self.all_boxes.append(refined_rect)
            self.box_types.append("field")

        self._cleanup_field_fragments()
        self._cleanup_line_field_conflicts()
        doc.close()

    # --------------------------------------------------------
    # Mapping helpers
    # --------------------------------------------------------
    def _mapping_rect_list(self, item):
        rects = item.get("rects")
        if rects:
            out = []
            for r in rects:
                out.append(r if isinstance(r, fitz.Rect) else fitz.Rect(*r))
            return sorted(out, key=lambda rr: (round(rr.y0, 3), rr.x0))

        r = item.get("rect")
        if r is None:
            return []
        return [r if isinstance(r, fitz.Rect) else fitz.Rect(*r)]

    def _rects_refer_to_same_target(self, a, b, tol=0.01):
        if (
            abs(a.x0 - b.x0) < tol and
            abs(a.y0 - b.y0) < tol and
            abs(a.x1 - b.x1) < tol and
            abs(a.y1 - b.y1) < tol
        ):
            return True

        acx = (a.x0 + a.x1) / 2.0
        acy = (a.y0 + a.y1) / 2.0
        bcx = (b.x0 + b.x1) / 2.0
        bcy = (b.y0 + b.y1) / 2.0

        if (b.x0 <= acx <= b.x1 and b.y0 <= acy <= b.y1):
            return True
        if (a.x0 <= bcx <= a.x1 and a.y0 <= bcy <= a.y1):
            return True

        inter = _rect_intersection_area(a, b)
        if inter <= 0:
            return False

        a_area = max(a.width * a.height, 1e-6)
        b_area = max(b.width * b.height, 1e-6)
        iou = inter / (a_area + b_area - inter)
        cover_small = inter / min(a_area, b_area)

        return iou >= 0.60 or cover_small >= 0.80

    def _mapping_match_score(self, target_rect, mapped_rect, tol=0.01):
        if (
            abs(target_rect.x0 - mapped_rect.x0) < tol and
            abs(target_rect.y0 - mapped_rect.y0) < tol and
            abs(target_rect.x1 - mapped_rect.x1) < tol and
            abs(target_rect.y1 - mapped_rect.y1) < tol
        ):
            return 9999.0

        inter = _rect_intersection_area(target_rect, mapped_rect)
        if inter <= 0:
            return -1.0

        t_area = max(target_rect.width * target_rect.height, 1e-6)
        iou = inter / (t_area + max(mapped_rect.width * mapped_rect.height, 1e-6) - inter)
        cover_target = inter / t_area

        cx = (target_rect.x0 + target_rect.x1) / 2.0
        cy = (target_rect.y0 + target_rect.y1) / 2.0
        center_inside = (mapped_rect.x0 <= cx <= mapped_rect.x1) and (mapped_rect.y0 <= cy <= mapped_rect.y1)

        score = 0.0
        if center_inside:
            score += 1.0
        score += iou * 10.0
        score += cover_target * 5.0
        return score

    def describe_box_mapping(self, box_idx, page_idx):
        if box_idx < 0 or box_idx >= len(self.all_boxes):
            return "EMPTY"

        target_rect = self.all_boxes[box_idx]
        best_hit = None
        best_score = -1.0

        for map_key, configs in self.custom_mappings.items():
            for c in configs:
                if c.get("page", 0) != page_idx:
                    continue

                rects = self._mapping_rect_list(c)
                if not rects:
                    continue

                for r in rects:
                    score = self._mapping_match_score(target_rect, r)
                    if score > best_score:
                        best_score = score
                        best_hit = c

        if best_hit is not None and best_score >= 1.0:
            col = str(best_hit.get("column", "")).strip()
            trig = str(best_hit.get("trigger", "")).strip()
            is_grid = bool(best_hit.get("g", False))
            grid_n = int(best_hit.get("n", 1))

            parts = [f"Mapped: {col}"]
            if trig:
                parts.append(f"trigger={trig}")
            if is_grid:
                parts.append(f"grid={grid_n}")

            return " | ".join(parts)

        return "EMPTY"

    def get_box_mapping_payload(self, box_id, page_idx):
        if box_id < 0 or box_id >= len(self.all_boxes):
            return {"box_id": box_id, "page": page_idx, "status": "EMPTY", "entries": []}

        target_rect = self.all_boxes[box_id]
        best_hit = None
        best_score = -1.0

        for map_key, configs in self.custom_mappings.items():
            for c in configs:
                if c.get("page", 0) != page_idx:
                    continue

                rects = self._mapping_rect_list(c)
                if not rects:
                    continue

                for r in rects:
                    score = self._mapping_match_score(target_rect, r)
                    if score > best_score:
                        best_score = score
                        best_hit = c

        if best_hit is not None and best_score >= 1.0:
            return {
                "box_id": box_id,
                "page": page_idx,
                "status": "MAPPED",
                "entries": [
                    {
                        "column": str(best_hit.get("column", "")),
                        "trigger": str(best_hit.get("trigger", "")),
                        "g": bool(best_hit.get("g", False)),
                        "n": int(best_hit.get("n", 1)),
                    }
                ]
            }

        return {
            "box_id": box_id,
            "page": page_idx,
            "status": "EMPTY",
            "entries": []
        }

    def assign_mapping(self, box_ids, column, trigger, is_grid, grid_n, page_idx):
        selected_rects = []
        for b_id in box_ids:
            if b_id < 0 or b_id >= len(self.all_boxes):
                raise ValueError(f"Box ID {b_id} is out of range.")
            selected_rects.append(self.all_boxes[b_id])

        selected_rects = sorted(selected_rects, key=lambda rr: (round(rr.y0, 3), rr.x0))
        map_key = f"{str(column).strip()}_{str(trigger).strip()}"

        def mapping_conflicts_with_selected(mapping_item, selected_rects_):
            existing_rects = self._mapping_rect_list(mapping_item)
            if not existing_rects:
                return False

            for er in existing_rects:
                for sr in selected_rects_:
                    if self._rects_refer_to_same_target(er, sr, tol=0.20):
                        return True
            return False

        for k in list(self.custom_mappings.keys()):
            kept = []
            for m in self.custom_mappings[k]:
                if m.get("page", 0) != page_idx:
                    kept.append(m)
                    continue

                if mapping_conflicts_with_selected(m, selected_rects):
                    continue
                kept.append(m)

            if kept:
                self.custom_mappings[k] = kept
            else:
                del self.custom_mappings[k]

        if map_key not in self.custom_mappings:
            self.custom_mappings[map_key] = []

        self.custom_mappings[map_key].append({
            "column": str(column).strip(),
            "trigger": str(trigger).strip(),
            "rects": selected_rects,
            "page": int(page_idx),
            "g": bool(is_grid),
            "n": int(grid_n),
        })

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

        root = BoxLayout(orientation="horizontal", spacing=10, padding=10)

        # =========================================
        # LEFT PANEL WRAP
        # =========================================
        left_wrap = BoxLayout(
            orientation="vertical",
            size_hint_x=0.40
        )

        left_scroll = ScrollView(
            do_scroll_x=False,
            do_scroll_y=True,
            bar_width=10,
            scroll_type=["bars", "content"]
        )

        left = GridLayout(
            cols=1,
            spacing=8,
            size_hint_y=None,
            padding=(0, 0, 4, 0)
        )
        left.bind(minimum_height=left.setter("height"))

        # -----------------------------
        # Status section
        # -----------------------------
        status_box = BoxLayout(
            orientation="vertical",
            size_hint_y=None,
            height=82,
            spacing=2
        )

        status_title = Label(
            text="[b]MediMap Pro[/b]",
            markup=True,
            size_hint_y=None,
            height=30,
            halign="left",
            valign="middle"
        )
        status_title.bind(size=self._sync_label_text_size)
        status_box.add_widget(status_title)

        self.status_lbl = Label(
            text="Load CSV/XLSX, Google Sheet URL, and PDF to begin",
            size_hint_y=None,
            height=50,
            halign="left",
            valign="top"
        )
        self.status_lbl.bind(size=self._sync_label_text_size)
        status_box.add_widget(self.status_lbl)

        left.add_widget(status_box)

        # -----------------------------
        # Data / files title
        # -----------------------------
        files_title = Label(
            text="[b]Data & Files[/b]",
            markup=True,
            size_hint_y=None,
            height=24,
            halign="left",
            valign="middle"
        )
        files_title.bind(size=self._sync_label_text_size)
        left.add_widget(files_title)

        # -----------------------------
        # File action buttons row 1
        # -----------------------------
        file_btn_row_1 = GridLayout(
            cols=3,
            size_hint_y=None,
            height=44,
            spacing=6
        )

        self.btn_load_pdf = Button(text="Load PDF")
        self.btn_load_pdf.bind(on_release=self.on_load_pdf)
        file_btn_row_1.add_widget(self.btn_load_pdf)

        self.btn_load_csv = Button(text="Load CSV/XLSX")
        self.btn_load_csv.bind(on_release=self.on_load_csv)
        file_btn_row_1.add_widget(self.btn_load_csv)

        self.btn_load_gsheet = Button(text="Load Google Sheet URL")
        self.btn_load_gsheet.bind(on_release=self.on_load_gsheet_url)
        file_btn_row_1.add_widget(self.btn_load_gsheet)

        left.add_widget(file_btn_row_1)

        # -----------------------------
        # File action buttons row 2
        # -----------------------------
        file_btn_row_2 = GridLayout(
            cols=3,
            size_hint_y=None,
            height=44,
            spacing=6
        )

        self.btn_load_cfg = Button(text="Load Config")
        self.btn_load_cfg.bind(on_release=self.on_load_config)
        file_btn_row_2.add_widget(self.btn_load_cfg)

        self.btn_merge_cfg = Button(text="Merge Config")
        self.btn_merge_cfg.bind(on_release=self.on_merge_config)
        file_btn_row_2.add_widget(self.btn_merge_cfg)

        self.btn_save_cfg = Button(text="Save Config")
        self.btn_save_cfg.bind(on_release=self.on_save_config)
        file_btn_row_2.add_widget(self.btn_save_cfg)

        left.add_widget(file_btn_row_2)

        # -----------------------------
        # Navigation title
        # -----------------------------
        nav_title = Label(
            text="[b]Navigation[/b]",
            markup=True,
            size_hint_y=None,
            height=24,
            halign="left",
            valign="middle"
        )
        nav_title.bind(size=self._sync_label_text_size)
        left.add_widget(nav_title)

        # -----------------------------
        # Navigation row 1
        # -----------------------------
        nav_row_1 = BoxLayout(size_hint_y=None, height=44, spacing=6)

        self.patient_spinner = Spinner(
            text="Select Patient",
            values=[],
            size_hint_x=0.56
        )
        self.patient_spinner.bind(text=self.on_patient_change)
        nav_row_1.add_widget(self.patient_spinner)

        self.page_input = TextInput(
            text="0",
            multiline=False,
            hint_text="Page",
            size_hint_x=0.12
        )
        nav_row_1.add_widget(self.page_input)

        self.btn_prev = Button(text="Prev", size_hint_x=0.16)
        self.btn_prev.bind(on_release=self.on_prev_page)
        nav_row_1.add_widget(self.btn_prev)

        self.btn_next = Button(text="Next", size_hint_x=0.16)
        self.btn_next.bind(on_release=self.on_next_page)
        nav_row_1.add_widget(self.btn_next)

        left.add_widget(nav_row_1)

        # -----------------------------
        # Navigation row 2
        # -----------------------------
        nav_row_2 = GridLayout(cols=2, size_hint_y=None, height=44, spacing=6)

        self.btn_detect = Button(text="Run Detect")
        self.btn_detect.bind(on_release=self.on_run_detect)
        nav_row_2.add_widget(self.btn_detect)

        self.btn_preview = Button(text="Refresh Preview")
        self.btn_preview.bind(on_release=self.on_preview)
        nav_row_2.add_widget(self.btn_preview)

        left.add_widget(nav_row_2)

        # -----------------------------
        # Detection Settings title
        # -----------------------------
        settings_title = Label(
            text="[b]Detection Settings[/b]",
            markup=True,
            size_hint_y=None,
            height=24,
            halign="left",
            valign="middle"
        )
        settings_title.bind(size=self._sync_label_text_size)
        left.add_widget(settings_title)

        # A
        ctl2 = GridLayout(cols=3, size_hint_y=None, height=44, spacing=6)
        self.f_area = TextInput(text=str(DEFAULTS["F_Area"]), multiline=False, hint_text="F_Area")
        self.f_minw = TextInput(text=str(DEFAULTS["F_MinW"]), multiline=False, hint_text="F_MinW")
        self.f_minh = TextInput(text=str(DEFAULTS["F_MinH"]), multiline=False, hint_text="F_MinH")
        ctl2.add_widget(self.f_area)
        ctl2.add_widget(self.f_minw)
        ctl2.add_widget(self.f_minh)
        left.add_widget(ctl2)

        # B
        ctl2b = GridLayout(cols=3, size_hint_y=None, height=44, spacing=6)
        self.f_close = TextInput(text=str(DEFAULTS["F_Close"]), multiline=False, hint_text="F_Close")
        self.line_minw = TextInput(text=str(DEFAULTS["Line_MinW"]), multiline=False, hint_text="Line_MinW")
        self.line_maxw = TextInput(text=str(DEFAULTS["Line_MaxW"]), multiline=False, hint_text="Line_MaxW")
        ctl2b.add_widget(self.f_close)
        ctl2b.add_widget(self.line_minw)
        ctl2b.add_widget(self.line_maxw)
        left.add_widget(ctl2b)

        # C
        ctl3 = GridLayout(cols=3, size_hint_y=None, height=44, spacing=6)
        self.c_strict = TextInput(text=str(DEFAULTS["C_Strict"]), multiline=False, hint_text="C_Strict")
        self.c_size_min = TextInput(text=str(DEFAULTS["C_Size"][0]), multiline=False, hint_text="C_Size_Min")
        self.c_size_max = TextInput(text=str(DEFAULTS["C_Size"][1]), multiline=False, hint_text="C_Size_Max")
        ctl3.add_widget(self.c_strict)
        ctl3.add_widget(self.c_size_min)
        ctl3.add_widget(self.c_size_max)
        left.add_widget(ctl3)

        # D
        ctl3b = GridLayout(cols=3, size_hint_y=None, height=44, spacing=6)
        self.c_border = TextInput(text=str(DEFAULTS["C_Border"]), multiline=False, hint_text="C_Border")
        self.c_inner = TextInput(text=str(DEFAULTS["C_Inner"]), multiline=False, hint_text="C_Inner")
        self.roi_max = TextInput(text=str(DEFAULTS["ROI_Max"]), multiline=False, hint_text="ROI_Max")
        ctl3b.add_widget(self.c_border)
        ctl3b.add_widget(self.c_inner)
        ctl3b.add_widget(self.roi_max)
        left.add_widget(ctl3b)

        # E
        ctl4 = GridLayout(cols=3, size_hint_y=None, height=44, spacing=6)
        self.c_open = TextInput(text=str(DEFAULTS["C_Open"]), multiline=False, hint_text="C_Open")
        self.c_close = TextInput(text=str(DEFAULTS["C_Close"]), multiline=False, hint_text="C_Close")
        self.c_band = TextInput(text=str(DEFAULTS["C_BandPct"]), multiline=False, hint_text="C_BandPct")
        ctl4.add_widget(self.c_open)
        ctl4.add_widget(self.c_close)
        ctl4.add_widget(self.c_band)
        left.add_widget(ctl4)

        # F
        ctl4b = GridLayout(cols=3, size_hint_y=None, height=44, spacing=6)
        self.c_aspect = TextInput(text=str(DEFAULTS["C_AspectTol"]), multiline=False, hint_text="C_AspectTol")
        self.ext_low = TextInput(text=str(DEFAULTS["Ext_Low"]), multiline=False, hint_text="Ext_Low")
        self.ext_high = TextInput(text=str(DEFAULTS["Ext_High"]), multiline=False, hint_text="Ext_High")
        ctl4b.add_widget(self.c_aspect)
        ctl4b.add_widget(self.ext_low)
        ctl4b.add_widget(self.ext_high)
        left.add_widget(ctl4b)

        # G
        ctl4c = GridLayout(cols=3, size_hint_y=None, height=44, spacing=6)
        self.c_fill = TextInput(text=str(DEFAULTS["C_FillMin"]), multiline=False, hint_text="C_FillMin")
        self.c_eps = TextInput(text=str(DEFAULTS["C_Eps"]), multiline=False, hint_text="C_Eps")
        self.use_extent = TextInput(text="0", multiline=False, hint_text="Use_Extent 0/1")
        ctl4c.add_widget(self.c_fill)
        ctl4c.add_widget(self.c_eps)
        ctl4c.add_widget(self.use_extent)
        left.add_widget(ctl4c)

        # -----------------------------
        # Mapping title
        # -----------------------------
        mapping_title = Label(
            text="[b]Mapping[/b]",
            markup=True,
            size_hint_y=None,
            height=24,
            halign="left",
            valign="middle"
        )
        mapping_title.bind(size=self._sync_label_text_size)
        left.add_widget(mapping_title)

        map_row_1 = BoxLayout(size_hint_y=None, height=44, spacing=6)

        self.box_ids_input = TextInput(
            multiline=False,
            hint_text="Box IDs",
            size_hint_x=0.24
        )
        map_row_1.add_widget(self.box_ids_input)

        self.column_spinner = Spinner(
            text="Select Column",
            values=[],
            size_hint_x=0.46
        )
        map_row_1.add_widget(self.column_spinner)

        self.trigger_input = TextInput(
            multiline=False,
            hint_text="Trigger",
            size_hint_x=0.30
        )
        map_row_1.add_widget(self.trigger_input)

        left.add_widget(map_row_1)

        map_row_2 = BoxLayout(size_hint_y=None, height=44, spacing=6)

        self.grid_flag_input = TextInput(
            text="0",
            multiline=False,
            hint_text="Grid 0/1",
            size_hint_x=0.18
        )
        map_row_2.add_widget(self.grid_flag_input)

        self.grid_n_input = TextInput(
            text="1",
            multiline=False,
            hint_text="Grid N",
            size_hint_x=0.18
        )
        map_row_2.add_widget(self.grid_n_input)

        self.btn_assign = Button(
            text="Assign Mapping",
            size_hint_x=0.64
        )
        self.btn_assign.bind(on_release=self.on_assign_mapping)
        map_row_2.add_widget(self.btn_assign)

        left.add_widget(map_row_2)

        # -----------------------------
        # Output title
        # -----------------------------
        output_title = Label(
            text="[b]Output[/b]",
            markup=True,
            size_hint_y=None,
            height=24,
            halign="left",
            valign="middle"
        )
        output_title.bind(size=self._sync_label_text_size)
        left.add_widget(output_title)

        out_row = GridLayout(cols=2, size_hint_y=None, height=46, spacing=6)

        self.btn_generate_one = Button(text="Generate Single PDF")
        self.btn_generate_one.bind(on_release=self.on_generate_single)
        out_row.add_widget(self.btn_generate_one)

        self.btn_generate_batch = Button(text="Generate Batch PDFs")
        self.btn_generate_batch.bind(on_release=self.on_generate_batch)
        out_row.add_widget(self.btn_generate_batch)

        left.add_widget(out_row)

        left_scroll.add_widget(left)
        left_wrap.add_widget(left_scroll)

        # =========================================
        # RIGHT PANEL
        # =========================================
        right = BoxLayout(
            orientation="vertical",
            size_hint_x=0.60,
            spacing=6
        )

        preview_title = Label(
            text="[b]Preview[/b]",
            markup=True,
            size_hint_y=None,
            height=24,
            halign="left",
            valign="middle"
        )
        preview_title.bind(size=self._sync_label_text_size)
        right.add_widget(preview_title)

        preview_wrap = ScrollView(
            do_scroll_x=True,
            do_scroll_y=True,
            bar_width=10,
            scroll_type=["bars", "content"]
        )

        self.preview = Image(
            size_hint=(None, None),
            allow_stretch=False,
            keep_ratio=True
        )
        self.preview.bind(texture=self._update_preview_size)

        preview_wrap.add_widget(self.preview)
        right.add_widget(preview_wrap)

        root.add_widget(left_wrap)
        root.add_widget(right)

        return root
    # --------------------------------------------------------
    # UI helpers
    # --------------------------------------------------------
    def _sync_label_text_size(self, instance, value):
        instance.text_size = value

    def _update_preview_size(self, instance, texture):
        if texture:
            self.preview.size = texture.size

    def set_status(self, text):
        self.status_lbl.text = text

    def get_selected_file(self):
        if not self.file_chooser.selection:
            return None
        return self.file_chooser.selection[0]

    def current_page_idx(self):
        try:
            idx = max(0, int(self.page_input.text.strip()))
        except Exception:
            idx = 0

        total = self.engine.total_pages()
        if total <= 0:
            return 0
        return min(idx, total - 1)

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

    def get_default_file_path(self):
        return "/sdcard/Download" if platform == "android" else os.path.expanduser("~")    
        
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
    def on_load_csv(self, instance):
        content = FileChooserListView(
            filters=["*.csv", "*.xlsx", "*.xls"],
            path=self.get_default_file_path()
        )
        popup = Popup(title="Select CSV/XLSX File", content=content, size_hint=(0.9, 0.9))
        content.bind(on_submit=lambda obj, sel, touch: self._handle_csv_selection(sel, popup))
        popup.open()
    
    def _handle_csv_selection(self, selection, popup):
        if selection:
            try:
                path = selection[0]
                self.engine.load_dataframe(path)
                self.refresh_patient_and_column_lists()
    
                self.set_status(
                    f"Data loaded:\n{os.path.basename(path)}\nRows: {len(self.engine.df)}"
                )
    
                if self.engine.pdf_path:
                    Clock.schedule_once(lambda dt: self.on_preview(None), 0.1)
    
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"Load data error:\n{e}")
        popup.dismiss()


    def on_load_config(self, instance):
        content = FileChooserListView(
            filters=["*.json"],
            path=self.get_default_file_path()
        )
        popup = Popup(title="Select Config File", content=content, size_hint=(0.9, 0.9))
        content.bind(on_submit=lambda obj, sel, touch: self._handle_load_config_selection(sel, popup))
        popup.open()
    
    def _handle_load_config_selection(self, selection, popup):
        if selection:
            try:
                path = selection[0]
                if not path.lower().endswith(".json"):
                    self.set_status("Please select a JSON config file.")
                    return
    
                self.engine.load_config(path)
                self.engine.all_boxes = []
                self.engine.box_types = []
                self.push_engine_settings_to_ui()
                self.set_status(f"Config loaded:\n{os.path.basename(path)}")
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"Load config error:\n{e}")
        popup.dismiss()
    
    
    def on_merge_config(self, instance):
        content = FileChooserListView(
            filters=["*.json"],
            path=self.get_default_file_path()
        )
        popup = Popup(title="Select Config File to Merge", content=content, size_hint=(0.9, 0.9))
        content.bind(on_submit=lambda obj, sel, touch: self._handle_merge_config_selection(sel, popup))
        popup.open()
    
    def _handle_merge_config_selection(self, selection, popup):
        if selection:
            try:
                path = selection[0]
                if not path.lower().endswith(".json"):
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
                self.engine.all_boxes = []
                self.engine.box_types = []
                self.push_engine_settings_to_ui()
                self.set_status(f"Config merged:\n{os.path.basename(path)}")
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"Merge config error:\n{e}")
        popup.dismiss()

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
    
            if self.engine.pdf_path:
                patient = self.selected_patient()
                if patient and self.engine.df is not None and not self.engine.df.empty:
                    self.on_preview(None)
                else:
                    raw_img = self.engine.get_raw_preview_pixmap(
                        page_idx=idx,
                        preview_zoom=PREVIEW_SCALE
                    )
                    self.render_preview_image(raw_img)
                    self.set_status(f"Raw PDF preview.\nPage: {idx}")
        except Exception as e:
            self.set_status(f"Prev page error:\n{e}")
    
    def on_next_page(self, instance):
        try:
            total = self.engine.total_pages()
            idx = min(self.current_page_idx() + 1, max(total - 1, 0))
            self.page_input.text = str(idx)
    
            if self.engine.pdf_path:
                patient = self.selected_patient()
                if patient and self.engine.df is not None and not self.engine.df.empty:
                    self.on_preview(None)
                else:
                    raw_img = self.engine.get_raw_preview_pixmap(
                        page_idx=idx,
                        preview_zoom=PREVIEW_SCALE
                    )
                    self.render_preview_image(raw_img)
                    self.set_status(f"Raw PDF preview.\nPage: {idx}")
        except Exception as e:
            self.set_status(f"Next page error:\n{e}")
    
    def on_patient_change(self, spinner, text):
        if not self.engine.pdf_path:
            return
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
    
            page_idx = self.current_page_idx()
            patient = self.selected_patient()
    
            if not patient or self.engine.df is None or self.engine.df.empty:
                raw_img = self.engine.get_raw_preview_pixmap(
                    page_idx=page_idx,
                    preview_zoom=PREVIEW_SCALE
                )
                self.render_preview_image(raw_img)
                self.set_status(
                    f"Raw PDF preview.\n"
                    f"Page: {page_idx}\n"
                    f"No patient selected yet."
                )
                return
    
            self.apply_ui_settings_to_engine()
    
            img = self.engine.get_preview_pixmap_with_boxes(
                patient_name=patient,
                page_idx=page_idx,
                preview_zoom=PREVIEW_SCALE
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

    def on_load_pdf(self, instance):
        content = FileChooserListView(
            filters=["*.pdf"],
            path=self.get_default_file_path()
        )
        popup = Popup(title="Select PDF Template", content=content, size_hint=(0.9, 0.9))
        content.bind(on_submit=lambda obj, sel, touch: self._handle_pdf_selection(sel, popup))
        popup.open()
    
    def _handle_pdf_selection(self, selection, popup):
        if selection:
            try:
                path = selection[0]
                total = self.engine.load_pdf(path)
    
                cur_idx = self.current_page_idx()
                max_idx = max(total - 1, 0)
                if cur_idx > max_idx:
                    self.page_input.text = "0"
    
                raw_img = self.engine.get_raw_preview_pixmap(
                    page_idx=self.current_page_idx(),
                    preview_zoom=PREVIEW_SCALE
                )
                self.render_preview_image(raw_img)
    
                self.set_status(
                    f"PDF Loaded: {os.path.basename(path)}\n"
                    f"Pages: {total}\n"
                    f"Showing raw template page: {self.current_page_idx()}"
                )
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"PDF Error: {e}")
        popup.dismiss()

    def open_text_input_popup(self, title, hint_text, on_submit_callback, default_text=""):
        wrap = BoxLayout(orientation="vertical", spacing=8, padding=8)
    
        txt = TextInput(
            text=default_text,
            hint_text=hint_text,
            multiline=False,
            size_hint_y=None,
            height=42
        )
        wrap.add_widget(txt)
    
        btn_row = GridLayout(cols=2, size_hint_y=None, height=42, spacing=6)
    
        btn_cancel = Button(text="Cancel")
        btn_ok = Button(text="OK")
    
        btn_row.add_widget(btn_cancel)
        btn_row.add_widget(btn_ok)
        wrap.add_widget(btn_row)
    
        popup = Popup(title=title, content=wrap, size_hint=(0.82, 0.32))
    
        btn_cancel.bind(on_release=lambda *_: popup.dismiss())
    
        def _submit(*_):
            try:
                on_submit_callback(txt.text.strip())
            finally:
                popup.dismiss()
    
        btn_ok.bind(on_release=_submit)
        txt.bind(on_text_validate=lambda *_: _submit())
    
        popup.open()
    
    
    def on_load_gsheet_url(self, instance):
        self.open_text_input_popup(
            title="Load Google Sheet URL",
            hint_text="Paste Google Sheet URL here",
            on_submit_callback=self._handle_gsheet_url_submit
        )
    
    
    def _handle_gsheet_url_submit(self, url):
        try:
            if not url:
                self.set_status("No Google Sheet URL provided.")
                return
    
            self.engine.load_dataframe(url)
            self.refresh_patient_and_column_lists()
    
            self.set_status(
                f"Google Sheet loaded.\n"
                f"Rows: {len(self.engine.df)}\n"
                f"Patients: {len(self.engine.patient_names)}"
            )
    
            if self.engine.pdf_path:
                Clock.schedule_once(lambda dt: self.on_preview(None), 0.1)
    
        except Exception as e:
            traceback.print_exc()
            self.set_status(f"Google Sheet load error:\n{e}")
    
    def on_generate_batch(self, instance):
        """Processes all rows in the data file and generates PDFs."""
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
                    safe_p_name = "".join(c if c.isalnum() or c in (" ", "-", "_") else "_" for c in patient_name).strip()
                    
                    out_path = os.path.join(out_dir, f"Filled_{safe_p_name or 'Unknown'}.pdf")
                    doc.save(out_path)
                    doc.close()
                    success += 1
                except Exception:
                    skipped += 1

            self.set_status(f"Batch done.\nFolder: {out_dir}\nSuccess: {success} | Skipped: {skipped}")
        except Exception as e:
            traceback.print_exc()
            self.set_status(f"Batch Error: {e}")


if __name__ == "__main__":
    MediMapProApp().run()
