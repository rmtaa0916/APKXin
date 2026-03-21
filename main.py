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
try:
    import fitz
except Exception:
    fitz = None
import numpy as np
import pandas as pd

from pypdf import PdfReader, PdfWriter

from kivy.app import App
from kivy.clock import Clock
from kivy.core.image import Image as CoreImage
from kivy.core.text import Label as CoreLabel
from kivy.core.window import Window
from kivy.graphics import Color, Line, Rectangle, RoundedRectangle
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
from kivy.uix.widget import Widget
from kivy.utils import platform

try:
    from kivymd.app import MDApp
    from kivymd.uix.card import MDCard
    from kivymd.uix.button import MDRaisedButton, MDFlatButton, MDRectangleFlatButton
    from kivymd.uix.textfield import MDTextField
    from kivymd.uix.label import MDLabel
    from kivymd.uix.selectioncontrol import MDCheckbox, MDSwitch
    KIVYMD_AVAILABLE = True
except Exception:
    MDApp = App
    MDCard = None
    MDRaisedButton = None
    MDFlatButton = None
    MDRectangleFlatButton = None
    MDTextField = None
    MDLabel = None
    MDCheckbox = None
    MDSwitch = None
    KIVYMD_AVAILABLE = False

FITZ_AVAILABLE = bool(fitz is not None and hasattr(fitz, "open") and hasattr(fitz, "Rect"))

if platform == "android":
    try:
        from jnius import autoclass
        from android import activity
        AndroidPythonActivity = autoclass("org.kivy.android.PythonActivity")
        AndroidIntent = autoclass("android.content.Intent")
        AndroidBitmap = autoclass("android.graphics.Bitmap")
        AndroidBitmapConfig = autoclass("android.graphics.Bitmap$Config")
        AndroidCompressFormat = autoclass("android.graphics.Bitmap$CompressFormat")
        AndroidPdfRenderer = autoclass("android.graphics.pdf.PdfRenderer")
        AndroidParcelFileDescriptor = autoclass("android.os.ParcelFileDescriptor")
        AndroidUri = autoclass("android.net.Uri")
        AndroidByteArrayOutputStream = autoclass("java.io.ByteArrayOutputStream")
        ANDROID_JAVA_AVAILABLE = True
    except Exception:
        AndroidPythonActivity = None
        AndroidIntent = None
        AndroidBitmap = None
        AndroidPdfRenderer = None
        AndroidParcelFileDescriptor = None
        AndroidUri = None
        AndroidByteArrayOutputStream = None
        AndroidBitmapConfig = None
        AndroidCompressFormat = None
        ANDROID_JAVA_AVAILABLE = False
else:
    ANDROID_JAVA_AVAILABLE = False


def android_render_pdf_page(path, page_idx=0, preview_zoom=1.5):
    """Render a PDF page on Android using the native PdfRenderer API."""
    if platform != "android" or not ANDROID_JAVA_AVAILABLE:
        raise RuntimeError("Android PdfRenderer is unavailable.")

    pfd = None
    renderer = None
    page = None
    try:
        file_obj = autoclass("java.io.File")(path)
        pfd = AndroidParcelFileDescriptor.open(file_obj, AndroidParcelFileDescriptor.MODE_READ_ONLY)
        renderer = AndroidPdfRenderer(pfd)
        total = renderer.getPageCount()
        if total <= 0:
            raise ValueError("PDF has no pages.")
        page_idx = max(0, min(int(page_idx), total - 1))
        page = renderer.openPage(page_idx)

        width = max(1, int(page.getWidth() * float(preview_zoom)))
        height = max(1, int(page.getHeight() * float(preview_zoom)))
        bitmap = AndroidBitmap.createBitmap(width, height, AndroidBitmapConfig.ARGB_8888)
        bitmap.eraseColor(-1)
        page.render(bitmap, None, None, AndroidPdfRenderer.Page.RENDER_MODE_FOR_DISPLAY)

        baos = AndroidByteArrayOutputStream()
        bitmap.compress(AndroidCompressFormat.PNG, 100, baos)
        png_bytes = bytes(baos.toByteArray())
        img = cv2.imdecode(np.frombuffer(png_bytes, np.uint8), cv2.IMREAD_COLOR)
        if img is None:
            raise ValueError("Android PdfRenderer returned an unreadable image.")
        return img
    finally:
        try:
            if page is not None:
                page.close()
        except Exception:
            pass
        try:
            if renderer is not None:
                renderer.close()
        except Exception:
            pass
        try:
            if pfd is not None:
                pfd.close()
        except Exception:
            pass


# ============================================================
# Defaults
# ============================================================
APP_TITLE = "MediMap Pro: Intelligent Form Automator"
CONFIG_FILENAME = "medimap_config.json"
if platform == "android":
    ZOOM = 2.4
    PREVIEW_SCALE = 1.35
else:
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
            total = self._pdf_page_count(path)
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
            return max(self._pdf_page_count(self.pdf_path), 1)
        except Exception:
            return 1

    def _pdf_page_count(self, path):
        """Return page count using the most reliable backend available."""
        if not path or not os.path.exists(path):
            raise FileNotFoundError(f"PDF file not found: {path}")

        # Prefer pure-Python counting first so Android does not depend on PyMuPDF
        try:
            with open(path, "rb") as f:
                reader = PdfReader(f)
                total = len(reader.pages)
            if total > 0:
                return total
        except Exception:
            pass

        if FITZ_AVAILABLE:
            with fitz.open(path) as doc:
                return len(doc)

        if platform == "android":
            if ANDROID_JAVA_AVAILABLE:
                pfd = None
                renderer = None
                try:
                    file_obj = autoclass("java.io.File")(path)
                    pfd = AndroidParcelFileDescriptor.open(file_obj, AndroidParcelFileDescriptor.MODE_READ_ONLY)
                    renderer = AndroidPdfRenderer(pfd)
                    return renderer.getPageCount()
                finally:
                    try:
                        if renderer is not None:
                            renderer.close()
                    except Exception:
                        pass
                    try:
                        if pfd is not None:
                            pfd.close()
                    except Exception:
                        pass

        raise RuntimeError("No PDF page-count backend is available.")

    def _render_pdf_page_bgr(self, path, page_idx=0, preview_zoom=1.5):
        """Render a PDF page to an OpenCV BGR image."""
        if not path or not os.path.exists(path):
            raise FileNotFoundError(f"PDF file not found: {path}")

        if FITZ_AVAILABLE:
            with fitz.open(path) as doc:
                total = len(doc)
                if total <= 0:
                    raise ValueError("PDF has no pages.")
                page_idx = max(0, min(int(page_idx), total - 1))
                page = doc.load_page(page_idx)
                mat = fitz.Matrix(float(preview_zoom), float(preview_zoom))
                pix = page.get_pixmap(matrix=mat, alpha=False)
                img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                if pix.n == 4:
                    img = cv2.cvtColor(img, cv2.COLOR_BGRA2BGR)
                else:
                    img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
                return img

        if platform == "android":
            return android_render_pdf_page(path, page_idx=page_idx, preview_zoom=preview_zoom)

        raise RuntimeError("No PDF rendering backend is available.")


    def supports_processing_backend(self):
        """
        Return True when full PDF processing is available.
        Full detection/fill/export currently depends on PyMuPDF (fitz).
        """
        return bool(FITZ_AVAILABLE)

    def get_raw_preview_pixmap(self, page_idx=0, preview_zoom=1.5):
        """Render the raw loaded PDF page without filling or boxes."""
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            raise FileNotFoundError("PDF path is missing or invalid.")
        return self._render_pdf_page_bgr(self.pdf_path, page_idx=page_idx, preview_zoom=preview_zoom)



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
        if not self.supports_processing_backend():
            raise RuntimeError("Android native preview is available, but detection still requires the PyMuPDF backend in this build.")

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
                    self._append_box_unique(
                        fitz.Rect(fx / ZOOM, fy / ZOOM, (fx + fw) / ZOOM, (fy + fh) / ZOOM),
                        "check",
                        iou_thresh=0.45
                    )

            elif max(w, h) <= int(red_roi_max_val) and min(w, h) >= checkbox_min_sz:
                for fx, fy, fw, fh in self.find_checkbox_rects_in_roi(bin_checks, x, y, w, h):
                    self._append_box_unique(
                        fitz.Rect(fx / ZOOM, fy / ZOOM, (fx + fw) / ZOOM, (fy + fh) / ZOOM),
                        "check",
                        iou_thresh=0.45
                    )

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
        if not self.supports_processing_backend():
            raise RuntimeError("Android native preview is available, but PDF filling/export still requires the PyMuPDF backend in this build.")
        if self.df is None or self.df.empty:
            raise ValueError("DataFrame is empty.")

        matches = self.df[self.df["_DISPLAY_NAME"] == patient_name]
        if matches.empty:
            raise ValueError(f"Patient not found: {patient_name}")
        if len(matches) > 1:
            raise ValueError(f"Multiple records found for '{patient_name}'. Please make the display name unique.")

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

    def get_processed_preview_pixmap(self, patient_name, page_idx=0, preview_zoom=1.5):
        if not self.all_boxes:
            self.run_detection(page_idx=page_idx)

        doc = self.process_doc(patient_name, page_idx=page_idx)
        page = doc[page_idx]
        pix = page.get_pixmap(matrix=fitz.Matrix(preview_zoom, preview_zoom))
        img = cv2.imdecode(np.frombuffer(pix.tobytes("png"), np.uint8), cv2.IMREAD_COLOR)

        doc.close()
        if img is None:
            raise ValueError("Failed to build preview image.")
        return img

    def get_preview_pixmap_with_boxes(self, patient_name, page_idx=0, preview_zoom=1.5):
        img = self.get_processed_preview_pixmap(patient_name, page_idx=page_idx, preview_zoom=preview_zoom)

        for i, r in enumerate(self.all_boxes):
            x0 = int(r.x0 * preview_zoom)
            y0 = int(r.y0 * preview_zoom)
            x1 = int(r.x1 * preview_zoom)
            y1 = int(r.y1 * preview_zoom)

            box_type = self.box_types[i]
            if box_type == "check":
                color = (0, 0, 255)
            elif box_type == "line":
                color = (0, 215, 255)
            else:
                color = (0, 255, 0)

            cv2.rectangle(img, (x0, y0), (x1, y1), color, 2)
            label = str(i)
            cv2.rectangle(img, (x0, max(0, y0 - 14)), (x0 + 20, y0), (0, 0, 0), -1)
            cv2.putText(img, label, (x0 + 2, y0 - 3), cv2.FONT_HERSHEY_SIMPLEX, 0.35, (255, 255, 255), 1, cv2.LINE_AA)

        return img

# =========================
# PART 4 / 4
# Kivy App UI + Wiring
# Paste this BELOW Part 3
# =========================

class InteractivePreview(Image):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.allow_stretch = False
        self.keep_ratio = True
        self.size = (dp(1), dp(1))
        self.boxes_payload = []
        self.selected_ids = set()
        self.hovered_box_id = None
        self.preview_zoom = 1.0
        self.page_idx = 0
        self.box_tap_callback = None
        self.hover_callback = None
        self._mouse_bound = False
        self.bind(texture=self._redraw_overlay, size=self._redraw_overlay, pos=self._redraw_overlay)
        if platform != "android":
            Window.bind(mouse_pos=self._on_mouse_pos)
            self._mouse_bound = True

    def on_parent(self, *args):
        self._redraw_overlay()

    def set_texture_from_bgr(self, img_bgr):
        if img_bgr is None:
            self.texture = None
            self.canvas.after.clear()
            return
        rgba = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGBA)
        texture = Texture.create(size=(rgba.shape[1], rgba.shape[0]), colorfmt="rgba")
        texture.blit_buffer(rgba.tobytes(), colorfmt="rgba", bufferfmt="ubyte")
        texture.flip_vertical()
        self.texture = texture
        self._redraw_overlay()

    def set_boxes_payload(self, boxes_payload, selected_ids=None, preview_zoom=1.0, page_idx=0):
        self.boxes_payload = list(boxes_payload or [])
        self.preview_zoom = float(preview_zoom or 1.0)
        self.page_idx = int(page_idx or 0)
        self.selected_ids = set(selected_ids or [])
        self.hovered_box_id = None
        self._redraw_overlay()

    def set_selected_ids(self, selected_ids):
        self.selected_ids = set(selected_ids or [])
        self._redraw_overlay()

    def clear_boxes(self):
        self.boxes_payload = []
        self.selected_ids = set()
        self.hovered_box_id = None
        self._redraw_overlay()

    def _get_display_rect(self):
        tex = self.texture
        if not tex or tex.width <= 0 or tex.height <= 0:
            return None
        widget_w, widget_h = self.width, self.height
        tex_w, tex_h = tex.width, tex.height
        scale = min(widget_w / float(tex_w), widget_h / float(tex_h))
        draw_w = tex_w * scale
        draw_h = tex_h * scale
        x = self.x + (widget_w - draw_w) / 2.0
        y = self.y + (widget_h - draw_h) / 2.0
        return x, y, draw_w, draw_h

    def _box_color(self, box_type, selected=False, hovered=False):
        if selected:
            return (0.15, 0.85, 1.0, 1.0)
        if hovered:
            return (1.0, 0.6, 0.0, 1.0)
        if box_type == "check":
            return (1.0, 0.15, 0.15, 1.0)
        if box_type == "line":
            return (1.0, 0.85, 0.1, 1.0)
        return (0.1, 1.0, 0.3, 1.0)

    def _draw_label(self, x, y, text):
        core = CoreLabel(text=str(text), font_size=12)
        core.refresh()
        tex = core.texture
        pad_x, pad_y = 4, 2
        bg_w = tex.width + pad_x * 2
        bg_h = tex.height + pad_y * 2
        Rectangle(pos=(x, y - bg_h), size=(bg_w, bg_h))
        Color(1, 1, 1, 1)
        Rectangle(texture=tex, pos=(x + pad_x, y - bg_h + pad_y), size=tex.size)

    def _redraw_overlay(self, *args):
        self.canvas.after.clear()
        disp = self._get_display_rect()
        if not disp:
            return
        dx, dy, dw, dh = disp
        tex = self.texture
        sx = dw / float(max(tex.width, 1))
        sy = dh / float(max(tex.height, 1))
        with self.canvas.after:
            for box in self.boxes_payload:
                x = dx + box["x"] * sx
                y = dy + (tex.height - (box["y"] + box["h"])) * sy
                w = max(1.0, box["w"] * sx)
                h = max(1.0, box["h"] * sy)
                selected = box["id"] in self.selected_ids
                hovered = (box["id"] == self.hovered_box_id)
                Color(*self._box_color(box.get("t", "field"), selected=selected, hovered=hovered))
                Line(rectangle=(x, y, w, h), width=3.2 if selected else (2.4 if hovered else 1.35))
                Color(0, 0, 0, 0.88)
                self._draw_label(x, y + h, box["id"])

    def _touch_to_image_point(self, touch):
        disp = self._get_display_rect()
        tex = self.texture
        if not disp or not tex:
            return None
        dx, dy, dw, dh = disp
        if not (dx <= touch.x <= dx + dw and dy <= touch.y <= dy + dh):
            return None
        img_x = (touch.x - dx) * tex.width / float(max(dw, 1))
        img_y = tex.height - ((touch.y - dy) * tex.height / float(max(dh, 1)))
        return img_x, img_y

    def _find_hit_box(self, img_x, img_y):
        hits = []
        for box in self.boxes_payload:
            if box["x"] <= img_x <= box["x"] + box["w"] and box["y"] <= img_y <= box["y"] + box["h"]:
                hits.append(box)
        if not hits:
            return None
        hits.sort(key=lambda b: (b["id"], b["w"] * b["h"]))
        return hits[0]

    def on_touch_down(self, touch):
        if not self.collide_point(*touch.pos):
            return super().on_touch_down(touch)
        pt = self._touch_to_image_point(touch)
        if pt is None:
            return super().on_touch_down(touch)
        hit = self._find_hit_box(*pt)
        if hit is not None:
            if callable(self.box_tap_callback):
                self.box_tap_callback(hit)
            return True
        return super().on_touch_down(touch)

    def _on_mouse_pos(self, _window, pos):
        if platform == "android" or not self.get_root_window():
            return
        local = self.to_widget(*pos)
        class Dummy: pass
        d = Dummy(); d.x=local[0]; d.y=local[1]
        pt = self._touch_to_image_point(d)
        new_hover = None
        if pt is not None:
            hit = self._find_hit_box(*pt)
            if hit is not None:
                new_hover = hit["id"]
                if callable(self.hover_callback):
                    self.hover_callback(hit)
        if new_hover != self.hovered_box_id:
            self.hovered_box_id = new_hover
            self._redraw_overlay()


class MediMapProLayout(BoxLayout):
    pass


class MediMapProApp(MDApp):

    def build(self):
        self.title = "MediMap Pro"
        self.engine = MediMapEngine()

        if KIVYMD_AVAILABLE and hasattr(self, "theme_cls"):
            try:
                self.theme_cls.theme_style = "Dark"
                self.theme_cls.primary_palette = "Blue"
                self.theme_cls.accent_palette = "Teal"
                if hasattr(self.theme_cls, "material_style"):
                    self.theme_cls.material_style = "M3"
            except Exception:
                pass

        is_mobile = (platform == "android")
        pad = dp(12) if is_mobile else dp(14)
        gap = dp(10) if is_mobile else dp(12)
        row_h = dp(52) if is_mobile else dp(46)
        input_h = dp(52) if is_mobile else dp(44)

        palette = {
            "bg": (0.055, 0.075, 0.11, 1),
            "surface": (0.085, 0.105, 0.145, 1),
            "surface_alt": (0.11, 0.135, 0.185, 1),
            "surface_soft": (0.135, 0.16, 0.215, 1),
            "primary": (0.24, 0.48, 0.96, 1),
            "primary_soft": (0.16, 0.24, 0.38, 1),
            "accent": (0.10, 0.78, 0.63, 1),
            "accent_soft": (0.10, 0.28, 0.24, 1),
            "text": (0.93, 0.96, 1.0, 1),
            "muted": (0.60, 0.68, 0.80, 1),
            "border": (0.18, 0.24, 0.34, 1),
            "preview_bg": (0.06, 0.08, 0.11, 1),
            "chip": (0.13, 0.18, 0.27, 1),
            "danger": (0.87, 0.33, 0.41, 1),
        }
        Window.clearcolor = palette["bg"]

        def style_card(widget, color=None, radius=dp(22)):
            color = color or palette["surface"]
            with widget.canvas.before:
                widget._card_bg_color = Color(*color)
                widget._card_bg = RoundedRectangle(radius=[radius] * 4)
                widget._card_border_color = Color(*palette["border"])
                widget._card_border = Line(rounded_rectangle=[0, 0, 0, 0, radius], width=1.1)
            def _upd(*_):
                widget._card_bg.pos = widget.pos
                widget._card_bg.size = widget.size
                widget._card_border.rounded_rectangle = [widget.x, widget.y, widget.width, widget.height, radius]
            widget.bind(pos=_upd, size=_upd)
            _upd()
            return widget

        def make_card(title, subtitle=None):
            if KIVYMD_AVAILABLE and MDCard is not None:
                card = MDCard(
                    orientation="vertical",
                    padding=dp(14),
                    spacing=dp(10),
                    radius=[dp(24)] * 4,
                    elevation=4,
                    size_hint_y=None,
                    md_bg_color=palette["surface"],
                )
                card.bind(minimum_height=card.setter("height"))

                head = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(2))
                head.height = dp(48) if subtitle else dp(28)
                ttl_cls = MDLabel if MDLabel is not None else Label
                ttl = ttl_cls(
                    text=title,
                    theme_text_color="Custom" if MDLabel is not None else None,
                    text_color=palette["text"] if MDLabel is not None else None,
                    size_hint_y=None,
                    height=dp(24),
                    halign="left",
                    valign="middle",
                    font_style="Subtitle1" if MDLabel is not None else None,
                )
                if hasattr(ttl, "bind") and not MDLabel is ttl_cls:
                    ttl.bind(size=self._sync_label_text_size)
                head.add_widget(ttl)
                if subtitle:
                    sub = ttl_cls(
                        text=subtitle,
                        theme_text_color="Custom" if MDLabel is not None else None,
                        text_color=palette["muted"] if MDLabel is not None else None,
                        size_hint_y=None,
                        height=dp(18),
                        halign="left",
                        valign="middle",
                        font_style="Caption" if MDLabel is not None else None,
                    )
                    if hasattr(sub, "bind") and not MDLabel is ttl_cls:
                        sub.bind(size=self._sync_label_text_size)
                    head.add_widget(sub)
                card.add_widget(head)

                body = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
                body.bind(minimum_height=body.setter("height"))
                card.add_widget(body)
                return card, body

            wrap = BoxLayout(orientation="vertical", spacing=dp(8), padding=dp(12), size_hint_y=None)
            wrap.bind(minimum_height=wrap.setter("height"))
            style_card(wrap, palette["surface"])

            head = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(2))
            head.height = dp(44) if subtitle else dp(28)
            ttl = Label(
                text=title, markup=False, bold=True, color=palette["text"],
                size_hint_y=None, height=dp(24), halign="left", valign="middle",
                font_size=dp(16) if is_mobile else dp(15)
            )
            ttl.bind(size=self._sync_label_text_size)
            head.add_widget(ttl)
            if subtitle:
                sub = Label(
                    text=subtitle, color=palette["muted"],
                    size_hint_y=None, height=dp(18), halign="left", valign="middle",
                    font_size=dp(12)
                )
                sub.bind(size=self._sync_label_text_size)
                head.add_widget(sub)
            wrap.add_widget(head)

            body = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
            body.bind(minimum_height=body.setter("height"))
            wrap.add_widget(body)
            return wrap, body

        def make_button(text, tone="primary", icon=None):
            colors = {
                "primary": palette["primary"],
                "soft": palette["primary_soft"],
                "accent": palette["accent"],
                "plain": palette["surface_soft"],
            }
            label = text if not icon else f"{icon}  {text}"
            if KIVYMD_AVAILABLE and MDRaisedButton is not None:
                if tone in ("primary", "accent"):
                    btn = MDRaisedButton(
                        text=label,
                        size_hint_y=None,
                        height=row_h,
                        md_bg_color=colors.get(tone, palette["primary"]),
                        text_color=(1, 1, 1, 1),
                        elevation=1,
                    )
                elif tone == "plain":
                    btn = MDRectangleFlatButton(
                        text=label,
                        size_hint_y=None,
                        height=row_h,
                        line_color=palette["border"],
                        text_color=palette["text"],
                    )
                else:
                    btn = MDFlatButton(
                        text=label,
                        size_hint_y=None,
                        height=row_h,
                        theme_text_color="Custom",
                        text_color=palette["text"],
                    )
                return btn

            txt_color = (1,1,1,1) if tone in ("primary","accent") else palette["text"]
            btn = Button(
                text=label,
                size_hint_y=None, height=row_h,
                background_normal="", background_down="",
                background_color=colors.get(tone, palette["primary"]),
                color=txt_color, font_size=dp(14) if is_mobile else dp(13),
            )
            return btn

        def make_input(default="", hint="", input_filter=None):
            if KIVYMD_AVAILABLE and MDTextField is not None:
                field = MDTextField(
                    text=str(default),
                    hint_text=hint,
                    mode="fill",
                    size_hint_y=None,
                    height=input_h + dp(8),
                    helper_text_mode="on_focus",
                )
                try:
                    field.fill_color_normal = palette["surface_alt"]
                    field.fill_color_focus = palette["surface_alt"]
                    field.line_color_normal = palette["border"]
                    field.line_color_focus = palette["primary"]
                    field.text_color_normal = palette["text"]
                    field.text_color_focus = palette["text"]
                    field.hint_text_color_normal = palette["muted"]
                    field.hint_text_color_focus = palette["primary"]
                except Exception:
                    pass
                if input_filter is not None:
                    try:
                        field.input_filter = input_filter
                    except Exception:
                        pass
                return field

            kwargs = dict(text=str(default), hint_text=hint, multiline=False,
                          size_hint_y=None, height=input_h, font_size=dp(14), write_tab=False,
                          background_normal="", background_active="",
                          background_color=palette["surface_alt"], foreground_color=palette["text"],
                          cursor_color=palette["primary"], padding=[dp(12), dp(14), dp(12), dp(14)])
            if input_filter is not None:
                kwargs["input_filter"] = input_filter
            return TextInput(**kwargs)

        def make_spinner(text, values=()):
            return Spinner(text=text, values=list(values), size_hint_y=None, height=input_h,
                           font_size=dp(14), background_normal="", background_color=palette["surface_alt"],
                           color=palette["text"])

        def labeled_field(label_text, widget, helper=None):
            box = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(4))
            extra = dp(18) if helper else 0
            box.height = dp(24) + widget.height + extra
            lbl = Label(text=label_text, color=palette["text"], size_hint_y=None,
                        height=dp(20), halign="left", valign="middle", font_size=dp(12))
            lbl.bind(size=self._sync_label_text_size)
            box.add_widget(lbl)
            box.add_widget(widget)
            if helper:
                hlp = Label(text=helper, color=palette["muted"], size_hint_y=None,
                            height=dp(16), halign="left", valign="middle", font_size=dp(10))
                hlp.bind(size=self._sync_label_text_size)
                box.add_widget(hlp)
            return box

        def labeled_checkbox(label_text, checkbox, helper=None):
            row = BoxLayout(orientation="horizontal", size_hint_y=None, height=dp(34), spacing=dp(10))
            checkbox.size_hint = (None, None)
            checkbox.size = (dp(36), dp(24))
            row.add_widget(checkbox)
            lbl_cls = MDLabel if KIVYMD_AVAILABLE and MDLabel is not None else Label
            lbl_kwargs = dict(text=label_text, halign="left", valign="middle")
            if lbl_cls is Label:
                lbl_kwargs.update(color=palette["text"], font_size=dp(13))
            else:
                lbl_kwargs.update(theme_text_color="Custom", text_color=palette["text"], font_style="Body1")
            lbl = lbl_cls(**lbl_kwargs)
            if lbl_cls is Label:
                lbl.bind(size=self._sync_label_text_size)
            row.add_widget(lbl)
            if helper:
                wrap = BoxLayout(orientation="vertical", size_hint_y=None, spacing=dp(2))
                wrap.height = dp(54)
                wrap.add_widget(row)
                sub = Label(text=helper, color=palette["muted"], size_hint_y=None, height=dp(14), halign="left", valign="middle", font_size=dp(10))
                sub.bind(size=self._sync_label_text_size)
                wrap.add_widget(sub)
                return wrap
            return row

        root = BoxLayout(orientation="vertical", spacing=gap, padding=pad)

        appbar = BoxLayout(
            orientation="vertical" if is_mobile else "horizontal",
            spacing=dp(10) if is_mobile else dp(14),
            size_hint_y=None,
            height=dp(108) if is_mobile else dp(74),
            padding=[dp(14), dp(10), dp(14), dp(10)],
        )
        style_card(appbar, palette["surface"], radius=dp(26))

        brand_wrap = BoxLayout(orientation="vertical", spacing=dp(2))
        hero_title = Label(
            text="MediMap Pro",
            color=palette["text"],
            bold=True,
            size_hint_y=None,
            height=dp(28),
            halign="left",
            valign="middle",
            font_size=dp(24) if is_mobile else dp(20),
        )
        hero_title.bind(size=self._sync_label_text_size)
        brand_wrap.add_widget(hero_title)

        hero_sub = Label(
            text="Medical-tech PDF automation with interactive mapping and Android-native fallback preview.",
            color=palette["muted"],
            size_hint_y=None,
            height=dp(20),
            halign="left",
            valign="middle",
            font_size=dp(11.5),
        )
        hero_sub.bind(size=self._sync_label_text_size)
        brand_wrap.add_widget(hero_sub)
        appbar.add_widget(brand_wrap)

        status_chip = Label(
            text="Ready • Load CSV/XLSX, Google Sheet URL, and PDF to begin",
            color=palette["text"],
            size_hint=(1 if is_mobile else 0.42, None),
            height=dp(36) if is_mobile else dp(42),
            halign="center",
            valign="middle",
            font_size=dp(10.5) if is_mobile else dp(11),
        )
        status_chip.bind(size=self._sync_label_text_size)
        style_card(status_chip, palette["chip"], radius=dp(18))
        self.status_lbl = status_chip
        appbar.add_widget(status_chip)
        root.add_widget(appbar)

        main = BoxLayout(orientation="vertical" if is_mobile else "horizontal", spacing=dp(10) if is_mobile else dp(14))

        self._mobile_section_buttons = {}
        self._mobile_section_cards = {}
        self._mobile_section_host = None
        self._mobile_active_section = None

        def _style_mobile_section_button(btn, active=False):
            bg = palette["primary"] if active else palette["surface_soft"]
            fg = (1, 1, 1, 1) if active else palette["text"]
            if hasattr(btn, "md_bg_color"):
                try:
                    btn.md_bg_color = bg
                except Exception:
                    pass
                try:
                    btn.text_color = fg
                except Exception:
                    pass
            else:
                btn.background_normal = ""
                btn.background_down = ""
                btn.background_color = bg
                btn.color = fg

        def _show_mobile_section(section_key):
            if not is_mobile or self._mobile_section_host is None:
                return
            section_card = self._mobile_section_cards.get(section_key)
            if section_card is None:
                return
            self._mobile_section_host.clear_widgets()
            self._mobile_section_host.add_widget(section_card)
            self._mobile_active_section = section_key
            for key, btn in self._mobile_section_buttons.items():
                _style_mobile_section_button(btn, active=(key == section_key))

        def _make_mobile_section_tabs(section_names):
            tabs_wrap = BoxLayout(orientation="vertical", spacing=dp(8), size_hint_y=None)
            rows = []
            if len(section_names) <= 3:
                rows = [section_names]
            else:
                rows = [section_names[:3], section_names[3:]]
            total_h = len(rows) * row_h + max(len(rows) - 1, 0) * dp(8)
            tabs_wrap.height = total_h

            for row_names in rows:
                row = GridLayout(cols=max(1, len(row_names)), spacing=dp(8), size_hint_y=None, height=row_h)
                for name in row_names:
                    btn = Button(
                        text=name,
                        size_hint_y=None,
                        height=row_h,
                        background_normal="",
                        background_down="",
                        background_color=palette["surface_soft"],
                        color=palette["text"],
                        font_size=dp(13),
                    )
                    btn.bind(on_release=lambda inst, n=name: _show_mobile_section(n))
                    self._mobile_section_buttons[name] = btn
                    row.add_widget(btn)
                tabs_wrap.add_widget(row)
            return tabs_wrap

        controls_wrap = BoxLayout(orientation="vertical", size_hint=(1, 0.56) if is_mobile else (0.44, 1))
        controls_scroll = ScrollView(do_scroll_x=False, do_scroll_y=True, bar_width=dp(6), scroll_type=["bars", "content"])
        controls = GridLayout(cols=1, spacing=gap, size_hint_y=None)
        controls.bind(minimum_height=controls.setter("height"))

        files_card, files_body = make_card("Workspace", "Open data, templates, and saved presets")
        file_grid = GridLayout(cols=2 if is_mobile else 3, spacing=dp(8), size_hint_y=None)
        file_buttons = [
            ("Load PDF", self.on_load_pdf, "primary"),
            ("Load CSV/XLSX", self.on_load_csv, "soft"),
            ("Load GSheet", self.on_load_gsheet_url, "soft"),
            ("Load Config", self.on_load_config, "plain"),
            ("Merge Mappings", self.on_merge_config, "plain"),
            ("Save Config", self.on_save_config, "accent"),
        ]
        rows = (len(file_buttons) + file_grid.cols - 1) // file_grid.cols
        file_grid.height = rows * row_h + max(rows-1,0) * dp(8)
        for txt, cb, tone in file_buttons:
            btn = make_button(txt, tone=tone)
            btn.bind(on_release=cb)
            file_grid.add_widget(btn)
        files_body.add_widget(file_grid)
        if is_mobile:
            self._mobile_section_cards["Files"] = files_card
        else:
            controls.add_widget(files_card)

        nav_card, nav_body = make_card("Session", "Choose a patient, move pages, detect, and refresh")
        self.patient_spinner = make_spinner("Select Patient")
        self.patient_spinner.bind(text=self.on_patient_change)
        nav_body.add_widget(labeled_field("Patient", self.patient_spinner, "Active data row for preview and export"))
        nav_grid = GridLayout(cols=3, spacing=dp(8), size_hint_y=None, height=row_h)
        self.page_input = make_input("0", "Page", input_filter="int")
        self.btn_prev = make_button("Prev", tone="plain")
        self.btn_next = make_button("Next", tone="plain")
        self.btn_prev.bind(on_release=self.on_prev_page)
        self.btn_next.bind(on_release=self.on_next_page)
        nav_grid.add_widget(self.page_input)
        nav_grid.add_widget(self.btn_prev)
        nav_grid.add_widget(self.btn_next)
        nav_body.add_widget(labeled_field("Page", nav_grid, "Preview uses zero-based page index"))
        action_grid = GridLayout(cols=2, spacing=dp(8), size_hint_y=None, height=row_h)
        self.btn_detect = make_button("Run Detect", tone="accent")
        self.btn_detect.bind(on_release=self.on_run_detect)
        self.btn_preview = make_button("Refresh Preview", tone="primary")
        self.btn_preview.bind(on_release=self.on_preview)
        action_grid.add_widget(self.btn_detect)
        action_grid.add_widget(self.btn_preview)
        nav_body.add_widget(action_grid)
        if is_mobile:
            self._mobile_section_cards["Session"] = nav_card
        else:
            controls.add_widget(nav_card)

        detect_card, detect_body = make_card("Detection", "Field, line, checkbox, and extent thresholds")
        explain = Label(text="F = field sizing  •  Line = answer lines  •  C = checkbox tuning  •  Ext = contour extent limits",
                        color=palette["muted"], size_hint_y=None, height=dp(18), halign="left", valign="middle", font_size=dp(11))
        explain.bind(size=self._sync_label_text_size)
        detect_body.add_widget(explain)

        self.f_area = make_input(DEFAULTS["F_Area"], "500", input_filter="int")
        self.f_minw = make_input(DEFAULTS["F_MinW"], "15", input_filter="int")
        self.f_minh = make_input(DEFAULTS["F_MinH"], "35", input_filter="int")
        self.f_close = make_input(DEFAULTS["F_Close"], "1", input_filter="int")
        self.line_minw = make_input(DEFAULTS["Line_MinW"], "100", input_filter="int")
        self.line_maxw = make_input(DEFAULTS["Line_MaxW"], "1680", input_filter="int")
        self.c_strict = make_input(DEFAULTS["C_Strict"], "40", input_filter="int")
        self.c_size_min = make_input(DEFAULTS["C_Size"][0], "14", input_filter="int")
        self.c_size_max = make_input(DEFAULTS["C_Size"][1], "65", input_filter="int")
        self.c_border = make_input(DEFAULTS["C_Border"], "0.10")
        self.c_inner = make_input(DEFAULTS["C_Inner"], "0.40")
        self.roi_max = make_input(DEFAULTS["ROI_Max"], "200", input_filter="int")
        self.c_open = make_input(DEFAULTS["C_Open"], "1", input_filter="int")
        self.c_close = make_input(DEFAULTS["C_Close"], "0", input_filter="int")
        self.c_band = make_input(DEFAULTS["C_BandPct"], "0.18")
        self.c_aspect = make_input(DEFAULTS["C_AspectTol"], "0.10")
        self.ext_low = make_input(DEFAULTS["Ext_Low"], "0.04")
        self.ext_high = make_input(DEFAULTS["Ext_High"], "0.60")
        self.c_fill = make_input(DEFAULTS["C_FillMin"], "0.45")
        self.c_eps = make_input(DEFAULTS["C_Eps"], "0.04")
        self.use_extent_chk = (MDSwitch(active=bool(DEFAULTS["Use_Extent"])) if KIVYMD_AVAILABLE and MDSwitch is not None else CheckBox(active=bool(DEFAULTS["Use_Extent"])))

        detection_fields = [
            labeled_field("Field Area", self.f_area, "Minimum contour area kept as a field"),
            labeled_field("Field Min Width", self.f_minw),
            labeled_field("Field Min Height", self.f_minh),
            labeled_field("Field Close", self.f_close, "Morph close kernel for field cleanup"),
            labeled_field("Line Min Width", self.line_minw),
            labeled_field("Line Max Width", self.line_maxw),
            labeled_field("Checkbox Strict", self.c_strict, "Higher values tighten checkbox rules"),
            labeled_field("Checkbox Min Size", self.c_size_min),
            labeled_field("Checkbox Max Size", self.c_size_max),
            labeled_field("Border Fill", self.c_border),
            labeled_field("Inner Fill", self.c_inner),
            labeled_field("ROI Max", self.roi_max),
            labeled_field("Checkbox Open", self.c_open),
            labeled_field("Checkbox Close", self.c_close),
            labeled_field("Band %", self.c_band),
            labeled_field("Aspect Tolerance", self.c_aspect),
            labeled_field("Extent Low", self.ext_low),
            labeled_field("Extent High", self.ext_high),
            labeled_field("Fill Min", self.c_fill),
            labeled_field("Approx Eps", self.c_eps),
            labeled_checkbox("Use Extent", self.use_extent_chk, "Use contour extent as an extra checkbox filter"),
        ]
        det_grid = GridLayout(cols=2 if is_mobile else 3, spacing=dp(8), size_hint_y=None)
        rows = (len(detection_fields)+det_grid.cols-1)//det_grid.cols
        max_cell_h = max(w.height for w in detection_fields)
        det_grid.height = rows * max_cell_h + max(rows-1,0)*dp(8)
        for w in detection_fields:
            det_grid.add_widget(w)
        detect_body.add_widget(det_grid)
        if is_mobile:
            self._mobile_section_cards["Detection"] = detect_card
        else:
            controls.add_widget(detect_card)

        map_card, map_body = make_card("Mapping", "Assign values to selected boxes directly from preview")
        self.box_ids_input = make_input("", "0,1,2")
        self.column_spinner = make_spinner("Select Column")
        self.trigger_input = make_input("", "Trigger")
        self.grid_flag_chk = (MDSwitch(active=False) if KIVYMD_AVAILABLE and MDSwitch is not None else CheckBox(active=False))
        self.grid_n_input = make_input("1", "1", input_filter="int")
        map_body.add_widget(labeled_field("Selected Box IDs", self.box_ids_input, "Click boxes in preview to add or remove them"))
        map_body.add_widget(labeled_field("Column", self.column_spinner, "Data column written into the selected target"))
        map_body.add_widget(labeled_field("Trigger", self.trigger_input, "Leave blank for text fill, use value for checkbox X mark"))
        map_opts = GridLayout(cols=2, spacing=dp(8), size_hint_y=None, height=max(dp(48), self.grid_n_input.height))
        map_opts.add_widget(labeled_checkbox("Is Grid?", self.grid_flag_chk, "Split characters across boxes or cells"))
        map_opts.add_widget(labeled_field("Grid N", self.grid_n_input, "Characters or cells to distribute"))
        map_body.add_widget(map_opts)
        self.btn_assign = make_button("Assign Mapping", tone="primary")
        self.btn_assign.bind(on_release=self.on_assign_mapping)
        map_body.add_widget(self.btn_assign)
        if is_mobile:
            self._mobile_section_cards["Mapping"] = map_card
        else:
            controls.add_widget(map_card)

        export_card, export_body = make_card("Export", "Generate patient output files")
        out_grid = GridLayout(cols=1 if is_mobile else 2, spacing=dp(8), size_hint_y=None, height=(2*row_h+dp(8)) if is_mobile else row_h)
        self.btn_generate_one = make_button("Generate Single PDF", tone="accent")
        self.btn_generate_one.bind(on_release=self.on_generate_single)
        self.btn_generate_batch = make_button("Generate Batch PDFs", tone="plain")
        self.btn_generate_batch.bind(on_release=self.on_generate_batch)
        out_grid.add_widget(self.btn_generate_one)
        out_grid.add_widget(self.btn_generate_batch)
        export_body.add_widget(out_grid)
        if is_mobile:
            self._mobile_section_cards["Export"] = export_card
        else:
            controls.add_widget(export_card)

        if is_mobile:
            mobile_controls_card, mobile_controls_body = make_card("Controls", "Switch sections to keep the screen lighter and easier to navigate")
            tabs = _make_mobile_section_tabs(["Files", "Session", "Detection", "Mapping", "Export"])
            mobile_controls_body.add_widget(tabs)
            self._mobile_section_host = GridLayout(cols=1, spacing=dp(8), size_hint_y=None)
            self._mobile_section_host.bind(minimum_height=self._mobile_section_host.setter("height"))
            mobile_controls_body.add_widget(self._mobile_section_host)
            controls.add_widget(mobile_controls_card)
            Clock.schedule_once(lambda dt: _show_mobile_section("Files"), 0)
        controls_scroll.add_widget(controls)
        controls_wrap.add_widget(controls_scroll)

        preview_outer = BoxLayout(orientation="vertical", spacing=gap, size_hint=(1, 0.44) if is_mobile else (0.56,1))
        preview_card, preview_body = make_card("Live Preview", "Interactive canvas for detection review, mapping, and patient output preview.")
        preview_card.size_hint_y = 1
        preview_card.bind(minimum_height=preview_card.setter("height"))
        preview_toolbar = GridLayout(cols=2 if is_mobile else 4, spacing=dp(8), size_hint_y=None, height=row_h if not is_mobile else 2*row_h+dp(8))
        for txt, cb, tone in [
            ("Refresh", self.on_preview, "primary"),
            ("Detect", self.on_run_detect, "accent"),
            ("Prev", self.on_prev_page, "plain"),
            ("Next", self.on_next_page, "plain"),
        ]:
            b = make_button(txt, tone=tone)
            b.bind(on_release=cb)
            preview_toolbar.add_widget(b)
        preview_body.add_widget(preview_toolbar)

        preview_shell = BoxLayout(orientation="vertical", spacing=dp(8), padding=dp(8), size_hint_y=None)
        preview_shell.height = max(dp(320), Window.height * 0.42) if is_mobile else dp(860)
        style_card(preview_shell, palette["preview_bg"], radius=dp(24))
        preview_wrap = ScrollView(do_scroll_x=True, do_scroll_y=True, bar_width=dp(6), scroll_type=["bars", "content"])
        preview_stack = BoxLayout(orientation="vertical", spacing=dp(8), size_hint_y=None)
        preview_stack.bind(minimum_height=preview_stack.setter("height"))
        self.preview = InteractivePreview(size_hint=(None, None))
        self.preview.bind(texture=self._update_preview_size)
        self.preview.box_tap_callback = self.on_preview_box_tap
        self.preview.hover_callback = self.on_preview_box_hover
        self.preview_info = Label(text="Interactive preview ready. Tap a box to select it.", color=palette["text"],
                                  size_hint_y=None, height=dp(44), halign="left", valign="middle", font_size=dp(12))
        self.preview_info.bind(size=self._sync_label_text_size)
        style_card(self.preview_info, palette["chip"], radius=dp(18))
        preview_stack.add_widget(self.preview_info)
        preview_stack.add_widget(self.preview)
        preview_wrap.add_widget(preview_stack)
        preview_shell.add_widget(preview_wrap)
        preview_body.add_widget(preview_shell)
        preview_outer.add_widget(preview_card)

        if is_mobile:
            main.add_widget(preview_outer)
            main.add_widget(controls_wrap)
        else:
            main.add_widget(controls_wrap)
            main.add_widget(preview_outer)
        root.add_widget(main)

        return root

    # --------------------------------------------------------
    # UI helpers
    # --------------------------------------------------------
    def _sync_label_text_size(self, instance, value):
        instance.text_size = value

    def _update_preview_size(self, instance, texture):
        if texture:
            if platform == "android":
                max_w = max(dp(220), Window.width - dp(56))
                scale = min(1.0, max_w / float(texture.width))
                self.preview.size = (texture.width * scale, texture.height * scale)
            else:
                self.preview.size = texture.size

    def set_status(self, text):
        self.status_lbl.text = text

    def get_selected_file(self):
        return None

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
            self.engine.settings["Use_Extent"] = bool(getattr(self, "use_extent_chk", None).active if hasattr(self, "use_extent_chk") else int(self.use_extent.text.strip() or "0"))
        except Exception as e:
            raise ValueError(f"Invalid settings input: {e}")

    def get_app_output_dir(self):
        """Return a writable app-controlled directory for generated files/configs."""
        base_dir = getattr(self, "user_data_dir", None)
        if not base_dir:
            if platform == "android":
                base_dir = "/sdcard/Download/MediMapPro"
            else:
                base_dir = os.path.join(os.path.expanduser("~"), "MediMapPro")

        out_dir = os.path.join(base_dir, "output")
        os.makedirs(out_dir, exist_ok=True)
        return out_dir

    def get_default_file_path(self):
        if platform == "android":
            shared_dir = "/sdcard/Download"
            if os.path.isdir(shared_dir):
                return shared_dir
            return self.get_app_output_dir()
        return os.path.expanduser("~")
        
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
        
        if hasattr(self, "use_extent_chk"):
            self.use_extent_chk.active = bool(s["Use_Extent"])
        elif hasattr(self, "use_extent"):
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

    def render_preview_image(self, img_bgr, boxes_payload=None, page_idx=0, preview_zoom=PREVIEW_SCALE):
        if img_bgr is None:
            return
        if hasattr(self.preview, "set_texture_from_bgr"):
            self.preview.set_texture_from_bgr(img_bgr)
            if boxes_payload is None:
                self.preview.clear_boxes()
            else:
                self.preview.set_boxes_payload(
                    boxes_payload,
                    selected_ids=self.engine.selected_box_ids,
                    preview_zoom=preview_zoom,
                    page_idx=page_idx,
                )
        else:
            rgba = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2RGBA)
            buf = rgba.tobytes()
            texture = Texture.create(size=(rgba.shape[1], rgba.shape[0]), colorfmt="rgba")
            texture.blit_buffer(buf, colorfmt="rgba", bufferfmt="ubyte")
            texture.flip_vertical()
            self.preview.texture = texture

    def _build_preview_boxes_payload(self, page_idx, preview_zoom):
        payload = []
        for i, r in enumerate(self.engine.all_boxes):
            payload.append({
                "id": i,
                "x": float(r.x0 * preview_zoom),
                "y": float(r.y0 * preview_zoom),
                "w": float(r.width * preview_zoom),
                "h": float(r.height * preview_zoom),
                "t": self.engine.box_types[i] if i < len(self.engine.box_types) else "field",
                "mapping": self.engine.describe_box_mapping(i, page_idx),
            })
        return payload

    def _find_mapping_for_box(self, box_id, page_idx):
        if box_id < 0 or box_id >= len(self.engine.all_boxes):
            return None
        target_rect = self.engine.all_boxes[box_id]
        best_hit = None
        best_score = -1.0
        for configs in self.engine.custom_mappings.values():
            for cfg in configs:
                if cfg.get("page", 0) != page_idx:
                    continue
                for r in self.engine._mapping_rect_list(cfg):
                    score = self.engine._mapping_match_score(target_rect, r)
                    if score > best_score:
                        best_score = score
                        best_hit = cfg
        if best_hit is not None and best_score >= 1.0:
            return best_hit
        return None

    def _sync_box_selection_ui(self):
        ids = sorted(set(int(x) for x in self.engine.selected_box_ids if isinstance(x, int) or str(x).isdigit()))
        self.engine.selected_box_ids = ids
        self.box_ids_input.text = ",".join(str(i) for i in ids)
        if hasattr(self.preview, "set_selected_ids"):
            self.preview.set_selected_ids(ids)

        if not ids:
            self.preview_info.text = "Interactive preview ready. Tap a box to select it."
            return

        first = ids[0]
        box_type = self.engine.box_types[first] if first < len(self.engine.box_types) else "field"
        mapping = self.engine.describe_box_mapping(first, self.current_page_idx())
        self.preview_info.text = f"Selected: {ids} | Type: {box_type} | {mapping}"

    def on_preview_box_hover(self, hit):
        if not hit:
            return
        mapping = hit.get("mapping") or "Unmapped"
        self.preview_info.text = f"Hover Box {hit['id']} | Type: {hit.get('t', 'field')} | {mapping}"

    def on_preview_box_tap(self, hit):
        box_id = int(hit["id"])
        ids = list(self.engine.selected_box_ids)
        if box_id in ids:
            ids.remove(box_id)
        else:
            ids.append(box_id)
        self.engine.selected_box_ids = sorted(set(ids))
        self._sync_box_selection_ui()

        mapping = self._find_mapping_for_box(box_id, self.current_page_idx())
        if mapping:
            col = str(mapping.get("column", "")).strip()
            trig = str(mapping.get("trigger", "")).strip()
            if col:
                self.column_spinner.text = col
            self.trigger_input.text = trig
            if hasattr(self, "grid_flag_chk"):
                self.grid_flag_chk.active = bool(mapping.get("g", False))
            else:
                self.grid_flag_input.text = "1" if bool(mapping.get("g", False)) else "0"
            self.grid_n_input.text = str(int(mapping.get("n", 1)))

        self.open_mapping_popup_for_selection(primary_box_id=box_id)

    def open_mapping_popup_for_selection(self, primary_box_id=None):
        ids = sorted(set(int(x) for x in self.engine.selected_box_ids if isinstance(x, int) or str(x).isdigit()))
        if not ids:
            self.set_status("Tap at least one preview box first.")
            return

        page_idx = self.current_page_idx()
        focus_id = ids[0] if primary_box_id is None else int(primary_box_id)
        existing = self._find_mapping_for_box(focus_id, page_idx)

        current_col = self.column_spinner.text if self.column_spinner.text != "Select Column" else ""
        current_trigger = self.trigger_input.text.strip()
        current_grid_flag = "1" if getattr(self, "grid_flag_chk", None).active else "0" if hasattr(self, "grid_flag_chk") else (self.grid_flag_input.text.strip() or "0")
        current_grid_n = self.grid_n_input.text.strip() or "1"

        if existing:
            current_col = str(existing.get("column", "")).strip() or current_col
            current_trigger = str(existing.get("trigger", "")).strip()
            current_grid_flag = "1" if bool(existing.get("g", False)) else "0"
            current_grid_n = str(int(existing.get("n", 1)))

        wrap = BoxLayout(orientation="vertical", spacing=8, padding=8)

        title_lbl = Label(
            text=f"Selected boxes: {', ' .join(str(i) for i in ids)}\nPage: {page_idx}",
            size_hint_y=None,
            height=46,
            halign="left",
            valign="middle"
        )
        title_lbl.bind(size=self._sync_label_text_size)
        wrap.add_widget(title_lbl)

        column_spinner = Spinner(
            text=current_col if current_col else "Select Column",
            values=self.column_spinner.values,
            size_hint_y=None,
            height=44
        )
        wrap.add_widget(column_spinner)

        trigger_input = TextInput(
            text=current_trigger,
            hint_text="Trigger",
            multiline=False,
            size_hint_y=None,
            height=44
        )
        wrap.add_widget(trigger_input)

        grid_row = GridLayout(cols=2, size_hint_y=None, height=64, spacing=8)
        popup_grid_chk = CheckBox(active=(str(current_grid_flag).strip() == "1"))
        grid_flag_wrap = BoxLayout(orientation="horizontal", spacing=8)
        grid_flag_wrap.add_widget(popup_grid_chk)
        grid_flag_lbl = Label(text="Is Grid?", halign="left", valign="middle")
        grid_flag_lbl.bind(size=self._sync_label_text_size)
        grid_flag_wrap.add_widget(grid_flag_lbl)
        grid_n_input = TextInput(
            text=current_grid_n,
            hint_text="Grid N",
            multiline=False
        )
        grid_row.add_widget(grid_flag_wrap)
        grid_row.add_widget(grid_n_input)
        wrap.add_widget(grid_row)

        btn_row = GridLayout(cols=3, size_hint_y=None, height=46, spacing=6)
        btn_assign = Button(text="Assign")
        btn_clear = Button(text="Clear")
        btn_close = Button(text="Close")
        btn_row.add_widget(btn_assign)
        btn_row.add_widget(btn_clear)
        btn_row.add_widget(btn_close)
        wrap.add_widget(btn_row)

        popup = Popup(title="Map Selected Box(es)", content=wrap, size_hint=(0.88, 0.52))

        def _assign(*_):
            try:
                column = column_spinner.text.strip()
                if not column or column == "Select Column":
                    raise ValueError("Please select a column.")
                trigger = trigger_input.text.strip()
                is_grid = bool(getattr(popup_grid_chk, "active", False))
                grid_n = int(grid_n_input.text.strip() or "1")

                self.engine.assign_mapping(ids, column, trigger, is_grid, grid_n, page_idx)

                self.column_spinner.text = column
                self.trigger_input.text = trigger
                
                if hasattr(self, "grid_flag_chk"):
                    self.grid_flag_chk.active = bool(is_grid)
                else:
                    self.grid_flag_input.text = "1" if is_grid else "0"
                self.grid_n_input.text = str(grid_n)
                self.box_ids_input.text = ",".join(str(i) for i in ids)
                self.engine.selected_box_ids = ids
                self._sync_box_selection_ui()
                self.set_status(
                    f"Mapping saved from preview.\nPage: {page_idx}\nColumn: {column}\nBoxes: {ids}"
                )
                popup.dismiss()
                self.on_preview(None)
            except Exception as e:
                self.set_status(f"Preview mapping error:\n{e}")

        def _clear(*_):
            try:
                self.clear_mapping_for_box_ids(ids, page_idx)
                self.engine.selected_box_ids = ids
                self._sync_box_selection_ui()
                self.set_status(f"Cleared mapping for boxes: {ids} on page {page_idx}")
                popup.dismiss()
                self.on_preview(None)
            except Exception as e:
                self.set_status(f"Clear mapping error:\n{e}")

        btn_assign.bind(on_release=_assign)
        btn_clear.bind(on_release=_clear)
        btn_close.bind(on_release=lambda *_: popup.dismiss())
        popup.open()

    def clear_mapping_for_box_ids(self, box_ids, page_idx):
        box_ids = sorted(set(int(x) for x in box_ids))
        target_rects = []
        for b_id in box_ids:
            if b_id < 0 or b_id >= len(self.engine.all_boxes):
                continue
            target_rects.append(self.engine.all_boxes[b_id])

        def overlaps_any(mapping_item):
            existing_rects = self.engine._mapping_rect_list(mapping_item)
            for er in existing_rects:
                for sr in target_rects:
                    if self.engine._rects_refer_to_same_target(er, sr, tol=0.20):
                        return True
            return False

        for k in list(self.engine.custom_mappings.keys()):
            kept = []
            for m in self.engine.custom_mappings[k]:
                if m.get("page", 0) != page_idx:
                    kept.append(m)
                    continue
                if overlaps_any(m):
                    continue
                kept.append(m)
            if kept:
                self.engine.custom_mappings[k] = kept
            else:
                del self.engine.custom_mappings[k]

    # --------------------------------------------------------
    # File loading
    # --------------------------------------------------------
    def _load_dataframe_from_path(self, path):
        self.engine.load_dataframe(path)
        self.refresh_patient_and_column_lists()

        cols = self.engine.df_columns()
        self.patient_col_spinner.values = cols
        self.patient_col_spinner.text = cols[0] if cols else ""

        names = self.engine.patient_names()
        self.patient_spinner.values = names
        self.patient_spinner.text = names[0] if names else ""

        self.set_status(f"Data loaded:\n{os.path.basename(path)}\nRows: {len(self.engine.df)}")
        if self.engine.total_pages() > 0:
            Clock.schedule_once(lambda dt: self.on_preview(None), 0.1)

    def _load_config_from_path(self, path):
        if not path.lower().endswith(".json"):
            raise ValueError("Please select a JSON config file.")

        self.engine.load_config(path)
        self.engine.all_boxes = []
        self.engine.box_types = []
        self.push_engine_settings_to_ui()
        self.set_status(f"Config loaded:\n{os.path.basename(path)}")

    def _merge_config_from_path(self, path):
        if not path.lower().endswith(".json"):
            raise ValueError("Please select a JSON config file.")

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

    def _copy_android_uri_to_local_file(self, uri, default_name="selected_input", required_suffix=None):
        if platform != "android" or not ANDROID_JAVA_AVAILABLE:
            raise RuntimeError("Android document picker is unavailable.")

        if uri is None:
            raise ValueError("Android returned no document URI.")

        activity_obj = AndroidPythonActivity.mActivity
        resolver = activity_obj.getContentResolver()

        if isinstance(uri, str):
            uri_string = uri.strip()
        else:
            try:
                uri_string = str(uri.toString()).strip()
            except Exception:
                uri_string = ""

        if not uri_string:
            raise ValueError("Android returned an empty document URI.")
        if not (uri_string.startswith("content://") or uri_string.startswith("file://")):
            raise ValueError(f"Android returned an invalid document URI: {uri_string}")

        uri = AndroidUri.parse(uri_string)
        input_stream = resolver.openInputStream(uri)
        if input_stream is None:
            raise ValueError("Unable to open the selected Android document.")

        display_name = None
        cursor = None
        try:
            OpenableColumns = autoclass("android.provider.OpenableColumns")
            cursor = resolver.query(uri, None, None, None, None)
            if cursor is not None and cursor.moveToFirst():
                idx = cursor.getColumnIndex(OpenableColumns.DISPLAY_NAME)
                if idx >= 0:
                    display_name = cursor.getString(idx)
        except Exception:
            display_name = None
        finally:
            try:
                if cursor is not None:
                    cursor.close()
            except Exception:
                pass

        filename = safe_name(display_name or default_name)
        if required_suffix:
            required_suffix = str(required_suffix).strip()
            if required_suffix and not filename.lower().endswith(required_suffix.lower()):
                filename += required_suffix

        out_dir = self.get_app_output_dir()
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, filename)

        BufferedInputStream = autoclass("java.io.BufferedInputStream")
        BufferedOutputStream = autoclass("java.io.BufferedOutputStream")
        FileOutputStream = autoclass("java.io.FileOutputStream")
        ByteArray = autoclass("java.lang.reflect.Array")
        JavaByte = autoclass("java.lang.Byte").TYPE

        bis = BufferedInputStream(input_stream)
        bos = BufferedOutputStream(FileOutputStream(out_path))
        buffer = ByteArray.newInstance(JavaByte, 65536)

        try:
            while True:
                count = bis.read(buffer)
                if count == -1:
                    break
                bos.write(buffer, 0, count)
            bos.flush()
        finally:
            try:
                bis.close()
            except Exception:
                pass
            try:
                bos.close()
            except Exception:
                pass
            try:
                input_stream.close()
            except Exception:
                pass

        return out_path

    def _open_android_document_picker(self, request_code, mime_type="*/*", mime_types=None, title="document", on_picked=None, cancel_message=None):
        if platform != "android" or not ANDROID_JAVA_AVAILABLE:
            raise RuntimeError("Android system document picker is unavailable.")

        if on_picked is None:
            raise ValueError("Android document picker callback is required.")

        def _on_activity_result(request_code_result, result_code, intent):
            if request_code_result != request_code:
                return
            activity.unbind(on_activity_result=_on_activity_result)
            if result_code != -1 or intent is None:
                self.set_status(cancel_message or f"{title.capitalize()} selection cancelled.")
                return
            try:
                uri = intent.getData()
                if uri is None:
                    raise ValueError(f"No {title} URI was returned by Android.")
                try:
                    flags = intent.getFlags()
                    take_flags = flags & (AndroidIntent.FLAG_GRANT_READ_URI_PERMISSION | AndroidIntent.FLAG_GRANT_WRITE_URI_PERMISSION)
                    AndroidPythonActivity.mActivity.getContentResolver().takePersistableUriPermission(uri, take_flags)
                except Exception:
                    pass
                on_picked(uri)
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"{title.capitalize()} Error: {e}")

        activity.bind(on_activity_result=_on_activity_result)
        intent = AndroidIntent(AndroidIntent.ACTION_OPEN_DOCUMENT)
        intent.addCategory(AndroidIntent.CATEGORY_OPENABLE)
        intent.setType(mime_type or "*/*")
        intent.addFlags(AndroidIntent.FLAG_GRANT_READ_URI_PERMISSION)
        intent.addFlags(AndroidIntent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION)

        if mime_types:
            try:
                StringClass = autoclass("java.lang.String")
                ByteArray = autoclass("java.lang.reflect.Array")
                java_mime_types = ByteArray.newInstance(StringClass, len(mime_types))
                for idx, value in enumerate(mime_types):
                    java_mime_types[idx] = value
                intent.putExtra(AndroidIntent.EXTRA_MIME_TYPES, java_mime_types)
            except Exception:
                pass

        AndroidPythonActivity.mActivity.startActivityForResult(intent, request_code)

    def on_load_csv(self, instance):
        if platform == "android":
            try:
                return self._open_android_document_picker(
                    request_code=42422,
                    mime_type="*/*",
                    mime_types=[
                        "text/csv",
                        "application/vnd.ms-excel",
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "application/octet-stream",
                    ],
                    title="data file",
                    cancel_message="Data file selection cancelled.",
                    on_picked=lambda uri: self._load_dataframe_from_path(
                        self._copy_android_uri_to_local_file(uri, default_name="selected_data")
                    ),
                )
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"Data file Error: {e}")
                return

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
                self._load_dataframe_from_path(path)
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"Load data error:\n{e}")
        popup.dismiss()


    def on_load_config(self, instance):
        if platform == "android":
            try:
                return self._open_android_document_picker(
                    request_code=42423,
                    mime_type="application/json",
                    mime_types=["application/json", "text/json", "text/plain"],
                    title="config file",
                    cancel_message="Config file selection cancelled.",
                    on_picked=lambda uri: self._load_config_from_path(
                        self._copy_android_uri_to_local_file(uri, default_name="selected_config", required_suffix=".json")
                    ),
                )
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"Config Error: {e}")
                return

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
                self._load_config_from_path(path)
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"Load config error:\n{e}")
        popup.dismiss()


    def on_merge_config(self, instance):
        if platform == "android":
            try:
                return self._open_android_document_picker(
                    request_code=42424,
                    mime_type="application/json",
                    mime_types=["application/json", "text/json", "text/plain"],
                    title="merge config file",
                    cancel_message="Merge config selection cancelled.",
                    on_picked=lambda uri: self._merge_config_from_path(
                        self._copy_android_uri_to_local_file(uri, default_name="merge_config", required_suffix=".json")
                    ),
                )
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"Merge Config Error: {e}")
                return

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
                self._merge_config_from_path(path)
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"Merge config error:\n{e}")
        popup.dismiss()

    def on_save_config(self, instance):
        try:
            self.apply_ui_settings_to_engine()
            out_path = os.path.join(self.get_app_output_dir(), CONFIG_FILENAME)
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
                    self.render_preview_image(raw_img, boxes_payload=[], page_idx=idx, preview_zoom=PREVIEW_SCALE)
                    self._sync_box_selection_ui()
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
                    self.render_preview_image(raw_img, boxes_payload=[], page_idx=idx, preview_zoom=PREVIEW_SCALE)
                    self._sync_box_selection_ui()
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
            self.engine.selected_box_ids = []
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

            if platform == "android" and not self.engine.supports_processing_backend():
                raw_img = self.engine.get_raw_preview_pixmap(
                    page_idx=page_idx,
                    preview_zoom=PREVIEW_SCALE
                )
                self.render_preview_image(
                    raw_img,
                    boxes_payload=[],
                    page_idx=page_idx,
                    preview_zoom=PREVIEW_SCALE,
                )
                self._sync_box_selection_ui()
                self.set_status("Android native PDF preview is available, but detection/filling still requires the PyMuPDF backend in this build.")
                return

            if not patient or self.engine.df is None or self.engine.df.empty:
                raw_img = self.engine.get_raw_preview_pixmap(
                    page_idx=page_idx,
                    preview_zoom=PREVIEW_SCALE
                )
                self.render_preview_image(
                    raw_img,
                    boxes_payload=[],
                    page_idx=page_idx,
                    preview_zoom=PREVIEW_SCALE,
                )
                self._sync_box_selection_ui()
                self.set_status(
                    f"Raw PDF preview.\n"
                    f"Page: {page_idx}\n"
                    f"No patient selected yet."
                )
                return

            self.apply_ui_settings_to_engine()

            img = self.engine.get_processed_preview_pixmap(
                patient_name=patient,
                page_idx=page_idx,
                preview_zoom=PREVIEW_SCALE
            )
            boxes_payload = self._build_preview_boxes_payload(
                page_idx=page_idx,
                preview_zoom=PREVIEW_SCALE,
            )
            self.render_preview_image(
                img,
                boxes_payload=boxes_payload,
                page_idx=page_idx,
                preview_zoom=PREVIEW_SCALE,
            )
            self._sync_box_selection_ui()

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

            trigger = self.trigger_input.text.strip()
            is_grid = bool(self.grid_flag_chk.active) if hasattr(self, "grid_flag_chk") else bool(int(self.grid_flag_input.text.strip() or "0"))
            grid_n = int(self.grid_n_input.text.strip() or "1")
            page_idx = self.current_page_idx()

            self.engine.assign_mapping(box_ids, column, trigger, is_grid, grid_n, page_idx)
            self.engine.selected_box_ids = sorted(set(box_ids))
            self._sync_box_selection_ui()

            self.set_status(
                f"Mapping saved.\nPage: {page_idx}\nColumn: {column}\nBoxes: {box_ids}"
            )
            self.on_preview(None)
        except Exception as e:
            self.set_status(f"Assign mapping error:\n{e}")

    # --------------------------------------------------------
    # Output generation
    # --------------------------------------------------------
    def on_generate_single(self, instance):
        try:
            if platform == "android" and not self.engine.supports_processing_backend():
                self.set_status("Android native PDF preview works in this build, but PDF filling/export still requires the PyMuPDF backend.")
                return
            if not self.engine.pdf_path:
                self.set_status("Load PDF first.")
                return

            patient = self.selected_patient()
            if not patient:
                self.set_status("Select a patient first.")
                return

            page_idx = self.current_page_idx()
            doc = self.engine.process_doc(patient, page_idx=page_idx)

            out_dir = self.get_app_output_dir()
            out_path = os.path.join(
                out_dir,
                f"Filled_{safe_name(patient)}_page_{page_idx + 1}.pdf"
            )
            

            doc.save(out_path)
            doc.close()

            self.set_status(f"Generated:\n{out_path}")
        except Exception as e:
            traceback.print_exc()
            self.set_status(f"Generate single error:\n{e}")

    def _load_pdf_from_path(self, path):
        total = self.engine.load_pdf(path)

        cur_idx = self.current_page_idx()
        max_idx = max(total - 1, 0)
        if cur_idx > max_idx:
            self.page_input.text = "0"

        raw_img = self.engine.get_raw_preview_pixmap(
            page_idx=self.current_page_idx(),
            preview_zoom=PREVIEW_SCALE
        )
        self.render_preview_image(raw_img, boxes_payload=[], page_idx=self.current_page_idx(), preview_zoom=PREVIEW_SCALE)
        self._sync_box_selection_ui()

        mode_note = "\nMode: Android native preview fallback" if self.engine.android_preview_only_mode() else ""
        self.set_status(
            f"PDF Loaded: {os.path.basename(path)}\n"
            f"Pages: {total}\n"
            f"Showing raw template page: {self.current_page_idx()}"
            f"{mode_note}"
        )

    def _copy_android_uri_to_local_pdf(self, uri):
        return self._copy_android_uri_to_local_file(uri, default_name="selected_input", required_suffix=".pdf")

    def _open_android_pdf_picker(self):
        if platform != "android" or not ANDROID_JAVA_AVAILABLE:
            self.set_status("Android system document picker is unavailable; falling back to the file chooser.")
            return self._open_legacy_pdf_picker()

        request_code = 42421

        def _on_activity_result(request_code_result, result_code, intent):
            if request_code_result != request_code:
                return
            activity.unbind(on_activity_result=_on_activity_result)
            if result_code != -1 or intent is None:
                self.set_status("PDF selection cancelled.")
                return
            try:
                uri = intent.getData()
                if uri is None:
                    raise ValueError("No PDF URI was returned by Android.")
                try:
                    flags = intent.getFlags()
                    take_flags = flags & (AndroidIntent.FLAG_GRANT_READ_URI_PERMISSION | AndroidIntent.FLAG_GRANT_WRITE_URI_PERMISSION)
                    AndroidPythonActivity.mActivity.getContentResolver().takePersistableUriPermission(uri, take_flags)
                except Exception:
                    pass
                local_path = self._copy_android_uri_to_local_pdf(uri)
                self._load_pdf_from_path(local_path)
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"PDF Error: {e}")

        activity.bind(on_activity_result=_on_activity_result)
        intent = AndroidIntent(AndroidIntent.ACTION_OPEN_DOCUMENT)
        intent.addCategory(AndroidIntent.CATEGORY_OPENABLE)
        intent.setType("application/pdf")
        intent.addFlags(AndroidIntent.FLAG_GRANT_READ_URI_PERMISSION)
        intent.addFlags(AndroidIntent.FLAG_GRANT_PERSISTABLE_URI_PERMISSION)
        AndroidPythonActivity.mActivity.startActivityForResult(intent, request_code)

    def _open_legacy_pdf_picker(self):
        content = FileChooserListView(
            filters=["*.pdf"],
            path=self.get_default_file_path()
        )
        popup = Popup(title="Select PDF Template", content=content, size_hint=(0.9, 0.9))
        content.bind(on_submit=lambda obj, sel, touch: self._handle_pdf_selection(sel, popup))
        popup.open()

    def on_load_pdf(self, instance):
        if platform == "android":
            return self._open_android_pdf_picker()
        return self._open_legacy_pdf_picker()
    
    def _handle_pdf_selection(self, selection, popup):
        popup.dismiss()

        if selection:
            try:
                path = selection[0]
                self.set_status(f"Loading PDF...\n{os.path.basename(path)}")

                total = self.engine.load_pdf(path)

                cur_idx = self.current_page_idx()
                max_idx = max(total - 1, 0)
                if cur_idx > max_idx:
                    self.page_input.text = "0"

                self.set_status(
                    f"PDF Loaded: {os.path.basename(path)}\n"
                    f"Pages: {total}\n"
                    f"Rendering preview..."
                )
                Clock.schedule_once(lambda dt: self._finish_pdf_load_preview(path, total), 0.05)
            except Exception as e:
                traceback.print_exc()
                self.set_status(f"PDF Error: {e}")

    def _finish_pdf_load_preview(self, path, total):
        try:
            page_idx = self.current_page_idx()
            raw_img = self.engine.get_raw_preview_pixmap(
                page_idx=page_idx,
                preview_zoom=PREVIEW_SCALE
            )
            self.render_preview_image(
                raw_img,
                boxes_payload=[],
                page_idx=page_idx,
                preview_zoom=PREVIEW_SCALE
            )
            self._sync_box_selection_ui()

            self.set_status(
                f"PDF Loaded: {os.path.basename(path)}\n"
                f"Pages: {total}\n"
                f"Showing raw template page: {page_idx}"
            )
        except Exception as e:
            traceback.print_exc()
            self.set_status(f"PDF Preview Error: {e}")


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
            if platform == "android" and not self.engine.supports_processing_backend():
                self.set_status("Android native PDF preview works in this build, but batch filling/export still requires the PyMuPDF backend.")
                return
            if self.engine.df is None or self.engine.df.empty:
                self.set_status("Load CSV/XLSX first.")
                return

            if not self.engine.pdf_path:
                self.set_status("Load PDF first.")
                return

            names = sorted(self.engine.df["_DISPLAY_NAME"].dropna().astype(str).unique())
            out_dir = os.path.join(self.get_app_output_dir(), "batch_output")
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
