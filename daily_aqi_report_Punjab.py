#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
===============================================================================
 DAILY AIR QUALITY INDEX REPORT OF PUNJAB  —  automated generator
 Directorate of Environmental Monitoring Center (EMC), EPA Punjab
===============================================================================

What it produces (all in the output folder):
  1. DAILY_AQI_REPORT_<dd.mm.yyyy>.pdf   Page 1  Punjab district ranking (Urdu)
                                          Page 2  Punjab district AQI map (filled by AQI band)
                                          Page 3  Lahore AQMS map + station table
  2. AQMS_Map_<District>_<dd.mm.yyyy>.png  High-resolution zoomed AQMS map (300 dpi)
  3. AQI_Summary_<dd.mm.yyyy>.xlsx        Districts, Stations, Hourly data, QA flags
  4. DAILY_AQI_REPORT_<dd.mm.yyyy>.html   The same report as a web page
  5. DAILY_AQI_REPORT_<dd.mm.yyyy>.docx   Word copy rendered from the PDF pages

Inputs:
  --csv   Station-level dashboard export ("graphs_periodic_*.csv"), with columns
          "<Station> • AQI" and "<Station> • Dominant Pollutant", 24 hourly rows.
  --shp   AQMS locations (point shapefile, any CRS). The station-name field is
          detected automatically and matched to the CSV names (fuzzy), or set it
          with --shp-name-field.
  --districts-shp  (optional) Punjab district boundaries. Used to highlight the
          focus district on the map and as the map background when offline.
          Page 2 (Punjab map) also finds it automatically: any *dist*.shp in the shp folder.

Street map (basemap) — tried in this order, first one that works is used:
  1. Your own file: basemap/<Focus>.tif, or basemap/<Focus>.png + .pgw world
     file (e.g. basemap/Lahore.png exported from QGIS), or --basemap-file PATH.
  2. The copy saved from an earlier download (basemap/cache/) — works offline.
  3. Download: --basemap (default osm), then BASEMAP_FALLBACKS.
  4. A clean offline style.
  Check your network with:  python daily_aqi_report_Punjab.py --test-basemap

How the numbers are calculated (reproduces the existing manual report):
  * Station 24-h AQI     = mean of the valid hourly AQI values for the day.
                           A station needs at least MIN_VALID_HOURS (default 18)
                           hourly AQI values, otherwise it is omitted from report
                           tables and maps (exclusions remain in QA). All present
                           hourly AQI values are included, including zeros and
                           isolated peaks; QA flags report unusual values without
                           excluding them. Missing AQI values are not counted.
  * District AQI         = mean of its stations' 24-h AQI (transboundary AQMS
                           are shown separately and excluded from district averages).
                           Stations count in the district they lie in: 11 in Lahore,
                           3 in Sheikhupura (DHQ Sheikhupura, Lathepur, BHU Jandiala); all 3 are
                           averaged in Sheikhupura. Lathepur and BHU Jandiala are also drawn on the
                           Lahore map (as transboundary AQMS, not averaged into Lahore).
  * District list        = the 36 districts that have AQMS (PUNJAB_DISTRICTS), always all
                           listed; a district without data is left blank (nothing written).
  * Lahore city AQI      = mean of the Lahore city AQMS (same as district value).
  * Punjab average AQI   = mean of the district AQI values that have data.
  * Dominant pollutants  = the (up to) three pollutants that were dominant in the
                           most hours, e.g. "PM2.5 > PM10 > O3".
  * Values are rounded half-up (122.5 -> 123).

Quick start (Windows / VS Code terminal):
  pip install -r requirements.txt
  python -m playwright install chromium
  python daily_aqi_report_Punjab.py --csv graphs_periodic_2026-09-30.csv --shp 56AQMS.shp

Everything you may want to change (colours, Urdu wording, station names,
thresholds, which district is mapped) is in the CONFIGURATION section below.
===============================================================================
"""
from __future__ import annotations

import argparse
import datetime as dt
import difflib
import html
import io
import math
import os
import re
import sys
import warnings
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)

SCRIPT_DIR = Path(__file__).resolve().parent

# =============================================================================
# CONFIGURATION
# =============================================================================

MIN_VALID_HOURS = 18            # station needs >= this many hourly AQI values
REPORT_TIME = "06:30AM"         # printed on page 2 ("Updated Time")
FOCUS_DISTRICT = "Lahore"       # district shown on the zoomed AQMS map (page 2)
RIGHT_COLUMN_ROWS = 22          # page 1: ranks 1..22 right column, rest left
BASEMAP = "osm"                 # street map: osm | voyager | esri-street | positron | satellite | none
BASEMAP_FALLBACKS = ["voyager", "esri-street", "osm"]   # tried in turn if the first one fails
BASEMAP_DIR = SCRIPT_DIR / "basemap"   # drop Lahore.tif (or Lahore.png + Lahore.pgw) here to use your own
BASEMAP_ZOOM = "auto"           # or a fixed tile zoom, e.g. 12
BASEMAP_DETAIL = 0.75           # lower = bigger street labels, fewer streets; 1.0 = most detail, smallest labels
BASEMAP_FADE = 0.12             # white veil over the street map so AQI labels stand out (0 = none)
BASEMAP_TIMEOUT = 20            # seconds per tile request
BASEMAP_USER_AGENT = "EPA-Punjab-Daily-AQI-Report/1.1 (Directorate of EMC, EPA Punjab)"
BASEMAP_FILE_ATTRIBUTION = "user-supplied basemap"

ASSETS = SCRIPT_DIR / "assets"
FONTS = SCRIPT_DIR / "fonts"
LOGO_PUNJAB = ASSETS / "logo_punjab.png"
LOGO_EPA = ASSETS / "logo_epa.png"
LOGO_HELPLINE = ASSETS / "helpline_1373.png"
BORDER_FILE = ASSETS / "border_pk_in.geojson"

DIRECTORATE = "Directorate of Environmental Monitoring Center, Environmental Protection Agency Punjab."

# AQI bands: (low, high, English, Urdu, fill colour, text colour)
BANDS = [
    (0,   50,  "Good",                           "اچھا",                          "#00B050", "#000000"),
    (51,  100, "Satisfactory",                   "تسلی بخش",                      "#92D050", "#000000"),
    (101, 150, "Moderate",                       "معتدل",                         "#FFFF00", "#000000"),
    (151, 200, "Unhealthy for Sensitive Groups", "حساس افراد کے لیے غیر صحت بخش", "#ED7D31", "#000000"),
    (201, 300, "Unhealthy",                      "غیر صحت بخش",                   "#FF0000", "#FFFFFF"),
    (301, 400, "Very Unhealthy",                 "انتہائی غیر صحت بخش",           "#7030A0", "#FFFFFF"),
    (401, 500, "Hazardous",                      "خطرناک",                        "#C00000", "#FFFFFF"),
]
BAND_LABELS = ["0-50", "51-100", "101-150", "151-200", "201-300", "301-400", "401+"]
NO_DATA_FILL = "#D9D9D9"

# The 36 districts listed on page 1 (the districts that have AQMS stations; always all shown, blank when no data).
# Bahawalnagar, Layyah, Lodhran, Toba Tek Singh and Taunsa are not listed.
PUNJAB_DISTRICTS = [
    "Attock", "Bahawalpur", "Bhakkar", "Chakwal", "Chiniot", "DG Khan", "Faisalabad", "Gujranwala",
    "Gujrat", "Hafizabad", "Jhang", "Jhelum", "Kasur", "Khanewal", "Khushab", "Kot Addu", "Lahore",
    "Mandi Bahauddin", "Mianwali", "Multan", "Murree", "Muzaffargarh", "Nankana Sahib", "Narowal",
    "Okara", "Pakpattan", "Rahim Yar Khan", "Rajanpur", "Rawalpindi", "Sahiwal", "Sargodha",
    "Sheikhupura", "Sialkot", "Talagang", "Vehari", "Wazirabad",
]

# Stations drawn on (and listed under) another district's AQMS map, as transboundary AQMS (never averaged)
MAP_EXTRA_STATIONS = {"Lahore": ("Lathepur", "BHU Jandiala")}
# Transboundary stations that still count in the average of the district they lie in (page 1 ranking)
COUNTED_IN_DISTRICT = ("Lathepur", "BHU Jandiala")          # -> Sheikhupura


def in_focus(s, focus):
    return s.district == focus or s.label in MAP_EXTRA_STATIONS.get(focus, ())


# District names in Urdu (as printed on page 1)
DISTRICT_URDU = {
    "Attock": "اٹک", "Bahawalnagar": "بہاولنگر", "Bahawalpur": "بہاولپور", "Bhakkar": "بھکر",
    "Chakwal": "چکوال", "Chiniot": "چنیوٹ", "DG Khan": "ڈیرہ غازی خان", "Faisalabad": "فیصل آباد",
    "Gujranwala": "گوجرانوالہ", "Gujrat": "گجرات", "Hafizabad": "حافظ آباد", "Jhang": "جھنگ",
    "Jhelum": "جہلم", "Kasur": "قصور", "Khanewal": "خانیوال", "Khushab": "خوشاب", "Kot Addu": "کوٹ ادو",
    "Lahore": "لاہور", "Layyah": "لیہ", "Lodhran": "لودھراں", "Mandi Bahauddin": "منڈی بہاؤالدین",
    "Mianwali": "میانوالی", "Multan": "ملتان", "Murree": "مری", "Muzaffargarh": "مظفر گڑھ",
    "Nankana Sahib": "ننکانہ صاحب", "Narowal": "نارووال", "Okara": "اوکاڑہ", "Pakpattan": "پاکپتن",
    "Rahim Yar Khan": "رحیم یار خان", "Rajanpur": "راجن پور", "Rawalpindi": "راولپنڈی",
    "Sahiwal": "ساہیوال", "Sargodha": "سرگودھا", "Sheikhupura": "شیخوپورہ", "Sialkot": "سیالکوٹ",
    "Talagang": "تلہ گنگ", "Toba Tek Singh": "ٹوبہ ٹیک سنگھ", "Vehari": "وہاڑی", "Wazirabad": "وزیر آباد",
}
# Spelling variants found in CSV / shapefiles -> standard district name
DISTRICT_ALIASES = {
    "attok": "Attock", "d.g. khan": "DG Khan", "d g khan": "DG Khan", "dera ghazi khan": "DG Khan",
    "muree": "Murree", "tala gang": "Talagang", "r.y. khan": "Rahim Yar Khan", "ryk": "Rahim Yar Khan",
    "m.b. din": "Mandi Bahauddin", "mandi bahaudin": "Mandi Bahauddin", "t.t. singh": "Toba Tek Singh",
    "lhr": "Lahore", "nankana": "Nankana Sahib",
}

# Station registry: CSV station name -> (short label, district, role)
#   role = "city" (counts in district/city average) | "transboundary" (shown
#   separately, not averaged).  Stations not listed here are assigned to a
#   district automatically from their name, with role "city".
STATION_REGISTRY = {
    "Safari Park-LHR":                      ("Safari Park",   "Lahore", "city"),
    "Kahna Nau Hospital-LHR":               ("Kahna Nau",     "Lahore", "city"),
    "PKLI-LHR":                             ("PKLI",          "Lahore", "city"),
    "FMDRC-LHR":                            ("FMDRC",         "Lahore", "city"),
    "UET-LHR":                              ("UET",           "Lahore", "city"),
    "LWMC-LHR":                             ("LWMC",          "Lahore", "city"),
    "Punjab University-LHR":                ("Punjab University", "Lahore", "city"),
    "Govt. Teaching Hospital Shahdara-LHR": ("Shahdara",      "Lahore", "city"),
    "Gulberg III Lahore - Mobile 1":        ("Gulberg III",   "Lahore", "city"),
    "Egerton Road - Mobile 4":              ("Egerton Road",  "Lahore", "city"),
    "Lathepur LHR - Mobile 2":              ("Lathepur",      "Sheikhupura", "transboundary"),
    "Wagha Border LHR - Mobile 3":          ("Wagha",         "Lahore", "transboundary"),
    "BHU Jandiala Kalsan LHR - Mobile 5":   ("BHU Jandiala",  "Sheikhupura", "transboundary"),
    "DHQ Sheikhupura":                      ("DHQ Sheikhupura", "Sheikhupura", "city"),
    "M. Nawaz Sharif University of Engineering & Technology Multan": ("MNSUET Multan", "Multan", "city"),
    "IUB (Baghdad Campus) Bahawalpur":      ("IUB Baghdad Campus", "Bahawalpur", "city"),
    "IUB (Khawaja Fareed Campus) Bahawalpur": ("IUB Khawaja Fareed Campus", "Bahawalpur", "city"),
    "Drug Testing Laboratory Rawalpindi":   ("DTL Rawalpindi", "Rawalpindi", "city"),
    "Attok":                                ("Attock",        "Attock", "city"),
}

# Shapefile point name -> CSV station name, for names too different to match automatically
SHP_ALIASES = {
    "Deputy Commissioner Office, Attock": "Attok",
    "Deputy Commissioner Office, MB Din": "Mandi Bahauddin",
    "Deputy Commissioner Office, RY Khan": "DC Office Rahim Yar Khan",
}

# Public message (page 1), one per AQI band: (public items, sensitive-group items)
# Each item = (icon key, Urdu text). Icons: aqi heart food nosmoke run kit doctor
# mask home window child elder purifier goggles outdoor
PUBLIC_MESSAGES = {
    0: ([("outdoor", "فضائی معیار اچھا ہے، معمول کی بیرونی سرگرمیاں جاری رکھیں۔")],
        [("outdoor", "کسی خاص احتیاط کی ضرورت نہیں۔")]),
    1: ([("outdoor", "فضائی معیار تسلی بخش ہے، معمول کی سرگرمیاں جاری رکھیں۔")],
        [("run", "غیر معمولی طور پر حساس افراد طویل یا سخت بیرونی سرگرمی کم کریں۔"),
         ("doctor", "سانس کی تکلیف کی صورت میں ڈاکٹر سے رجوع کریں۔")]),
    2: ([("aqi", "باہر کی سرگرمیوں کی منصوبہ بندی کے لیے AQI کو مدنظر رکھیں۔")],
        [("heart", "صحت کی باقاعدگی سے نگرانی کریں (آکسیجن، بی پی، وغیرہ)۔"),
         ("food", "قوتِ مدافعت بڑھانے کے لیے صحت بخش غذا کھائیں۔"),
         ("nosmoke", "تمباکو نوشی سے گریز کریں۔"),
         ("run", "باہر کی سرگرمیوں کو محدود کریں۔"),
         ("kit", "ہنگامی امداد کے آلات (جیسے نیبولائزر) تیار رکھیں۔"),
         ("doctor", "اگر سانس کے مسائل پیدا ہوں تو ڈاکٹر سے مشورہ کریں۔")]),
    3: ([("mask", "باہر نکلتے وقت فیس ماسک پہنیں۔"),
         ("home", "غیر ضروری سفر سے گریز کریں۔")],
        [("child", "بچوں کو باہر کھیلنے سے روکیں۔"),
         ("elder", "بزرگ افراد باہر کم سے کم وقت گزاریں۔"),
         ("window", "دروازے اور کھڑکیاں بند رکھیں۔"),
         ("doctor", "دمہ اور دل کے مریض ڈاکٹر کے مشورے سے ماسک استعمال کریں۔")]),
    4: ([("mask", "باہر نکلتے وقت N95 ماسک پہنیں۔"),
         ("home", "جہاں تک ممکن ہو گھر پر رہیں۔"),
         ("run", "باہر جسمانی مشقت سے گریز کریں۔")],
        [("heart", "AQI اور صحت کی باقاعدگی سے نگرانی کریں۔"),
         ("child", "بچوں کو بیرونی سرگرمیوں سے روکیں۔"),
         ("doctor", "سی او پی ڈی اور دل کے مریض تجویز کردہ ماسک استعمال کریں۔")]),
    5: ([("home", "گھر کے اندر رہیں۔"),
         ("run", "ورزش محدود کریں اور اندرونِ خانہ ورزش کریں۔"),
         ("goggles", "باہر جانا ہو تو N95 ماسک اور چشمہ استعمال کریں۔")],
        [("heart", "AQI اور صحت کی باقاعدگی سے نگرانی کریں۔"),
         ("doctor", "سی او پی ڈی اور دل کے مریض تجویز کردہ ماسک استعمال کریں۔")]),
    6: ([("home", "گھر کے اندر رہیں۔"),
         ("goggles", "باہر جانا ناگزیر ہو تو N95 ماسک اور حفاظتی چشمہ استعمال کریں۔"),
         ("purifier", "گھر میں ایئر پیوریفائر استعمال کریں۔")],
        [("heart", "صحت (آکسیجن، بی پی وغیرہ) کی بار بار نگرانی کریں۔"),
         ("doctor", "طبیعت خراب ہونے پر فوراً ڈاکٹر سے رجوع کریں۔")]),
}

# Health advisory (page 1 message box), one per AQI band: (band name used in the title, [(section heading, [(icon, Urdu text), ...]), ...])
# Bands 0-4 follow the advisory sheets supplied by EMC; 301-400 and 401+ reuse the page-1 PUBLIC_MESSAGES text.
ADVISORY_TEXT = {
    0: ("بہتر", [
        ("عام عوام:", [("outdoor", "ہوا کا معیار اچھا ہے۔"),
                       ("run", "تمام بیرونی سرگرمیوں کے لیے مثالی حالات ہیں۔"),
                       ("outdoor", "کوئی پابندی نہیں ہے۔")]),
        ("حساس گروپس:", [("outdoor", "ہوا کے اچھے معیار کا فائدہ اٹھائیں۔")])]),
    1: ("اطمینان بخش", [
        ("عام عوام:", [("outdoor", "ہوا کا معیار اچھا ہے۔"),
                       ("doctor", "صحت مند افراد کے لیے کسی خاص احتیاط کی ضرورت نہیں ہے۔")]),
        ("حساس گروہ:", [("outdoor", "منصوبے کے مطابق بیرونی سرگرمیاں جاری رکھیں۔")])]),
    2: ("معتدل", [
        ("عوامُ النَّاسِ:", [("aqi", "باہر کی سرگرمیوں کی منصوبہ بندی کے لیے AQI کو مدنظر رکھیں۔")]),
        ("حساس گروہ:", [("heart", "صحت کی باقاعدگی سے نگرانی کریں (آکسیجن، بی پی، وغیرہ)۔"),
                        ("food", "قوتِ مدافعت بڑھانے کے لیے صحت بخش غذا کھائیں۔"),
                        ("nosmoke", "تمباکو نوشی سے گریز کریں۔"),
                        ("run", "باہر کی سرگرمیوں کو محدود کریں۔"),
                        ("kit", "ہنگامی امداد کے آلات (جیسے نیبولائزر) تیار رکھیں۔"),
                        ("doctor", "اگر سانس کے مسائل پیدا ہوں تو ڈاکٹر سے مشورہ کریں۔")])]),
    3: ("نازک گروہوں کے لئے غیر صحت بخش", [
        ("عام عوام کے لئے:", [("run", "بیرون خانہ طویل یا سخت مشقت کو کم کریں۔")]),
        ("نازک گروہ:", [("aqi", "باہر جانے سے پہلے AQI کی جانچ کریں۔"),
                        ("mask", "باہر جاتے ہوئے ماسک پہنیں۔"),
                        ("child", "بچوں کو گھر کے اندر رکھیں۔"),
                        ("home", "ناقص AQI والے علاقوں میں سفر سے گریز کریں۔"),
                        ("elder", "بزرگوں کو گھر کے اندر رہنا چاہئے۔"),
                        ("window", "دروازے/کھڑکیاں بند رکھیں۔"),
                        ("run", "بیرون خانہ سخت سرگرمی سے گریز کریں۔"),
                        ("doctor", "پھیپھڑوں اور دل کی پرانی بیماری کے مریض ماسک کے استعمال کے لئے ڈاکٹر سے مشورہ کریں۔")])]),
    4: ("مضر صحت", [
        ("عام عوام کے لئے:", [("aqi", "باہر جانے سے پہلے AQI کی جانچ کریں۔"),
                              ("run", "بیرون خانہ طویل یا سخت مشقت کو کم کریں۔")]),
        ("نازک گروہ (بشمول بچے اور بزرگ):", [("home", "زیادہ سے زیادہ وقت گھر پر گزاریں۔"),
                                              ("mask", "N95 ماسک کا استعمال کریں۔"),
                                              ("window", "دروازے اور کھڑکیاں بند رکھیں۔"),
                                              ("child", "بچوں کو گھر کے اندر رکھیں۔"),
                                              ("doctor", "CVD اور COPD کے مریض اپنے معالج کے مشورے سے ماسک کا انتخاب کریں۔")])]),
}


def advisory_sections(bi):
    """(title, sections) for the page-1 message box, where sections = [(heading, [(icon, text), ...]), ...]."""
    if bi is None:
        return None, []
    if bi in ADVISORY_TEXT:
        return ADVISORY_TEXT[bi]
    pub, sen = PUBLIC_MESSAGES[bi]
    return BANDS[bi][3], [(UR["public"], pub), (UR["sensitive"], sen)]


URDU_MONTHS = ["جنوری", "فروری", "مارچ", "اپریل", "مئی", "جون", "جولائی", "اگست",
               "ستمبر", "اکتوبر", "نومبر", "دسمبر"]
UR = {
    "title": "پنجاب کے فضائی معیار کی رپورٹ",
    "last24": "گزشتہ چوبیس گھنٹے",
    "message": "عوام کے نام پیغام",
    "avg": "اوسط AQI",
    "rank": "درجہ", "district": "ضلع", "causes": "بنیادی وجوہات",
    "public": "عوام الناس:", "sensitive": "حساس گروہ:",
    "nodata": "ڈیٹا دستیاب نہیں",
}

# QA thresholds (flags only — values are never altered)
QA_HIGH_AQI = 300               # flag any hour above this
QA_GAS_AQI = 150                # flag hours above this driven by a gas
QA_SPIKE_JUMP = 150             # isolated jump vs both neighbouring hours
GASES = {"O3", "CO", "SO2", "NO2", "NO", "NOX"}


# =============================================================================
# HELPERS
# =============================================================================

def round_half_up(x: float) -> int:
    return int(math.floor(x + 0.5))


def band_index(v: float | None) -> int | None:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    v = round_half_up(v)
    for i, (lo, hi, *_rest) in enumerate(BANDS):
        if v <= hi:
            return i
    return len(BANDS) - 1


def band_fill(v):
    i = band_index(v)
    return NO_DATA_FILL if i is None else BANDS[i][4]


def band_text(v):
    i = band_index(v)
    return "#000000" if i is None else BANDS[i][5]


def norm(s: str) -> str:
    s = str(s).lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def canonical_district(name: str) -> str | None:
    n = norm(name)
    if n in DISTRICT_ALIASES:
        return DISTRICT_ALIASES[n]
    for d in DISTRICT_URDU:
        if norm(d) == n:
            return d
    return None


def infer_district(station: str) -> str | None:
    """Find a district name inside a station name (longest match wins)."""
    n = " " + norm(station) + " "
    hits = []
    for d in DISTRICT_URDU:
        if f" {norm(d)} " in n:
            hits.append(d)
    for alias, d in DISTRICT_ALIASES.items():
        if f" {norm(alias)} " in n:
            hits.append(d)
    return max(hits, key=len) if hits else None


def short_label(station: str) -> str:
    s = re.sub(r"\s*-\s*Mobile\s*\d+\s*$", "", station)
    s = re.sub(r"[-\s]+LHR$", "", s)
    return s.strip()


def esc(s) -> str:
    return html.escape(str(s))


# =============================================================================
# 1. READ THE DASHBOARD CSV
# =============================================================================

@dataclass
class Station:
    name: str
    label: str
    district: str | None
    role: str
    hours: pd.DataFrame            # columns: time, aqi, dom, valid
    n_valid: int = 0
    mean: float | None = None
    value: int | None = None
    dominant: str = ""
    dom_counts: dict = field(default_factory=dict)
    max_aqi: float | None = None
    max_time: pd.Timestamp | None = None
    sufficient: bool = False
    x: float | None = None          # map coordinates (EPSG:3857)
    y: float | None = None


def registry_lookup(name: str):
    """STATION_REGISTRY entry for a CSV station name: exact match first, then the same name ignoring
    case / punctuation / spacing, then the registry short label found as whole words in the name
    (so 'Lathepur', 'LATHEPUR LHR Mobile-2' or 'Lathepur (Mobile 2)' still resolve to Lathepur)."""
    if name in STATION_REGISTRY:
        return STATION_REGISTRY[name]
    n = norm(name)
    for key, val in STATION_REGISTRY.items():
        if norm(key) == n:
            return val
    padded = f" {n} "
    hits = [val for key, val in STATION_REGISTRY.items() if f" {norm(val[0])} " in padded]
    return max(hits, key=lambda v: len(v[0])) if hits else None


def read_dashboard_csv(path: Path, data_date: dt.date | None):
    raw = pd.read_csv(path, encoding="utf-8-sig", dtype=str)
    tcol = raw.columns[0]
    t = pd.to_datetime(raw[tcol], errors="coerce", format="mixed")
    raw = raw[t.notna()].copy()
    raw["__t"] = t[t.notna()]
    if raw.empty:
        sys.exit("No hourly rows found in the CSV.")
    if data_date is None:  # the calendar day with the most hourly rows (latest if tie)
        counts = raw["__t"].dt.date.value_counts()
        data_date = max(counts[counts == counts.max()].index)
    raw = raw[raw["__t"].dt.date == data_date].sort_values("__t")

    names = [c[: -len(" • AQI")] for c in raw.columns if c.endswith(" • AQI")]
    if not names:
        sys.exit("CSV has no '<Station> • AQI' columns — is this the station-level export?")
    stations = []
    for name in names:
        aqi = pd.to_numeric(raw[f"{name} • AQI"], errors="coerce")
        dcol = f"{name} • Dominant Pollutant"
        dom = raw[dcol].fillna("").str.strip() if dcol in raw else pd.Series("", index=raw.index)
        valid = aqi.notna()
        hours = pd.DataFrame({"time": raw["__t"].values, "aqi": aqi.values,
                              "dom": dom.values, "valid": valid.values})
        reg = registry_lookup(name)
        if reg:
            label, district, role = reg
        else:
            label, district, role = short_label(name), infer_district(name), "city"
        stations.append(Station(name, label, district, role, hours))
    return stations, data_date


# =============================================================================
# 2. CALCULATIONS
# =============================================================================

def dominant_string(counter: Counter, n=3) -> str:
    # Counter.most_common keeps first-seen order for ties (station order, then hour)
    return " > ".join(p for p, _ in counter.most_common(n))


def isolated_aqi_spikes(hours: pd.DataFrame) -> list[int]:
    valid_indices = hours.index[hours.valid].tolist()
    spikes = []
    for position, hour_index in enumerate(valid_indices):
        if position == 0:
            comparison_indices = valid_indices[1:3]
        elif position == len(valid_indices) - 1:
            comparison_indices = valid_indices[-3:-1]
        else:
            comparison_indices = [valid_indices[position - 1], valid_indices[position + 1]]
        if len(comparison_indices) == 2 and all(
            hours.at[hour_index, "aqi"] - hours.at[index, "aqi"] > QA_SPIKE_JUMP
            for index in comparison_indices
        ):
            spikes.append(hour_index)
    return spikes


def compute_station(s: Station):
    spike_indices = isolated_aqi_spikes(s.hours)
    s.hours["spike"] = s.hours.index.isin(spike_indices)
    v = s.hours[s.hours.valid]
    s.n_valid = len(v)
    s.sufficient = s.n_valid >= MIN_VALID_HOURS
    c = Counter(d for d in v.dom if d)
    s.dom_counts = dict(c)
    if s.n_valid:
        i = v.aqi.idxmax()
        s.max_aqi, s.max_time = float(v.aqi[i]), v.time[i]
    if s.sufficient:
        s.mean = float(v.aqi.mean())
        s.value = round_half_up(s.mean)
        s.dominant = dominant_string(c)


@dataclass
class District:
    name: str
    stations: list
    mean: float | None = None
    value: int | None = None
    dominant: str = ""
    rank: int | None = None


def compute_districts(stations):
    groups: dict[str, list[Station]] = {name: [] for name in PUNJAB_DISTRICTS}   # all 36, even without stations
    for s in stations:
        if s.district in groups and (s.role != "transboundary" or s.label in COUNTED_IN_DISTRICT):
            groups[s.district].append(s)
    districts = []
    for name, sts in groups.items():
        d = District(name, sts)
        ok = [s for s in sts if s.sufficient]
        if ok:
            d.mean = float(np.mean([s.mean for s in ok]))
            d.value = round_half_up(d.mean)
            c = Counter()
            for s in ok:
                for p in s.hours[s.hours.valid].dom:
                    if p:
                        c[p] += 1
            d.dominant = dominant_string(c)
        districts.append(d)
    ranked = sorted([d for d in districts if d.value is not None], key=lambda d: (-d.value, -d.mean))
    for i, d in enumerate(ranked, 1):
        d.rank = i
    nodata = sorted([d for d in districts if d.value is None], key=lambda d: d.name)
    for i, d in enumerate(nodata, len(ranked) + 1):
        d.rank = i
    return ranked + nodata


def qa_flags(stations):
    rows = []
    for s in stations:
        h = s.hours
        n_zero = int(((h.aqi == 0) & (h.dom == "")).sum())
        n_missing = int(h.aqi.isna().sum())
        if len(h) < 24 or n_missing:
            rows.append((s.name, s.district, "", "Missing hours", f"{max(24 - len(h), 0) + n_missing} of 24 hours have no AQI"))
        if n_zero:
            rows.append((s.name, s.district, "", "AQI = 0, no pollutant",
                         f"{n_zero} hour(s) retained; not filtered from AQI calculations"))
        if not s.sufficient:
            rows.append((s.name, s.district, "", "Insufficient data",
                         f"{s.n_valid} hourly AQI values (< {MIN_VALID_HOURS}); station omitted from report"))
        for _, r in h[h.spike].iterrows():
            ts = pd.Timestamp(r.time).strftime("%d.%m.%Y %H:%M")
            rows.append((s.name, s.district, ts, "Isolated AQI peak",
                         f"AQI {r.aqi:.0f} ({r.dom or 'no pollutant'}); included in AQI calculations"))
        v = h[h.valid].reset_index(drop=True)
        for i, r in v.iterrows():
            ts = pd.Timestamp(r.time).strftime("%d.%m.%Y %H:%M")
            if r.aqi > QA_HIGH_AQI:
                rows.append((s.name, s.district, ts, "Very high hourly AQI", f"AQI {r.aqi:.0f} ({r.dom or 'no pollutant'})"))
            elif r.aqi > QA_GAS_AQI and r.dom in GASES:
                rows.append((s.name, s.district, ts, "Gas-driven high AQI", f"AQI {r.aqi:.0f} driven by {r.dom} — check analyser"))
        if s.district is None:
            rows.append((s.name, "", "", "Unknown district", "Add the station to STATION_REGISTRY"))
    return pd.DataFrame(rows, columns=["Station", "District", "Time", "Flag", "Detail"])


# =============================================================================
# 3. AQMS LOCATIONS (shapefile) -> match to CSV stations
# =============================================================================

def _match_score(a: str, b: str) -> float:
    na, nb = norm(a), norm(b)
    drop = {"lhr", "lahore", "mobile", "dc", "office", "govt", "the", "of", "and", "aqms", "station"}
    ta = {w for w in na.split() if w not in drop and not w.isdigit()}
    tb = {w for w in nb.split() if w not in drop and not w.isdigit()}
    seq = difflib.SequenceMatcher(None, na, nb).ratio()
    jac = len(ta & tb) / len(ta | tb) if (ta | tb) else 0
    contain = 1.0 if (ta and tb and (ta <= tb or tb <= ta)) else 0
    # same name written with different spacing ('Lathe Pur' vs 'Lathepur', 'Wahga Border' vs 'Wagha Border')
    ca = "".join(w for w in na.split() if w not in drop and not w.isdigit())
    cb = "".join(w for w in nb.split() if w not in drop and not w.isdigit())
    compact = 0.97 * difflib.SequenceMatcher(None, ca, cb).ratio() if (ca and cb and ca != cb and (" " in na or " " in nb)) else 0
    if ca and ca == cb:
        compact = 0.97
    return max(seq, jac, 0.9 * contain, compact)


def load_station_locations(shp: Path, stations, name_field: str | None, aliases: dict, log):
    import geopandas as gpd
    gdf = gpd.read_file(shp)
    if gdf.crs is None:
        minx, miny, maxx, maxy = gdf.total_bounds
        gdf = gdf.set_crs(4326 if (-180 <= minx <= 180 and -90 <= miny <= 90) else 32643)
        log(f"  ! {shp.name} has no CRS — assumed {gdf.crs.to_string()}")
    gdf = gdf[gdf.geometry.notna()].to_crs(3857)
    gdf["geometry"] = gdf.geometry.representative_point()

    csv_names = [s.name for s in stations]
    text_cols = [c for c in gdf.columns if c != "geometry" and (gdf[c].dtype == object or pd.api.types.is_string_dtype(gdf[c]))]
    if not text_cols:
        sys.exit(f"{shp.name} has no text attribute to match station names.")
    if name_field is None:   # pick the column whose values best match the CSV names
        best = max(text_cols, key=lambda c: np.mean([max(_match_score(str(v), n) for n in csv_names)
                                                      for v in gdf[c].fillna("")]))
        name_field = best
        log(f"  Station-name field detected: '{name_field}'  (override with --shp-name-field)")

    pairs = []
    for i, v in gdf[name_field].fillna("").items():
        v = str(v)
        if v in aliases:
            pairs.append((1.0, i, aliases[v]))
            continue
        for n in csv_names:
            pairs.append((_match_score(v, n), i, n))
    pairs.sort(reverse=True)
    used_rows, used_names, matched = set(), set(), {}
    for score, i, n in pairs:
        if score < 0.55 or i in used_rows or n in used_names:
            continue
        used_rows.add(i); used_names.add(n); matched[n] = (i, score)
    by_name = {s.name: s for s in stations}
    for n, (i, score) in matched.items():
        p = gdf.geometry[i]
        by_name[n].x, by_name[n].y = p.x, p.y
    unmatched_shp = gdf.loc[[i for i in gdf.index if i not in used_rows]]
    extra = [(str(r[name_field]), r.geometry.x, r.geometry.y) for _, r in unmatched_shp.iterrows()]
    match_table = pd.DataFrame(
        [(n, gdf[name_field][i], round(sc, 2)) for n, (i, sc) in matched.items()] +
        [(n, "— not in shapefile —", None) for n in csv_names if n not in matched] +
        [("— not in CSV —", e[0], None) for e in extra],
        columns=["CSV station", "Shapefile name", "Match score"])
    return extra, match_table


# =============================================================================
# 4. THE ZOOMED AQMS MAP
# =============================================================================

def _pick_font():
    from matplotlib import font_manager
    have = {f.name for f in font_manager.fontManager.ttflist}
    for f in ["Segoe UI", "Arial", "Liberation Sans", "DejaVu Sans"]:
        if f in have:
            return f
    return "sans-serif"


# -----------------------------------------------------------------------------
# BASEMAP (street map) — robust loader
#   Order tried:  1) a basemap file you supply (GeoTIFF, or PNG/JPG + world file)
#                 2) the saved copy from an earlier run (basemap/cache/)
#                 3) download from the chosen provider, then the fallbacks
#                 4) clean offline style (never fails)
#   Station locations do not move, so after ONE successful download the same
#   street map is reused every day without internet (use --refresh-basemap to
#   download a fresh copy).
# -----------------------------------------------------------------------------
_R = 6378137.0
_HALF = math.pi * _R

TILE_PROVIDERS = {
    # name: (URL template, tile size px, max zoom, attribution)
    "osm":         ("https://tile.openstreetmap.org/{z}/{x}/{y}.png", 256, 19,
                    "© OpenStreetMap contributors"),
    "voyager":     ("https://a.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}@2x.png", 512, 20,
                    "© OpenStreetMap contributors © CARTO"),
    "positron":    ("https://a.basemaps.cartocdn.com/light_all/{z}/{x}/{y}@2x.png", 512, 20,
                    "© OpenStreetMap contributors © CARTO"),
    "esri-street": ("https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}",
                    256, 19, "Esri, HERE, Garmin, © OpenStreetMap contributors"),
    "satellite":   ("https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",
                    256, 19, "Esri, Maxar, Earthstar Geographics"),
}


@dataclass
class Basemap:
    image: np.ndarray            # H x W x 3/4, uint8
    extent: tuple                # (left, right, bottom, top) in EPSG:3857
    attribution: str
    source: str


def _http_session():
    import requests
    try:                          # use the Windows / macOS certificate store
        import truststore         # (fixes SSL errors behind office proxies)
        truststore.inject_into_ssl()
    except Exception:
        pass
    s = requests.Session()
    s.headers.update({"User-Agent": BASEMAP_USER_AGENT, "Accept": "image/png,image/*;q=0.9,*/*;q=0.5"})
    return s


def _auto_zoom(extent_w_m: float, fig_px_w: int, max_zoom: int) -> int:
    # one tile pixel ~ 1/BASEMAP_DETAIL output pixels (street labels stay readable when printed)
    z = math.log2(2 * _HALF * fig_px_w * BASEMAP_DETAIL / (256 * extent_w_m))
    return int(max(3, min(max_zoom, round(z))))


def _tile_range(x0, y0, x1, y1, z):
    n = 2 ** z
    tx = lambda x: int(math.floor((x + _HALF) / (2 * _HALF) * n))
    ty = lambda y: int(math.floor((_HALF - y) / (2 * _HALF) * n))
    return max(tx(x0), 0), max(ty(y1), 0), min(tx(x1), n - 1), min(ty(y0), n - 1)


def _diagnose(err: Exception) -> str:
    msg = f"{type(err).__name__}: {str(err)[:160]}"
    low = msg.lower()
    if "ssl" in low or "certificate" in low:
        hint = "SSL check failed (office proxy?) -> run: pip install truststore"
    elif " 403" in low or "forbidden" in low or "blocked" in low:
        hint = "provider refused the request -> try --basemap voyager or esri-street"
    elif "proxy" in low or "connection" in low or "resolve" in low or "timed out" in low or "timeout" in low:
        hint = "no internet / proxy blocked -> use --basemap-file (see README) or run once on an open connection"
    else:
        hint = ""
    return msg + (f"\n      hint: {hint}" if hint else "")


def _download_tiles(name, url, ts, z, rng, log):
    from PIL import Image
    import io
    tx0, ty0, tx1, ty1 = rng
    n_tiles = (tx1 - tx0 + 1) * (ty1 - ty0 + 1)
    if n_tiles > 600:
        raise RuntimeError(f"{n_tiles} tiles needed at zoom {z} — extent too large")
    sess = _http_session()
    canvas = Image.new("RGBA", ((tx1 - tx0 + 1) * ts, (ty1 - ty0 + 1) * ts), (240, 240, 240, 255))
    for tx in range(tx0, tx1 + 1):
        for ty in range(ty0, ty1 + 1):
            u = url.format(z=z, x=tx, y=ty, s="a", r="")
            last = None
            for attempt in range(3):
                try:
                    r = sess.get(u, timeout=BASEMAP_TIMEOUT)
                    if r.status_code != 200:
                        raise RuntimeError(f"HTTP {r.status_code} from {u.split('/')[2]}")
                    tile = Image.open(io.BytesIO(r.content)).convert("RGBA")
                    if tile.size != (ts, ts):
                        tile = tile.resize((ts, ts), Image.LANCZOS)
                    canvas.paste(tile, ((tx - tx0) * ts, (ty - ty0) * ts))
                    last = None
                    break
                except Exception as e:  # retry, then give up on this provider
                    last = e
            if last is not None:
                raise last
    n = 2 ** z
    left = tx0 / n * 2 * _HALF - _HALF
    right = (tx1 + 1) / n * 2 * _HALF - _HALF
    top = _HALF - ty0 / n * 2 * _HALF
    bottom = _HALF - (ty1 + 1) / n * 2 * _HALF
    return canvas, (left, right, bottom, top), n_tiles


def _read_world_file(img_path: Path):
    for ext in (".pgw", ".pngw", ".jgw", ".jpgw", ".tfw", ".wld"):
        wf = img_path.with_suffix(ext)
        if wf.exists():
            v = [float(x) for x in wf.read_text().split()[:6]]
            return v  # A, D, B, E, C, F
    return None


def _load_basemap_file(path: Path, log) -> Basemap | None:
    from PIL import Image
    path = Path(path)
    if not path.exists():
        log(f"  ! Basemap file not found: {path}")
        return None
    suffix = path.suffix.lower()
    if suffix in (".tif", ".tiff"):
        try:
            import rasterio
            from rasterio.warp import calculate_default_transform, reproject, Resampling
        except ImportError:
            log("  ! Reading a GeoTIFF basemap needs rasterio:  pip install rasterio")
            return None
        with rasterio.open(path) as src:
            bands = min(src.count, 4)
            if src.crs and src.crs.to_epsg() != 3857:
                tr, w, h = calculate_default_transform(src.crs, "EPSG:3857", src.width, src.height, *src.bounds)
                data = np.zeros((bands, h, w), dtype=np.uint8)
                for b in range(bands):
                    reproject(rasterio.band(src, b + 1), data[b], src_transform=src.transform, src_crs=src.crs,
                              dst_transform=tr, dst_crs="EPSG:3857", resampling=Resampling.bilinear)
                left, top = tr.c, tr.f
                right, bottom = left + tr.a * w, top + tr.e * h
            else:
                data = src.read(list(range(1, bands + 1)))
                if data.dtype != np.uint8:
                    data = np.clip(data / max(data.max(), 1) * 255, 0, 255).astype(np.uint8)
                left, bottom, right, top = src.bounds
        img = np.moveaxis(data, 0, -1)
        if img.shape[2] == 1:
            img = np.repeat(img, 3, axis=2)
        return Basemap(img, (left, right, bottom, top), BASEMAP_FILE_ATTRIBUTION, f"file {path.name}")
    # PNG / JPG + world file (QGIS: Project > Import/Export > Export Map to Image, tick "world file")
    wf = _read_world_file(path)
    if wf is None:
        log(f"  ! {path.name} has no world file (.pgw/.jgw/.wld) next to it — cannot place it on the map")
        return None
    A, D, B, E, C, F = wf
    im = Image.open(path).convert("RGBA")
    w, h = im.size
    left, top = C - A / 2, F - E / 2
    right, bottom = left + A * w, top + E * h
    prj = path.with_suffix(".prj")
    if prj.exists():     # not Web Mercator? convert the corner coordinates
        from pyproj import CRS, Transformer
        crs = CRS.from_wkt(prj.read_text())
        if crs.to_epsg() != 3857:
            t = Transformer.from_crs(crs, 3857, always_xy=True)
            (left, right), (bottom, top) = t.transform([left, right], [bottom, top])
    elif abs(left) <= 180 and abs(top) <= 90:   # looks like degrees
        from pyproj import Transformer
        t = Transformer.from_crs(4326, 3857, always_xy=True)
        (left, right), (bottom, top) = t.transform([left, right], [bottom, top])
    return Basemap(np.asarray(im), (left, right, bottom, top), BASEMAP_FILE_ATTRIBUTION, f"file {path.name}")


def load_basemap(bounds, fig_px_w, style, basemap_file, refresh, focus, log) -> Basemap | None:
    """bounds = (x0, y0, x1, y1) in EPSG:3857."""
    from PIL import Image
    import json
    x0, y0, x1, y1 = bounds

    # 1) basemap file: explicit, or dropped into basemap/ as <Focus>.tif / <Focus>.png (+ .pgw)
    cands = [Path(basemap_file)] if basemap_file else [
        BASEMAP_DIR / f"{focus}{ext}" for ext in (".tif", ".tiff", ".png", ".jpg")]
    for c in cands:
        if c.exists() or basemap_file:
            bm = _load_basemap_file(c, log)
            if bm is not None:
                l, r, b, t = bm.extent
                if l > x0 or r < x1 or b > y0 or t < y1:
                    log(f"  ! {c.name} does not fully cover the map area — edges will be blank")
                log(f"  Basemap: {bm.source}")
                return bm
    if style == "none":
        return None

    # 2/3) tiles: chosen provider first, then fallbacks
    order = [style] + [p for p in BASEMAP_FALLBACKS if p != style]
    cache_dir = BASEMAP_DIR / "cache"
    for name in order:
        if name in TILE_PROVIDERS:
            url, ts, maxz, attr = TILE_PROVIDERS[name]
        elif "{z}" in str(name):
            url, ts, maxz, attr = name, 256, 19, "custom tiles"
        else:
            log(f"  ! Unknown basemap '{name}' — skipped")
            continue
        z = _auto_zoom(x1 - x0, fig_px_w, maxz) if BASEMAP_ZOOM == "auto" else int(BASEMAP_ZOOM)
        rng = _tile_range(x0, y0, x1, y1, z)
        key = re.sub(r"[^A-Za-z0-9]+", "_", name if name in TILE_PROVIDERS else "custom")[:20]
        cpng = cache_dir / f"{key}_z{z}_{rng[0]}_{rng[1]}_{rng[2]}_{rng[3]}.png"
        cjson = cpng.with_suffix(".json")
        if cpng.exists() and cjson.exists() and not refresh:
            meta = json.loads(cjson.read_text())
            log(f"  Basemap: {name} (saved copy from {meta.get('downloaded', '?')}, zoom {z})")
            return Basemap(np.asarray(Image.open(cpng).convert("RGBA")), tuple(meta["extent"]), meta["attribution"],
                           f"{name} (cached)")
        try:
            img, ext, n = _download_tiles(name, url, ts, z, rng, log)
            cache_dir.mkdir(parents=True, exist_ok=True)
            img.save(cpng)
            cjson.write_text(json.dumps({"extent": ext, "attribution": attr, "provider": name, "zoom": z,
                                         "downloaded": dt.date.today().isoformat()}))
            log(f"  Basemap: {name} downloaded ({n} tiles, zoom {z}) — saved for offline reuse")
            return Basemap(np.asarray(img), ext, attr, name)
        except Exception as e:
            log(f"  ! Basemap '{name}' failed — {_diagnose(e)}")
            # an older saved copy of this provider is better than nothing
            old = sorted(cache_dir.glob(f"{key}_z*.png")) if cache_dir.exists() else []
            for o in old:
                j = o.with_suffix(".json")
                if j.exists():
                    meta = json.loads(j.read_text())
                    l, r, b, t = meta["extent"]
                    if l <= x0 and r >= x1 and b <= y0 and t >= y1:
                        log(f"  Basemap: {name} (older saved copy {o.name})")
                        return Basemap(np.asarray(Image.open(o).convert("RGBA")), tuple(meta["extent"]),
                                       meta["attribution"], f"{name} (cached)")
    log("  ! No street map available — using the offline map style")
    return None


def test_basemaps(log):
    """--test-basemap: try every provider on a small Lahore area and report."""
    lon0, lat0, lon1, lat1 = 74.25, 31.45, 74.40, 31.60
    fx = lambda lon: math.radians(lon) * _R
    fy = lambda lat: _R * math.log(math.tan(math.pi / 4 + math.radians(lat) / 2))
    x0, y0, x1, y1 = fx(lon0), fy(lat0), fx(lon1), fy(lat1)
    log("Testing street-map providers (Lahore, zoom 12):")
    for name, (url, ts, maxz, attr) in TILE_PROVIDERS.items():
        try:
            _, _, n = _download_tiles(name, url, ts, 12, _tile_range(x0, y0, x1, y1, 12), log)
            log(f"  OK    {name:<12} ({n} tiles)")
        except Exception as e:
            log(f"  FAIL  {name:<12} {_diagnose(e)}")


def _plot_lines(ax, geoms, **kw):
    """Plot (multi)line / polygon boundaries with plain matplotlib (supports dash tuples)."""
    for g in geoms:
        if g is None or g.is_empty:
            continue
        if g.geom_type in ("Polygon", "MultiPolygon"):
            g = g.boundary
        parts = getattr(g, "geoms", [g])
        for part in parts:
            if part.geom_type in ("LineString", "LinearRing"):
                x, y = part.xy
                ax.plot(x, y, **kw)
            elif hasattr(part, "geoms"):
                _plot_lines(ax, part.geoms, **kw)


LAHORE_MAP_SIZE_MM = (196, 145)    # landscape AQMS map; must match .p2 .mapimg in the CSS


def render_map(stations, focus: str, extra_points, districts_shp: Path | None,
               out_png: Path, data_date: dt.date, basemap: str, log, size_mm=LAHORE_MAP_SIZE_MM,
               basemap_file: Path | None = None, refresh_basemap: bool = False):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch, Rectangle, Circle, PathPatch
    from matplotlib.path import Path as MPath
    import matplotlib.patheffects as pe
    import geopandas as gpd
    from shapely.geometry import box

    font = _pick_font()
    plt.rcParams["font.family"] = font

    pts = [s for s in stations if in_focus(s, focus) and s.x is not None]
    if not pts:
        log(f"  ! No located stations for {focus}; map skipped")
        return None
    extent_pts = pts   # city + transboundary stations all stay in view
    xs = np.array([s.x for s in extent_pts]); ys = np.array([s.y for s in extent_pts])

    W, H = size_mm[0] / 25.4, size_mm[1] / 25.4
    fig = plt.figure(figsize=(W, H), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()

    # extent: tight fit on all stations incl. transboundary (+2% padding), matched to the figure aspect ratio
    cx, cy = (xs.min() + xs.max()) / 2, (ys.min() + ys.max()) / 2
    w = (xs.max() - xs.min()) * 1.18 + 1500      # margin keeps edge stations (e.g. Wagha) clear of the corner panels
    h = (ys.max() - ys.min()) * 1.18 + 1500
    if w / h > W / H:
        h = w * H / W
    else:
        w = h * W / H
    x0, x1, y0, y1 = cx - w / 2, cx + w / 2, cy - h / 2, cy + h / 2
    ax.set_xlim(x0, x1); ax.set_ylim(y0, y1)
    extent = box(x0, y0, x1, y1)
    pts = [s for s in pts if x0 <= s.x <= x1 and y0 <= s.y <= y1]

    bm = load_basemap((x0, y0, x1, y1), int(fig.bbox.width), basemap, basemap_file, refresh_basemap, focus, log)
    online = bm is not None
    if online:
        ax.imshow(bm.image, extent=bm.extent, origin="upper", interpolation="lanczos", aspect="auto", zorder=0)
        ax.set_xlim(x0, x1); ax.set_ylim(y0, y1)
    if not online:
        ax.add_patch(Rectangle((x0, y0), w, h, facecolor="#EEF1EC", zorder=0))
        # faint graticule
        lat0 = math.degrees(2 * math.atan(math.exp(y0 / 6378137)) - math.pi / 2)
        lat1 = math.degrees(2 * math.atan(math.exp(y1 / 6378137)) - math.pi / 2)
        lon0, lon1 = math.degrees(x0 / 6378137), math.degrees(x1 / 6378137)
        step = 0.05
        for lon in np.arange(math.ceil(lon0 / step) * step, lon1, step):
            X = math.radians(lon) * 6378137
            ax.plot([X, X], [y0, y1], color="#D5DAD2", lw=0.4, zorder=1)
        for lat in np.arange(math.ceil(lat0 / step) * step, lat1, step):
            Y = 6378137 * math.log(math.tan(math.pi / 4 + math.radians(lat) / 2))
            ax.plot([x0, x1], [Y, Y], color="#D5DAD2", lw=0.4, zorder=1)
    else:
        if BASEMAP_FADE > 0:
            ax.add_patch(Rectangle((x0, y0), w, h, facecolor="white", alpha=BASEMAP_FADE, lw=0, zorder=1))

    # district boundaries + focus highlight
    dg_all, name_col = None, None
    focus_geom, border_geom = None, None
    if districts_shp and Path(districts_shp).exists():
        dg = gpd.read_file(districts_shp)
        if dg.crs is None:
            dg = dg.set_crs(4326)
        dg = dg.to_crs(3857)
        name_col = max([c for c in dg.columns if c != "geometry" and (dg[c].dtype == object or pd.api.types.is_string_dtype(dg[c]))] or [None],
                       key=lambda c: 0 if c is None else dg[c].astype(str).map(canonical_district).notna().sum())
        dg_all = dg.copy()
        dg = dg[dg.intersects(extent)]
        if not online:
            dg.plot(ax=ax, facecolor="#F7F8F4", edgecolor="#B9BFB5", lw=0.6, zorder=2)
        else:
            _plot_lines(ax, dg.geometry, color="#5F6B73", lw=0.6, ls=(0, (4, 2)), zorder=3)
        if name_col:
            fd = dg[dg[name_col].astype(str).map(canonical_district) == focus]
            if len(fd):
                geom = fd.union_all() if hasattr(fd, "union_all") else fd.unary_union
                focus_geom = geom
                mask = extent.difference(geom)
                gpd.GeoSeries([mask], crs=3857).plot(ax=ax, color="white", alpha=0.45 if online else 0.6, lw=0, zorder=3)
                _plot_lines(ax, [geom], color="white", lw=4.5, alpha=0.9, zorder=4)
                _plot_lines(ax, [geom], color="#1F4E79", lw=1.6, zorder=4)

    # international border
    if BORDER_FILE.exists():
        bd = gpd.read_file(BORDER_FILE).to_crs(3857)
        bd = bd[bd.intersects(extent)]
        if len(bd):
            _plot_lines(ax, bd.geometry, color="white", lw=3.2, alpha=0.8, zorder=4)
            _plot_lines(ax, bd.geometry, color="#8B1E1E", lw=1.3, ls=(0, (6, 2, 1, 2)), zorder=4)
            border_geom = bd.union_all() if hasattr(bd, "union_all") else bd.unary_union

    # geopandas may have changed limits/aspect: lock the extent again
    ax.set_aspect("auto"); ax.set_xlim(x0, x1); ax.set_ylim(y0, y1)

    # SHP stations without data in the CSV (grey)
    extra_artists = []
    for nm, ex, ey in extra_points:
        if x0 < ex < x1 and y0 < ey < y1:
            ax.scatter([ex], [ey], s=26, color="#9E9E9E", edgecolor="white", lw=0.8, zorder=6)
            extra_artists.append(ax.annotate(f"{nm} (no data)", (ex, ey), xytext=(5, -3), textcoords="offset points",
                                 fontsize=5.5, color="#6B6B6B", zorder=6,
                                 path_effects=[pe.withStroke(linewidth=2, foreground="white")]))

    # station markers
    for s in pts:
        col = band_fill(s.value)
        if s.role == "transboundary":
            ax.scatter([s.x], [s.y], s=230, color="#E03C31", alpha=0.18, lw=0, zorder=7)
            ax.scatter([s.x], [s.y], s=70, color="#E03C31", edgecolor="white", lw=1.4, zorder=8)
            ax.text(s.x, s.y, "T", ha="center", va="center_baseline", fontsize=5.5, color="white",
                    fontweight="bold", zorder=9)
        else:
            ax.scatter([s.x], [s.y], s=260, color=col, alpha=0.30, lw=0, zorder=7)
            ax.scatter([s.x], [s.y], s=62, color=col, edgecolor="#1B1B1B", lw=1.0, zorder=8)
            ax.scatter([s.x], [s.y], s=6, color="#1B1B1B", lw=0, zorder=9)

    # ---- legend panel (corner with fewest stations) ----
    fig.canvas.draw()
    R = fig.canvas.get_renderer()
    def to_disp(x, y):
        return ax.transData.transform((x, y))
    pix = np.array([to_disp(s.x, s.y) for s in pts])
    fw, fh = fig.bbox.width, fig.bbox.height
    lw_, lh_ = 46 / size_mm[0], 60 / size_mm[1]          # legend panel: 46 x 60 mm whatever the map shape
    iw, ih = 40 / size_mm[0], 54 / size_mm[1]            # locator inset: 40 x 54 mm

    def _panel(c, wf, hf, bottom):
        x = 0.012 * fw if c[1] == "l" else fw * (1 - wf - 0.012)
        y = fh * (1 - hf - 0.015) if c[0] == "t" else fh * bottom
        return (x, y, x + wf * fw, y + hf * fh)

    def _arrow_panel(c):
        axf = 0.955 if c[1] == "r" else 0.045
        ayf = 0.90 if c[0] == "t" else 0.16
        return ((axf - 0.03) * fw, (ayf - 0.03) * fh, (axf + 0.03) * fw, (ayf + 0.10) * fh)

    def _stations_under(rect, pad):
        r0, r1, r2, r3 = rect
        return int(((pix[:, 0] > r0 - pad) & (pix[:, 0] < r2 + pad) & (pix[:, 1] > r1 - pad) & (pix[:, 1] < r3 + pad)).sum())

    names_ = ["tl", "bl", "tr", "br"]
    has_locator = dg_all is not None and name_col
    if has_locator:
        ix, iy = 1 - iw - 0.012, 1 - ih - 0.015
        inset_rect = (ix * fw, iy * fh, (ix + iw) * fw, (iy + ih) * fh)
        best_asg = None
        for leg_corner_ in names_:
            lx = 0.012 if leg_corner_[1] == "l" else 1 - lw_ - 0.012
            ly = 1 - lh_ - 0.015 if leg_corner_[0] == "t" else 0.06
            leg_rect = (lx * fw, ly * fh, (lx + lw_) * fw, (ly + lh_) * fh)
            for arrow_corner in names_:
                arrow_rect = _arrow_panel(arrow_corner)
                rects = [leg_rect, inset_rect, arrow_rect]
                cost = sum(10000 * _stations_under(r, 15) + 300 * _stations_under(r, 40) for r in rects)
                for first, second in ((leg_rect, inset_rect), (leg_rect, arrow_rect), (inset_rect, arrow_rect)):
                    overlap_w = max(0, min(first[2], second[2]) - max(first[0], second[0]))
                    overlap_h = max(0, min(first[3], second[3]) - max(first[1], second[1]))
                    cost += 10000 * overlap_w * overlap_h
                cost += 10 * names_.index(leg_corner_) + 3 * names_.index(arrow_corner)
                if best_asg is None or cost < best_asg[0]:
                    best_asg = (cost, leg_corner_, arrow_corner, lx, ly)
        _, leg_corner, arrow_corner, lx, ly = best_asg
    else:
        import itertools
        best_asg = None
        for lc_, ic_, ac_ in itertools.permutations(names_, 3):
            rects_ = [(_panel(lc_, lw_, lh_, 0.06), 70), (_panel(ic_, iw, ih, 0.075), 70), (_arrow_panel(ac_), 40)]
            cost_ = sum(10000 * _stations_under(r_, 15) + 300 * _stations_under(r_, pad_) for r_, pad_ in rects_)
            cost_ += 10 * names_.index(lc_) + 3 * names_.index(ic_) + 3 * ["tr", "tl", "br", "bl"].index(ac_) + (500 if lc_ == "br" else 0)
            if best_asg is None or cost_ < best_asg[0]:
                best_asg = (cost_, lc_, ic_, ac_)
        _, leg_corner, ins_corner, arrow_corner = best_asg
        lx = 0.012 if leg_corner[1] == "l" else 1 - lw_ - 0.012
        ly = 1 - lh_ - 0.015 if leg_corner[0] == "t" else 0.06

    lax = fig.add_axes([lx, ly, lw_, lh_]); lax.set_xlim(0, 1); lax.set_ylim(0, 1); lax.set_axis_off()
    lax.add_patch(FancyBboxPatch((0.02, 0.02), 0.96, 0.96, boxstyle="round,pad=0,rounding_size=0.04",
                                 facecolor="white", edgecolor="#9AA5AE", lw=0.8, alpha=0.94))
    lax.text(0.5, 0.93, "AQI limits of EPA Punjab", ha="center", va="center", fontsize=7, fontweight="bold", color="#1F2D3A")
    yy = 0.83
    for i, (lo, hi, en, _ur, fill, _tc) in enumerate(BANDS):
        lax.add_patch(FancyBboxPatch((0.07, yy - 0.035), 0.16, 0.07, boxstyle="round,pad=0,rounding_size=0.015",
                                     facecolor=fill, edgecolor="#555", lw=0.3))
        lax.text(0.27, yy, BAND_LABELS[i], va="center", fontsize=6, color="#222", fontweight="bold")
        lax.text(0.50, yy, en if i != 3 else "Unhealthy (Sensitive)", va="center", fontsize=5.6, color="#333")
        yy -= 0.083
    yy -= 0.01
    lax.plot([0.07, 0.93], [yy + 0.035, yy + 0.035], color="#D0D5DA", lw=0.5)
    lax.scatter([0.15], [yy - 0.01], s=28, color="#FFFF00", edgecolor="#1B1B1B", lw=0.7)
    lax.text(0.27, yy - 0.01, "AQMS", va="center", fontsize=5.8)
    if any(s.role == "transboundary" for s in pts):
        lax.scatter([0.15], [yy - 0.085], s=34, color="#E03C31", edgecolor="white", lw=0.7)
        lax.text(0.15, yy - 0.085, "T", ha="center", va="center_baseline", fontsize=4, color="white", fontweight="bold")
        lax.text(0.27, yy - 0.085, "Transboundary AQMS", va="center", fontsize=5.8)
    leg_box = lax.get_window_extent(R)


    # Punjab locator inset over the India-side map area, with Lahore highlighted
    if has_locator:
        iax = fig.add_axes([ix, iy, iw, ih]); iax.set_axis_off()
        iax.set_zorder(10)
        iax.patch.set_facecolor("white")
        dg_all.plot(ax=iax, facecolor="#F4F6F2", edgecolor="#A7B0A4", lw=0.25)
        foc = dg_all[dg_all[name_col].astype(str).map(canonical_district) == focus]
        if len(foc):
            foc.plot(ax=iax, facecolor="#1F4E79", edgecolor="#12334F", lw=0.65)
            focus_geom_inset = (foc.geometry.union_all() if hasattr(foc.geometry, "union_all")
                                else foc.geometry.unary_union)
            focus_point = focus_geom_inset.representative_point()
            iax.scatter([focus_point.x], [focus_point.y], s=22, color="#E03C31", edgecolor="white",
                        linewidth=1.0, zorder=6)
        iax.add_patch(Rectangle((x0, y0), w, h, fill=False, ec="#C0392B", lw=0.9))
        bx0, by0, bx1, by1 = dg_all.total_bounds
        pad = 0.04 * max(bx1 - bx0, by1 - by0)
        iax.set_xlim(bx0 - pad, bx1 + pad); iax.set_ylim(by0 - pad, by1 + pad); iax.set_aspect("equal")
        iax.add_patch(Rectangle((0, 0.86), 1, 0.14, transform=iax.transAxes, facecolor="white",
                                edgecolor="none", alpha=0.94, zorder=7))
        iax.text(0.5, 0.93, "PUNJAB", transform=iax.transAxes, ha="center", va="center",
                 fontsize=7, fontweight="bold", color="#1F2D3A", zorder=8)
        iax.add_patch(Rectangle((0, 0), 1, 0.13, transform=iax.transAxes, facecolor="white",
                                edgecolor="none", alpha=0.94, zorder=7))
        iax.scatter([0.18], [0.065], transform=iax.transAxes, s=19, color="#E03C31",
                    edgecolor="white", linewidth=0.8, zorder=8)
        iax.text(0.29, 0.065, "Lahore focus", transform=iax.transAxes, ha="left", va="center",
                 fontsize=5.8, color="#1F2D3A", zorder=8)
        iax.add_patch(FancyBboxPatch((0.01, 0.01), 0.98, 0.98, transform=iax.transAxes,
                                     boxstyle="round,pad=0,rounding_size=0.03", fc="none",
                                     ec="#526273", lw=1.0, zorder=9, clip_on=False))
        fig.canvas.draw()
        inset_box = iax.get_window_extent(R)
    else:
        inset_box = None

    # north arrow
    ax_x = 0.955 if arrow_corner[1] == "r" else 0.045
    ax_y = 0.90 if arrow_corner[0] == "t" else 0.16
    ax.annotate("", xy=(ax_x, ax_y + 0.055), xytext=(ax_x, ax_y - 0.01), xycoords="axes fraction",
                arrowprops=dict(arrowstyle="-|>,head_width=0.35,head_length=0.7", color="#1F2D3A", lw=1.4))
    ax.text(ax_x, ax_y + 0.07, "N", transform=ax.transAxes, ha="center", fontsize=8, fontweight="bold", color="#1F2D3A",
            path_effects=[pe.withStroke(linewidth=2.5, foreground="white")])
    arrow_box = ax.transAxes.transform([[ax_x - 0.03, ax_y - 0.03], [ax_x + 0.03, ax_y + 0.1]])

    # scale bar (bottom-left unless legend is there)
    lat_c = math.degrees(2 * math.atan(math.exp(cy / 6378137)) - math.pi / 2)
    m_per_unit = math.cos(math.radians(lat_c))
    km = [1, 2, 5, 10, 20][int(np.argmin([abs(k * 1000 / m_per_unit - w * 0.18) for k in [1, 2, 5, 10, 20]]))]
    L = km * 1000 / m_per_unit
    sb_left = leg_corner != "bl"
    sx = x0 + w * (0.03 if sb_left else 0.73); sy = y0 + h * 0.035
    for k in range(4):
        ax.add_patch(Rectangle((sx + k * L / 4, sy), L / 4, h * 0.010,
                               facecolor="#1F2D3A" if k % 2 == 0 else "white", edgecolor="#1F2D3A", lw=0.6, zorder=12))
    ax.text(sx, sy + h * 0.018, "0", fontsize=5.5, ha="center", zorder=12, path_effects=[pe.withStroke(linewidth=2, foreground="white")])
    ax.text(sx + L, sy + h * 0.018, f"{km} km", fontsize=5.5, ha="center", zorder=12, path_effects=[pe.withStroke(linewidth=2, foreground="white")])
    sb_box = ax.transData.transform([[sx - 500, sy - 300], [sx + L + 1500, sy + h * 0.05]])

    # title tag + attribution
    tag = ax.text(0.985 if sb_left else 0.015, 0.042, f"AQMS Map — {focus}  |  24-h AQI {data_date:%d.%m.%Y}",
                  transform=ax.transAxes, ha="right" if sb_left else "left", va="bottom", fontsize=6.6,
                  fontweight="bold", color="#1F2D3A", zorder=12,
                  bbox=dict(boxstyle="round,pad=0.35", fc="white", ec="#1F2D3A", lw=0.8))
    if online:
        attr = f"Basemap: {bm.attribution}"
        ax.text(0.985 if sb_left else 0.015, 0.005, attr, transform=ax.transAxes, ha="right" if sb_left else "left",
                va="bottom", fontsize=4, color="#555", zorder=12)
    fig.canvas.draw()
    tag_box = tag.get_window_extent(R)

    # ---- AQI callout labels with collision avoidance ----
    # Label geometry is computed analytically (fast, and independent of the leader line),
    # then each label is drawn once at the best free position.
    from matplotlib.transforms import Bbox
    PT = fig.dpi / 72.0
    NUM_FS, NAME_FS, NUM_PAD, NAME_PAD = 10, 5.6, 0.28, 0.22

    def measure(txt, fs):
        t = ax.text(0, 0, txt, fontsize=fs, fontweight="bold", transform=None)
        e = t.get_window_extent(R)
        t.remove()
        return e.width, e.height

    # ---- context labels: adjoining districts and INDIA, set in the free space around the stations ----
    ctx_boxes = []
    fixed = [leg_box, tag_box, Bbox(arrow_box), Bbox(sb_box)] + ([inset_box] if inset_box is not None else [])
    axpad = ax.get_window_extent(R).padded(-8)

    def _ov(b1, b2):
        return max(0, min(b1.x1, b2.x1) - max(b1.x0, b2.x0)) * max(0, min(b1.y1, b2.y1) - max(b1.y0, b2.y0))

    def spaced(txt):
        return "   ".join(" ".join(word) for word in txt.split())

    def place_label(txt, cands, fs, rot=0, color="#56626E", anchor=None):
        """Write `txt` at the candidate (data x, y) that clears stations, panels and other labels."""
        t = ax.text(0, 0, txt, fontsize=fs, fontweight="bold", fontstyle="italic", rotation=rot, ha="center",
                    va="center", color=color, zorder=5, path_effects=[pe.withStroke(linewidth=2.4, foreground="white")])
        e = t.get_window_extent(R)
        tw, th = e.width, e.height
        inv = ax.transData.inverted()
        lines_ = ([focus_geom.boundary] if focus_geom is not None else []) + ([border_geom] if border_geom is not None else [])
        best = None
        for x, y in cands:
            px_, py_ = to_disp(x, y)
            b = Bbox([[px_ - tw / 2 - 4, py_ - th / 2 - 4], [px_ + tw / 2 + 4, py_ + th / 2 + 4]])
            if not (axpad.x0 <= b.x0 and b.x1 <= axpad.x1 and axpad.y0 <= b.y0 and b.y1 <= axpad.y1):
                continue
            cost = 10 * sum(_ov(b, o) for o in fixed + ctx_boxes)
            (bx_a, by_a), (bx_b, by_b) = inv.transform((b.x0, b.y0)), inv.transform((b.x1, b.y1))
            bpoly = box(min(bx_a, bx_b), min(by_a, by_b), max(bx_a, bx_b), max(by_a, by_b))
            cost += 4000 * sum(1 for g in lines_ if bpoly.intersects(g))      # never write across a boundary line
            if anchor is not None:
                ax_pt = to_disp(*anchor)
                cost += 1.5 * math.hypot(px_ - ax_pt[0], py_ - ax_pt[1])
            for sx_, sy_ in pix:                              # keep clear of every station and its call-out room
                dx = max(b.x0 - sx_, 0, sx_ - b.x1); dy = max(b.y0 - sy_, 0, sy_ - b.y1)
                cost += max(0.0, 110 - math.hypot(dx, dy)) ** 2
            if best is None or cost < best[0]:
                best = (cost, x, y, b)
        if best is None:
            t.remove()
            return False
        t.set_position((best[1], best[2]))
        ctx_boxes.append(best[3])
        return True

    if dg_all is not None and name_col and focus_geom is not None:
        from shapely.geometry import Point
        for _, r in dg_all[dg_all.geometry.intersects(focus_geom.buffer(500))].iterrows():
            nm = canonical_district(str(r[name_col])) or str(r[name_col])
            if nm == focus or nm in {"Kasur", "Nankana Sahib"}:
                continue
            vis = r.geometry.intersection(extent)
            part = None if vis.is_empty else (max(vis.geoms, key=lambda g: g.area) if hasattr(vis, "geoms") else vis)
            if part is None or part.area < 0.02 * extent.area:
                rp = r.geometry.representative_point()                  # outside the frame: name it at the frame edge
                ddx, ddy = (rp.x - cx) / w, (rp.y - cy) / h
                if abs(ddy) >= abs(ddx):
                    word = "north" if ddy > 0 else "south"
                    yy_ = (y1 - 0.045 * h, y1 - 0.075 * h) if ddy > 0 else (y0 + 0.075 * h, y0 + 0.105 * h)
                    cands = [(x0 + f * w, yv) for yv in yy_ for f in np.linspace(0.12, 0.88, 17)]
                    place_label(f"{spaced(nm.upper())}   ({word})", cands, 7.5)
                else:
                    word = "east" if ddx > 0 else "west"
                    xx_ = (x1 - 0.035 * w, x1 - 0.07 * w) if ddx > 0 else (x0 + 0.035 * w, x0 + 0.07 * w)
                    cands = [(xv, y0 + f * h) for xv in xx_ for f in np.linspace(0.12, 0.88, 17)]
                    place_label(f"{spaced(nm.upper())}   ({word})", cands, 7.5, rot=90)
                continue
            inner = part.buffer(-0.025 * min(w, h))
            inner = part if inner.is_empty else inner
            mnx, mny, mxx, mxy = inner.bounds
            grid = [Point(mnx + (i + 0.5) * (mxx - mnx) / 16, mny + (j + 0.5) * (mxy - mny) / 16)
                    for i in range(16) for j in range(16)]
            rpt = part.representative_point()
            place_label(spaced(nm.upper()), [(q.x, q.y) for q in grid if inner.contains(q)], 7.5, anchor=(rpt.x, rpt.y))

    if border_geom is not None:
        from shapely.geometry import LineString
        cands = []
        for f in np.linspace(0.12, 0.88, 25):
            yb = y0 + f * h
            hit = LineString([(x0 - 10, yb), (x1 + 10, yb)]).intersection(border_geom)
            xs_hit = [g.x for g in (hit.geoms if hasattr(hit, "geoms") else [hit]) if g.geom_type == "Point"]
            if not xs_hit:
                continue
            xb = max(xs_hit)
            for fr in (0.35, 0.5, 0.65):
                if xb + (x1 - xb) * fr < x1 - 0.03 * w:
                    cands.append((xb + (x1 - xb) * fr, yb))
        cands += [(x1 - 0.05 * w, y0 + f * h) for f in np.linspace(0.2, 0.8, 13)]     # fallback: right margin
        place_label(spaced("INDIA"), cands, 13, rot=90, color="#7A1F1F")

    obstacles = [leg_box, tag_box, Bbox(arrow_box), Bbox(sb_box)]
    obstacles.extend(ctx_boxes)
    for ea in extra_artists:
        obstacles.append(ea.get_window_extent(R))
    if inset_box is not None:
        obstacles.append(inset_box)
    for px, py in pix:
        mr = 5.5 * PT   # marker radius incl. halo
        obstacles.append(Bbox([[px - mr, py - mr], [px + mr, py + mr]]))
    axbox = ax.get_window_extent(R).padded(-6)

    def overlap(b1, b2):
        x = max(0, min(b1.x1, b2.x1) - max(b1.x0, b2.x0))
        y = max(0, min(b1.y1, b2.y1) - max(b1.y0, b2.y0))
        return x * y

    def segment_hits(p, q, b):
        # does the leader line p->q cross box b? (sampled)
        for t in np.linspace(0.15, 0.85, 8):
            x, y = p[0] + (q[0] - p[0]) * t, p[1] + (q[1] - p[1]) * t
            if b.x0 < x < b.x1 and b.y0 < y < b.y1:
                return True
        return False

    angles = [45, 135, 315, 225, 0, 180, 90, 270, 20, 160, 340, 200, 65, 115, 245, 295]
    radii = [16, 22, 30, 40, 52, 66, 82]
    placed_lines = []
    for idx in sorted(range(len(pts)), key=lambda i: -(pts[i].value or -1)):
        s = pts[idx]
        px, py = pix[idx]
        val = "–" if s.value is None else str(s.value)
        nw, nh = measure(val, NUM_FS); nw += 2 * NUM_PAD * NUM_FS * PT; nh += 2 * NUM_PAD * NUM_FS * PT
        mw, mh = measure(s.label, NAME_FS); mw += 2 * NAME_PAD * NAME_FS * PT; mh += 2 * NAME_PAD * NAME_FS * PT
        best = None
        for r in radii:
            for a in angles:
                dx, dy = r * math.cos(math.radians(a)), r * math.sin(math.radians(a))
                cxp, cyp = px + dx * PT, py + dy * PT
                nb = Bbox([[cxp - nw / 2, cyp - nh / 2], [cxp + nw / 2, cyp + nh / 2]])
                top = nb.y0 - 1.5 * PT
                mb = Bbox([[cxp - mw / 2, top - mh], [cxp + mw / 2, top]])
                b = Bbox.union([nb, mb]).padded(3)
                inside = axbox.x0 <= b.x0 and b.x1 <= axbox.x1 and axbox.y0 <= b.y0 and b.y1 <= axbox.y1
                cost = sum(overlap(b, o) for o in obstacles) + (0 if inside else 1e7)
                cost += sum(2500 for (p0, q0) in placed_lines if segment_hits(p0, q0, b))
                cost += r * 2  # prefer short leaders
                if best is None or cost < best[0]:
                    best = (cost, dx, dy, b)
            if best and best[0] < r * 2 + 1:
                break
        _, dx, dy, b = best
        num = ax.annotate(val, (s.x, s.y), xytext=(dx, dy), textcoords="offset points",
                          ha="center", va="center", fontsize=NUM_FS, fontweight="bold",
                          color=band_text(s.value), zorder=14,
                          bbox=dict(boxstyle=f"round,pad={NUM_PAD},rounding_size=0.25",
                                    fc=band_fill(s.value), ec="#1B1B1B", lw=0.8),
                          arrowprops=dict(arrowstyle="-", color="#1B1B1B", lw=0.8, shrinkA=0, shrinkB=4))
        ax.annotate(s.label, xy=(0.5, 0), xycoords=num, xytext=(0, -1.5), textcoords="offset points",
                    ha="center", va="top", fontsize=NAME_FS, color="#1B1B1B", zorder=14, fontweight="bold",
                    bbox=dict(boxstyle=f"round,pad={NAME_PAD}", fc="white", ec="#9AA5AE", lw=0.4, alpha=0.95))
        obstacles.append(b)
        placed_lines.append(((px, py), (px + dx * PT, py + dy * PT)))

    # neat frame
    ax.add_patch(Rectangle((0, 0), 1, 1, transform=ax.transAxes, fill=False, ec="#1F2D3A", lw=1.2, zorder=20))
    fig.savefig(out_png, dpi=300)
    plt.close(fig)
    return out_png


# =============================================================================
# 4b. PUNJAB DISTRICT AQI MAP (page 3)
# =============================================================================

PUNJAB_MAP_SIZE_MM = (194, 220)     # must match .mapimg3 in the CSS
PUNJAB_CRS = "+proj=aea +lat_1=29 +lat_2=33 +lat_0=31 +lon_0=72 +datum=WGS84 +units=m +no_defs"
PUNJAB_LABEL_FONTS = [7.6, 6.8, 6.0, 5.3, 4.6]   # tried largest-first so each district name fits inside its polygon


def find_districts_shp(shp_dir: Path) -> Path | None:
    """Pick the Punjab district boundary shapefile from the shp folder (file name containing 'dist')."""
    if not shp_dir.exists():
        return None
    for p in sorted(shp_dir.glob("*.shp")):
        if re.search(r"dist", p.stem, re.I) and "aqms" not in p.stem.lower():
            return p
    return None


def _canon_loose(name: str) -> str | None:
    """canonical_district() that also tolerates spacing differences (e.g. 'Bahawal Nagar')."""
    c = canonical_district(name)
    if c:
        return c
    key = norm(name).replace(" ", "")
    for d in DISTRICT_URDU:
        if norm(d).replace(" ", "") == key:
            return d
    for a, d in DISTRICT_ALIASES.items():
        if norm(a).replace(" ", "") == key:
            return d
    return None


def render_punjab_map(districts, districts_shp: Path | None, out_png: Path, data_date: dt.date, log,
                      size_mm=PUNJAB_MAP_SIZE_MM):
    """Punjab map, every district filled with its AQI band colour and labelled with name + AQI.
    Labels are fitted inside each district (font shrinks to fit); districts too small for a
    label get a callout placed in free space with a leader line."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.patches import FancyBboxPatch, Rectangle
    import matplotlib.patheffects as pe
    from matplotlib.transforms import Bbox
    import geopandas as gpd
    from shapely.geometry import box, Point

    if not districts_shp or not Path(districts_shp).exists():
        log("  ! Punjab district shapefile not found; Punjab map skipped (use --districts-shp or put it in the shp folder)")
        return None
    plt.rcParams["font.family"] = _pick_font()

    dg = gpd.read_file(districts_shp)
    if dg.crs is None:
        dg = dg.set_crs(4326)
    dg = dg.to_crs(PUNJAB_CRS)
    dg = dg[dg.geometry.notna() & ~dg.geometry.is_empty].copy()
    str_cols = [c for c in dg.columns if c != "geometry" and pd.api.types.is_string_dtype(dg[c])]
    if not str_cols:
        log("  ! No text field with district names in the district shapefile; Punjab map skipped")
        return None
    name_col = max(str_cols, key=lambda c: dg[c].astype(str).map(_canon_loose).notna().sum())
    if dg[name_col].astype(str).map(_canon_loose).notna().sum() == 0:
        log(f"  ! Could not match any district names in field '{name_col}'; Punjab map skipped")
        return None
    dg["_name"] = [(_canon_loose(n) or str(n).strip()) for n in dg[name_col].astype(str)]
    dg = dg.dissolve(by="_name", as_index=False)          # one polygon per district (also works for tehsil-level files)
    vals = {d.name: d.value for d in districts}
    dg["_val"] = dg["_name"].map(vals)
    dg["_fill"] = [band_fill(v) if v == v and v is not None else NO_DATA_FILL for v in dg["_val"]]
    in_shp = set(dg["_name"])
    missing = [d.name for d in districts if d.value is not None and d.name not in in_shp]
    if missing:
        log(f"  ! districts with AQI but no polygon in the shapefile (not drawn): {', '.join(missing)}")
    log(f"  Punjab map: {int(dg['_val'].notna().sum())} of {len(dg)} districts have AQI data")

    W, H = size_mm[0] / 25.4, size_mm[1] / 25.4
    fig = plt.figure(figsize=(W, H), dpi=300)
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_axis_off()
    union = dg.union_all() if hasattr(dg, "union_all") else dg.unary_union

    # extent: whole Punjab + 2.5% padding, matched to the frame aspect ratio
    bx0, by0, bx1, by1 = dg.total_bounds
    cx, cy = (bx0 + bx1) / 2, (by0 + by1) / 2
    w, h = (bx1 - bx0) * 1.05, (by1 - by0) * 1.05
    if w / h > W / H:
        h = w * H / W
    else:
        w = h * W / H
    x0, x1, y0, y1 = cx - w / 2, cx + w / 2, cy - h / 2, cy + h / 2
    ax.set_xlim(x0, x1); ax.set_ylim(y0, y1)
    ax.add_patch(Rectangle((x0, y0), w, h, facecolor="#F3F5F7", lw=0, zorder=0))

    dg.plot(ax=ax, color=dg["_fill"].tolist(), edgecolor="#3B4650", linewidth=0.55, zorder=2)
    _plot_lines(ax, [union], color="#1F2D3A", lw=1.5, zorder=4)

    if BORDER_FILE.exists():                                  # international border, as on the Lahore map
        bd = gpd.read_file(BORDER_FILE).to_crs(PUNJAB_CRS)
        bd = bd[bd.intersects(box(x0, y0, x1, y1))]
        if len(bd):
            _plot_lines(ax, bd.geometry, color="white", lw=3.0, alpha=0.8, zorder=4)
            _plot_lines(ax, bd.geometry, color="#8B1E1E", lw=1.2, ls=(0, (6, 2, 1, 2)), zorder=5)
    ax.set_aspect("auto"); ax.set_xlim(x0, x1); ax.set_ylim(y0, y1)

    fig.canvas.draw()
    R = fig.canvas.get_renderer()
    PT = fig.dpi / 72.0
    fw, fh = fig.bbox.width, fig.bbox.height

    # ---- legend panel + north arrow / scale bar in the two emptiest corners ----
    lw_, lh_ = 0.27, 0.25
    cfrac = {"tl": (0.012, 1 - lh_ - 0.012), "tr": (1 - lw_ - 0.012, 1 - lh_ - 0.012),
             "bl": (0.012, 0.012), "br": (1 - lw_ - 0.012, 0.012)}
    def cover(k):
        fx, fy = cfrac[k]
        b = box(x0 + fx * w, y0 + fy * h, x0 + (fx + lw_) * w, y0 + (fy + lh_) * h)
        return union.intersection(b).area / b.area
    order = sorted(cfrac, key=lambda k: (cover(k), ["bl", "br", "tl", "tr"].index(k)))
    leg_c, aux_c = order[0], order[1]

    lx, ly = cfrac[leg_c]
    lax = fig.add_axes([lx, ly, lw_, lh_]); lax.set_xlim(0, 1); lax.set_ylim(0, 1); lax.set_axis_off()
    lax.add_patch(FancyBboxPatch((0.02, 0.02), 0.96, 0.96, boxstyle="round,pad=0,rounding_size=0.04",
                                 facecolor="white", edgecolor="#9AA5AE", lw=0.8, alpha=0.95))
    lax.text(0.5, 0.94, "AQI limits of EPA Punjab", ha="center", va="center", fontsize=7.2, fontweight="bold", color="#1F2D3A")
    yy = 0.85
    for i, (lo, hi, en, _ur, fill, _tc) in enumerate(BANDS):
        lax.add_patch(FancyBboxPatch((0.07, yy - 0.032), 0.16, 0.064, boxstyle="round,pad=0,rounding_size=0.015",
                                     facecolor=fill, edgecolor="#555", lw=0.3))
        lax.text(0.27, yy, BAND_LABELS[i], va="center", fontsize=6.4, color="#222", fontweight="bold")
        lax.text(0.50, yy, en if i != 3 else "Unhealthy (Sensitive)", va="center", fontsize=6.0, color="#333")
        yy -= 0.092
    lax.add_patch(FancyBboxPatch((0.07, yy - 0.032), 0.16, 0.064, boxstyle="round,pad=0,rounding_size=0.015",
                                 facecolor=NO_DATA_FILL, edgecolor="#555", lw=0.3))
    lax.text(0.27, yy, "No data", va="center", fontsize=6.4, color="#222", fontweight="bold")
    fig.canvas.draw()
    leg_box = lax.get_window_extent(R)

    ax_x = x0 + w * (cfrac[aux_c][0] + lw_ / 2)
    top_side = aux_c[0] == "t"
    ay_f = cfrac[aux_c][1] + (lh_ * 0.80 if top_side else lh_ * 0.52)
    ax.annotate("", xy=(ax_x, y0 + h * (ay_f + 0.06)), xytext=(ax_x, y0 + h * ay_f),
                arrowprops=dict(arrowstyle="-|>,head_width=0.4,head_length=0.8", color="#1F2D3A", lw=1.5), zorder=12)
    ax.text(ax_x, y0 + h * (ay_f + 0.072), "N", ha="center", va="bottom", fontsize=9, fontweight="bold", color="#1F2D3A",
            zorder=12, path_effects=[pe.withStroke(linewidth=2.5, foreground="white")])
    km = [25, 50, 100, 200][int(np.argmin([abs(k * 1000 - w * 0.22) for k in [25, 50, 100, 200]]))]
    L = km * 1000
    sx = ax_x - L / 2; sy = y0 + h * (cfrac[aux_c][1] + (lh_ * 0.22 if not top_side else lh_ * 0.45))
    for k in range(4):
        ax.add_patch(Rectangle((sx + k * L / 4, sy), L / 4, h * 0.0075,
                               facecolor="#1F2D3A" if k % 2 == 0 else "white", edgecolor="#1F2D3A", lw=0.6, zorder=12))
    ax.text(sx, sy + h * 0.013, "0", fontsize=6, ha="center", zorder=12, path_effects=[pe.withStroke(linewidth=2, foreground="white")])
    ax.text(sx + L, sy + h * 0.013, f"{km} km", fontsize=6, ha="center", zorder=12, path_effects=[pe.withStroke(linewidth=2, foreground="white")])
    aux_box = Bbox(ax.transData.transform([[sx - L * 0.15, sy - h * 0.01], [sx + L * 1.25, y0 + h * (ay_f + 0.10)]]))

    # ---- district labels ----
    inv = ax.transData.inverted()
    axbox = ax.get_window_extent(R).padded(-5)

    def measure(txt, fs, bold=True):
        t = ax.text(0, 0, txt, fontsize=fs, fontweight="bold" if bold else "normal", transform=None, multialignment="center")
        e = t.get_window_extent(R)
        t.remove()
        return e.width, e.height

    def disp_to_geom(b):
        (a0, b0), (a1, b1) = inv.transform([[b.x0, b.y0], [b.x1, b.y1]])
        return box(a0, b0, a1, b1)

    def overlap(b1, b2):
        x = max(0, min(b1.x1, b2.x1) - max(b1.x0, b2.x0))
        y = max(0, min(b1.y1, b2.y1) - max(b1.y0, b2.y0))
        return x * y

    def wrap2(name):
        if " " not in name:
            return None
        parts = name.split(" ")
        best = min(range(1, len(parts)), key=lambda i: abs(len(" ".join(parts[:i])) - len(" ".join(parts[i:]))))
        return " ".join(parts[:best]) + "\n" + " ".join(parts[best:])

    rows = []
    for _, r in dg.iterrows():
        if r["_val"] is None or r["_val"] != r["_val"]:
            continue                                  # no data: district stays grey, nothing is written on it
        g = r.geometry
        part = max(g.geoms, key=lambda p: p.area) if hasattr(g, "geoms") else g
        rows.append(dict(name=r["_name"], val=r["_val"], geom=part, fill=r["_fill"], area=part.area))
    rows.sort(key=lambda d: -d["area"])

    obstacles = [leg_box, aux_box]
    placed, unplaced = [], []
    for d in rows:
        val = "–" if d["val"] is None or d["val"] != d["val"] else str(int(d["val"]))
        d["vs"] = val
        part = d["geom"]
        cands = [part.representative_point()]
        if part.contains(part.centroid):
            cands.insert(0, part.centroid)
        mnx, mny, mxx, mxy = part.bounds
        grid = [Point(mnx + (i + 0.5) * (mxx - mnx) / 9, mny + (j + 0.5) * (mxy - mny) / 9) for i in range(9) for j in range(9)]
        grid = sorted([p for p in grid if part.contains(p)], key=lambda p: p.distance(part.centroid))
        cands += grid[:40]
        tc = band_text(d["val"]) if val != "–" else "#000000"
        hit = None
        for fs in PUNJAB_LABEL_FONTS:
            nm_opts = [d["name"]] + ([wrap2(d["name"])] if wrap2(d["name"]) else [])
            for nm in nm_opts:
                nw, nh = measure(nm, fs)
                vw, vh = measure(val, fs + 2.4)
                gap = 0.18 * fs * PT
                bw, bh = max(nw, vw) + 2 * PT, nh + vh + gap
                for c in cands:
                    px, py = ax.transData.transform((c.x, c.y))
                    b = Bbox([[px - bw / 2, py - bh / 2], [px + bw / 2, py + bh / 2]])
                    if not (axbox.x0 <= b.x0 and b.x1 <= axbox.x1 and axbox.y0 <= b.y0 and b.y1 <= axbox.y1):
                        continue
                    if any(overlap(b, o) > 0 for o in obstacles):
                        continue
                    if not part.contains(disp_to_geom(b)):
                        continue
                    hit = (nm, fs, nh, vh, gap, b, px, py)
                    break
                if hit:
                    break
            if hit:
                break
        if hit:
            nm, fs, nh, vh, gap, b, px, py = hit
            ny = b.y1 - nh / 2
            vy = b.y0 + vh / 2
            (dx_, ny_d), = inv.transform([[px, ny]]); (_, vy_d), = inv.transform([[px, vy]])
            ax.text(dx_, ny_d, nm, ha="center", va="center", fontsize=fs, fontweight="bold", color=tc, zorder=10,
                    multialignment="center", linespacing=0.95)
            ax.text(dx_, vy_d, val, ha="center", va="center", fontsize=fs + 2.4, fontweight="bold", color=tc, zorder=10)
            obstacles.append(b)
            placed.append(d)
        else:
            unplaced.append(d)

    # districts too small for an inside label -> callout in free space with a leader line
    NUM_FS, NAME_FS, NUM_PAD, NAME_PAD = 9, 6.0, 0.28, 0.22
    angles = [0, 180, 90, 270, 45, 135, 315, 225, 20, 160, 340, 200, 65, 115, 245, 295]
    radii = [22, 32, 44, 58, 74, 92, 112, 134]
    placed_lines = []

    def segment_hits(p, q, b):
        for t in np.linspace(0.15, 0.85, 8):
            x, y = p[0] + (q[0] - p[0]) * t, p[1] + (q[1] - p[1]) * t
            if b.x0 < x < b.x1 and b.y0 < y < b.y1:
                return True
        return False

    for d in unplaced:
        part = d["geom"]
        rp = part.representative_point()
        px, py = ax.transData.transform((rp.x, rp.y))
        val = d["vs"]
        nw, nh = measure(val, NUM_FS); nw += 2 * NUM_PAD * NUM_FS * PT; nh += 2 * NUM_PAD * NUM_FS * PT
        mw, mh = measure(d["name"], NAME_FS); mw += 2 * NAME_PAD * NAME_FS * PT; mh += 2 * NAME_PAD * NAME_FS * PT
        best = None
        for r_ in radii:
            for a in angles:
                dx, dy = r_ * math.cos(math.radians(a)), r_ * math.sin(math.radians(a))
                cxp, cyp = px + dx * PT, py + dy * PT
                nb = Bbox([[cxp - nw / 2, cyp - nh / 2], [cxp + nw / 2, cyp + nh / 2]])
                top = nb.y0 - 1.5 * PT
                mb = Bbox([[cxp - mw / 2, top - mh], [cxp + mw / 2, top]])
                b = Bbox.union([nb, mb]).padded(3)
                inside = axbox.x0 <= b.x0 and b.x1 <= axbox.x1 and axbox.y0 <= b.y0 and b.y1 <= axbox.y1
                gb = disp_to_geom(b)
                on_map = union.intersection(gb).area / gb.area
                cost = sum(overlap(b, o) for o in obstacles) + (0 if inside else 1e7) + 6000 * on_map
                cost += sum(2500 for (p0, q0) in placed_lines if segment_hits(p0, q0, b))
                cost += r_ * 2
                if best is None or cost < best[0]:
                    best = (cost, dx, dy, b)
            if best and best[0] < r_ * 2 + 1:
                break
        _, dx, dy, b = best
        ax.scatter([rp.x], [rp.y], s=7, color="#1B1B1B", edgecolor="white", lw=0.5, zorder=12)
        tcol = band_text(d["val"]) if val != "–" else "#000000"
        num = ax.annotate(val, (rp.x, rp.y), xytext=(dx, dy), textcoords="offset points", ha="center", va="center",
                          fontsize=NUM_FS, fontweight="bold", color=tcol, zorder=14,
                          bbox=dict(boxstyle=f"round,pad={NUM_PAD},rounding_size=0.25", fc=d["fill"], ec="#1B1B1B", lw=0.8),
                          arrowprops=dict(arrowstyle="-", color="#1B1B1B", lw=0.8, shrinkA=0, shrinkB=2))
        ax.annotate(d["name"], xy=(0.5, 0), xycoords=num, xytext=(0, -1.5), textcoords="offset points",
                    ha="center", va="top", fontsize=NAME_FS, color="#1B1B1B", zorder=14, fontweight="bold",
                    bbox=dict(boxstyle=f"round,pad={NAME_PAD}", fc="white", ec="#9AA5AE", lw=0.4, alpha=0.95))
        obstacles.append(b)
        placed_lines.append(((px, py), (px + dx * PT, py + dy * PT)))

    ax.add_patch(Rectangle((0, 0), 1, 1, transform=ax.transAxes, fill=False, ec="#1F2D3A", lw=1.2, zorder=20))
    fig.savefig(out_png, dpi=300)
    plt.close(fig)
    log(f"  Punjab map: {len(placed)} labels inside districts, {len(unplaced)} as callouts")
    return out_png


# =============================================================================
# 5. HTML REPORT  ->  PDF
# =============================================================================

ICONS = {  # 24x24 flat icons
    "aqi": '<rect x="3" y="13" width="4" height="8" rx="1" fill="#2E86C1"/><rect x="10" y="8" width="4" height="13" rx="1" fill="#F1C40F"/><rect x="17" y="3" width="4" height="18" rx="1" fill="#E74C3C"/>',
    "heart": '<path d="M12 21s-8-5.2-8-11a4.5 4.5 0 0 1 8-2.8A4.5 4.5 0 0 1 20 10c0 5.8-8 11-8 11z" fill="#E74C3C"/><path d="M5 12h4l1.5-3 2 6 1.5-3h5" stroke="#fff" stroke-width="1.4" fill="none"/>',
    "food": '<circle cx="12" cy="14" r="7" fill="#E74C3C"/><path d="M12 7c0-2 1-4 3-4" stroke="#6E2C00" stroke-width="1.5" fill="none"/><ellipse cx="15.5" cy="5.5" rx="2.5" ry="1.3" fill="#27AE60"/>',
    "nosmoke": '<rect x="3" y="11" width="15" height="3" fill="#BDC3C7"/><rect x="15" y="11" width="3" height="3" fill="#E67E22"/><circle cx="12" cy="12.5" r="9" stroke="#C0392B" stroke-width="2" fill="none"/><path d="M5.6 6.1l12.8 12.8" stroke="#C0392B" stroke-width="2"/>',
    "run": '<circle cx="14" cy="4" r="2.2" fill="#2C3E50"/><path d="M9 21l3-6 3 3v4M12 15l1-6 4 3h3M13 9l-4 1-2 3" stroke="#2C3E50" stroke-width="2" fill="none" stroke-linecap="round"/>',
    "kit": '<rect x="3" y="7" width="18" height="13" rx="2" fill="#ECF0F1" stroke="#2C3E50" stroke-width="1.2"/><path d="M9 7V5h6v2" stroke="#2C3E50" stroke-width="1.2" fill="none"/><path d="M12 10v7M8.5 13.5h7" stroke="#E74C3C" stroke-width="2.4"/>',
    "doctor": '<circle cx="12" cy="7" r="4" fill="#F5CBA7"/><path d="M4 22c0-5 3.5-8 8-8s8 3 8 8z" fill="#5DADE2"/><path d="M12 15v4M10 17h4" stroke="#fff" stroke-width="1.6"/>',
    "mask": '<path d="M4 9c4-2 12-2 16 0v5c-2 4-6 5-8 5s-6-1-8-5z" fill="#AED6F1" stroke="#2874A6" stroke-width="1"/><path d="M4 10H1M20 10h3M4 13H1.5M20 13h2.5" stroke="#2874A6" stroke-width="1"/><path d="M7 12h10M8 15h8" stroke="#2874A6" stroke-width=".8"/>',
    "home": '<path d="M3 11l9-8 9 8" stroke="#2C3E50" stroke-width="2" fill="none"/><path d="M5 10v11h14V10" fill="#F9E79F" stroke="#2C3E50" stroke-width="1.2"/><rect x="10" y="14" width="4" height="7" fill="#A04000"/>',
    "window": '<rect x="4" y="3" width="16" height="18" rx="1" fill="#D6EAF8" stroke="#2C3E50" stroke-width="1.4"/><path d="M12 3v18M4 12h16" stroke="#2C3E50" stroke-width="1.4"/>',
    "child": '<circle cx="12" cy="6" r="3" fill="#F5CBA7"/><path d="M8 21v-5l-2-3 6-3 6 3-2 3v5" fill="#F1948A"/><circle cx="12" cy="12.5" r="9.5" stroke="#C0392B" stroke-width="1.6" fill="none"/>',
    "elder": '<circle cx="11" cy="5" r="2.6" fill="#D5DBDB"/><path d="M8 21l2-7-1-4 4-1 2 5 3 1M17 13v8" stroke="#566573" stroke-width="2" fill="none" stroke-linecap="round"/>',
    "purifier": '<rect x="6" y="2" width="12" height="20" rx="3" fill="#EBF5FB" stroke="#2874A6" stroke-width="1.2"/><circle cx="12" cy="10" r="4" fill="none" stroke="#2874A6" stroke-width="1.2"/><path d="M9 17h6M9 19h6" stroke="#2874A6" stroke-width="1"/>',
    "goggles": '<rect x="2" y="8" width="9" height="7" rx="3.5" fill="#85C1E9" stroke="#1B4F72" stroke-width="1.2"/><rect x="13" y="8" width="9" height="7" rx="3.5" fill="#85C1E9" stroke="#1B4F72" stroke-width="1.2"/><path d="M11 11h2" stroke="#1B4F72" stroke-width="1.5"/>',
    "outdoor": '<circle cx="8" cy="8" r="4" fill="#F4D03F"/><path d="M15 21v-5" stroke="#6E2C00" stroke-width="2"/><circle cx="15" cy="12" r="5" fill="#58D68D"/>',
}


def svg_icon(key):
    return f'<svg class="ic" viewBox="0 0 24 24">{ICONS.get(key, ICONS["aqi"])}</svg>'


def svg_face(bi):
    fill = BANDS[bi][4] if bi is not None else NO_DATA_FILL
    mouth = {0: "M8 14q4 5 8 0", 1: "M8 14.5q4 3 8 0", 2: "M8 15.5h8", 3: "M8 16.5q4-2.5 8 0",
             4: "M8 17q4-4 8 0", 5: "M8 17.5q4-5 8 0", 6: "M8 18q4-6 8 0"}.get(bi, "M8 15.5h8")
    return (f'<svg class="face" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10.5" fill="{fill}" '
            f'fill-opacity=".35" stroke="#333" stroke-width="1.1"/><circle cx="8.5" cy="9.5" r="1.2" fill="#333"/>'
            f'<circle cx="15.5" cy="9.5" r="1.2" fill="#333"/><path d="{mouth}" stroke="#333" stroke-width="1.3" '
            f'fill="none" stroke-linecap="round"/></svg>')


def svg_gauge(v):
    n = len(BANDS)
    cx, cy, r0, r1 = 60, 58, 30, 52
    parts = []
    for i, b in enumerate(BANDS):
        a0 = math.pi - i * math.pi / n
        a1 = math.pi - (i + 1) * math.pi / n
        p = lambda a, r: (cx + r * math.cos(a), cy - r * math.sin(a))
        (x0, y0), (x1, y1), (x2, y2), (x3, y3) = p(a0, r1), p(a1, r1), p(a1, r0), p(a0, r0)
        parts.append(f'<path d="M{x0:.1f},{y0:.1f} A{r1},{r1} 0 0 1 {x1:.1f},{y1:.1f} L{x2:.1f},{y2:.1f} '
                     f'A{r0},{r0} 0 0 0 {x3:.1f},{y3:.1f}Z" fill="{b[4]}" stroke="#fff" stroke-width="1.2"/>')
    if v is not None:
        bi = band_index(v); lo, hi = BANDS[bi][0], BANDS[bi][1]
        frac = (bi + (min(max(v, lo), hi) - lo) / max(hi - lo, 1)) / n
        a = math.pi - frac * math.pi
        tx, ty = cx + (r1 - 4) * math.cos(a), cy - (r1 - 4) * math.sin(a)
        parts.append(f'<line x1="{cx}" y1="{cy}" x2="{tx:.1f}" y2="{ty:.1f}" stroke="#333" stroke-width="4" stroke-linecap="round"/>')
        parts.append(f'<circle cx="{cx}" cy="{cy}" r="6" fill="#333"/>')
    return f'<svg class="gauge" viewBox="0 0 120 64">{"".join(parts)}</svg>'


def font_faces():
    def uri(p):
        return p.resolve().as_uri()
    out = []
    for fam, f, wt in [("NotoNastaliqUrduLocal", "NotoNastaliqUrdu-400.woff2", 400),
                       ("NotoNastaliqUrduLocal", "NotoNastaliqUrdu-700.woff2", 700),
                       ("NotoSerifLocal", "NotoSerif-400.woff2", 400), ("NotoSerifLocal", "NotoSerif-700.woff2", 700),
                       ("NotoSansLocal", "NotoSans-400.woff2", 400), ("NotoSansLocal", "NotoSans-700.woff2", 700)]:
        p = FONTS / f
        if p.exists():
            out.append(f"@font-face{{font-family:'{fam}';src:url('{uri(p)}') format('woff2');font-weight:{wt};}}")
    return "\n".join(out)


def img_tag(p: Path, cls: str):
    return f'<img class="{cls}" src="{p.resolve().as_uri()}">' if p.exists() else f'<div class="{cls} ph"></div>'


CSS = r"""
@page p1 { size: A4 portrait; margin: 0 }
@page p2 { size: A4 landscape; margin: 0 }
* { box-sizing: border-box; margin: 0; padding: 0 }
html, body { background: #fff; -webkit-print-color-adjust: exact; print-color-adjust: exact }
body { font-family: 'NotoSansLocal', Arial, sans-serif; color: #111 }
.ur { font-family: 'Jameel Noori Nastaleeq', 'NotoNastaliqUrduLocal', 'Urdu Typesetting', serif; direction: rtl }
.page { position: relative; overflow: hidden; break-after: page }
.page:last-child { break-after: auto }
.p1 { page: p1; width: 210mm; height: 297mm; padding: 7mm 8mm 5mm }
.p2 { page: p2; width: 297mm; height: 210mm; padding: 5mm 8mm 4mm; font-family: 'Times New Roman', 'NotoSerifLocal', serif }

/* ---------- page 1 ---------- */
.hdr { display: flex; align-items: center; justify-content: space-between; background: #B9E2A6;
       border-radius: 5mm; height: 30mm; padding: 0 5mm }
.hdr .logos { display: flex; align-items: center; gap: 3mm; width: 52mm }
.hdr .lp { height: 20mm } .hdr .le { height: 15mm }
.hdr .hl { height: 21mm; width: 52mm; object-fit: contain; object-position: right }
.hdr .ttl { text-align: center; flex: 1 }
.hdr .ttl .t1 { font-size: 17pt; font-weight: 700; line-height: 1.9; white-space: nowrap }
.hdr .ttl .t2 { font-size: 12pt; font-weight: 700; line-height: 1.9; margin-top: 1.5mm }
.hdr .ttl .t2 .n { font-family: 'NotoSansLocal', Arial, sans-serif; font-size: 11pt }
.band { display: flex; align-items: center; justify-content: space-between; background: #F8DCD2;
        border-radius: 4mm; height: 19mm; margin-top: 3mm; padding: 0 6mm }
.msgpill { background: #CFE6F7; border: 0.35mm solid #444; padding: 0 6mm; font-size: 13pt; font-weight: 700; line-height: 2.1 }
.avg { display: flex; align-items: center; gap: 4mm }
.avgbox { min-width: 30mm; text-align: center; font-size: 22pt; font-weight: 700; border: 0.35mm solid #444; line-height: 12mm }
.avglbl { font-size: 15pt; font-weight: 700; line-height: 2 }
.avglbl b { font-family: 'NotoSansLocal', Arial, sans-serif; font-size: 14pt }
.gauge { height: 16mm }
.main { display: flex; gap: 4mm; margin-top: 4mm; height: 195mm }
.col { flex: 1; display: flex; flex-direction: column; justify-content: flex-end }
.msg { flex: 1; min-height: 0; overflow: hidden; margin-bottom: 3mm; padding: 2mm 3mm 1mm; border: 0.3mm solid #E5E5E5; border-radius: 2mm }
.msg .cat { display: flex; align-items: center; gap: 2mm; font-size: 11pt; line-height: 1.8; white-space: nowrap }
.msg .cat.long { font-size: 9pt; gap: 1.2mm } .msg .cat.long .nm { font-size: 10pt }
.msg .cat .nm { font-weight: 700; font-size: 13pt }
.face { width: 9mm; height: 9mm; flex: none }
.msg h4 { font-size: 12pt; font-weight: 700; line-height: 1.7; margin-top: 0.8mm }
.msg li { list-style: none; display: flex; align-items: center; gap: 1.8mm; font-size: 9.6pt; line-height: 1.6 }
.ic { width: 5.2mm; height: 5.2mm; flex: none }
table.rk { width: 100%; border-collapse: collapse; direction: rtl }
table.rk th { font-size: 11pt; font-weight: 700; height: 9mm; border-bottom: 0.3mm solid #999 }
table.rk th.en { font-family: 'NotoSansLocal', Arial, sans-serif; font-size: 10.5pt }
table.rk td { border-bottom: 0.2mm solid #E3E3E3; text-align: center; vertical-align: middle }
table.rk td.r { width: 11mm; font-family: 'NotoSansLocal', Arial, sans-serif; font-size: 9.5pt; direction: ltr }
table.rk td.d { width: 30mm; font-size: 10.5pt; font-weight: 700; line-height: 1 }
table.rk td.a { width: 15mm; font-family: 'NotoSansLocal', Arial, sans-serif; font-weight: 700; font-size: 12.5pt }
table.rk td.c { font-family: 'NotoSansLocal', Arial, sans-serif; font-size: 8.4pt; direction: ltr }
table.rk td.c.ur { font-family: 'Jameel Noori Nastaleeq', 'NotoNastaliqUrduLocal', serif; direction: rtl; color: #777 }
.legend { display: flex; gap: 1.2mm; margin-top: 3mm }
.legend div { flex: 1; text-align: center; font-size: 7.2pt; font-weight: 700; line-height: 1.2; display: flex; flex-direction: column; justify-content: flex-end }
.legend div small { display: block; font-weight: 700; font-size: 6.4pt }
.legend i { display: block; height: 5mm; border-radius: 1.5mm; margin-top: 1mm; flex: none }
.legend div small { min-height: 5.5mm }
.foot { font-size: 8pt; margin-top: 2.5mm; padding-left: 4mm }
.dir { text-align: center; color: #1E8B3C; font-weight: 700; font-size: 9pt; margin-top: 1.8mm }

/* ---------- focus-district AQMS map (physical page 3) ---------- */
.h2 { display: flex; align-items: center; justify-content: space-between; height: 28mm }
.h2 .lp { height: 25mm } .h2 .le { height: 25mm }
.h2 .tb { text-align: center }
.h2 .tbox { border: 0.5mm solid #111; padding: 1mm 10mm; font-size: 26pt; font-weight: 700; display: inline-block }
.h2 .tbox small { font-size: 13pt }
.h2 .upd { font-size: 14pt; font-weight: 700; margin-top: 2mm }
.body2 { display: flex; gap: 3mm; margin-top: 3mm; align-items: flex-start }
.mapimg { width: 126mm; height: 214mm; border: 0.3mm solid #444; flex: none }
table.st { border-collapse: collapse; width: 65mm; font-size: 8.6pt }
table.st th, table.st td { border: 0.25mm solid #555; text-align: center; padding: 0 1mm }
table.st th { background: #F8DCD2; height: 10.5mm; font-size: 9pt }
table.st td { height: var(--rh, 11.4mm) }
table.st td.a { font-weight: 700; font-size: 9.5pt; width: 10mm }
table.st td.rk { font-weight: 700; width: 9mm }
table.st tr.sec td { background: #F8DCD2; font-weight: 700; font-size: 9pt }
.tdot { display: inline-block; width: 4mm; height: 4mm; border-radius: 50%; background: #E03C31; color: #fff;
        font: 700 7pt/4mm Arial, sans-serif; text-align: center; vertical-align: middle; margin-right: 1mm }
.foot2 { font-size: 8.5pt; margin-top: 3mm }
.dir2 { text-align: center; color: #1E8B3C; font-weight: 700; font-size: 10pt; margin-top: 2mm; font-family: Arial, sans-serif }
.ph { background: #eee }

/* ---------- Punjab district AQI map (physical page 2) ---------- */
.p3 { page: p1; width: 210mm; height: 297mm; padding: 6mm 8mm 5mm; font-family: 'Times New Roman', 'NotoSerifLocal', serif }
.p2 .h2 .tbox, .p3 .h2 .tbox { font-size: 19pt; padding: 1mm 6mm }
.p2 .h2 .tbox small, .p3 .h2 .tbox small { font-size: 11pt }
.p2 .h2 { height: 25mm }
.p2 .h2 .lp, .p2 .h2 .le { height: 22mm }
.p2 .body2 { gap: 3mm; margin-top: 2mm }
.p2 .mapimg { width: 196mm; height: 145mm }
.p2 table.st { width: 80mm; font-size: 8pt }
.p2 table.st th { height: 8mm; font-size: 8.2pt }
.p2 .foot2 { margin-top: 2mm }
.p2 .dir2 { margin-top: 1mm }
.mapimg3 { display: block; width: 194mm; height: 220mm; margin-top: 3mm; border: 0.3mm solid #444 }
"""


def build_html(districts, punjab_avg, lahore, lahore_city, lahore_tb, map_png, data_date, report_date,
               map_focus, punjab_map_png=None) -> str:
    bi = band_index(punjab_avg)
    rd = f'{report_date.day} {URDU_MONTHS[report_date.month - 1]} {report_date.year}'
    span = f"{data_date:%d.%m.%Y}, 12:00AM to 11:00PM"

    def rk_rows(ds):
        out = []
        for d in ds:
            ur = DISTRICT_URDU.get(d.name, d.name)
            if d.value is None:
                out.append(f'<tr><td class="r">{d.rank}.</td><td class="d ur">{esc(ur)}</td>'
                           f'<td class="a" style="background:{NO_DATA_FILL}"></td><td class="c"></td></tr>')
            else:
                out.append(f'<tr><td class="r">{d.rank}.</td><td class="d ur">{esc(ur)}</td>'
                           f'<td class="a" style="background:{band_fill(d.value)};color:{band_text(d.value)}">{d.value}</td>'
                           f'<td class="c">{esc(d.dominant)}</td></tr>')
        return "".join(out)

    head = (f'<thead><tr><th class="ur">{UR["rank"]}</th><th class="ur">{UR["district"]}</th>'
            f'<th class="en">AQI</th><th class="ur">{UR["causes"]}</th></tr></thead>')
    right, left = districts[:RIGHT_COLUMN_ROWS], districts[RIGHT_COLUMN_ROWS:]
    nrows = max(len(right), 1)
    row_h = min(8.2, (195 - 9) / nrows)
    left_h = min(row_h, 7.2)

    if bi is not None:
        lo, hi = BANDS[bi][0], BANDS[bi][1]
        adv_name, adv_secs = advisory_sections(bi)
        secs = "".join(
            f'<h4>{esc(h)}</h4><ul>{"".join(f"<li>{svg_icon(k)}<span>{esc(t)}</span></li>" for k, t in items)}</ul>'
            for h, items in adv_secs)
        msg = (f'<div class="msg ur"><div class="cat{" long" if len(adv_name) > 12 else ""}">{svg_face(bi)}<span class="nm" style="color:#B7950B">'
               f'{esc(adv_name)}</span><span>:</span><span dir="ltr">({lo}—{hi}) (AQI)</span></div>{secs}</div>')
        # estimated height (mm) of the message box -> the left-hand ranking rows shrink just enough to leave room for it
        est = 12 + 7.5 * len(adv_secs) + sum(5.6 * (2 if len(t) > 58 else 1) for _h, items in adv_secs for _k, t in items)
        left_h = max(5.0, min(left_h, (195 - est - 3 - 9) / max(len(left), 1)))
    else:
        msg = '<div class="msg"></div>'

    legend = "".join(
        f'<div>{BAND_LABELS[i]}<small>({esc(b[2])})</small><i style="background:{b[4]}"></i></div>'
        for i, b in enumerate(BANDS))

    page1 = f"""
<section class="page p1">
  <div class="hdr">
    <div class="logos">{img_tag(LOGO_PUNJAB, 'lp')}{img_tag(LOGO_EPA, 'le')}</div>
    <div class="ttl ur"><div class="t1">{UR['title']}</div>
      <div class="t2"><span class="n">{report_date.day}</span> {URDU_MONTHS[report_date.month - 1]} <span class="n">{report_date.year}</span> &nbsp;&nbsp;&nbsp; {UR['last24']}</div></div>
    {img_tag(LOGO_HELPLINE, 'hl')}
  </div>
  <div class="band">
    <div class="msgpill ur">{UR['message']}</div>
    <div class="avg">
      <div class="avgbox" style="background:{band_fill(punjab_avg)};color:{band_text(punjab_avg)}">{'' if punjab_avg is None else round_half_up(punjab_avg)}</div>
      <div class="avglbl ur">اوسط <b>AQI</b></div>
      {svg_gauge(punjab_avg)}
    </div>
  </div>
  <div class="main">
    <div class="col">
      {msg}
      <table class="rk" style="--h:{left_h}mm">{head}<tbody>{rk_rows(left)}</tbody></table>
    </div>
    <div class="col">
      <table class="rk">{head}<tbody>{rk_rows(right)}</tbody></table>
    </div>
  </div>
  <div class="legend">{legend}</div>
  <div class="foot">AQI calculations is performed for last 24 hours ({span}).</div>
  <div class="dir">{esc(DIRECTORATE)}</div>
</section>
<style>.p1 .main .col:last-child table.rk td {{ height: {row_h:.2f}mm }} .p1 .main .col:first-child table.rk td {{ height: {left_h:.2f}mm }}</style>
"""

    def st_rows(sts, start=1):
        out = []
        for i, s in enumerate(sts, start):
            val = "" if s.value is None else s.value
            dom = esc(s.dominant) if s.value is not None else ""
            out.append(f'<tr><td class="rk">{i:02d}</td><td>{esc(s.label)}</td>'
                       f'<td class="a" style="background:{band_fill(s.value)};color:{band_text(s.value)}">{val}</td>'
                       f'<td>{dom}</td></tr>')
        return "".join(out)

    page2 = ""
    if lahore is not None:
        n_ok = sum(1 for s in lahore_city if s.sufficient)
        tb_rows = (f'<tr class="sec"><td colspan="4"><span class="tdot">T</span>Transboundary AQMS</td></tr>'
                   + st_rows(lahore_tb)) if lahore_tb else ""
        table_rows = len(lahore_city) + (len(lahore_tb) + 1 if lahore_tb else 0)
        station_row_height = (LAHORE_MAP_SIZE_MM[1] - 8) / max(table_rows, 1)
        mp = (f'<img class="mapimg" src="{map_png.resolve().as_uri()}">' if map_png
              else '<div class="mapimg ph"></div>')
        page2 = f"""
<section class="page p2">
  <div class="h2">
    {img_tag(LOGO_PUNJAB, 'lp')}
    <div class="tb">
      <div class="tbox" style="background:{band_fill(lahore.value)};color:{band_text(lahore.value)}">Average AQI of {esc(map_focus)} <small>({n_ok:02d} AQMS)</small> ({'' if lahore.value is None else lahore.value})</div>
      <div class="upd">Updated Time: {report_date:%d.%m.%Y} ({REPORT_TIME}) 24 Hourly Report</div>
    </div>
    {img_tag(LOGO_EPA, 'le')}
  </div>
  <div class="body2">
    {mp}
    <table class="st" style="--rh:{station_row_height:.2f}mm">
      <thead><tr><th>Rank</th><th>Location</th><th>AQI</th><th>Dominant Pollutant</th></tr></thead>
      <tbody>{st_rows(lahore_city)}{tb_rows}</tbody>
    </table>
  </div>
  <div class="foot2">AQI calculations is performed for last 24 hours ({span}).</div>
  <div class="dir2">{esc(DIRECTORATE)}</div>
</section>"""

    page3 = ""
    if punjab_map_png:
        n_d = sum(1 for d in districts if d.value is not None)
        pavg = "" if punjab_avg is None else round_half_up(punjab_avg)
        page3 = f"""
<section class="page p3">
  <div class="h2">
    {img_tag(LOGO_PUNJAB, 'lp')}
    <div class="tb">
      <div class="tbox" style="background:{band_fill(punjab_avg)};color:{band_text(punjab_avg)}">Average AQI of Punjab <small>({n_d:02d} Districts)</small> ({pavg})</div>
      <div class="upd">Updated Time: {report_date:%d.%m.%Y} ({REPORT_TIME}) 24 Hourly Report</div>
    </div>
    {img_tag(LOGO_EPA, 'le')}
  </div>
  <div class="m3"><img class="mapimg3" src="{punjab_map_png.resolve().as_uri()}"></div>
  <div class="foot2">AQI calculations is performed for last 24 hours ({span}).</div>
  <div class="dir2">{esc(DIRECTORATE)}</div>
</section>"""

    return f"""<!doctype html><html lang="ur"><head><meta charset="utf-8">
<title>Daily AQI Report {report_date:%d.%m.%Y}</title>
<style>{font_faces()}{CSS}</style></head><body>{page1}{page3}{page2}</body></html>"""


def html_to_pdf(html_path: Path, pdf_path: Path, log):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        log("  ! Playwright not installed — open the HTML file and print to PDF, or run: "
            "pip install playwright && python -m playwright install chromium")
        return False
    with sync_playwright() as p:
        exe = os.environ.get("CHROMIUM_PATH")
        try:
            browser = p.chromium.launch(executable_path=exe) if exe else p.chromium.launch()
        except Exception:
            alt = "/opt/pw-browsers/chromium"
            browser = p.chromium.launch(executable_path=alt) if Path(alt).exists() else None
            if browser is None:
                log("  ! Chromium not found — run: python -m playwright install chromium")
                return False
        page = browser.new_page()
        page.goto(html_path.resolve().as_uri(), wait_until="load")
        page.evaluate("document.fonts.ready")
        page.wait_for_timeout(300)
        page.pdf(path=str(pdf_path), prefer_css_page_size=True, print_background=True)
        browser.close()
    return True


def write_docx_from_pdf(pdf_path: Path, docx_path: Path, dpi: int = 300):
    """Embed each rendered PDF page as a full-page image so Word matches the PDF layout."""
    import tempfile
    import pymupdf
    from docx import Document
    from docx.enum.section import WD_ORIENT, WD_SECTION
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Mm, Pt

    doc = Document()
    with tempfile.TemporaryDirectory(prefix="aqi_pdf_pages_") as temp_dir:
        with pymupdf.open(pdf_path) as pdf:
            for index, page in enumerate(pdf):
                width_mm = page.rect.width * 25.4 / 72
                height_mm = page.rect.height * 25.4 / 72
                section = doc.sections[0] if index == 0 else doc.add_section(WD_SECTION.NEW_PAGE)
                section.orientation = WD_ORIENT.LANDSCAPE if width_mm > height_mm else WD_ORIENT.PORTRAIT
                section.page_width = Mm(width_mm)
                section.page_height = Mm(height_mm)
                section.left_margin = section.right_margin = Mm(4)
                section.top_margin = section.bottom_margin = Mm(4)
                section.header_distance = section.footer_distance = Mm(0)

                page_image = Path(temp_dir) / f"page_{index + 1}.png"
                page.get_pixmap(matrix=pymupdf.Matrix(dpi / 72, dpi / 72), alpha=False).save(page_image)
                paragraph = doc.add_paragraph()
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.add_run().add_picture(str(page_image), width=Mm(width_mm - 10),
                                                height=Mm(height_mm - 10))

    doc.save(str(docx_path))


# =============================================================================
# 6. EXCEL SUMMARY
# =============================================================================

def write_excel(path, districts, stations, qa, match_table, punjab_avg):
    drows = [dict(Rank=d.rank, District=d.name, District_Urdu=DISTRICT_URDU.get(d.name, ""),
                  AQI=d.value, AQI_unrounded=None if d.mean is None else round(d.mean, 3),
                  Category=None if d.value is None else BANDS[band_index(d.value)][2],
                  Stations_used=sum(s.sufficient for s in d.stations), Stations_total=len(d.stations),
                  Dominant=d.dominant) for d in districts]
    drows.append(dict(Rank=None, District="PUNJAB AVERAGE", AQI=None if punjab_avg is None else round_half_up(punjab_avg),
                      AQI_unrounded=None if punjab_avg is None else round(punjab_avg, 3)))
    srows = [dict(Station=s.name, Label=s.label, District=s.district, Role=s.role, Valid_hours=s.n_valid,
                  Used=s.sufficient, AQI=s.value, AQI_unrounded=None if s.mean is None else round(s.mean, 3),
                  Max_hourly_AQI=s.max_aqi, Max_time=None if s.max_time is None else pd.Timestamp(s.max_time).strftime("%H:%M"),
                  Dominant=s.dominant, Dominant_counts=", ".join(f"{k}:{v}" for k, v in s.dom_counts.items()),
                  Located_on_map=s.x is not None) for s in stations]
    hourly = pd.concat([s.hours.assign(Station=s.name) for s in stations])[
        ["Station", "time", "aqi", "dom", "valid", "spike"]]
    hourly.columns = ["Station", "Time", "AQI", "Dominant", "Valid", "Isolated_AQI_Peak_Flag"]
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame(drows).to_excel(xw, sheet_name="Districts", index=False)
        pd.DataFrame(srows).to_excel(xw, sheet_name="Stations", index=False)
        qa.to_excel(xw, sheet_name="QA_Flags", index=False)
        if match_table is not None:
            match_table.to_excel(xw, sheet_name="Shapefile_Match", index=False)
        hourly.to_excel(xw, sheet_name="Hourly", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 2, 8), 60)
            ws.freeze_panes = "A2"


# =============================================================================
# 7. WORD REPORT
# =============================================================================

def render_svg_pngs(svgs, out_dir, log=print):
    """Draw small SVGs (gauge, face, icons) to transparent PNGs with the same Chromium that makes the PDF.
    svgs = {name: (svg_markup, width_px, height_px)}  ->  {name: Path}.  Returns {} if Chromium is unavailable."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return {}
    done = {}
    try:
        with sync_playwright() as p:
            exe = os.environ.get("CHROMIUM_PATH")
            try:
                browser = p.chromium.launch(executable_path=exe) if exe else p.chromium.launch()
            except Exception:
                alt = "/opt/pw-browsers/chromium"
                if not Path(alt).exists():
                    return {}
                browser = p.chromium.launch(executable_path=alt)
            page = browser.new_page(device_scale_factor=4)
            for name, (svg, w, h) in svgs.items():
                svg = svg.replace("<svg ", f'<svg width="{w}" height="{h}" ', 1)
                page.set_content(f'<html><body style="margin:0;background:transparent">{svg}</body></html>')
                png = Path(out_dir) / f"{name}.png"
                page.locator("svg").first.screenshot(path=str(png), omit_background=True)
                done[name] = png
            browser.close()
    except Exception as exc:
        log(f"  ! could not draw the icons for the Word report ({exc}); the Word file is made without them")
    return done


def embed_fonts(docx_path, families):
    """Embed TrueType fonts in the .docx (fonts/<file>.ttf) so Urdu shows in Nastaliq on any PC.
    families = {"Noto Sans": ("NotoSans-400.ttf", "NotoSans-700.ttf"), ...}  (regular, bold)"""
    import shutil
    import uuid
    import zipfile
    from lxml import etree
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    wq = lambda t: f"{{{W}}}{t}"
    tmp = Path(str(docx_path) + ".tmp")
    with zipfile.ZipFile(docx_path) as zin:
        files = {n: zin.read(n) for n in zin.namelist()}
    ft = etree.fromstring(files["word/fontTable.xml"])
    rels, new_parts, n = [], {}, 0
    for fam, (reg, bold) in families.items():
        old = [f for f in ft.findall(wq("font")) if f.get(wq("name")) == fam]
        for f in old:
            ft.remove(f)
        font = etree.SubElement(ft, wq("font")); font.set(wq("name"), fam)
        etree.SubElement(font, wq("charset")).set(wq("val"), "00")
        etree.SubElement(font, wq("family")).set(wq("val"), "auto")
        etree.SubElement(font, wq("pitch")).set(wq("val"), "variable")
        for tag, fname in (("embedRegular", reg), ("embedBold", bold)):
            src = FONTS / fname
            if not src.exists():
                continue
            n += 1
            guid = str(uuid.uuid4()).upper()
            key = bytes.fromhex(guid.replace("-", ""))[::-1]
            data = bytearray(src.read_bytes())
            for i in range(32):
                data[i] ^= key[i % 16]
            new_parts[f"word/fonts/font{n}.odttf"] = bytes(data)
            rid = f"rIdFont{n}"
            rels.append(f'<Relationship Id="{rid}" Type="{R}/font" Target="fonts/font{n}.odttf"/>')
            el = etree.SubElement(font, wq(tag))
            el.set(f"{{{R}}}id", rid); el.set(wq("fontKey"), "{" + guid + "}")
    if not n:
        return
    files["word/fontTable.xml"] = etree.tostring(ft, xml_declaration=True, encoding="UTF-8", standalone=True)
    files["word/_rels/fontTable.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "".join(rels) + "</Relationships>"
    ).encode("utf-8")
    ct = files["[Content_Types].xml"].decode("utf-8")
    if 'Extension="odttf"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/><Default ', 1)
    files["[Content_Types].xml"] = ct.encode("utf-8")
    st = etree.fromstring(files["word/settings.xml"])
    if st.find(wq("embedTrueTypeFonts")) is None:
        el = etree.Element(wq("embedTrueTypeFonts"))
        zoom = st.find(wq("zoom"))
        if zoom is not None:
            zoom.addnext(el)
        else:
            st.insert(0, el)
    files["word/settings.xml"] = etree.tostring(st, xml_declaration=True, encoding="UTF-8", standalone=True)
    files.update(new_parts)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in files.items():
            zout.writestr(name, content)
    shutil.move(str(tmp), str(docx_path))


def write_docx(path, districts, punjab_avg, focus_d, city, tb, map_png, punjab_map_png, data_date, report_date, focus):
    """Legacy editable layout builder; final report output uses write_docx_from_pdf for visual parity."""
    import tempfile
    from docx import Document
    from docx.enum.section import WD_ORIENT, WD_SECTION
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor

    CENTER, LEFT, RIGHT = WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.RIGHT
    SANS, SERIF, URDU_FONT = "Noto Sans", "Noto Serif", "Noto Nastaliq Urdu"
    GREEN_HDR, PINK_BAND, BLUE_PILL, PINK_HEAD = "#B9E2A6", "#F8DCD2", "#CFE6F7", "#F8DCD2"
    tmpdir = Path(tempfile.mkdtemp(prefix="aqi_docx_"))

    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = SANS
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(0)
    normal.element.get_or_add_rPr().find(qn("w:rFonts")).set(qn("w:cs"), URDU_FONT)

    # ---------------------------------------------------------------- pictures drawn from the PDF's own SVGs
    bi = band_index(punjab_avg)
    adv_name, adv_secs = advisory_sections(bi)
    svgs = {"gauge": (svg_gauge(punjab_avg), 120, 64)}
    if bi is not None:
        svgs["face"] = (svg_face(bi), 24, 24)
    for _h, items in adv_secs:
        for key, _t in items:
            svgs.setdefault(f"ic_{key}", (svg_icon(key), 24, 24))
    pics = render_svg_pngs(svgs, tmpdir)

    # ---------------------------------------------------------------- helpers
    def new_page(first, landscape=False):
        sec = doc.sections[0] if first else doc.add_section(WD_SECTION.NEW_PAGE)
        w, h = (297, 210) if landscape else (210, 297)
        sec.page_width, sec.page_height = Mm(w), Mm(h)
        sec.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
        sec.left_margin = sec.right_margin = Mm(8)
        sec.top_margin = sec.bottom_margin = Mm(7)

    def shade(cell, fill):
        tcPr = cell._tc.get_or_add_tcPr()
        for old in tcPr.findall(qn("w:shd")):
            tcPr.remove(old)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), fill.lstrip("#"))
        tcPr.append(shd)

    def valign(cell):
        tcPr = cell._tc.get_or_add_tcPr()
        v = OxmlElement("w:vAlign"); v.set(qn("w:val"), "center"); tcPr.append(v)

    def borders(table, color="#9A9A9A", sz=4):
        tblPr = table._tbl.tblPr
        b = OxmlElement("w:tblBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement(f"w:{edge}")
            e.set(qn("w:val"), "single"); e.set(qn("w:sz"), str(sz)); e.set(qn("w:space"), "0"); e.set(qn("w:color"), color.lstrip("#"))
            b.append(e)
        tblPr.append(b)

    def widths(table, mms):
        table.autofit = False
        for row in table.rows:
            for c, w in zip(row.cells, mms):
                c.width = Mm(w)

    def row_h(row, mm):
        row.height = Mm(mm)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    def run(par, text, size=9, bold=False, color="#000000", urdu=False, serif=False):
        r = par.add_run(str(text))
        r.bold = bold
        r.font.size = Pt(size)
        r.font.color.rgb = RGBColor.from_string(color.lstrip("#").upper())
        if serif:
            r.font.name = SERIF
        if urdu:
            rpr = r._r.get_or_add_rPr()
            fonts = rpr.find(qn("w:rFonts"))
            if fonts is None:
                fonts = OxmlElement("w:rFonts"); rpr.insert(0, fonts)
            fonts.set(qn("w:cs"), URDU_FONT)
            rpr.append(OxmlElement("w:rtl"))
            szcs = OxmlElement("w:szCs"); szcs.set(qn("w:val"), str(int(size * 2))); rpr.append(szcs)
            if bold:
                rpr.append(OxmlElement("w:bCs"))
        return r

    def put(cell, text, size=9, bold=False, fill=None, color="#000000", align=CENTER, urdu=False, serif=False,
            para=None, line=None):
        """Write text into a cell (its first paragraph, or `para`). An empty text writes nothing."""
        if para is None:
            cell.text = ""
            para = cell.paragraphs[0]
        para.alignment = align
        para.paragraph_format.space_before = para.paragraph_format.space_after = Pt(0)
        if line:
            para.paragraph_format.line_spacing = Pt(line)
        if urdu:
            para._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
        if text not in (None, ""):
            run(para, text, size, bold, color, urdu, serif)
        if fill:
            shade(cell, fill)
        valign(cell)
        return para

    def picture(par, path, height_mm=None, width_mm=None):
        r = par.add_run()
        if height_mm:
            r.add_picture(str(path), height=Mm(height_mm))
        else:
            r.add_picture(str(path), width=Mm(width_mm))

    def gap(pt=4):
        par = doc.add_paragraph()
        par.paragraph_format.line_spacing = Pt(pt)
        run(par, "", 1)

    def drop_lead(cell):
        first = cell.paragraphs[0]._p
        if first.getnext() is not None and first.getnext().tag == qn("w:tbl"):
            first.getparent().remove(first)

    span = f"{data_date:%d.%m.%Y}, 12:00AM to 11:00PM"
    avg = "" if punjab_avg is None else round_half_up(punjab_avg)
    n_d = sum(1 for d in districts if d.value is not None)
    ur_date = f"{report_date.day} {URDU_MONTHS[report_date.month - 1]} {report_date.year}"

    def footer():
        par = doc.add_paragraph(); par.alignment = LEFT
        par.paragraph_format.space_before = Pt(3)
        run(par, f"AQI calculations is performed for last 24 hours ({span}).", 8.5)
        par = doc.add_paragraph(); par.alignment = CENTER
        par.paragraph_format.space_before = Pt(3)
        run(par, DIRECTORATE, 10, True, "#1E8B3C")

    # ================================================================ PAGE 1 : Urdu district ranking
    new_page(True)
    hdr = doc.add_table(rows=1, cols=3); hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths(hdr, [62, 84, 48]); row_h(hdr.rows[0], 26)
    c0, c1, c2 = hdr.rows[0].cells
    for c in (c0, c1, c2):
        shade(c, GREEN_HDR); valign(c)
    par = c0.paragraphs[0]; par.alignment = LEFT
    for logo in (LOGO_PUNJAB, LOGO_EPA):
        if logo.exists():
            picture(par, logo, height_mm=19); run(par, "  ", 6)
    put(c1, UR["title"], 16, True, GREEN_HDR, "#14304A", urdu=True, line=32)
    put(c1, f"{ur_date}     {UR['last24']}", 11, True, None, "#000000", urdu=True, para=c1.add_paragraph(), line=22)
    par = c2.paragraphs[0]; par.alignment = RIGHT
    if LOGO_HELPLINE.exists():
        picture(par, LOGO_HELPLINE, height_mm=19)
    gap(5)

    band = doc.add_table(rows=1, cols=4); band.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths(band, [62, 32, 36, 64]); row_h(band.rows[0], 17)
    b0, b1, b2, b3 = band.rows[0].cells
    for c in (b0, b1, b2, b3):
        shade(c, PINK_BAND); valign(c)
    put(b0, UR["message"], 13, True, BLUE_PILL, "#14304A", urdu=True, line=26)
    put(b1, avg, 22, True, band_fill(punjab_avg), band_text(punjab_avg))
    lp = put(b2, "", 12, align=CENTER)
    run(lp, "AQI ", 12, True); run(lp, "اوسط", 13, True, urdu=True)
    gp = b3.paragraphs[0]; gp.alignment = CENTER
    if "gauge" in pics:
        picture(gp, pics["gauge"], width_mm=36)
    gap(5)

    def rank_table(container, ds, row_mm):
        t = container.add_table(rows=1, cols=4); t.alignment = WD_TABLE_ALIGNMENT.CENTER
        borders(t)
        for c, h in zip(t.rows[0].cells, [UR["causes"], "AQI", UR["district"], UR["rank"]]):
            put(c, h, 10, True, "#FFFFFF", "#000000", urdu=(h != "AQI"), line=18)
        row_h(t.rows[0], 8)
        for d in ds:
            r = t.add_row(); row_h(r, row_mm)
            has = d.value is not None
            cs = r.cells
            put(cs[0], d.dominant if has else "", 8.5)
            put(cs[1], d.value if has else "", 11, True, band_fill(d.value) if has else NO_DATA_FILL,
                band_text(d.value) if has else "#000000")
            put(cs[2], DISTRICT_URDU.get(d.name, d.name), 11, False, urdu=True, line=17)
            put(cs[3], f"{d.rank}.", 9.5)
        widths(t, [40, 16, 31, 11])
        return t

    right, left = districts[:RIGHT_COLUMN_ROWS], districts[RIGHT_COLUMN_ROWS:]
    main_t = doc.add_table(rows=1, cols=2); main_t.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths(main_t, [99, 99])
    lc, rc = main_t.rows[0].cells
    if bi is not None:
        lo, hi = BANDS[bi][0], BANDS[bi][1]
        mt = lc.add_table(rows=1, cols=1); mt.alignment = WD_TABLE_ALIGNMENT.CENTER
        borders(mt, "#B0B0B0", 6)
        widths(mt, [95])
        mcell = mt.rows[0].cells[0]
        tp = put(mcell, "", 11, align=RIGHT, urdu=True, line=24)
        if "face" in pics:
            picture(tp, pics["face"], height_mm=8); run(tp, " ", 8)
        run(tp, adv_name, 13, True, "#B7950B", urdu=True)
        run(tp, " : ", 11, False, "#000000", urdu=True)
        run(tp, f"({lo}—{hi}) (AQI)", 10, False, "#000000")
        for heading_txt, items in adv_secs:
            put(mcell, heading_txt, 12, True, None, "#000000", align=RIGHT, urdu=True, para=mcell.add_paragraph(), line=22)
            for key, text in items:
                ip = put(mcell, "", 9.5, align=RIGHT, urdu=True, para=mcell.add_paragraph(), line=17)
                if f"ic_{key}" in pics:
                    picture(ip, pics[f"ic_{key}"], height_mm=4.2); run(ip, " ", 6, urdu=True)
                run(ip, text, 9.5, False, "#000000", urdu=True)
        lc.add_paragraph().paragraph_format.line_spacing = Pt(5)
        drop_lead(lc)
    rank_table(lc, left, 7.2)
    rank_table(rc, right, 7.9)
    drop_lead(rc)

    gap(4)
    leg = doc.add_table(rows=2, cols=len(BANDS)); leg.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths(leg, [27.5] * len(BANDS))
    for i, bnd in enumerate(BANDS):
        put(leg.rows[0].cells[i], f"{BAND_LABELS[i]}\n({bnd[2]})", 6.5, True)
        put(leg.rows[1].cells[i], "", 8, False, bnd[4]); row_h(leg.rows[1], 5)
    footer()

    # ================================================================ pages 2 and 3 : shared header
    def page_header(parts, border_fill, border_text, size):
        t = doc.add_table(rows=1, cols=3); t.alignment = WD_TABLE_ALIGNMENT.CENTER
        row_h(t.rows[0], 26)
        a, b, c = t.rows[0].cells
        pa = a.paragraphs[0]; pa.alignment = LEFT
        if LOGO_PUNJAB.exists():
            picture(pa, LOGO_PUNJAB, height_mm=24)
        bt = b.add_table(rows=1, cols=1); bt.alignment = CENTER
        borders(bt, "#111111", 12)
        bc = bt.rows[0].cells[0]
        bp = put(bc, "", size, True, border_fill, border_text, serif=True)
        for text, sz in parts:
            run(bp, text, sz, True, border_text, serif=True)
        up = b.add_paragraph(); up.alignment = CENTER
        up.paragraph_format.space_before = Pt(3)
        run(up, f"Updated Time: {report_date:%d.%m.%Y} ({REPORT_TIME}) 24 Hourly Report", 12, True, serif=True)
        drop_lead(b)
        pc = c.paragraphs[0]; pc.alignment = RIGHT
        if LOGO_EPA.exists():
            picture(pc, LOGO_EPA, height_mm=24)
        return t, bt

    if punjab_map_png and Path(punjab_map_png).exists():
        new_page(False)
        t, bt = page_header([("Average AQI of Punjab ", 17), (f"({n_d:02d} Districts) ", 9), (f"({avg})", 17)],
                            band_fill(punjab_avg), band_text(punjab_avg), 17)
        widths(t, [38, 118, 38]); widths(bt, [114])
        par = doc.add_paragraph(); par.alignment = CENTER
        picture(par, punjab_map_png, height_mm=208)
        footer()

    if focus_d is not None:
        new_page(False, landscape=True)
        n_ok = sum(1 for s in city if s.sufficient)
        fv = "" if focus_d.value is None else focus_d.value
        t, bt = page_header([(f"Average AQI of {focus} ", 17), (f"({n_ok:02d} AQMS) ", 9), (f"({fv})", 17)],
                            band_fill(focus_d.value), band_text(focus_d.value), 17)
        widths(t, [52, 177, 52]); widths(bt, [173])
        body = doc.add_table(rows=1, cols=2); body.alignment = WD_TABLE_ALIGNMENT.CENTER
        widths(body, [196, 85])
        mc, sc = body.rows[0].cells
        if map_png and Path(map_png).exists():
            picture(mc.paragraphs[0], map_png, width_mm=196)
        st = sc.add_table(rows=1, cols=4); borders(st, "#555555", 4)
        for c, h in zip(st.rows[0].cells, ["Rank", "Location", "AQI", "Dominant"]):
            put(c, h, 7.5, True, PINK_HEAD, serif=True)
        row_h(st.rows[0], 8)
        table_rows = len(city) + (len(tb) + 1 if tb else 0)
        station_row_height = (LAHORE_MAP_SIZE_MM[1] - 8) / max(table_rows, 1)

        def add_rows(sts):
            for i, x in enumerate(sts, 1):
                r = st.add_row(); row_h(r, station_row_height)
                put(r.cells[0], f"{i:02d}", 8, False, serif=True)
                put(r.cells[1], x.label, 8, False, serif=True)
                put(r.cells[2], "" if x.value is None else x.value, 9, True, band_fill(x.value), band_text(x.value), serif=True)
                put(r.cells[3], x.dominant if x.value is not None else "", 7, False, serif=True)

        add_rows(city)
        if tb:
            r = st.add_row(); row_h(r, station_row_height)
            m = r.cells[0].merge(r.cells[3])
            mp = put(m, "", 9, align=LEFT, fill=PINK_HEAD)
            run(mp, "● ", 9, True, "#E03C31"); run(mp, "Transboundary AQMS", 9, True, serif=True)
            add_rows(tb)
        widths(st, [12, 22, 12, 39])
        drop_lead(sc)
        footer()

    doc.save(str(path))
    try:
        embed_fonts(path, {"Noto Sans": ("NotoSans-400.ttf", "NotoSans-700.ttf"),
                           "Noto Serif": ("NotoSerif-400.ttf", "NotoSerif-700.ttf"),
                           "Noto Nastaliq Urdu": ("NotoNastaliqUrdu-400.ttf", "NotoNastaliqUrdu-700.ttf")})
    except Exception as exc:
        print(f"  ! fonts were not embedded in the Word file ({exc})")


# =============================================================================
# MAIN
# =============================================================================

def main(argv=None):
    ap = argparse.ArgumentParser(description="Daily AQI Report of Punjab — PDF + AQMS map generator")
    ap.add_argument("--csv", type=Path, default=SCRIPT_DIR / "data",
                    help="station-level dashboard export (graphs_periodic_*.csv), or a folder (newest CSV is used; default: data/)")
    ap.add_argument("--shp", type=Path, default=SCRIPT_DIR / "shp" / "56AQMS.shp",
                    help="AQMS locations point shapefile (default: shp/56AQMS.shp)")
    ap.add_argument("--shp-name-field", help="attribute holding station names (auto-detected if omitted)")
    ap.add_argument("--districts-shp", type=Path, help="optional Punjab district boundaries shapefile")
    ap.add_argument("--out", type=Path, default=Path("output"), help="output folder (default: ./output)")
    ap.add_argument("--data-date", help="day to report, YYYY-MM-DD (default: the full day in the CSV)")
    ap.add_argument("--report-date", help="date printed on the report, YYYY-MM-DD (default: data date + 1)")
    ap.add_argument("--focus", default=FOCUS_DISTRICT, help="district shown on the zoomed map (default: Lahore)")
    ap.add_argument("--basemap", default=BASEMAP,
                    help="osm | voyager | esri-street | positron | satellite | none | an XYZ tile URL with {z}/{x}/{y}")
    ap.add_argument("--basemap-file", type=Path,
                    help="your own street map: GeoTIFF, or PNG/JPG with a world file (.pgw/.jgw) — e.g. exported from QGIS")
    ap.add_argument("--refresh-basemap", action="store_true", help="download a fresh street map instead of the saved copy")
    ap.add_argument("--test-basemap", action="store_true", help="test which street-map providers work on this network, then exit")
    ap.add_argument("--no-pdf", action="store_true",
                    help="do not keep a PDF file; a temporary PDF is still used for the matching Word report")
    a = ap.parse_args(argv)

    log = print
    if a.test_basemap:
        test_basemaps(log)
        return
    a.out.mkdir(parents=True, exist_ok=True)
    data_date = dt.date.fromisoformat(a.data_date) if a.data_date else None
    if a.csv.is_dir():   # a folder was given: use the newest dashboard export in it
        cands = sorted(a.csv.glob("graphs_periodic*.csv")) or sorted(a.csv.glob("*.csv"))
        if not cands:
            sys.exit(f"No CSV files in {a.csv}")
        a.csv = max(cands, key=lambda p: p.stat().st_mtime)
    log(f"Reading {a.csv.name} ...")
    stations, data_date = read_dashboard_csv(a.csv, data_date)
    report_date = dt.date.fromisoformat(a.report_date) if a.report_date else data_date + dt.timedelta(days=1)
    for s in stations:
        compute_station(s)
    report_stations = [s for s in stations if s.sufficient]
    flagged_spikes = sum(int(s.hours.spike.sum()) for s in stations)
    districts = compute_districts(report_stations)
    vals = [d.mean for d in districts if d.mean is not None]
    punjab_avg = float(np.mean(vals)) if vals else None
    qa = qa_flags(stations)
    log(f"  {len(report_stations)}/{len(stations)} stations included, {flagged_spikes} isolated peak hour(s) flagged (included), "
        f"{len(districts)} districts, data date {data_date:%d.%m.%Y}, "
        f"Punjab average AQI {'' if punjab_avg is None else round_half_up(punjab_avg)}")

    extra, match_table = [], None
    if a.shp:
        log(f"Reading AQMS locations {a.shp.name} ...")
        extra, match_table = load_station_locations(a.shp, report_stations, a.shp_name_field, SHP_ALIASES, log)
        miss = [s.name for s in report_stations if s.x is None]
        log(f"  matched {len(report_stations) - len(miss)}/{len(report_stations)} reported stations to the shapefile")
        if miss:
            log(f"    not located ({len(miss)}): {', '.join(miss[:6])}{' ...' if len(miss) > 6 else ''}"
                "  (full list: sheet 'Shapefile_Match')")

    tag = f"{report_date:%d.%m.%Y}"
    focus = canonical_district(a.focus) or a.focus
    city = sorted([s for s in report_stations if in_focus(s, focus) and s.role != "transboundary"],
                  key=lambda s: (s.value is None, -(s.mean or 0)))
    tb = sorted([s for s in report_stations if in_focus(s, focus) and s.role == "transboundary"],
                key=lambda s: (s.value is None, -(s.mean or 0)))
    focus_d = next((d for d in districts if d.name == focus), None)
    districts_shp = a.districts_shp or find_districts_shp(SCRIPT_DIR / "shp")
    map_png = None
    if a.shp:
        log("Drawing AQMS map ...")
        map_png = render_map(report_stations, focus, [], districts_shp,
                             a.out / f"AQMS_Map_{focus.replace(' ', '_')}_{tag}.png", data_date, a.basemap, log,
                             basemap_file=a.basemap_file, refresh_basemap=a.refresh_basemap)

    log("Drawing Punjab district AQI map ...")
    punjab_map_png = render_punjab_map(districts, districts_shp,
                                       a.out / f"Punjab_District_AQI_Map_{tag}.png", data_date, log)

    html_path = a.out / f"DAILY_AQI_REPORT_{tag}.html"
    html_path.write_text(build_html(districts, punjab_avg, focus_d, city, tb, map_png, data_date, report_date, focus,
                                    punjab_map_png), encoding="utf-8")
    pdf_path = a.out / f"DAILY_AQI_REPORT_{tag}.pdf"
    temporary_pdf_dir = None
    docx_source_pdf = pdf_path
    if a.no_pdf:
        import tempfile
        temporary_pdf_dir = tempfile.TemporaryDirectory(prefix="aqi_word_pdf_")
        docx_source_pdf = Path(temporary_pdf_dir.name) / pdf_path.name
        log("Rendering temporary PDF for the matching Word report ...")
    else:
        log("Rendering PDF ...")
    pdf_ready = html_to_pdf(html_path, docx_source_pdf, log)
    xlsx = a.out / f"AQI_Summary_{tag}.xlsx"
    write_excel(xlsx, districts, stations, qa, match_table, punjab_avg)
    docx_path = a.out / f"DAILY_AQI_REPORT_{tag}.docx"
    try:
        if pdf_ready:
            write_docx_from_pdf(docx_source_pdf, docx_path)
        else:
            log("  ! Word report skipped because the matching PDF could not be rendered")
    except ImportError as exc:
        log(f"  ! Word report skipped: install python-docx and PyMuPDF ({exc})")
    finally:
        if temporary_pdf_dir is not None:
            temporary_pdf_dir.cleanup()

    log("\nDistrict ranking:")
    for d in districts:
        log(f"  {d.rank:>2}. {d.name:<16} {'' if d.value is None else d.value:>4}  {d.dominant}")
    if len(qa):
        log(f"\nQA: {len(qa)} flag(s) — see sheet 'QA_Flags' in {xlsx.name}")
    log(f"\nDone. Output in: {a.out.resolve()}")


if __name__ == "__main__":
    main()
