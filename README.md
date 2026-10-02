# Daily AQI Report of Punjab — generator

Generates the **Daily AQI Report** from the station-level dashboard CSV: a Punjab district ranking in Urdu, a Punjab district AQI map, and a zoomed AQMS map with its station table.

## 1. One-time setup (Windows, VS Code terminal)

```bat
pip install -r requirements.txt
python -m playwright install chromium
```

Playwright's Chromium turns the HTML page into the PDF. It lays out Urdu Nastaliq correctly, which matplotlib and ReportLab cannot do.

## 2. Folder layout

```
daily_aqi_report/
  daily_aqi_report_Punjab.py <- the script (all settings at the top)
  run_daily_report.bat     <- double-click to run
  requirements.txt
  assets/                  <- logos, helpline banner, Pakistan–India border line
  fonts/                   <- Noto Nastaliq Urdu + Noto Serif/Sans (bundled, OFL)
  data/                    <- put graphs_periodic_YYYY-MM-DD.csv here (newest is used)
  shp/                     <- 56AQMS.shp and District.shp (Punjab district boundaries)
  output/                  <- results
```

## 3. Run

```bat
python daily_aqi_report_Punjab.py --csv data --shp shp\56AQMS.shp --districts-shp shp\District.shp
```

You can also edit the paths in `run_daily_report.bat` and double-click it.

## 3a. Upload app

Run the upload app locally with:

```bat
streamlit run app.py
```

Upload the station-level dashboard CSV. The app uses the bundled AQMS shapefile
and generates the HTML, PDF, PNG map, and Excel summary. Download everything as
one ZIP or download individual files.

Useful options:

| Option | Meaning |
|---|---|
| `--csv FILE or FOLDER` | Station-level export. If you give a folder, the newest CSV in it is used. |
| `--shp FILE` | AQMS point shapefile, in any CRS. The name field is detected automatically. |
| `--shp-name-field NAME` | Sets the station-name attribute yourself if the detected one is wrong. |
| `--districts-shp FILE` | Punjab district boundaries used for the district AQI map and the focus-district map. |
| `--basemap osm/voyager/esri-street/positron/satellite/none` | Street-map provider. The default is `osm`; you can also pass your own `https://.../{z}/{x}/{y}.png`. |
| `--basemap-file FILE` | Your own street map: a GeoTIFF, or a PNG/JPG with a `.pgw`/`.jgw` world file. |
| `--refresh-basemap` / `--test-basemap` | Download a fresh copy of the street map / test which providers work on this network. |
| `--focus "Lahore"` | District shown on the zoomed map (page 3). |
| `--report-date 2026-09-30` | Date printed on the report. The default is the data date + 1. |
| `--min-hours 18` | Minimum valid hours for a station to be reported. |
| `--keep-zero-aqi` | Counts "AQI 0 with no pollutant" hours the way the dashboard Average row does. |

## 4. Outputs (in `output/`)

- `DAILY_AQI_REPORT_dd.mm.yyyy.pdf`: page 1 (Urdu ranking), page 2 (Punjab district AQI map), and page 3 (focus-district AQMS map)
- `Punjab_District_AQI_Map_dd.mm.yyyy.png`: Punjab district AQI map
- `AQMS_Map_Lahore_dd.mm.yyyy.png`: the zoomed map at 300 dpi, for WhatsApp or slides
- `AQI_Summary_dd.mm.yyyy.xlsx`: sheets Districts, Stations, QA_Flags, Shapefile_Match, Hourly
- `DAILY_AQI_REPORT_dd.mm.yyyy.html`: the same report as a web page

## 5. How the figures are calculated

- **Station AQI** is the mean of the day's valid hourly AQI. A station needs at least 18 valid hours; with fewer it is shown as "ڈیٹا دستیاب نہیں". An hour with AQI 0 and no dominant pollutant counts as a gap.
- **District AQI** is the mean of its stations. Transboundary stations (Lathepur, Wagha, BHU Jandiala) are listed separately and are not averaged in.
- **Punjab average** is the mean of the district values.
- **Dominant pollutants** are the three pollutants that were dominant in the most hours. Ties go to the one seen first.
- Values are rounded half-up. Districts are ranked by value; ties are broken by the unrounded value.

## 6. Customising

Everything is in the CONFIGURATION block at the top of the script:

- `STATION_REGISTRY`: short names, district and city/transboundary role. A new station not listed there is given its district from its name automatically.
- `DISTRICT_URDU`: Urdu district names.
- `PUBLIC_MESSAGES`: the Urdu public message for each AQI band. Page 1 picks the message that matches the Punjab average. The Moderate text is copied from the current report; please have the other bands vetted.
- `BANDS` / `BAND_LABELS`: colours and labels.
- Logos: replace the files in `assets/` with high-resolution versions, keeping the same names.

## 7. Street map (basemap)

The map tries these sources in order and uses the first one that works:

1. **Your own file** in `basemap/`, named after the focus district: `Lahore.tif`, or `Lahore.png` with `Lahore.pgw` beside it. You can also point to any file with `--basemap-file`.
2. **The copy saved from an earlier download** in `basemap/cache/`. The AQMS locations don't move, so after one successful download the report works offline every day. Use `--refresh-basemap` to fetch a new copy.
3. **A download** from `--basemap` (default `osm`), then each provider in `BASEMAP_FALLBACKS` (`voyager`, then `esri-street`).
4. **A clean offline style**, used only when nothing above works.

To see which providers work on your network, run:

```bat
python daily_aqi_report_Punjab.py --test-basemap
```

Each provider is reported as OK or FAIL, with the reason and a hint. For example, an SSL error means you need `pip install truststore`, and HTTP 403 means you should try another provider.

**To make your own street map in QGIS** (one time, and it then works with no internet):

1. Set the Project CRS to EPSG:3857.
2. In the Browser panel, open XYZ Tiles and add OpenStreetMap.
3. Zoom to an area slightly larger than all the Lahore AQMS.
4. Go to Project > Import/Export > Export Map to Image. Set about 300 dpi and tick **"Append georeference information (embedded or via world file)"**.
5. Save the image as `basemap/Lahore.png`. QGIS writes `Lahore.pgw` next to it, and the script picks both up automatically.

## 8. Troubleshooting

- **A station is missing from the map.** Open the `Shapefile_Match` sheet. It lists every CSV name next to the shapefile name it was matched with.
- **You prefer Jameel Noori Nastaleeq.** If that font is installed on the PC, it is used automatically. Otherwise the bundled Noto Nastaliq Urdu is used.
- **Chromium is not found.** Run `python -m playwright install chromium`, or set the `CHROMIUM_PATH` environment variable to a Chrome or Edge executable.
