from __future__ import annotations

import inspect
import io
import os
import tempfile
import zipfile
from contextlib import redirect_stdout
from pathlib import Path

import streamlit as st

try:                                    # the Punjab version of the report script ...
    import daily_aqi_report_Punjab as daily_aqi_report
except ImportError:                     # ... or the same script saved under the original name
    import daily_aqi_report


ROOT = Path(__file__).resolve().parent
SHP_DIR = ROOT / "shp"
DEFAULT_SHP = SHP_DIR / "56AQMS.shp"                     # AQMS station locations (Lahore map)
DISTRICTS_SHP = (daily_aqi_report.find_districts_shp(SHP_DIR)   # Punjab district boundaries (Punjab map, page 2)
                 if hasattr(daily_aqi_report, "find_districts_shp") else None)

MIME = {".pdf": "application/pdf", ".html": "text/html", ".png": "image/png",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
ORDER = {".pdf": 0, ".xlsx": 1, ".png": 2, ".html": 3}


def wide() -> dict:
    """Full-width button argument that works on old and new Streamlit versions."""
    if "width" in inspect.signature(st.button).parameters:
        return {"width": "stretch"}
    return {"use_container_width": True}


def run_report(uploaded_csv, focus: str, basemap: str, keep_zero: bool) -> tuple[dict[str, bytes], str]:
    with tempfile.TemporaryDirectory() as work_dir:
        work = Path(work_dir)
        csv_path = work / uploaded_csv.name
        csv_path.write_bytes(uploaded_csv.getvalue())
        output_dir = work / "output"
        args = [
            "--csv", str(csv_path),
            "--shp", str(DEFAULT_SHP),
            "--out", str(output_dir),
            "--focus", focus,
            "--basemap", basemap,
        ]
        if DISTRICTS_SHP is not None:
            args += ["--districts-shp", str(DISTRICTS_SHP)]
        if keep_zero:
            args.append("--keep-zero-aqi")

        if "CHROMIUM_PATH" not in os.environ:
            for candidate in ("/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome"):
                if Path(candidate).exists():
                    os.environ["CHROMIUM_PATH"] = candidate
                    break

        log = io.StringIO()
        try:
            with redirect_stdout(log):
                daily_aqi_report.main(args)
        except SystemExit as exc:
            message = str(exc) or "The report generator stopped unexpectedly."
            raise RuntimeError(message) from exc

        files = {
            path.name: path.read_bytes()
            for path in sorted(output_dir.iterdir(), key=lambda p: (ORDER.get(p.suffix.lower(), 9), p.name))
            if path.is_file()
        }
        return files, log.getvalue()


def zip_files(files: dict[str, bytes]) -> bytes:
    archive = io.BytesIO()
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content)
    return archive.getvalue()


st.set_page_config(page_title="Daily AQI Report", page_icon="🌿", layout="centered")
st.title("Daily AQI Report")
st.caption("Upload the station-level dashboard CSV to generate the complete report package.")

uploaded_csv = st.file_uploader("Station-level CSV", type=["csv"], help="Use the graphs_periodic export with hourly '<Station> • AQI' columns.")

with st.expander("Report options", expanded=True):
    focus = st.text_input("Map district", value=daily_aqi_report.FOCUS_DISTRICT)
    basemap = st.selectbox("Basemap", ["osm", "voyager", "esri-street", "positron", "satellite", "none"], index=0)
    keep_zero = st.checkbox("Count AQI = 0 rows with no pollutant", value=False)

if DISTRICTS_SHP is None:
    st.warning("No Punjab district shapefile found in the shp folder (a file with 'dist' in its name), "
               "so the Punjab district map page will be left out of the report.")

if st.button("Generate report", type="primary", disabled=uploaded_csv is None, **wide()):
    if not DEFAULT_SHP.exists():
        st.error("The bundled AQMS shapefile is missing from the deployment.")
    else:
        with st.spinner("Calculating AQI and building report files..."):
            try:
                files, log = run_report(uploaded_csv, focus, basemap, keep_zero)
            except Exception as exc:
                st.error(str(exc))
            else:
                st.session_state["report_files"] = files
                st.session_state["report_log"] = log

files = st.session_state.get("report_files")
if files:
    log_text = st.session_state.get("report_log", "")
    st.success(f"Generated {len(files)} output files.")
    problems = [ln.strip().lstrip("! ").strip() for ln in log_text.splitlines() if ln.strip().startswith("!")]
    if problems:
        st.warning("Please check:\n\n" + "\n".join(f"- {p}" for p in problems))

    st.download_button(
        "Download all outputs (ZIP)",
        data=zip_files(files),
        file_name="daily_aqi_report_outputs.zip",
        mime="application/zip",
        **wide(),
    )

    maps = {n: c for n, c in files.items() if n.lower().endswith(".png")}
    if maps:
        st.subheader("Map previews")
        labels = ["Punjab districts" if "Punjab_District" in n else f"{focus} stations" for n in maps]
        for tab, (name, content) in zip(st.tabs(labels), maps.items()):
            with tab:
                st.image(content, caption=name)

    st.subheader("Individual files")
    for name, content in files.items():
        mime = MIME.get(Path(name).suffix.lower(), "application/octet-stream")
        st.download_button(name, content, file_name=name, mime=mime, key=f"download-{name}")
    with st.expander("Generation log"):
        st.code(log_text)
