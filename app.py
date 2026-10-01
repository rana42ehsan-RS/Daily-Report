from __future__ import annotations

import io
import shutil
import sys
import tempfile
import zipfile
from contextlib import redirect_stdout
from pathlib import Path

import streamlit as st

import daily_aqi_report


ROOT = Path(__file__).resolve().parent
DEFAULT_SHP = ROOT / "shp" / "56AQMS.shp"


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
        if keep_zero:
            args.append("--keep-zero-aqi")

        log = io.StringIO()
        try:
            with redirect_stdout(log):
                daily_aqi_report.main(args)
        except SystemExit as exc:
            message = str(exc) or "The report generator stopped unexpectedly."
            raise RuntimeError(message) from exc

        files = {
            path.name: path.read_bytes()
            for path in sorted(output_dir.iterdir())
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

if st.button("Generate report", type="primary", disabled=uploaded_csv is None, use_container_width=True):
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
    st.success(f"Generated {len(files)} output files.")
    st.download_button(
        "Download all outputs (ZIP)",
        data=zip_files(files),
        file_name="daily_aqi_report_outputs.zip",
        mime="application/zip",
        use_container_width=True,
    )
    st.subheader("Individual files")
    for name, content in files.items():
        suffix = Path(name).suffix.lower()
        mime = {".pdf": "application/pdf", ".html": "text/html", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", ".png": "image/png"}.get(suffix, "application/octet-stream")
        st.download_button(name, content, file_name=name, mime=mime, key=f"download-{name}")
    with st.expander("Generation log"):
        st.code(st.session_state.get("report_log", ""))