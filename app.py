from __future__ import annotations

import contextlib
import importlib.util
import io
import traceback
from pathlib import Path
from tempfile import TemporaryDirectory

import streamlit as st


APP_DIR = Path(__file__).resolve().parent
PIPELINE_PATH = APP_DIR / "build_maela_epi_quarterly_report.py"
DEFAULT_OUTPUT_NAME = "MaeLa_Camp_EPI_Quarterly_Report.xlsx"


def load_pipeline_module(script_path: Path):
    if not script_path.exists():
        raise FileNotFoundError(f"Pipeline script not found: {script_path}")

    spec = importlib.util.spec_from_file_location("maela_epi_pipeline", str(script_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"Could not import pipeline from: {script_path}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_pipeline(input_bytes: bytes, source_name: str, output_name: str, sheet_arg: str) -> tuple[bytes, str, str]:
    module = load_pipeline_module(PIPELINE_PATH)
    if not hasattr(module, "build_maela_report"):
        raise AttributeError("build_maela_report function was not found in pipeline script.")

    with TemporaryDirectory(prefix="maela_streamlit_") as temp_dir:
        temp_dir_path = Path(temp_dir)
        input_path = temp_dir_path / source_name
        output_path = temp_dir_path / output_name

        input_path.write_bytes(input_bytes)

        parsed_sheet: str | int
        parsed_sheet = int(sheet_arg) if str(sheet_arg).strip().isdigit() else str(sheet_arg).strip()

        log_buffer = io.StringIO()
        with contextlib.redirect_stdout(log_buffer):
            final_output_path = module.build_maela_report(
                input_path=input_path,
                output_path=output_path,
                sheet_arg=parsed_sheet,
            )

        final_output_path = Path(final_output_path)
        if not final_output_path.exists():
            raise FileNotFoundError("Expected output workbook was not created.")

        return final_output_path.read_bytes(), log_buffer.getvalue(), final_output_path.name


def parse_pipeline_messages(logs: str) -> tuple[list[str], list[str], list[str], list[str]]:
    warnings: list[str] = []
    infos: list[str] = []
    oks: list[str] = []
    others: list[str] = []

    for raw_line in logs.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("[WARNING]") or line.startswith("WARNING:"):
            warnings.append(line)
        elif line.startswith("[INFO]") or line.startswith("INFO:"):
            infos.append(line)
        elif line.startswith("[OK]"):
            oks.append(line)
        else:
            others.append(line)

    return warnings, infos, oks, others


st.set_page_config(page_title="MaeLa EPI Report", layout="wide")
st.title("MaeLa Camp EPI Quarterly Report Builder")
st.caption("Upload MaeLa_EPI_report, run the pipeline, and download the output workbook.")

if not PIPELINE_PATH.exists():
    st.error(f"Required file not found: {PIPELINE_PATH.name}")
    st.stop()

upload = st.file_uploader("Upload source workbook", type=["xls", "xlsx", "xlsm"])
output_name = st.text_input("Output filename", value=DEFAULT_OUTPUT_NAME).strip() or DEFAULT_OUTPUT_NAME
sheet_arg = st.text_input("Source sheet (name or index)", value="0").strip() or "0"

run_clicked = st.button("Run report", type="primary", use_container_width=True)

if upload is None:
    st.info("Please upload an Excel workbook to continue.")
    st.stop()

if not run_clicked:
    st.success(f"Ready to process: {upload.name}")
    st.stop()

if not output_name.lower().endswith(".xlsx"):
    output_name = f"{output_name}.xlsx"

progress = st.progress(10, text="Loading pipeline...")

try:
    progress.progress(45, text="Processing workbook...")
    output_bytes, logs, produced_name = run_pipeline(
        input_bytes=upload.getvalue(),
        source_name=upload.name,
        output_name=output_name,
        sheet_arg=sheet_arg,
    )
    progress.progress(100, text="Done")

    st.success("Report created successfully.")

    warnings, infos, oks, others = parse_pipeline_messages(logs)

    st.subheader("Warning Signs")
    if warnings:
        for line in warnings:
            st.warning(line)
    else:
        st.info("No warning signs were found during processing.")

    st.subheader("Checks")
    if oks:
        for line in oks:
            st.success(line)
    else:
        st.info("No check lines were generated.")

    st.subheader("Notes")
    if infos:
        for line in infos:
            st.info(line)
    else:
        st.info("No info lines were generated.")

    if logs.strip():
        st.subheader("Pipeline logs")
        if others:
            st.caption("Additional log lines")
        st.code(logs, language="text")

    st.download_button(
        label="Download report",
        data=output_bytes,
        file_name=produced_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
except Exception as exc:
    progress.empty()
    st.error(f"Pipeline failed: {exc}")
    st.code(traceback.format_exc(), language="text")
