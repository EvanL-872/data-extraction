"""Streamlit interface for in-memory PR extraction."""

from __future__ import annotations

import re

import pandas as pd
import streamlit as st

from extractor import (
    TemplateError,
    combine_results,
    discover_pr_sheets,
    error_issue,
    extract_pr,
    read_pr_metadata,
)
from reporting import analysis_frames, build_excel_export


st.set_page_config(page_title="PR Consolidator", page_icon="📊", layout="wide")

st.title("Purchase Request Consolidator")
st.write(
    "Upload PR workbooks, choose the worksheet to extract from each file, and download one "
    "analysis-ready Excel workbook. Hardware/items and professional services are separated automatically."
)
st.info(
    "Privacy: files are processed only in this running session. The app uses no database and does not write "
    "uploaded PRs or the generated report to disk. Refreshing or closing the session clears its data."
)

uploads = st.file_uploader(
    "Purchase request workbooks (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
    help="Select one or more PR workbooks. You will select a PR worksheet for every workbook below.",
)


def display_name(upload: object) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", getattr(upload, "name", "workbook"))[:80]


selections: list[dict[str, object]] = []
if uploads:
    st.subheader("Choose the PR worksheet")
    for index, upload in enumerate(uploads):
        file_bytes = upload.getvalue()
        try:
            sheets = discover_pr_sheets(file_bytes)
        except Exception as error:  # Invalid or protected workbooks should not affect other uploads.
            st.error(f"{upload.name}: could not read the workbook ({error}).")
            continue
        if not sheets:
            st.warning(f"{upload.name}: no supported PR worksheet was detected. It will be reported in the review sheet if selected manually is unavailable.")
            continue
        default_index = 0
        chosen = st.selectbox(
            f"Worksheet for {upload.name}",
            sheets,
            index=default_index,
            key=f"sheet_{index}_{display_name(upload)}",
        )
        metadata = read_pr_metadata(file_bytes, chosen)
        values = [
            f"Customer: {metadata.get('customer') or 'not detected'}",
            f"Sales rep: {metadata.get('sales_rep') or 'not detected'}",
            f"Sell currency: {metadata.get('sell_currency') or 'not detected'}",
        ]
        st.caption(" • ".join(values))
        selections.append({"name": upload.name, "content": file_bytes, "sheet": chosen})

if selections:
    st.divider()
    if st.button("Extract and build Excel report", type="primary", use_container_width=True):
        extracted = []
        failures = []
        with st.spinner("Extracting supplier, customer, item, price and category information…"):
            for selection in selections:
                try:
                    extracted.append(extract_pr(selection["content"], selection["name"], selection["sheet"]))
                except (TemplateError, KeyError, ValueError) as error:
                    failures.append(error_issue(selection["name"], selection["sheet"], str(error)))
                except Exception as error:  # Continue extracting unaffected PRs while retaining an auditable error.
                    failures.append(error_issue(selection["name"], selection["sheet"], f"Unexpected extraction error: {error}"))
        if not extracted:
            st.error("No selected worksheet could be extracted. Review the errors below and select a PR worksheet with the expected A:I layout.")
            if failures:
                st.dataframe(pd.concat(failures, ignore_index=True), use_container_width=True, hide_index=True)
        else:
            frames = combine_results(extracted)
            if failures:
                frames["issues"] = pd.concat([frames["issues"], *failures], ignore_index=True)
            st.session_state["pr_frames"] = frames
            st.session_state["pr_export"] = build_excel_export(frames)

if "pr_frames" in st.session_state:
    frames = st.session_state["pr_frames"]
    analyses = analysis_frames(frames)
    st.divider()
    st.subheader("Extraction result")
    metrics = st.columns(4)
    metrics[0].metric("Items extracted", len(frames["items"]))
    metrics[1].metric("Professional services", int((frames["items"]["segment"] == "Professional Services").sum()))
    metrics[2].metric("Hardware / other items", int((frames["items"]["segment"] == "Hardware / Items").sum()))
    metrics[3].metric("Review issues", len(frames["issues"]))
    st.download_button(
        "Download consolidated Excel report",
        data=st.session_state["pr_export"],
        file_name="pr_consolidated_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    st.caption("Use the exported All Items, Project Summary, Customer Revenue and Category Summary worksheets directly in Power BI or Tableau.")
    summary_tab, items_tab, services_tab, review_tab = st.tabs([
        "Project & customer summary", "All items", "Professional services", "Review issues",
    ])
    with summary_tab:
        st.dataframe(analyses["project_summary"], use_container_width=True, hide_index=True)
        st.dataframe(analyses["customer_revenue"], use_container_width=True, hide_index=True)
        st.dataframe(analyses["category_summary"], use_container_width=True, hide_index=True)
    with items_tab:
        st.dataframe(frames["items"], use_container_width=True, hide_index=True)
    with services_tab:
        st.dataframe(frames["items"][frames["items"]["segment"] == "Professional Services"], use_container_width=True, hide_index=True)
    with review_tab:
        st.dataframe(frames["issues"], use_container_width=True, hide_index=True)
elif uploads:
    st.caption("Select a worksheet for each uploaded file, then run the extraction.")
else:
    st.caption("Start by uploading one or more .xlsx purchase-request files.")
