import streamlit as st
from converter import convert_apc_to_oit
from qa_compare import create_qa_report

st.set_page_config(
    page_title="Engineering Data Tools",
    page_icon="🧰",
    layout="wide",
)

st.title("Engineering Data Tools")
st.caption("APC1 → OIT Converter  |  OIT vs webDM QA Comparison")

tab1, tab2 = st.tabs(["Tab 1 · APC1 → OIT", "Tab 2 · QA PSV Compare"])

with tab1:
    st.header("APC1 → OIT Converter")
    st.write("Upload an APC1 workbook and convert it to the OIT2 template.")
    apc_file = st.file_uploader(
        "Upload APC1 Excel",
        type=["xlsx", "xlsm"],
        key="apc1_upload",
        help="Select the APC1 source workbook (.xlsx or .xlsm).",
    )
    if apc_file:
        st.caption(f"Selected: {apc_file.name} · {apc_file.size / 1024 / 1024:.2f} MB")
        if st.button("Convert APC1 to OIT", type="primary", key="convert_apc"):
            try:
                with st.spinner("Converting APC1 workbook..."):
                    result, filename, rows = convert_apc_to_oit(apc_file.getvalue(), apc_file.name)
                st.session_state["apc_output"] = result
                st.session_state["apc_output_name"] = filename
                st.session_state["apc_rows"] = rows
            except Exception as exc:
                st.error(f"APC1 conversion failed: {exc}")
    if "apc_output" in st.session_state:
        st.success(f"Conversion completed · Rows converted: {st.session_state['apc_rows']}")
        st.download_button(
            "Download OIT Excel",
            data=st.session_state["apc_output"],
            file_name=st.session_state["apc_output_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_apc",
        )

with tab2:
    st.header("QA PSV Comparison")
    st.write("Compare OIT1 Feature IDs against webDM2 Device Listing Device Names.")
    oit_file = st.file_uploader(
        "Upload OIT1 Excel",
        type=["xlsx", "xlsm"],
        key="oit_upload",
    )
    webdm_file = st.file_uploader(
        "Upload webDM2 file",
        type=["psv", "xlsx", "xlsm"],
        key="webdm_upload",
        help="Supports raw .psv, PSV-export Excel (pipe-separated rows in column A), or normal Excel Device Listing.",
    )
    if oit_file and webdm_file:
        st.caption(
            f"OIT1: {oit_file.name} · webDM2: {webdm_file.name}"
        )
        if st.button("Run QA Comparison", type="primary", key="run_qa"):
            try:
                with st.spinner("Comparing OIT1 and webDM2..."):
                    result, filename, stats = create_qa_report(
                        oit_file.getvalue(), oit_file.name,
                        webdm_file.getvalue(), webdm_file.name,
                    )
                st.session_state["qa_output"] = result
                st.session_state["qa_output_name"] = filename
                st.session_state["qa_stats"] = stats
            except Exception as exc:
                st.error(f"QA comparison failed: {exc}")
    if "qa_stats" in st.session_state:
        stats = st.session_state["qa_stats"]
        st.success("QA comparison completed.")
        cols = st.columns(4)
        cols[0].metric("Matched IDs", stats["matched"])
        cols[1].metric("Different fields", stats["differences"])
        cols[2].metric("OIT-only IDs", stats["oit_only"])
        cols[3].metric("webDM-only IDs", stats["webdm_only"])
        st.download_button(
            "Download QA Excel Report",
            data=st.session_state["qa_output"],
            file_name=st.session_state["qa_output_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_qa",
        )

st.divider()
st.caption(
    "Files are processed in application memory. Use only company-approved data "
    "and hosting environments; a public app URL is accessible to anyone."
)
