"""表格清洗助手（Excel & CSV Cleaner）的 Streamlit 本地界面。"""

from __future__ import annotations

import hashlib
import os
from pathlib import Path

os.environ.setdefault("STREAMLIT_BROWSER_GATHER_USAGE_STATS", "false")

import pandas as pd
import streamlit as st

from cleaner import (
    CSV_DELIMITERS,
    CleaningOptions,
    CleaningStats,
    clean_dataframe,
    decode_csv_bytes,
    detect_csv_delimiter,
    export_csv_bytes,
    export_excel_bytes,
    list_excel_sheets,
    output_filename,
    read_csv_bytes,
    read_excel_bytes,
)


MAX_FILE_BYTES = 50 * 1024 * 1024
LARGE_FILE_BYTES = 15 * 1024 * 1024
LARGE_ROW_COUNT = 100_000
PREVIEW_ROWS = 200
PROJECT_ROOT = Path(__file__).resolve().parent
SAMPLE_PATH = PROJECT_ROOT / "samples" / "sample.csv"


st.set_page_config(
    page_title="表格清洗助手",
    page_icon="🧹",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      .block-container { padding-top: 2rem; padding-bottom: 3rem; }
      .hero {
        padding: 1.5rem 1.7rem;
        border-radius: 18px;
        color: white;
        background: linear-gradient(120deg, #123b63 0%, #176b87 60%, #2f9e8f 100%);
        box-shadow: 0 10px 30px rgba(18, 59, 99, .18);
        margin-bottom: 1.35rem;
      }
      .hero h1 { color: white; margin: 0 0 .5rem 0; font-size: 2rem; }
      .hero p { margin: 0; opacity: .92; font-size: 1rem; }
      .step-card {
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 1rem 1.1rem;
        background: #fbfdff;
        min-height: 112px;
      }
      .step-card strong { color: #155e75; }
      div[data-testid="stMetric"] {
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: .7rem .85rem;
        background: white;
      }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(show_spinner=False)
def file_digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


@st.cache_data(show_spinner=False, max_entries=8)
def cached_excel_sheets(data: bytes) -> list[str]:
    return list_excel_sheets(data)


@st.cache_data(show_spinner=False, max_entries=8)
def cached_read_csv(data: bytes, encoding: str, delimiter: str) -> pd.DataFrame:
    return read_csv_bytes(data, encoding=encoding, delimiter=delimiter)


@st.cache_data(show_spinner=False, max_entries=8)
def cached_read_excel(data: bytes, sheet_name: str) -> pd.DataFrame:
    return read_excel_bytes(data, sheet_name=sheet_name)


def clear_cleaned_result() -> None:
    for key in ("cleaned_data", "cleaning_stats", "cleaned_options_signature"):
        st.session_state.pop(key, None)


def load_sample_data() -> None:
    clear_cleaned_result()
    st.session_state["use_sample"] = True
    st.session_state["uploader_version"] = (
        st.session_state.get("uploader_version", 0) + 1
    )


def reset_application() -> None:
    st.session_state["use_sample"] = False
    st.session_state["uploader_version"] = (
        st.session_state.get("uploader_version", 0) + 1
    )
    clear_cleaned_result()
    for key in list(st.session_state.keys()):
        if key.startswith(
            (
                "sheet_choice_",
                "csv_delimiter_",
                "clean_",
                "export_format",
                "csv_utf8_bom",
                "protect_formulas",
            )
        ):
            st.session_state.pop(key, None)
    st.session_state.pop("source_signature", None)


if "uploader_version" not in st.session_state:
    st.session_state["uploader_version"] = 0
if "use_sample" not in st.session_state:
    st.session_state["use_sample"] = False


def show_intro() -> None:
    st.markdown(
        """
        <div class="hero">
          <h1>表格清洗助手</h1>
          <p>在本地整理 Excel 与 CSV：去重、清理空行空列和多余空格，然后安全导出。</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown(
            '<div class="step-card"><strong>1. 导入</strong><br>'
            "选择 CSV、XLSX 或 XLSM 文件；Excel 可切换工作表。</div>",
            unsafe_allow_html=True,
        )
    with col2:
        st.markdown(
            '<div class="step-card"><strong>2. 清洗</strong><br>'
            "逐项勾选需要的操作，预览统计后再下载结果。</div>",
            unsafe_allow_html=True,
        )
    with col3:
        st.markdown(
            '<div class="step-card"><strong>3. 导出</strong><br>'
            "另存为新的 CSV 或 XLSX，原文件始终不会被覆盖。</div>",
            unsafe_allow_html=True,
        )

    st.info(
        "文件只在本机浏览器与 Python 进程之间处理，不会上传到第三方服务。"
        "重要文件请先另做备份。",
        icon="🔒",
    )


def render_import_controls() -> tuple[bytes, str] | None:
    uploaded_file = st.file_uploader(
        "选择要清洗的文件",
        type=["csv", "xlsx", "xlsm"],
        help="单个文件最大 50 MB。文件仅在本机内存中读取。",
        key=f"uploader_{st.session_state['uploader_version']}",
    )

    if uploaded_file is not None:
        return uploaded_file.getvalue(), uploaded_file.name

    if st.session_state["use_sample"]:
        try:
            return SAMPLE_PATH.read_bytes(), SAMPLE_PATH.name
        except OSError:
            st.error("示例数据文件缺失，请检查 samples/sample.csv 是否存在。")

    return None


def render_raw_preview(dataframe: pd.DataFrame) -> None:
    st.subheader("原始数据预览")
    row_metric, column_metric, preview_metric = st.columns(3)
    row_metric.metric("数据行数", f"{len(dataframe):,}")
    column_metric.metric("数据列数", f"{len(dataframe.columns):,}")
    preview_metric.metric(
        "当前预览",
        f"前 {min(len(dataframe), PREVIEW_ROWS):,} 行",
    )

    if dataframe.empty:
        st.info("表格没有数据行。你仍可查看列名，或执行空表清洗测试。")
    else:
        st.dataframe(
            dataframe.head(PREVIEW_ROWS),
            width="stretch",
            hide_index=True,
        )
    if len(dataframe) > PREVIEW_ROWS:
        st.caption(f"为保持页面流畅，仅展示前 {PREVIEW_ROWS} 行。")
    if len(dataframe) > LARGE_ROW_COUNT:
        st.warning(
            "数据超过 10 万行，清洗和预览可能需要较长时间。建议先关闭其他程序，"
            "并在操作完成前不要重复点击。",
            icon="⏳",
        )


def render_cleaning_controls(
    columns: list[object],
) -> tuple[CleaningOptions, bool]:
    st.subheader("清洗选项")
    st.caption("所有选项默认关闭。点击“执行清洗”前，原始预览不会被改动。")

    with st.form("cleaning_form"):
        left, right = st.columns(2)
        with left:
            trim_whitespace = st.checkbox(
                "清理文本单元格首尾空格",
                key="clean_trim_whitespace",
            )
            drop_blank_rows = st.checkbox(
                "删除完全空白的行",
                key="clean_blank_rows",
            )
            drop_blank_columns = st.checkbox(
                "删除完全空白的列",
                key="clean_blank_columns",
            )
        with right:
            drop_duplicates = st.checkbox(
                "删除完全重复的行",
                key="clean_duplicates",
            )
            deduplicate_by_columns = st.checkbox(
                "按勾选列删除重复记录",
                key="clean_subset_duplicates",
            )
            case_sensitive = st.checkbox(
                "区分大小写（仅适用于文本重复判断）",
                value=True,
                key="clean_case_sensitive",
            )

        selected_indexes: list[int] = []
        if deduplicate_by_columns:
            selected_indexes = st.multiselect(
                "用于重复判断的列",
                options=list(range(len(columns))),
                format_func=lambda index: f"{index + 1}. {columns[index]}",
                key="clean_duplicate_columns",
                help="只有所选列的值都相同时，后出现的记录才会被删除。",
            )

        submitted = st.form_submit_button(
            "执行清洗",
            type="primary",
            width="stretch",
        )

    selected_columns = tuple(columns[index] for index in selected_indexes)
    options = CleaningOptions(
        drop_duplicate_rows=drop_duplicates,
        drop_blank_rows=drop_blank_rows,
        drop_blank_columns=drop_blank_columns,
        trim_whitespace=trim_whitespace,
        deduplicate_by_columns=deduplicate_by_columns,
        duplicate_columns=selected_columns,
        case_sensitive=case_sensitive,
    )
    return options, submitted


def render_result(
    source_name: str,
    options: CleaningOptions,
) -> None:
    cleaned_data = st.session_state.get("cleaned_data")
    stats: CleaningStats | None = st.session_state.get("cleaning_stats")
    stored_signature = st.session_state.get("cleaned_options_signature")

    if cleaned_data is None or stats is None:
        return

    st.divider()
    st.subheader("清洗结果")
    if stored_signature != options.signature():
        st.info("清洗选项已更改，请重新点击“执行清洗”以查看最新结果。")
        return

    row1 = st.columns(4)
    row1[0].metric("原始行数", f"{stats.original_rows:,}")
    row1[1].metric("原始列数", f"{stats.original_columns:,}")
    row1[2].metric("清洗后行数", f"{stats.result_rows:,}")
    row1[3].metric("清洗后列数", f"{stats.result_columns:,}")

    row2 = st.columns(3)
    row2[0].metric("删除重复行", f"{stats.removed_duplicate_rows:,}")
    row2[1].metric("删除空白行", f"{stats.removed_blank_rows:,}")
    row2[2].metric("删除空白列", f"{stats.removed_blank_columns:,}")

    st.dataframe(
        cleaned_data.head(PREVIEW_ROWS),
        width="stretch",
        hide_index=True,
    )
    if len(cleaned_data) > PREVIEW_ROWS:
        st.caption(f"结果共 {len(cleaned_data):,} 行，页面仅预览前 {PREVIEW_ROWS} 行。")

    st.subheader("导出结果")
    export_format = st.radio(
        "选择导出格式",
        options=("CSV", "XLSX"),
        horizontal=True,
        key="export_format",
    )

    utf8_bom = True
    if export_format == "CSV":
        utf8_bom = st.checkbox(
            "使用 UTF-8 BOM（推荐，Windows Excel 打开中文更稳定）",
            value=True,
            key="csv_utf8_bom",
        )

    st.caption(
        "安全提示：如果单元格文本或列标题以 =、+、-、@、Tab 或回车开头，"
        "用 Excel 打开导出文件时可能被当作公式。默认不修改这些文本。"
    )
    protect_formulas = st.checkbox(
        "启用公式注入保护（在危险文本前添加英文单引号）",
        value=False,
        key="protect_formulas",
        help="这会改变危险文本的外观，但可防止 Excel 将其作为公式执行。",
    )
    if protect_formulas:
        st.info("保护已启用：危险文本会被写成普通文本，下载后内容会多一个前导单引号。")

    try:
        if export_format == "CSV":
            output_data = export_csv_bytes(
                cleaned_data,
                utf8_bom=utf8_bom,
                protect_formulas=protect_formulas,
            )
            extension = "csv"
            mime = "text/csv;charset=utf-8"
        else:
            output_data = export_excel_bytes(
                cleaned_data,
                protect_formulas=protect_formulas,
            )
            extension = "xlsx"
            mime = (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        filename = output_filename(source_name, extension)
    except Exception as exc:
        st.error(f"导出准备失败：{exc}")
        return

    st.download_button(
        f"下载 {filename}",
        data=output_data,
        file_name=filename,
        mime=mime,
        type="primary",
        width="stretch",
    )


def main() -> None:
    show_intro()

    with st.sidebar:
        st.header("本地清洗工具")
        st.caption("无需登录 · 不上传文件 · 不覆盖原文件")
        st.button(
            "重置本次任务",
            on_click=reset_application,
            width="stretch",
            help="清除已导入文件、清洗结果和选项。",
        )
        st.divider()
        st.markdown("**使用建议**")
        st.write("1. 先备份重要原文件。")
        st.write("2. 检查预览和清洗统计。")
        st.write("3. 下载后打开结果复核。")

    sample_col, note_col = st.columns([1, 3])
    with sample_col:
        st.button(
            "加载示例数据",
            on_click=load_sample_data,
            width="stretch",
        )
    with note_col:
        st.caption("没有表格？可先加载示例数据体验完整流程。")

    source = render_import_controls()
    if source is None:
        return

    file_bytes, source_name = source
    file_size = len(file_bytes)
    extension = Path(source_name).suffix.lower()

    if file_size > MAX_FILE_BYTES:
        st.error("文件超过 50 MB。为避免页面长时间无响应，请先拆分文件后重试。")
        return
    if file_size > LARGE_FILE_BYTES:
        st.warning(
            f"文件约 {file_size / 1024 / 1024:.1f} MB，读取和清洗可能需要一些时间。",
            icon="⏳",
        )
    if file_size == 0:
        st.error("文件为空，请选择包含表格内容的文件。")
        return
    if extension not in {".csv", ".xlsx", ".xlsm"}:
        st.error("不支持该文件格式，请选择 CSV、XLSX 或 XLSM 文件。")
        return

    st.caption(f"当前文件：**{source_name}**（{file_size / 1024:.1f} KB）")

    try:
        digest = file_digest(file_bytes)
        if extension == ".csv":
            decoded = decode_csv_bytes(file_bytes)
            detected_delimiter = detect_csv_delimiter(decoded.text)
            delimiter_options = ["自动识别", *CSV_DELIMITERS.keys()]
            default_index = 0
            delimiter_choice = st.selectbox(
                "CSV 分隔符",
                options=delimiter_options,
                index=default_index,
                format_func=lambda value: (
                    value if value == "自动识别" else CSV_DELIMITERS[value]
                ),
                key=f"csv_delimiter_{digest[:12]}",
            )
            if delimiter_choice == "自动识别":
                selected_delimiter = detected_delimiter
            else:
                selected_delimiter = delimiter_choice

            if selected_delimiter is None:
                st.warning(
                    "未能自动识别 CSV 分隔符，请从上方列表手动选择逗号、制表符、"
                    "分号或竖线。",
                    icon="⚠️",
                )
                return

            dataframe = cached_read_csv(
                file_bytes,
                decoded.encoding,
                selected_delimiter,
            )
            encoding_note = decoded.encoding
            delimiter_note = CSV_DELIMITERS[selected_delimiter]
            sheet_name = ""
        else:
            sheet_names = cached_excel_sheets(file_bytes)
            if not sheet_names:
                st.error("Excel 文件中没有可读取的工作表。")
                return
            sheet_name = st.selectbox(
                "选择工作表",
                options=sheet_names,
                key=f"sheet_choice_{digest[:12]}",
            )
            if extension == ".xlsm":
                st.warning(
                    "XLSM 中的宏不会被读取或执行，导出文件也不会保留宏。"
                    "导出格式固定为普通 XLSX；请保留原文件作为宏版本备份。",
                    icon="⚠️",
                )
            dataframe = cached_read_excel(file_bytes, sheet_name)
            encoding_note = ""
            delimiter_note = ""
    except Exception as exc:
        st.error(f"文件读取失败：{exc}")
        st.caption(
            "请检查文件是否损坏、加密、扩展名与实际格式是否一致，"
            "并参阅 README 中的兼容性说明。"
        )
        return

    source_signature = (
        digest,
        extension,
        sheet_name,
        encoding_note,
        delimiter_note,
    )
    if st.session_state.get("source_signature") != source_signature:
        clear_cleaned_result()
        st.session_state["source_signature"] = source_signature

    summary_parts = [
        f"{len(dataframe):,} 行",
        f"{len(dataframe.columns):,} 列",
    ]
    if sheet_name:
        summary_parts.append(f"工作表：{sheet_name}")
    if encoding_note:
        summary_parts.append(f"编码：{encoding_note}")
    if delimiter_note:
        summary_parts.append(f"分隔符：{delimiter_note}")
    st.success("文件读取成功：" + " · ".join(summary_parts))

    render_raw_preview(dataframe)
    st.divider()
    columns = list(dataframe.columns)
    options, submitted = render_cleaning_controls(columns)

    if submitted:
        try:
            with st.spinner("正在清洗数据，请稍候..."):
                cleaned_data, stats = clean_dataframe(dataframe, options)
            st.session_state["cleaned_data"] = cleaned_data
            st.session_state["cleaning_stats"] = stats
            st.session_state["cleaned_options_signature"] = options.signature()
        except Exception as exc:
            st.error(f"清洗失败：{exc}")

    render_result(source_name, options)

    st.divider()
    st.caption(
        "兼容性提醒：复杂公式、日期显示格式、合并单元格、数据验证、图表和样式"
        "可能无法完整保留。CSV 只保存单元格值，不保存公式或 Excel 格式。"
    )


if __name__ == "__main__":
    main()
