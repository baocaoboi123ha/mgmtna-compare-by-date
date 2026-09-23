# -*- coding: utf-8 -*-
"""Public web UI for MGMTNA Compare By Date — multi-version, no login."""

from __future__ import annotations

import importlib
import shutil
import tempfile
from pathlib import Path
from types import ModuleType

import streamlit as st

VERSIONS: dict[str, dict[str, str]] = {
    "7.0": {
        "module": "engine_v7",
        "suffix": " - ByDate 7.0.xlsx",
        "blurb": "Bản gốc: layout 4.0, bảng Unit cơ bản dưới Compare.",
    },
    "8.0": {
        "module": "engine_v8",
        "suffix": " - ByDate 8.0.xlsx",
        "blurb": "Qty bán thực tế, công thức quy đổi, checkqty Base; viền bảng Compare/Unit/TongHop.",
    },
}

XLSX = ["xlsx"]


def _load_engine(version: str) -> ModuleType:
    meta = VERSIONS[version]
    return importlib.import_module(meta["module"])


st.set_page_config(
    page_title="MGMTNA Compare By Date",
    page_icon="📑",
    layout="centered",
)


def _write_upload(uploaded, folder: Path) -> Path:
    dest = folder / Path(uploaded.name).name
    dest.write_bytes(uploaded.getvalue())
    return dest


def _run_compare(data_file, unit_file, version: str):
    engine = _load_engine(version)
    suffix = VERSIONS[version]["suffix"]
    logs: list[str] = []
    work = Path(tempfile.mkdtemp(prefix="mgmtna-"))
    try:
        source = _write_upload(data_file, work)
        unit_path = _write_upload(unit_file, work) if unit_file is not None else None
        output = work / (source.stem + suffix)

        def log(msg: str) -> None:
            logs.append(str(msg))

        result = engine.compare_workbook(
            source,
            output,
            0.5,
            log,
            "v7",
            unit_path,
        )
        payload = Path(result["OutputPath"]).read_bytes()
        name = Path(result["OutputPath"]).name
        summary = (
            f"Xong {source.name}: {result.get('DayCount', 0)} ngày, "
            f"{result.get('TotalLech', 0)} mã lệch "
            f"({result.get('TotalQtyLech', 0)} lệch SL)."
        )
        return payload, name, "\n".join(logs), summary
    finally:
        shutil.rmtree(work, ignore_errors=True)


header_left, header_right = st.columns([4, 1])
with header_left:
    st.title("MGMTNA Compare By Date")
with header_right:
    version = st.selectbox(
        "Phiên bản",
        options=list(VERSIONS.keys()),
        index=list(VERSIONS.keys()).index("8.0"),
        key="engine_version",
        help="Chọn engine xử lý; file tải về ghi hậu tố ByDate tương ứng.",
    )

st.caption(
    f"**Phiên bản {version}** — {VERSIONS[version]['blurb']} "
    "Không đăng nhập. File chỉ xử lý tạm trên máy chủ rồi xóa — không lưu dữ liệu."
)

tab_basic, tab_unit = st.tabs(["So mã", "So mã + Unit"])

with tab_basic:
    st.write(
        "Đối chiếu nguyên thủy theo ngày. Chỉ dùng cột cần thiết; "
        "cột khác thừa hoặc thiếu thì bỏ qua."
    )
    data_basic = st.file_uploader(
        "File dữ liệu gốc (.xlsx)",
        type=XLSX,
        key="basic_data",
    )
    if st.button("Chạy đối chiếu", type="primary", key="run_basic"):
        if data_basic is None:
            st.warning("Hãy chọn file dữ liệu gốc.")
        else:
            engine_mod = _load_engine(version)
            try:
                with st.spinner("Đang đối chiếu… file lớn có thể mất khoảng 1 phút."):
                    payload, name, logs, summary = _run_compare(
                        data_basic, None, version
                    )
                st.success(summary)
                st.download_button(
                    "Tải file kết quả",
                    data=payload,
                    file_name=name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_basic",
                )
                with st.expander("Nhật ký xử lý"):
                    st.code(logs or "(trống)", language="text")
            except engine_mod.CompareError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Lỗi: {exc}")

with tab_unit:
    st.write(
        "Cùng rule so mã, thêm đối chiếu Unit Base. "
        "Phải tải riêng file Unit convert."
    )
    data_unit = st.file_uploader(
        "1. File dữ liệu gốc (.xlsx)",
        type=XLSX,
        key="unit_data",
    )
    unit_file = st.file_uploader(
        "2. File Unit convert (.xlsx)",
        type=XLSX,
        key="unit_file",
    )
    if st.button("Chạy đối chiếu + Unit", type="primary", key="run_unit"):
        if data_unit is None:
            st.warning("Hãy chọn file dữ liệu gốc.")
        elif unit_file is None:
            st.warning("Hãy chọn file Unit convert.")
        else:
            engine_mod = _load_engine(version)
            try:
                with st.spinner("Đang đối chiếu + Unit…"):
                    payload, name, logs, summary = _run_compare(
                        data_unit, unit_file, version
                    )
                st.success(summary)
                st.download_button(
                    "Tải file kết quả",
                    data=payload,
                    file_name=name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_unit",
                )
                with st.expander("Nhật ký xử lý"):
                    st.code(logs or "(trống)", language="text")
            except engine_mod.CompareError as exc:
                st.error(str(exc))
            except Exception as exc:
                st.error(f"Lỗi: {exc}")
