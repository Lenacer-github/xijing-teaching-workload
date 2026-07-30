"""Shared business rules for the Bootstrap web application.

The calculation catalogue is read from the literal ``WORKLOAD_DATA`` in the
existing Streamlit app so both interfaces keep one source of truth while the
new frontend is evaluated in parallel.
"""

from __future__ import annotations

import ast
import copy
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any


BASE_DIR = Path(__file__).resolve().parent
APP_VERSION = "1.0.0"
DEPARTMENT_OPTIONS = [
    "数字商务系",
    "大数据管理应用系",
    "国际贸易系",
    "金融科技系",
]
ACADEMIC_TERM_OPTIONS = [
    "2025-2026学年第一学期",
    "2025-2026学年第二学期",
]
DEFAULT_ACADEMIC_TERM = "2025-2026学年第二学期"
OWNER_EXPORT_ORDER = [
    "学院",
    "王丹",
    "李思仪",
    "史高峰",
    "钟慧娟",
    "惠媛",
    "郭文霞",
    "王秀芹",
]


def _load_workload_catalog() -> dict[str, list[str]]:
    """Load the catalogue without importing and executing the Streamlit app."""
    source = (BASE_DIR / "app.py").read_text(encoding="utf-8")
    module = ast.parse(source)
    for node in module.body:
        if not isinstance(node, ast.Assign):
            continue
        if any(
            isinstance(target, ast.Name) and target.id == "WORKLOAD_DATA"
            for target in node.targets
        ):
            value = ast.literal_eval(node.value)
            if isinstance(value, dict):
                return value
    raise RuntimeError("未能读取教学工作量计算标准清单")


WORKLOAD_DATA = _load_workload_catalog()


def split_standard_owner(standard: str) -> tuple[str, str]:
    standard_text = str(standard or "").strip()
    match = re.search(r"【([^】]+)】\s*$", standard_text)
    if not match:
        return standard_text, ""
    return standard_text[: match.start()].rstrip(), match.group(1).strip()


def get_standard_text(standard: str) -> str:
    return split_standard_owner(standard)[0]


def get_standard_owner(standard: str) -> str:
    return split_standard_owner(standard)[1]


def get_primary_standard_owner(standard: str) -> str:
    owner_text = get_standard_owner(standard)
    positions = [
        (owner_text.find(owner), owner)
        for owner in OWNER_EXPORT_ORDER
        if owner in owner_text
    ]
    return min(positions)[1] if positions else ""


def format_workload_value(value: Any) -> str:
    decimal_value = Decimal(str(value or 0))
    return format(decimal_value, "f").rstrip("0").rstrip(".") or "0"


def make_excel_filename(data: dict[str, Any]) -> str:
    parts = [
        str(data.get("dept", "")).strip(),
        str(data.get("name", "")).strip(),
        str(data.get("title", "")).strip(),
        "教师工作量填报表",
    ]
    return re.sub(r'[\\/:*?"<>|]', "_", "-".join(parts)) + ".xlsx"


def _match_catalog_standard(category: str, standard: str) -> tuple[str, str]:
    """Match current and historical progress files by their visible text."""
    standard_text = get_standard_text(standard)
    if not standard_text:
        return category, ""

    if category in WORKLOAD_DATA:
        for option in WORKLOAD_DATA[category]:
            if option == standard or get_standard_text(option) == standard_text:
                return category, option

    for current_category, options in WORKLOAD_DATA.items():
        for option in options:
            if get_standard_text(option) == standard_text:
                return current_category, option
    return category, ""


def normalize_payload(payload: Any, *, require_complete: bool) -> dict[str, Any]:
    """Normalize an incoming progress/export payload and validate its fields."""
    if not isinstance(payload, dict):
        raise ValueError("数据格式不正确")

    dept = str(payload.get("dept", "") or "").strip()
    if dept == "电子商务":
        dept = "数字商务系"
    if dept not in DEPARTMENT_OPTIONS:
        raise ValueError("请选择有效的所属系部")

    name = str(payload.get("name", "") or "").strip()
    title = str(payload.get("title", "") or "").strip()
    term = str(payload.get("term", DEFAULT_ACADEMIC_TERM) or "").strip()
    if term not in ACADEMIC_TERM_OPTIONS:
        raise ValueError("请选择有效的学期")
    if require_complete and not name:
        raise ValueError("请填写姓名")
    if require_complete and not title:
        raise ValueError("请填写职称")

    rows = payload.get("rows", [])
    if not isinstance(rows, list):
        raise ValueError("工作量明细必须是列表")
    if require_complete and not rows:
        raise ValueError("请至少添加一条教学工作量记录")

    normalized_rows: list[dict[str, Any]] = []
    selected_standards: dict[str, int] = {}
    for index, source_row in enumerate(rows, start=1):
        if not isinstance(source_row, dict):
            raise ValueError(f"记录 {index} 的数据结构不正确")

        category = str(source_row.get("category", "") or "").strip()
        standard = str(source_row.get("standard", "") or "").strip()
        if category and category not in WORKLOAD_DATA:
            raise ValueError(f"记录 {index} 的项目类别不在当前清单中")
        category, standard = _match_catalog_standard(category, standard)

        if require_complete and not category:
            raise ValueError(f"记录 {index} 未选择项目类别")
        if require_complete and not standard:
            raise ValueError(f"记录 {index} 未选择计算标准")
        if standard:
            if standard in selected_standards:
                raise ValueError(
                    f"记录 {index} 与记录 {selected_standards[standard]} "
                    "选择了重复的成果类别"
                )
            selected_standards[standard] = index

        workload_raw = source_row.get("workload", "")
        if workload_raw in ("", None):
            workload: float | str = ""
            if require_complete:
                raise ValueError(f"记录 {index} 未填写工作量")
        else:
            try:
                workload_decimal = Decimal(str(workload_raw))
                if not workload_decimal.is_finite() or workload_decimal < 0:
                    raise InvalidOperation
                workload = float(workload_decimal)
            except (InvalidOperation, ValueError) as exc:
                raise ValueError(
                    f"记录 {index} 的工作量必须是不小于 0 的数字"
                ) from exc

        completion_time = str(source_row.get("time", "") or "").strip()
        if completion_time:
            try:
                datetime.strptime(completion_time, "%Y-%m-%d")
            except ValueError as exc:
                raise ValueError(
                    f"记录 {index} 的实际取得时间格式不正确"
                ) from exc
        elif require_complete:
            raise ValueError(f"记录 {index} 未填写实际取得时间")

        remark = str(source_row.get("remark", "") or "").strip()
        if require_complete and not remark:
            raise ValueError(f"记录 {index} 未填写完成说明")

        normalized_rows.append(
            {
                "term": term,
                "category": category,
                "standard": standard,
                "workload": workload,
                "time": completion_time,
                "remark": remark,
            }
        )

    return {
        "schema_version": 1,
        "dept": dept,
        "name": name,
        "title": title,
        "term": term,
        "fill_date": date.today().isoformat(),
        "rows": normalized_rows,
    }


def prepare_export_data(payload: dict[str, Any]) -> dict[str, Any]:
    data = normalize_payload(payload, require_complete=True)
    export_data = copy.deepcopy(data)
    owner_rank = {
        owner: index for index, owner in enumerate(OWNER_EXPORT_ORDER)
    }
    indexed_rows = list(enumerate(export_data["rows"]))
    indexed_rows.sort(
        key=lambda item: (
            owner_rank.get(
                get_primary_standard_owner(item[1]["standard"]),
                len(OWNER_EXPORT_ORDER),
            ),
            item[0],
        )
    )
    export_data["rows"] = [row for _, row in indexed_rows]
    for row in export_data["rows"]:
        standard = row["standard"]
        row["reviewer"] = get_standard_owner(standard)
        row["standard"] = get_standard_text(standard)
    return export_data


def public_catalog() -> dict[str, list[dict[str, str]]]:
    return {
        category: [
            {
                "value": standard,
                "text": get_standard_text(standard),
                "owner": get_standard_owner(standard),
            }
            for standard in options
        ]
        for category, options in WORKLOAD_DATA.items()
    }
