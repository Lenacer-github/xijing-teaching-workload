import streamlit as st
import streamlit.components.v1 as components
import base64
import html
import json
import logging
import re
import time
from decimal import Decimal, InvalidOperation
from datetime import date, datetime
from pathlib import Path

from excel_export import build_workload_excel

APP_VERSION = "1.0.0"
RUN_STARTED_AT = time.perf_counter()
PERFORMANCE_LOGGER = logging.getLogger("workload_app.performance")

# ================= 1. 完整基础数据配置 =================
WORKLOAD_DATA = {
    "一、特殊贡献": [
        "1.取得教学成果，支撑学院目标责任书中基础工作目标、重点工作目标、发展工作目标的考核，分别计10、20、30学时/项。（不重复计算）【学院】",
    ],
    "二、专业建设": [
        "1.完成本系专业建设及教学工作任务【王丹】",
        "2.主办或承办应用型人才培养、专业建设、教学改革等研讨会20学时/次（需教务处认定）；【李思仪】",
        "3.担任微专业专业负责人，学生10人以下20学时/学期，10人以上30学时/学期（需考核）；【李思仪】",
        "4.开展专业人才培养质量评价指标数据试点采集工作，30学时/专业；【王丹】",
        "5.获评省级优秀产学合作协同育人项目，100学时/项，获批校级产学合作协同育人项目立项，50学时/项。【史高峰】",
    ],
    "三、教研室建设": [
        "1.立项省级虚拟教研室，100学时/项；【李思仪】",
        "2.校级虚拟教研室验收通过，60学时/项；【李思仪】",
        "3.担任学院平台课程负责人，完成课程安排、课程评价等工作，20学时/学年。【王丹】",
        "4.教研室（团队）主任（负责人），20学时/学年【王丹】",
    ],
    "四、教学改革": [
        "1.立项省级课程思政示范课，100学时/项；【李思仪】",
        "2.立项省级教学改革研究重点项目，100学时/项；【李思仪】",
        "3.获省级优秀教材一等奖100学时/项，二等奖80学时/项，三等奖60学时/项；【李思仪】",
        "4.教学成果奖，20学时/项（按时申报并被学校成功受理）【李思仪】",
        "5.教学改革研究项目，15学时/项（按时申报并被学校成功受理）【李思仪】",
        "6.出版数字化教材100学时/本（需教务处、研究生认定）；【李思仪】",
        "7.立项知识图谱课程、产教融合课程、AI赋能课程，50学时/项；【李思仪】",
        "8.一流课程、课程思政示范课、一师一优课结题验收认定校级“优秀”，50学时/项；结题验收认定校级“良好”，10学时/项；【李思仪】",
        "9.课程建设（重点课程、精品课程、“课程思政”示范课(团队)、一流课程等），10学时/项（按时申报并被学校成功受理）【李思仪】",
        "10.教材建设，8学时/部（校级自编教材项目申报或各类优秀教材按时申报并被学校成功受理）【李思仪】",
        "11.立项校企共建课程，提供配套资料，并开展共建教学活动，20学时/项；【李思仪】",
        "12.立项校级自编教材建设项目，50学时/项；【李思仪】",
        "13.校级及以上课程思政、产教融合等优秀教学案例获(批)奖，30学时/项；【李思仪】",
        "14.参加学校认定的教师A/B类竞赛以及研究生处组织的教师竞赛，获校级奖项50学时/项；获省级奖项计100学时/项。【李思仪】",
        "15.发表教研论文或指导学生专利授权、发表期刊论文，5学时/件（篇）【李思仪】",
    ],
    "五、实践教学": [
        "1.毕业设计（论文），10学时/生（指导学生不超过6人）【史高峰】",
        "2.完成产业学院建设工作，通过学校考核80学时/项；获批省级及以上产业学院并挂牌，150学时/项；【史高峰】",
        "3.新增校企共建实验室1个，明确建设内容并开展实践教学活动，认定30学时/项；【史高峰】",
        "4.校外实践教学基地承担毕业实习占比高于30%或各毕业实习集中率高于30%或实习基地利用率高于40%，认定40学时/专业；【史高峰】",
        "5.新增校外实习基地并开展实践教学活动（安排毕业实习和认知实习，提供认定相关材料），认定20学时/项（各专业限1项）；【史高峰】",
        "6.新增研究生实践基地，并有研究生在实践基地实习，提供研究生实习相关证明材料，认定20学时/个；【王秀芹】",
        "7.担任产业学院定向班/订单班班主任，学生10人以下15学时/学期，10人以上30学时/学期（需考核）；【史高峰】",
        "8.承担开放实验室项目，10学时/项。【史高峰】",
        "9.研究生校外实践，10学时/生（校外实践导师符合要求，实践资料齐全）【王秀芹】",
    ],
    "六、课堂教学": [
        "1.课程考核采取多元化考核，5学时/门；【王丹】",
        "2.学校教学文件、督导评价表扬，10学时/次；【王丹】",
        "3.新开课程编写大纲，2学时/门【李思仪】",
        "4.非导师给研究生课程编写大纲，5学时/门；【王秀芹】",
        "5.承担复学的专科生教学任务，5学时/门；【王丹】",
        "6.承担复学已停招的本科专业学生的教学任务，5学时/门。【王丹】",
        "7.校内考试监考，1学时/场（每位教师不少于8场/学年；英语四六级、研究生入学考试及各类社会考试不在其内）【王丹】",
        "8.校内考试试卷命题，4学时/套（无差错命题并按时提交，包括期末以及各种重修、补考、转专业考试等试卷命题）【王丹】",
        "9.课程达成度分析或课程评价报告，2学时/份【王丹】",
    ],
    "七、招生与就业": [
        "1.高质量完成招生工作，30学时/专业；【钟慧娟】",
        "2.参加各地招生咨询会，5学时/人·场；【钟慧娟】",
        "3.就业率达到90%及以上，20学时/专业；【惠媛】",
        "4.被学校遴选参与暑假留校招生咨询，10学时/人；【钟慧娟】",
        "5.推荐招聘单位（至少招聘我院1个毕业生），2个学时/企业；【惠媛】",
        "6.推荐毕业生就业，3学时/生。【惠媛】",
    ],
    "八、公共性事务服务": [
        "1.参与学院内部各类教科研项目评审、竞赛项目评审、职称评审、人才引进面试、第二课堂活动评委(活动指导)等，3-20学时/次；【科研：郭文霞；教研：李思仪；竞赛：史高峰；职称评审与人才引进：钟慧娟；第二课堂：惠媛；教材审核与基本功大赛：王丹】",
        "2.承担学院人才引进任务，引进符合要求教师，10-20学时/名；【钟慧娟】",
        "3.撰写高质量专业建设/人才培养宣传稿件被学院采纳发布1学时/篇，被学校基层动态采纳发布2学时/篇，被学校新闻动态采纳发布5学时/篇，被陕西省教育厅采纳发布10学时/篇。【钟慧娟】",
        "4.作为主讲人开展报告，20-30学时/次。【郭文霞】",
        "5.被学校聘为班主任，20学时/学年（需考核），被学校评为优秀班主任5学时/学年。【惠媛】",
        "6.高级职称教师指导青年教师成效良好，15学时/人·学期；【钟慧娟】",
        "7.被学校遴选参与企业挂职锻炼10学时/人（需考核）。【钟慧娟】",
    ],
    "九、创新创业与第二课堂": [
        "1.担任校级学科竞赛项目负责人，20学时/项·学年（需考核）；【史高峰】",
        "2.A类重点赛事省级及以上奖项获奖数量较本学院上学年增长10个百分点，获奖等级较本学院往届有突破，10学时/项；【史高峰】",
        "3.立项国家级大创项目，20学时/项；【史高峰】",
        "4.负责在学校备案的第二课堂社团，10学时/学期；【惠媛】",
        "5.组织并带领学生开展第二课堂活动，2-4学时/次（原则上5人以上）。【惠媛】",
        "6.学生学科竞赛指导参与（不包括“大创”），重点赛事20学时/项（参赛学生50人以下）、50学时/项（参赛学生50人以上）；非重点赛事15学时/项（参赛学生50人以下）、30学时/项（参赛学生50人以上）；中国研究生创新实践系列大赛5学时/队【史高峰】",
    ],
    "十、特色人才培养项目": [
        "1.成立“数字乡村实践与服务平台”人才培养平台，吸纳学生不少于20人，组建电商助农实践团队，全年开展数字乡村专题实践活动不少于10场，服务乡村品牌或合作社不少于5个，形成实践调研报告或成果案例不少于3份。需提交材料，完成认定30学时。【学院】",
    ]
}

BASE_DIR = Path(__file__).resolve().parent
LOGO_PATH = BASE_DIR / "assets" / "college_logo.png"
FAVICON_PATH = BASE_DIR / "assets" / "favicon.png"
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

@st.cache_data(show_spinner=False)
def load_image_data_uri(image_path):
    """缓存静态图片编码，避免每次 Streamlit 重跑都重复读取和转换。"""
    path = Path(image_path)
    if not path.exists():
        return ""
    return (
        "data:image/png;base64,"
        + base64.b64encode(path.read_bytes()).decode("ascii")
    )


LOGO_DATA_URI = load_image_data_uri(str(LOGO_PATH))


def generate_excel_bytes(app_data, export_date):
    """按需生成 Excel；不缓存不同用户的文件，避免云端内存持续增长。"""
    export_source = {
        **app_data,
        "fill_date": export_date,
    }
    return build_workload_excel(prepare_export_data(export_source))


def make_excel_filename(app_data):
    """生成跨平台安全的 Excel 文件名。"""
    parts = [
        str(app_data.get("dept", "")).strip(),
        str(app_data.get("name", "")).strip(),
        str(app_data.get("title", "")).strip(),
        "教师工作量填报表",
    ]
    filename = "-".join(parts)
    return re.sub(r'[\\/:*?"<>|]', "_", filename) + ".xlsx"


def format_workload_value(value):
    """整数不补零，小数保留用户需要的有效位。"""
    decimal_value = Decimal(str(value or 0))
    return format(decimal_value, "f").rstrip("0").rstrip(".") or "0"


def split_standard_owner(standard):
    """拆分计算标准正文及末尾【负责人】。"""
    standard_text = str(standard or "").strip()
    match = re.search(r"【([^】]+)】\s*$", standard_text)
    if not match:
        return standard_text, ""
    return standard_text[:match.start()].rstrip(), match.group(1).strip()


def get_standard_text(standard):
    """返回不包含负责人标注的正式计算标准。"""
    return split_standard_owner(standard)[0]


def get_standard_owner(standard):
    """返回计算标准末尾标注的负责人。"""
    return split_standard_owner(standard)[1]


def get_primary_standard_owner(standard):
    """多人分工时取负责人标注中最先出现的姓名用于导出归组。"""
    owner_text = get_standard_owner(standard)
    positions = [
        (owner_text.find(owner), owner)
        for owner in OWNER_EXPORT_ORDER
        if owner in owner_text
    ]
    return min(positions)[1] if positions else ""


def prepare_export_data(app_data):
    """复制并按负责人顺序整理导出数据，不改变页面中的填写顺序。"""
    export_data = json.loads(json.dumps(app_data, ensure_ascii=False))
    owner_rank = {
        owner: index
        for index, owner in enumerate(OWNER_EXPORT_ORDER)
    }
    indexed_rows = list(enumerate(export_data.get("rows", [])))
    indexed_rows.sort(
        key=lambda item: (
            owner_rank.get(
                get_primary_standard_owner(item[1].get("standard", "")),
                len(OWNER_EXPORT_ORDER),
            ),
            item[0],
        )
    )
    export_data["rows"] = [row for _, row in indexed_rows]
    export_term = export_data.get("term", DEFAULT_ACADEMIC_TERM)
    for row in export_data["rows"]:
        row["term"] = export_term
        standard = row.get("standard", "")
        row["reviewer"] = get_standard_owner(standard)
        row["standard"] = get_standard_text(standard)
    return export_data


def normalize_restored_data(data):
    """校验并规范化暂存文件，兼容未携带版本号的历史文件。"""
    if not isinstance(data, dict):
        raise ValueError("文件内容必须是一个完整的数据对象")

    schema_version = data.get("schema_version", 1)
    if schema_version != 1:
        raise ValueError(f"暂不支持版本 {schema_version} 的暂存文件")

    dept = data.get("dept", DEPARTMENT_OPTIONS[0])
    if dept == "电子商务":
        dept = "数字商务系"
    if dept not in DEPARTMENT_OPTIONS:
        raise ValueError("所属系部不在当前系统选项中")

    name = data.get("name", "")
    title = data.get("title", "")
    if not isinstance(name, str) or not isinstance(title, str):
        raise ValueError("姓名和职称必须是文本")

    fill_date = data.get("fill_date", date.today().isoformat())
    try:
        datetime.strptime(str(fill_date), "%Y-%m-%d")
    except ValueError as exc:
        raise ValueError("填报日期格式应为 YYYY-MM-DD") from exc

    rows = data.get("rows", [])
    if not isinstance(rows, list):
        raise ValueError("工作量明细必须是列表")

    term = data.get(
        "term",
        rows[0].get("term", DEFAULT_ACADEMIC_TERM)
        if rows and isinstance(rows[0], dict)
        else DEFAULT_ACADEMIC_TERM,
    )
    if term not in ACADEMIC_TERM_OPTIONS:
        raise ValueError("学期不在当前系统选项中")

    normalized_rows = []
    for index, row in enumerate(rows, start=1):
        if not isinstance(row, dict):
            raise ValueError(f"记录 {index} 的数据结构不正确")

        term = row.get("term", DEFAULT_ACADEMIC_TERM)
        if term not in ACADEMIC_TERM_OPTIONS:
            raise ValueError(f"记录 {index} 的学期不在当前系统选项中")

        category = row.get("category", "")
        standard = row.get("standard", "")
        standard_text = get_standard_text(standard)
        matched_standard = None
        if not standard_text:
            if category and category not in WORKLOAD_DATA:
                raise ValueError(f"记录 {index} 的项目类别不在当前清单中")
        elif category in WORKLOAD_DATA:
            matched_standard = next(
                (
                    option
                    for option in WORKLOAD_DATA[category]
                    if get_standard_text(option) == standard_text
                ),
                None,
            )
        if standard_text and matched_standard is None:
            for current_category, options in WORKLOAD_DATA.items():
                matched_standard = next(
                    (
                        option
                        for option in options
                        if get_standard_text(option) == standard_text
                    ),
                    None,
                )
                if matched_standard is not None:
                    category = current_category
                    break
        if standard_text and matched_standard is None:
            raise ValueError(f"记录 {index} 的计算标准不在当前清单中")
        standard = matched_standard or ""

        workload_raw = row.get("workload", "")
        if workload_raw in ("", None):
            workload = ""
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

        completion_time = str(row.get("time", "") or "")
        if completion_time:
            try:
                datetime.strptime(completion_time, "%Y-%m-%d")
            except ValueError as exc:
                raise ValueError(f"记录 {index} 的完成时间格式应为 YYYY-MM-DD") from exc

        remark = row.get("remark", "")
        if not isinstance(remark, str):
            raise ValueError(f"记录 {index} 的完成说明必须是文本")

        normalized_rows.append({
            "term": term,
            "category": category,
            "standard": standard,
            "workload": workload,
            "time": completion_time,
            "remark": remark,
        })

    return {
        "dept": dept,
        "name": name.strip(),
        "title": title.strip(),
        "term": term,
        "fill_date": str(fill_date),
        "rows": normalized_rows,
    }


def clear_record_widget_state():
    """记录增删或整体恢复后，清理按序号生成的旧组件状态。"""
    prefixes = (
        "term_", "cat_", "std_", "wl_", "time_", "rmk_", "del_",
        "confirm_delete_", "cancel_delete_"
    )
    for key in list(st.session_state.keys()):
        if key.startswith(prefixes):
            del st.session_state[key]


def delete_record(index):
    """直接删除指定记录，并清理按序号绑定的组件状态。"""
    rows = st.session_state.app_data["rows"]
    if 0 <= index < len(rows):
        rows.pop(index)
    st.session_state.pop("new_record_target", None)
    clear_record_widget_state()


def add_empty_record():
    """新增一条待用户选择成果类别和计算标准的记录。"""
    st.session_state.app_data["rows"].append({
        "term": st.session_state.app_data.get(
            "term", DEFAULT_ACADEMIC_TERM
        ),
        "category": "",
        "standard": "",
        "workload": "",
        "time": "",
        "remark": "",
    })
    st.session_state.new_record_target = len(
        st.session_state.app_data["rows"]
    )


def format_standard_option(standard):
    """计算标准直接显示正式正文，不附加展示前缀。"""
    return get_standard_text(standard)


def standard_display_map(options):
    """返回计算标准展示文本到原始业务值的映射。"""
    return {format_standard_option(option): option for option in options}


# 基础清单在运行期间不会变化，模块加载时一次性生成展示映射。
WORKLOAD_CATEGORY_OPTIONS = tuple(WORKLOAD_DATA)
STANDARD_DISPLAY_MAPS = {
    category: standard_display_map(options)
    for category, options in WORKLOAD_DATA.items()
}
STANDARD_DISPLAY_OPTIONS = {
    category: tuple(display_map)
    for category, display_map in STANDARD_DISPLAY_MAPS.items()
}


def get_effective_workload_text(index, row):
    """读取当前组件中的工作量文本，组件尚未创建时回退到数据值。"""
    widget_key = f"wl_{index}"
    if widget_key in st.session_state:
        return str(st.session_state[widget_key]).strip()
    value = row.get("workload", "")
    return "" if value in (None, "") else format_workload_value(value)


def collect_validation_state(app_data):
    """统一生成导出拦截、即时校验、完成状态所需的数据。"""
    rows = app_data["rows"]
    missing_basic_fields = [
        label
        for label, value in (
            ("所属系部", app_data.get("dept", "")),
            ("姓名", app_data.get("name", "")),
            ("职称", app_data.get("title", "")),
            ("学期", app_data.get("term", "")),
        )
        if not str(value).strip()
    ]

    standard_rows = {}
    for index, row in enumerate(rows, start=1):
        standard_rows.setdefault(row.get("standard", ""), []).append(index)
    duplicate_row_numbers = {
        row_number
        for standard, row_numbers in standard_rows.items()
        if standard and len(row_numbers) > 1
        for row_number in row_numbers
    }
    duplicate_groups = [
        row_numbers
        for standard, row_numbers in standard_rows.items()
        if standard and len(row_numbers) > 1
    ]

    row_errors = {}
    missing_category_rows = []
    missing_standard_rows = []
    missing_workload_rows = []
    invalid_workload_rows = []
    missing_time_rows = []
    missing_remark_rows = []
    for index, row in enumerate(rows, start=1):
        errors = []
        if not str(row.get("category", "")).strip():
            missing_category_rows.append(index)
            errors.append("请选择项目类别")
        if not str(row.get("standard", "")).strip():
            missing_standard_rows.append(index)
            errors.append("请选择计算标准")
        workload_text = get_effective_workload_text(index - 1, row)
        if not workload_text:
            missing_workload_rows.append(index)
            errors.append("请填写工作量")
        else:
            try:
                workload = Decimal(workload_text)
                if not workload.is_finite() or workload < 0:
                    raise InvalidOperation
            except (InvalidOperation, ValueError):
                invalid_workload_rows.append(index)
                errors.append("请输入不小于 0 的数字")

        if not str(row.get("time", "")).strip():
            missing_time_rows.append(index)
            errors.append("请选择完成时间")
        if not str(row.get("remark", "")).strip():
            missing_remark_rows.append(index)
            errors.append("请填写具体成果或业绩内容")
        if index in duplicate_row_numbers:
            conflicting_rows = next(
                (
                    group
                    for group in duplicate_groups
                    if index in group
                ),
                [],
            )
            other_row = next(
                (row_number for row_number in conflicting_rows if row_number != index),
                None,
            )
            if other_row is not None:
                errors.append(f"该成果类别与记录{other_row}重复")
        row_errors[index] = errors

    issues = []
    if missing_basic_fields:
        issues.append(f"基本信息缺少：{'、'.join(missing_basic_fields)}")
    if not rows:
        issues.append("尚未新增教学工作量明细")
    if missing_category_rows:
        issues.append(
            f"记录 {'、'.join(map(str, missing_category_rows))} 未选择项目类别"
        )
    if missing_standard_rows:
        issues.append(
            f"记录 {'、'.join(map(str, missing_standard_rows))} 未选择计算标准"
        )
    if missing_workload_rows:
        issues.append(
            f"记录 {'、'.join(map(str, missing_workload_rows))} 未填写工作量"
        )
    if invalid_workload_rows:
        issues.append(
            f"记录 {'、'.join(map(str, invalid_workload_rows))} 的工作量格式不正确"
        )
    if duplicate_groups:
        duplicate_text = "；".join(
            f"记录 {'、'.join(map(str, group))}"
            for group in duplicate_groups
        )
        issues.append(f"{duplicate_text} 选择了重复的成果类别")
    if missing_time_rows:
        issues.append(
            f"记录 {'、'.join(map(str, missing_time_rows))} 未填写完成时间"
        )
    if missing_remark_rows:
        issues.append(
            f"记录 {'、'.join(map(str, missing_remark_rows))} 未填写完成说明"
        )

    completed_count = sum(not errors for errors in row_errors.values())
    first_error_target = None
    if missing_basic_fields:
        first_error_target = "basic-info"
    elif rows:
        first_invalid_row = next(
            (index for index, errors in row_errors.items() if errors),
            None,
        )
        if first_invalid_row is not None:
            first_error_target = f"record-{first_invalid_row}"
    elif not rows:
        first_error_target = "workload-details"

    return {
        "issues": issues,
        "row_errors": row_errors,
        "completed_count": completed_count,
        "first_error_target": first_error_target,
    }


def request_export_validation():
    """请求完整校验，并在下一次渲染时定位第一处错误。"""
    st.session_state.validation_requested = True
    st.session_state.scroll_to_first_error = True


def request_excel_export():
    """仅在用户明确点击导出时生成 Excel，避免录入期间重复计算。"""
    st.session_state.pending_excel_export = True


def set_action_notice(message):
    """记录下载类操作的成功反馈。"""
    st.session_state.action_notice = message


# ================= 2. 页面与UI初始化及深度CSS汉化 =================
st.set_page_config(
    page_title="科技商学院教师教学工作量填报系统",
    page_icon=str(FAVICON_PATH),
    layout="wide"
)

st.markdown("""
    <style>
    :root {
        --brand: #0B3D91;
        --brand-strong: #072E72;
        --brand-soft: #EAF2FF;
        --ink: #172033;
        --muted: #64748B;
        --canvas: #F4F7FB;
        --border: #DCE5F0;
        --success: #15803D;
        --warning: #D97706;
        --danger: #DC2626;
        --shadow-sm: 0 7px 22px rgba(15, 48, 92, .065);
        --shadow-md: 0 18px 44px rgba(15, 48, 92, .105);
    }
    html { scroll-behavior: smooth; }
    .stApp,
    .stApp *,
    .stApp *::before,
    .stApp *::after {
        box-sizing: border-box;
    }
    body, .stApp {
        color: var(--ink);
        font-family: "Inter", "PingFang SC", "Microsoft YaHei", sans-serif;
    }
    .stApp {
        overflow-x: clip;
        background:
            radial-gradient(circle at 7% 0%, rgba(191,219,254,.42), transparent 29rem),
            linear-gradient(180deg, #F9FBFE 0%, var(--canvas) 48%, #F8FAFC 100%);
    }
    .block-container {
        width: min(1440px, calc(100% - 3.5rem)) !important;
        max-width: 1440px;
        margin-right: auto !important;
        margin-left: auto !important;
        padding: 1.1rem 0 2.4rem !important;
    }
    .block-container > [data-testid="stVerticalBlock"] {
        width: 100%;
        max-width: 100%;
        min-width: 0;
    }
    .app-hero,
    [data-testid="stExpander"],
    [data-testid="stVerticalBlockBorderWrapper"],
    [data-testid="stHorizontalBlock"] {
        width: auto !important;
        max-width: 100%;
        min-width: 0;
    }
    /*
     * Streamlit 的带边框容器会在内容块外再生成一层带内边距的 flex 包装。
     * 内容块默认 flex-grow 后会按包装层的边框宽度计算，导致把左右内边距
     * 再加到总宽度上。固定其基准为包装层的内容宽度，避免卡片向右溢出。
     */
    [data-testid="stVerticalBlockBorderWrapper"]
    > div:not([data-testid])
    > [data-testid="stVerticalBlock"] {
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
        flex: 0 1 100% !important;
    }
    [data-testid="stVerticalBlock"] > [data-testid="element-container"] {
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
    }
    [data-testid="element-container"] > [data-testid="stMarkdown"],
    [data-testid="stMarkdown"] > [data-testid="stMarkdownContainer"] {
        width: 100% !important;
        max-width: 100% !important;
        min-width: 0 !important;
    }

    /* 顶部品牌区 */
    .app-hero {
        position: relative;
        display: grid;
        grid-template-columns: minmax(320px, 1.15fr) minmax(320px, .85fr);
        align-items: center;
        gap: 1.5rem;
        min-height: 126px;
        margin-bottom: 1rem;
        padding: .85rem 1.5rem;
        overflow: hidden;
        border: 1px solid rgba(180,204,238,.78) !important;
        border-radius: 20px !important;
        background: linear-gradient(112deg, #FFFFFF 0%, #F8FBFF 58%, #E9F2FF 100%) !important;
        box-shadow: var(--shadow-md) !important;
    }
    .app-hero::after {
        width: 22rem;
        height: 22rem;
        right: -7rem;
        top: -14rem;
        background: radial-gradient(circle, rgba(37,99,235,.22), rgba(59,130,246,.025) 62%);
    }
    .brand-logo {
        position: relative;
        z-index: 1;
        display: block;
        width: min(100%, 520px);
        max-height: 78px;
        object-fit: contain;
        object-position: left center;
    }
    .hero-copy {
        position: relative;
        z-index: 1;
        padding-left: 1.5rem;
        border-left: 1px solid rgba(148,163,184,.34);
    }
    .app-hero .college-name {
        color: #3B6AAA;
        font-size: .78rem;
        letter-spacing: .16em;
    }
    .app-hero .system-name {
        margin: .22rem 0 .3rem;
        color: var(--brand-strong);
        font-size: clamp(1.65rem, 2.25vw, 2.2rem);
        line-height: 1.12;
        letter-spacing: -.02em;
    }
    .app-hero .system-subtitle {
        color: var(--muted);
        font-size: .92rem;
        letter-spacing: .08em;
    }

    /* 模块标题与卡片 */
    .section-heading {
        display: flex;
        align-items: center;
        gap: .8rem;
        margin: 0 0 1.05rem;
    }
    .section-icon {
        display: grid;
        place-items: center;
        width: 2.3rem;
        height: 2.3rem;
        flex: 0 0 2.3rem;
        border-radius: 10px;
        color: #FFFFFF;
        background: linear-gradient(145deg, #2563EB, var(--brand));
        box-shadow: 0 7px 16px rgba(37,99,235,.22);
    }
    .section-heading h2 {
        margin: 0;
        padding: 0 !important;
        color: var(--ink);
        font-size: 1.22rem;
        font-weight: 760;
        line-height: 1.2;
    }
    [data-testid="element-container"]:has(.section-heading) {
        min-height: 2.3rem !important;
        height: auto !important;
        flex: 0 0 auto !important;
    }
    .section-heading p {
        margin: .2rem 0 0;
        color: var(--muted);
        font-size: .81rem;
    }
    /* 模块标题仅作信息展示，不显示锚点链接或悬停交互 */
    .section-heading,
    .section-heading * {
        cursor: default !important;
        text-decoration: none !important;
    }
    .section-heading h2 a,
    .section-heading [data-testid="stHeaderActionElements"] {
        display: none !important;
        pointer-events: none !important;
    }
    .basic-section-marker,
    .details-section-marker,
    .operations-section-marker,
    .operation-primary-marker,
    .operation-secondary-marker,
    .operation-progress-marker,
    .record-card-marker {
        height: 0;
        overflow: hidden;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .basic-section-marker),
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .details-section-marker) {
        margin-top: 1rem;
        border: 1px solid var(--border) !important;
        border-radius: 18px !important;
        background: rgba(255,255,255,.97);
        box-shadow: var(--shadow-sm);
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .basic-section-marker) > div,
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .details-section-marker) > div {
        padding: .75rem 1.4rem 1rem;
    }
    [data-testid="element-container"]:has(.basic-section-marker),
    [data-testid="element-container"]:has(.details-section-marker) {
        display: none !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .details-section-marker)
    [data-testid="stHorizontalBlock"]:has(.section-heading) {
        align-items: center !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .details-section-marker)
    [data-testid="stHorizontalBlock"]:has(.section-heading)
    > [data-testid="column"] {
        align-self: center !important;
    }

    /* 数据管理：固定展开的紧凑型静态模块 */
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker) {
        margin-top: 1rem;
        border: 1px solid var(--border) !important;
        border-radius: 18px !important;
        background: rgba(255,255,255,.97);
        box-shadow: var(--shadow-sm);
        overflow: hidden;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker) > div {
        padding: .65rem .85rem .75rem;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    > div:not([data-testid])
    > [data-testid="stVerticalBlock"] {
        gap: .55rem !important;
    }
    [data-testid="element-container"]:has(.operations-section-marker) {
        display: none !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    .section-heading {
        margin-bottom: 0;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"] {
        display: flex;
        align-items: center;
        padding: .42rem .58rem;
        border: 1px solid #E8EEF6;
        border-radius: 13px;
        background: #FBFDFF;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]
    > [data-testid="stVerticalBlockBorderWrapper"] {
        width: 100% !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-primary-marker) {
        border-color: #BFD3EF;
        background: linear-gradient(145deg, #F2F7FF, #F8FBFF);
        box-shadow: inset 0 0 0 1px rgba(37,99,235,.035);
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker) {
        border-color: #D7E1EC;
        border-top: 3px solid #91A6BD;
        background: linear-gradient(180deg, #F7F9FC, #FBFCFE);
        box-shadow: none;
    }
    .operation-pair-title {
        display: flex;
        align-items: center;
        justify-content: flex-start;
        gap: .45rem;
        min-height: 1.35rem;
        padding: 0 .25rem;
        color: #40536A;
        font-size: .9rem;
        font-weight: 760;
        line-height: 1.2;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    > [data-testid="stVerticalBlockBorderWrapper"]
    [data-testid="stVerticalBlock"] {
        gap: .35rem !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    [data-testid="stHorizontalBlock"] {
        gap: .65rem !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    [data-testid="stHorizontalBlock"] > [data-testid="column"] {
        padding: 0 !important;
        border: 0 !important;
        border-radius: 0;
        background: transparent !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    [data-testid="stFileDropzoneInstructions"],
    [data-testid="stFileUploaderDropzoneInstructions"] {
        display: none !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    [data-testid="stFileUploadDropzone"],
    [data-testid="stFileUploaderDropzone"] {
        min-height: 38px;
        height: 38px;
        padding: 0 !important;
        border: 0 !important;
        background: transparent !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    [data-testid="stFileUploadDropzone"] > button,
    [data-testid="stFileUploaderDropzone"] > button {
        width: 100% !important;
        min-width: 0 !important;
        min-height: 38px !important;
        height: 38px !important;
        font-size: .9rem !important;
        font-weight: 680 !important;
        line-height: 1.25 !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    .stDownloadButton > button {
        min-height: 38px !important;
        height: 38px !important;
        font-size: .9rem !important;
        font-weight: 680 !important;
        line-height: 1.25 !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    .stDownloadButton > button p {
        margin: 0 !important;
        font-size: .9rem !important;
        font-weight: 680 !important;
        line-height: 1.25 !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-primary-marker)
    button {
        min-height: 38px !important;
        height: 38px !important;
    }
    .operation-help {
        width: calc(100% - .5rem);
        margin: 0 .25rem;
        box-sizing: border-box;
        color: #718096;
        text-align: left;
        font-size: .76rem;
        line-height: 1.45;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    [data-testid="element-container"]:has(.operation-help) {
        min-height: 1.15rem !important;
        height: auto !important;
        flex: 0 0 auto !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]
    > [data-testid="stVerticalBlockBorderWrapper"]
    > div:not([data-testid])
    > [data-testid="stVerticalBlock"] {
        min-height: 56px;
        justify-content: center;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    > [data-testid="stVerticalBlockBorderWrapper"]
    > div:not([data-testid])
    > [data-testid="stVerticalBlock"] {
        min-height: 86px !important;
        flex-shrink: 0 !important;
        justify-content: flex-start !important;
    }

    /* 表单 */
    [data-testid="stWidgetLabel"] p {
        color: #334155;
        font-size: .87rem;
        font-weight: 650;
    }
    div.stTextInput input,
    div.stTextArea textarea,
    div.stDateInput input,
    div.stNumberInput input {
        color: var(--ink) !important;
        border: 1px solid var(--border) !important;
        border-radius: 10px !important;
        background: #FFFFFF !important;
        box-shadow: 0 1px 2px rgba(15,23,42,.025);
        transition: border-color .16s ease, box-shadow .16s ease;
    }
    div.stTextInput input:focus,
    div.stTextArea textarea:focus,
    div.stDateInput input:focus {
        outline: none !important;
        border-color: #A8B7C8 !important;
        box-shadow: none !important;
    }
    [data-testid="stTextInput"] [data-baseweb="input"]:focus-within,
    [data-testid="stDateInput"] [data-baseweb="input"]:focus-within,
    [data-testid="stTextArea"] [data-baseweb="textarea"]:focus-within,
    [data-testid="stSelectbox"] [data-baseweb="select"] > div:focus-within {
        border-color: #A8B7C8 !important;
        box-shadow: 0 0 0 2px rgba(100,116,139,.08) !important;
    }
    div[data-baseweb="select"] > div {
        min-height: 42px !important;
        height: 42px !important;
        max-height: 42px !important;
        border: 1px solid var(--border) !important;
        border-radius: 10px !important;
        background: #FFFFFF !important;
        box-shadow: 0 1px 2px rgba(15,23,42,.025);
        overflow: hidden !important;
    }
    [data-testid="stSelectbox"] [data-baseweb="select"] div[value] {
        min-width: 0 !important;
        height: 40px !important;
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        display: flex !important;
        align-items: center !important;
        color: #1E293B !important;
        line-height: 1.2 !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    [data-testid="stSelectbox"] [data-baseweb="select"] div[value] > div {
        min-width: 0 !important;
        height: 100% !important;
        display: flex !important;
        align-items: center !important;
        line-height: 1.2 !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    /* 基本信息区四个控件保持完全一致的高度 */
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .basic-section-marker)
    [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .basic-section-marker)
    [data-baseweb="input"] > div {
        height: 42px !important;
        min-height: 42px !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .basic-section-marker)
    [data-testid="stSelectbox"] [data-baseweb="select"] div[value] {
        padding-top: 0 !important;
        padding-bottom: 0 !important;
        white-space: nowrap !important;
        line-height: 1.2 !important;
    }
    [role="option"] {
        min-height: 40px !important;
        line-height: 1.5 !important;
    }

    /* 主要按钮 */
    .stButton > button,
    .stDownloadButton > button {
        min-height: 2.65rem;
        border-radius: 10px;
        font-weight: 680;
        transition: transform .16s ease, box-shadow .16s ease, border-color .16s ease;
    }
    .stButton > button:hover,
    .stDownloadButton > button:hover {
        transform: translateY(-1px);
        border-color: #7EA5DD;
        box-shadow: 0 8px 20px rgba(37,99,235,.13);
    }
    .stButton > button[kind="primary"] {
        border-color: var(--brand) !important;
        background: linear-gradient(135deg, #1D5FBF, var(--brand)) !important;
        box-shadow: 0 8px 18px rgba(11,61,145,.2);
    }

    /* 单条工作量卡片 */
    [data-testid="stVerticalBlockBorderWrapper"]:has(.record-card-marker):not(:has(.details-section-marker)) {
        margin: .4rem 0 .6rem;
        border: 1px solid #D9E4F2 !important;
        border-radius: 15px !important;
        background: #FFFFFF;
        box-shadow: 0 7px 20px rgba(30,58,91,.055);
        transition: border-color .18s ease, box-shadow .18s ease;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(.record-card-marker):not(:has(.details-section-marker)):hover {
        border-color: #B9CDE8 !important;
        box-shadow: 0 12px 28px rgba(30,58,91,.085);
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(.record-card-marker):not(:has(.details-section-marker)) > div {
        padding: .55rem .75rem .65rem;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(.record-card-marker):not(:has(.details-section-marker))
    > div:not([data-testid]) > [data-testid="stVerticalBlock"] {
        gap: .55rem !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(.record-card-marker):not(:has(.details-section-marker))
    [data-testid="stTextArea"] {
        width: calc(100% - .2rem) !important;
        margin-right: .1rem !important;
        margin-left: .1rem !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(.record-card-marker):not(:has(.details-section-marker))
    [data-testid="stTextArea"] textarea {
        min-height: 58px !important;
        height: 58px !important;
    }
    .bottom-add-record-marker {
        height: 0;
    }
    [data-testid="column"]:has(.bottom-add-record-marker)
    .stButton > button {
        min-height: 40px;
        border-color: #9CB9E0 !important;
        color: var(--brand) !important;
        background: #F8FBFF !important;
        box-shadow: 0 5px 14px rgba(11,61,145,.08);
    }
    [data-testid="column"]:has(.bottom-add-record-marker)
    .stButton > button:hover {
        border-color: var(--brand) !important;
        background: var(--brand-soft) !important;
    }
    .record-title {
        display: flex;
        align-items: center;
        justify-content: space-between;
        min-height: 1.95rem;
        margin-top: 0;
        padding: .36rem .7rem;
        border-left: 4px solid var(--record-accent, var(--brand));
        border-radius: 9px;
        color: var(--brand-strong);
        background: linear-gradient(90deg, #EDF5FF, #F8FBFF 72%, #FFFFFF);
        font-weight: 760;
        box-shadow: none;
    }
    .accent-teaching { --record-accent: #2563EB; }
    .accent-practice { --record-accent: #0F8A74; }
    .accent-achievement { --record-accent: #7C3AED; }
    .accent-service { --record-accent: #C26A13; }
    .record-category-pill {
        display: inline-flex;
        align-items: center;
        justify-content: flex-end;
        flex: 0 1 auto;
        max-width: calc(100% - 6rem);
        margin-left: auto;
        padding: .2rem .55rem;
        border: 1px solid currentColor;
        border-radius: 8px;
        color: var(--record-accent, var(--brand));
        background: #FFFFFF;
        font-size: .73rem;
        font-weight: 700;
        line-height: 1.35;
        text-align: right;
        white-space: normal;
        overflow-wrap: anywhere;
        word-break: break-word;
        opacity: .88;
    }

    [data-testid="stHorizontalBlock"]:has(.record-title)
    [data-testid="column"]:last-child .stButton {
        padding-top: .15rem;
    }
    [data-testid="stHorizontalBlock"]:has(.record-title)
    [data-testid="column"]:last-child .stButton button {
        border-color: #F2B7B7 !important;
        color: #C91F1F !important;
        background: #FFF5F5 !important;
        box-shadow: none !important;
        font-size: 1.15rem !important;
    }
    [data-testid="stHorizontalBlock"]:has(.record-title)
    [data-testid="column"]:last-child .stButton button:hover {
        color: #FFFFFF !important;
        border-color: var(--danger) !important;
        background: var(--danger) !important;
    }

    /* 提示与汇总 */
    [data-testid="stAlert"] {
        border-radius: 11px;
        box-shadow: none;
    }
    .total-card {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
        margin-top: .8rem;
        padding: 1rem 1.15rem;
        border: 1px solid #C8DCF7;
        border-radius: 13px;
        color: var(--brand-strong);
        background: linear-gradient(100deg, #EFF6FF, #F9FCFF);
    }
    .total-label {
        color: #475569;
        font-size: .9rem;
        font-weight: 650;
    }
    .total-value {
        font-size: 1.65rem;
        font-weight: 820;
        letter-spacing: -.02em;
    }
    .total-unit {
        margin-left: .25rem;
        color: #475569;
        font-size: .9rem;
        font-weight: 650;
    }
    /* 隐藏文本输入后的 Streamlit 英文提交提示 */
    [data-testid="InputInstructions"] {
        display: none !important;
    }

    /* 恢复文件组件 */
    [data-testid="stFileUploadDropzone"],
    [data-testid="stFileUploaderDropzone"] {
        min-height: 74px;
        border-color: #CFDBEB;
        border-radius: 10px;
        background: #F7FAFE;
    }
    [data-testid="stFileDropzoneInstructions"] > div > span::before,
    [data-testid="stFileUploaderDropzoneInstructions"] > div > span::before {
        color: #334155;
        font-size: 13px !important;
    }
    [data-testid="stFileDropzoneInstructions"] small::before,
    [data-testid="stFileUploaderDropzoneInstructions"] small::before {
        color: #718096;
        font-size: 11px !important;
    }
    [data-testid="stFileUploadDropzone"] > button,
    [data-testid="stFileUploaderDropzone"] > button {
        width: 110px !important;
        min-width: 110px !important;
        height: 36px !important;
    }
    [data-testid="stFileUploadDropzone"] > button::after,
    [data-testid="stFileUploaderDropzone"] > button::after {
        color: var(--brand) !important;
        font-size: 13px;
    }
    .app-footer {
        margin-top: 1.25rem;
        padding: 1rem 0 .25rem;
        border-top: 1px solid #DFE7F0;
        color: #7A8798;
        text-align: center;
        font-size: .82rem;
    }
    .app-footer .footer-address {
        margin-top: .35rem;
        color: #8A96A6;
        font-size: .78rem;
    }

    /* 深色系统偏好兼容：本工具保持统一的高校浅色工作界面 */
    :root,
    html,
    body,
    .stApp,
    [data-testid="stAppViewContainer"] {
        color-scheme: light !important;
    }
    html,
    body,
    #root,
    [data-testid="stAppViewContainer"] {
        background-color: #F4F7FB !important;
    }
    /* 隐藏 Streamlit 自带的顶部状态栏及其彩色装饰线 */
    header[data-testid="stHeader"],
    [data-testid="stHeader"],
    header.stAppHeader,
    .stAppHeader {
        display: none !important;
        height: 0 !important;
        min-height: 0 !important;
    }
    [data-testid="stDecoration"],
    .stDecoration {
        display: none !important;
    }
    .stButton > button:not([kind="primary"]),
    .stDownloadButton > button,
    button[data-testid="baseButton-secondary"] {
        border: 1px solid #C9D7E8 !important;
        color: var(--brand) !important;
        background: #FFFFFF !important;
    }
    .stButton > button:not([kind="primary"]):hover,
    .stDownloadButton > button:hover,
    button[data-testid="baseButton-secondary"]:hover {
        border-color: #7EA5DD !important;
        color: var(--brand-strong) !important;
        background: #F5F9FF !important;
    }
    .stButton > button[kind="primary"],
    button[data-testid="baseButton-primary"] {
        border-color: var(--brand) !important;
        color: #FFFFFF !important;
        background: linear-gradient(135deg, #1D5FBF, var(--brand)) !important;
    }
    .stButton > button:disabled,
    .stDownloadButton > button:disabled {
        border-color: #DCE5F0 !important;
        color: #94A3B8 !important;
        background: #F8FAFC !important;
        opacity: 1 !important;
    }
    input,
    textarea,
    [data-baseweb="input"],
    [data-baseweb="textarea"],
    [data-baseweb="select"] > div {
        color-scheme: light !important;
        color: var(--ink) !important;
        background-color: #FFFFFF !important;
    }
    [data-baseweb="input"] > div,
    [data-baseweb="textarea"] > div {
        border-color: var(--border) !important;
        background-color: #FFFFFF !important;
    }
    [data-baseweb="popover"],
    [data-baseweb="menu"],
    [role="listbox"],
    [role="option"] {
        color: var(--ink) !important;
        background-color: #FFFFFF !important;
    }
    [role="option"]:hover,
    [role="option"][aria-selected="true"] {
        background-color: #EDF5FF !important;
    }
    [data-testid="stFileUploadDropzone"] > button,
    [data-testid="stFileUploaderDropzone"] > button {
        border: 1px solid #C9D7E8 !important;
        background: #FFFFFF !important;
    }
    [data-testid="stAlert"] p,
    [data-testid="stCaptionContainer"] p {
        color: inherit !important;
    }

    @media (max-width: 900px) {
        .block-container {
            width: calc(100% - 2rem) !important;
            padding: .8rem 0 2rem !important;
        }
        .app-hero {
            grid-template-columns: 1fr;
            gap: 1rem;
            min-height: auto;
            padding: 1.15rem 1.25rem;
        }
        .brand-logo {
            width: min(100%, 580px);
            max-height: 82px;
        }
        .hero-copy {
            padding: .9rem 0 0;
            border-top: 1px solid rgba(148,163,184,.28);
            border-left: 0;
        }
        [data-testid="stHorizontalBlock"] {
            flex-wrap: wrap !important;
            gap: .65rem !important;
        }
        [data-testid="column"] {
            min-width: min(100%, 290px) !important;
            flex: 1 1 290px !important;
        }
        [data-testid="stHorizontalBlock"]:has(.record-title) {
            flex-wrap: nowrap !important;
        }
        [data-testid="stHorizontalBlock"]:has(.record-title)
        [data-testid="column"]:first-child {
            min-width: 0 !important;
            flex: 1 1 auto !important;
        }
        [data-testid="stHorizontalBlock"]:has(.record-title)
        [data-testid="column"]:last-child {
            min-width: 2.25rem !important;
            flex: 0 0 2.25rem !important;
        }
    }
    @media (max-width: 560px) {
        .block-container {
            width: calc(100% - 1rem) !important;
        }
        .app-hero { border-radius: 15px !important; }
        .system-name { font-size: 1.65rem !important; }
        .system-subtitle { font-size: .82rem !important; }
        [data-testid="column"] {
            min-width: 100% !important;
            flex-basis: 100% !important;
        }
        [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .basic-section-marker) > div,
        [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .details-section-marker) > div {
            padding: 1rem;
        }
        .record-title {
            align-items: flex-start;
            gap: .5rem;
            flex-direction: column;
        }
        .record-category-pill {
            justify-content: flex-start;
            max-width: 100%;
            margin-left: 0;
            text-align: left;
        }
        .total-card {
            align-items: flex-start;
            flex-direction: column;
        }
    }

    /* Streamlit 1.37 上传组件兼容：保持“恢复进度”为紧凑的整宽次级按钮。 */
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    [data-testid="stFileUploaderDropzone"] {
        min-height: 38px !important;
        height: 38px !important;
        padding: 0 !important;
        border: 0 !important;
        background: transparent !important;
    }
    [data-testid="stVerticalBlockBorderWrapper"]:has(> div:not([data-testid]) > [data-testid="stVerticalBlock"] > [data-testid="element-container"] .operations-section-marker)
    [data-testid="column"]:has(.operation-progress-marker)
    [data-testid="stFileUploaderDropzone"] > button {
        width: 100% !important;
        min-width: 0 !important;
        min-height: 38px !important;
        height: 38px !important;
        font-size: .9rem !important;
        font-weight: 680 !important;
        line-height: 1.25 !important;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown(f"""
    <div class="app-hero">
        <img class="brand-logo" src="{LOGO_DATA_URI}" alt="西京学院科技商学院">
        <div class="hero-copy">
            <div class="college-name">西京学院 · 科技商学院</div>
            <div class="system-name">教师教学工作量填报系统</div>
            <p class="system-subtitle">规范填报・精准核算・高效审核</p>
        </div>
    </div>
""", unsafe_allow_html=True)

if 'app_data' not in st.session_state:
    st.session_state.app_data = {
        "dept": "", "name": "", "title": "",
        "term": DEFAULT_ACADEMIC_TERM,
        "fill_date": date.today().isoformat(),
        "rows": []
    }
st.session_state.app_data.setdefault("term", DEFAULT_ACADEMIC_TERM)
if "restore_uploader_version" not in st.session_state:
    st.session_state.restore_uploader_version = 0
if "validation_requested" not in st.session_state:
    st.session_state.validation_requested = False
if "scroll_to_first_error" not in st.session_state:
    st.session_state.scroll_to_first_error = False
if "pending_excel_export" not in st.session_state:
    st.session_state.pending_excel_export = False
if "pending_restore_data" in st.session_state:
    restored_data = st.session_state.pop("pending_restore_data")
    st.session_state.app_data = restored_data
    clear_record_widget_state()
    for basic_key in (
        "basic_dept", "basic_name", "basic_title", "basic_term"
    ):
        st.session_state.pop(basic_key, None)
    st.session_state.validation_requested = False
    st.session_state.scroll_to_first_error = False
    st.session_state.pop("new_record_target", None)
    st.session_state.pending_excel_export = False
    st.session_state.restore_notice = (
        f"已成功恢复 {len(restored_data['rows'])} 条工作量记录。"
    )


def sync_widget_values_to_app_data():
    """在顶部导出区渲染前同步本次交互产生的最新表单值。"""
    basic_widget_fields = (
        ("basic_dept", "dept"),
        ("basic_name", "name"),
        ("basic_title", "title"),
        ("basic_term", "term"),
    )
    for widget_key, data_key in basic_widget_fields:
        if widget_key in st.session_state:
            st.session_state.app_data[data_key] = (
                st.session_state[widget_key] or ""
            )

    current_term = st.session_state.app_data.get(
        "term", DEFAULT_ACADEMIC_TERM
    )
    for i, row in enumerate(st.session_state.app_data["rows"]):
        row["term"] = current_term
        for field, prefix in (
            ("category", "cat"),
            ("remark", "rmk"),
        ):
            widget_key = f"{prefix}_{i}"
            if widget_key in st.session_state:
                row[field] = st.session_state[widget_key] or ""

        standard_key = f"std_{i}"
        if standard_key in st.session_state:
            display_map = STANDARD_DISPLAY_MAPS.get(
                row.get("category", ""), {}
            )
            selected_standard = st.session_state[standard_key]
            if selected_standard in display_map:
                row["standard"] = display_map[selected_standard]
            else:
                row["standard"] = ""

        workload_key = f"wl_{i}"
        if workload_key in st.session_state:
            try:
                workload_value = Decimal(str(st.session_state[workload_key]).strip())
                if workload_value.is_finite() and workload_value >= 0:
                    row["workload"] = float(workload_value)
            except InvalidOperation:
                pass

        time_key = f"time_{i}"
        if time_key in st.session_state:
            time_value = st.session_state[time_key]
            row["time"] = time_value.strftime("%Y-%m-%d") if time_value else ""


@st.fragment
def render_basic_info():
    sync_widget_values_to_app_data()
    validation_state = collect_validation_state(st.session_state.app_data)
    # ================= 3. 基础信息填报 =================
    with st.container(border=True):
        basic_error_class = (
            " invalid-section-marker"
            if (
                st.session_state.validation_requested
                and validation_state["first_error_target"] == "basic-info"
            )
            else ""
        )
        st.markdown(
            f"<div id='basic-info' class='basic-section-marker{basic_error_class}'></div>",
            unsafe_allow_html=True,
        )
        st.markdown("""
            <div class="section-heading">
                <span class="section-icon">👨‍🏫</span>
                <div>
                    <h2>基本信息</h2>
                </div>
            </div>
        """, unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            saved_dept = st.session_state.app_data.get("dept", "")
            if saved_dept == "电子商务":
                saved_dept = "数字商务系"
            dept_index = DEPARTMENT_OPTIONS.index(saved_dept) if saved_dept in DEPARTMENT_OPTIONS else 0
            dept = st.selectbox(
                "所属系部",
                DEPARTMENT_OPTIONS,
                index=dept_index,
                key="basic_dept"
            )
            st.session_state.app_data["dept"] = dept
        with col2:
            name = st.text_input(
                "姓名",
                value=st.session_state.app_data["name"],
                key="basic_name"
            )
            st.session_state.app_data["name"] = name
            if st.session_state.validation_requested and not name.strip():
                st.error("请输入姓名")
        with col3:
            title = st.text_input(
                "职称",
                value=st.session_state.app_data["title"],
                key="basic_title"
            )
            st.session_state.app_data["title"] = title
            if st.session_state.validation_requested and not title.strip():
                st.error("请输入职称")
        with col4:
            saved_term = st.session_state.app_data.get(
                "term", DEFAULT_ACADEMIC_TERM
            )
            term = st.selectbox(
                "学期",
                ACADEMIC_TERM_OPTIONS,
                index=ACADEMIC_TERM_OPTIONS.index(saved_term),
                key="basic_term",
            )
            st.session_state.app_data["term"] = term
            for row in st.session_state.app_data["rows"]:
                row["term"] = term



render_basic_info()

@st.fragment
def render_workload_details():
    sync_widget_values_to_app_data()
    validation_state = collect_validation_state(st.session_state.app_data)
    # ================= 4. 工作量明细展示区 =================
    with st.container(border=True):
        st.markdown(
            "<div id='workload-details' class='details-section-marker'></div>",
            unsafe_allow_html=True,
        )
        detail_title_col, add_col = st.columns([5, 1])
        with detail_title_col:
            st.markdown("""
                <div class="section-heading">
                    <span class="section-icon">📝</span>
                    <div>
                        <h2>教学工作量明细</h2>
                    </div>
                </div>
        """, unsafe_allow_html=True)
        with add_col:
            st.button(
                "＋ 新增明细条目",
                type="primary",
                use_container_width=True,
                on_click=add_empty_record,
            )

        st.info(
            "同一项目、成果或事项涉及多个类别或阶段时，按最高标准认定，"
            "不重复计算，不得跨类别重复申报。"
        )

        selected_standards = [
            row["standard"]
            for row in st.session_state.app_data["rows"]
            if "standard" in row
        ]
        total_workload = 0.0
        invalid_workload_rows = []

        for i, row in enumerate(st.session_state.app_data["rows"]):
            row_number = i + 1
            row_errors = validation_state["row_errors"].get(row_number, [])
            row_complete = not row_errors
            row_status = "已完成" if row_complete else "待完善"
            category = row.get("category", "")
            if category in ("五、实践教学", "九、创新创业与第二课堂"):
                accent_class = "accent-practice"
            elif category in ("一、特殊贡献", "二、专业建设", "四、教学改革", "十、特色人才培养项目"):
                accent_class = "accent-achievement"
            elif category in ("三、教研室建设", "七、招生与就业", "八、公共性事务服务"):
                accent_class = "accent-service"
            else:
                accent_class = "accent-teaching"
            standard_owner = get_standard_owner(row.get("standard", ""))
            owner_label = (
                html.escape(standard_owner)
                if standard_owner
                else "未标注"
            )

            with st.container(border=True):
                invalid_class = (
                    " invalid-record-marker"
                    if st.session_state.validation_requested and row_errors
                    else ""
                )
                st.markdown(
                    (
                        f"<div id='record-{row_number}' "
                        f"class='record-card-marker{invalid_class}'></div>"
                    ),
                    unsafe_allow_html=True,
                )
                title_col, delete_col = st.columns([20, 0.8])
                with title_col:
                    st.markdown(
                        f"""
                        <div class="record-title {accent_class}">
                            <span>记录 {row_number}</span>
                            <span class="record-category-pill">{owner_label}</span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )
                with delete_col:
                    st.button(
                        "×",
                        key=f"del_{i}",
                        help=f"删除记录 {row_number}",
                        on_click=delete_record,
                        args=(i,),
                    )

                c1, c2, c3, c4 = st.columns([1.75, 4.5, 1.05, 1.9])
                with c1:
                    cat_index = (
                        WORKLOAD_CATEGORY_OPTIONS.index(row["category"])
                        if row["category"] in WORKLOAD_DATA
                        else None
                    )
                    selected_category = st.selectbox(
                        "项目类别",
                        WORKLOAD_CATEGORY_OPTIONS,
                        key=f"cat_{i}",
                        index=cat_index,
                        placeholder="请选择项目类别",
                    )
                    row["category"] = selected_category or ""
                    if (
                        st.session_state.validation_requested
                        and not row["category"]
                    ):
                        st.error("请选择项目类别")
                with c2:
                    display_map = STANDARD_DISPLAY_MAPS.get(
                        row["category"], {}
                    )
                    display_options = STANDARD_DISPLAY_OPTIONS.get(
                        row["category"], ()
                    )
                    current_display = format_standard_option(row["standard"])
                    default_idx = (
                        display_options.index(current_display)
                        if current_display in display_options
                        else None
                    )
                    selected_standard_display = st.selectbox(
                        "计算标准（输入关键词可检索）",
                        display_options,
                        key=f"std_{i}",
                        index=default_idx,
                        placeholder=(
                            "请选择计算标准"
                            if row["category"]
                            else "请先选择项目类别"
                        ),
                        disabled=not row["category"],
                        help="展开后可直接输入项目名称或核算规则中的关键词进行检索。",
                    )
                    row["standard"] = display_map.get(selected_standard_display, "")
                    if (
                        st.session_state.validation_requested
                        and not row["standard"]
                    ):
                        st.error("请选择计算标准")
                    elif (
                        row["standard"]
                        and selected_standards.count(row["standard"]) > 1
                    ):
                        duplicate_record = next(
                            (
                                number
                                for number, other_row in enumerate(
                                    st.session_state.app_data["rows"], start=1
                                )
                                if number != row_number
                                and other_row.get("standard") == row["standard"]
                            ),
                            None,
                        )
                        if duplicate_record is not None:
                            st.error(
                                f"该成果类别与记录{duplicate_record}重复"
                            )
                with c3:
                    workload_text = st.text_input(
                        "工作量（学时）",
                        key=f"wl_{i}",
                        value=(
                            ""
                            if row.get("workload", "") in ("", None)
                            else format_workload_value(row.get("workload"))
                        ),
                        placeholder="请输入工作量"
                    )
                    if not workload_text.strip():
                        row["workload"] = ""
                        if st.session_state.validation_requested:
                            st.error("请填写工作量")
                    else:
                        try:
                            workload_value = Decimal(workload_text.strip())
                            if not workload_value.is_finite() or workload_value < 0:
                                raise InvalidOperation
                            row["workload"] = float(workload_value)
                            total_workload += row["workload"]
                        except (InvalidOperation, ValueError):
                            invalid_workload_rows.append(str(i + 1))
                            st.error("请输入不小于 0 的数字")
                with c4:
                    time_str = row.get("time", "")
                    try:
                        current_date = (
                            datetime.strptime(time_str, "%Y-%m-%d").date()
                            if time_str
                            else None
                        )
                    except ValueError:
                        current_date = None

                    selected_date = st.date_input(
                        "实际取得时间",
                        value=current_date,
                        key=f"time_{i}"
                    )
                    row["time"] = selected_date.strftime("%Y-%m-%d") if selected_date else ""
                    if (
                        st.session_state.validation_requested
                        and not row["time"]
                    ):
                        st.error("请选择完成时间")

                row["remark"] = st.text_area(
                    "完成说明",
                    key=f"rmk_{i}",
                    value=row.get("remark", ""),
                    placeholder="请填写具体成果/业绩内容，不得为空",
                    height=68
                )
                if (
                    st.session_state.validation_requested
                    and not row["remark"].strip()
                ):
                    st.error("请填写具体成果或业绩内容")

        if st.session_state.app_data["rows"]:
            bottom_add_left, bottom_add_col, bottom_add_right = st.columns(
                [4, 1.7, 4]
            )
            with bottom_add_col:
                st.markdown(
                    "<div class='bottom-add-record-marker'></div>",
                    unsafe_allow_html=True,
                )
                st.button(
                    "＋ 继续新增一条",
                    use_container_width=True,
                    on_click=add_empty_record,
                    key="add_record_bottom",
                )

        total_workload_display = format_workload_value(f"{total_workload:.10f}")
        st.markdown(
            f"""
            <div class="total-card">
                <span class="total-label">工作量合计</span>
                <span>
                    <span class="total-value">{total_workload_display}</span>
                    <span class="total-unit">学时</span>
                </span>
            </div>
            """,
            unsafe_allow_html=True
        )

        new_record_target = st.session_state.pop(
            "new_record_target", None
        )
        if new_record_target is not None:
            new_record_id = json.dumps(f"record-{new_record_target}")
            components.html(
                f"""
                <script>
                window.setTimeout(() => {{
                    const target = parent.document.getElementById({new_record_id});
                    if (!target) return;
                    const recordCard =
                        target.closest('[data-testid="stVerticalBlockBorderWrapper"]') ||
                        target.parentElement;
                    recordCard.scrollIntoView({{
                        behavior: "smooth",
                        block: "start"
                    }});
                    window.setTimeout(() => {{
                        const categorySelect = recordCard.querySelector(
                            '[data-testid="stSelectbox"] [role="combobox"]'
                        );
                        if (categorySelect) {{
                            categorySelect.focus({{preventScroll: true}});
                        }}
                    }}, 420);
                }}, 160);
                </script>
                """,
                height=0,
            )


render_workload_details()

sync_widget_values_to_app_data()
validation_state = collect_validation_state(st.session_state.app_data)

# ================= 5. 文件生成与进度管理 =================
with st.container(border=True):
    st.markdown(
        "<div id='data-operations' class='operations-section-marker'></div>",
        unsafe_allow_html=True,
    )
    st.markdown("""
        <div class="section-heading">
            <span class="section-icon">⚙️</span>
            <div>
                <h2>数据导入导出</h2>
            </div>
        </div>
    """, unsafe_allow_html=True)
    excel_col, progress_col = st.columns([1, 1.55])

    with excel_col:
        st.markdown("<div class='operation-primary-marker'></div>", unsafe_allow_html=True)
        validation_issues = validation_state["issues"]
        excel_is_ready = not validation_issues

        if excel_is_ready:
            st.button(
                "📊 导出 Excel 报表",
                type="primary",
                use_container_width=True,
                on_click=request_excel_export,
            )
        else:
            st.button(
                "📊 导出 Excel 报表",
                type="primary",
                use_container_width=True,
                on_click=request_export_validation,
            )

        should_generate_excel = st.session_state.pending_excel_export
        st.session_state.pending_excel_export = False
        if should_generate_excel:
            try:
                excel_bytes = generate_excel_bytes(
                    st.session_state.app_data,
                    date.today().isoformat(),
                )
                excel_payload = base64.b64encode(excel_bytes).decode("ascii")
                excel_filename = json.dumps(
                    make_excel_filename(st.session_state.app_data),
                    ensure_ascii=False,
                )
                components.html(
                    f"""
                    <script>
                    const binary = atob("{excel_payload}");
                    const bytes = new Uint8Array(binary.length);
                    for (let i = 0; i < binary.length; i += 1) {{
                        bytes[i] = binary.charCodeAt(i);
                    }}
                    const blob = new Blob(
                        [bytes],
                        {{type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}}
                    );
                    const url = URL.createObjectURL(blob);
                    const link = parent.document.createElement("a");
                    link.href = url;
                    link.download = {excel_filename};
                    link.style.display = "none";
                    parent.document.body.appendChild(link);
                    link.click();
                    link.remove();
                    window.setTimeout(() => URL.revokeObjectURL(url), 1500);
                    </script>
                    """,
                    height=0,
                )
                st.session_state.action_notice = "Excel 报表已生成并开始下载。"
            except Exception as e:
                st.error(f"Excel 生成失败：{e}")

        if st.session_state.validation_requested and validation_issues:
            issue_list = "\n".join(
                f"- {issue}"
                for issue in validation_issues
            )
            st.error(f"导出前请完成以下项目：\n\n{issue_list}")

    with progress_col:
        st.markdown("<div class='operation-progress-marker'></div>", unsafe_allow_html=True)
        st.markdown(
            "<div class='operation-pair-title'>🔄 <span>进度管理</span></div>",
            unsafe_allow_html=True
        )
        save_action_col, restore_action_col = st.columns(2)

        with save_action_col:
            export_data = {
                "schema_version": 1,
                "dept": st.session_state.app_data["dept"],
                "name": st.session_state.app_data["name"],
                "title": st.session_state.app_data["title"],
                "term": st.session_state.app_data["term"],
                "fill_date": st.session_state.app_data["fill_date"],
                "rows": st.session_state.app_data["rows"]
            }
            json_str = json.dumps(export_data, ensure_ascii=False)
            st.download_button(
                "💾 暂存填报进度",
                data=json_str,
                file_name="工作量暂存.json",
                mime="application/json",
                use_container_width=True,
                on_click=set_action_notice,
                args=("暂存文件已生成并开始下载。",),
            )

        with restore_action_col:
            uploaded_file = st.file_uploader(
                "恢复填报进度",
                type="json",
                label_visibility="collapsed",
                key=f"restore_file_{st.session_state.restore_uploader_version}"
            )
            if uploaded_file is not None:
                try:
                    uploaded_bytes = uploaded_file.getvalue()
                    if len(uploaded_bytes) > 2 * 1024 * 1024:
                        raise ValueError("暂存文件不能超过 2MB")
                    restored_json = json.loads(
                        uploaded_bytes.decode("utf-8-sig")
                    )
                    restored_data = normalize_restored_data(restored_json)
                except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as exc:
                    st.error(f"恢复失败：{exc}")
                else:
                    st.session_state.pending_restore_data = restored_data
                    st.session_state.restore_uploader_version += 1
                    st.rerun()
        st.markdown(
            "<div class='operation-help'>下载当前填报内容作为本地暂存文件，后续可重新导入并继续填写。</div>",
            unsafe_allow_html=True
        )
        action_notice = st.session_state.pop("action_notice", None)
        if action_notice:
            st.success(action_notice)
        restore_notice = st.session_state.pop("restore_notice", None)
        if restore_notice:
            st.success(restore_notice)

new_record_target = st.session_state.pop("new_record_target", None)
if new_record_target is not None:
    new_record_id = json.dumps(f"record-{new_record_target}")
    components.html(
        f"""
        <script>
        window.setTimeout(() => {{
            const target = parent.document.getElementById({new_record_id});
            if (!target) return;
            const recordCard =
                target.closest('[data-testid="stVerticalBlockBorderWrapper"]') ||
                target.parentElement;
            recordCard.scrollIntoView({{
                behavior: "smooth",
                block: "start"
            }});
            window.setTimeout(() => {{
                const categorySelect = recordCard.querySelector(
                    '[data-testid="stSelectbox"] [role="combobox"]'
                );
                if (categorySelect) {{
                    categorySelect.focus({{preventScroll: true}});
                }}
            }}, 420);
        }}, 160);
        </script>
        """,
        height=0,
    )
elif (
    st.session_state.scroll_to_first_error
    and validation_state.get("first_error_target")
):
    target_id = json.dumps(validation_state["first_error_target"])
    components.html(
        f"""
        <script>
        window.setTimeout(() => {{
            const target = parent.document.getElementById({target_id});
            if (target) {{
                const visibleTarget =
                    target.closest('[data-testid="stVerticalBlockBorderWrapper"]') ||
                    target.closest('[data-testid="stExpander"]') ||
                    target.parentElement;
                visibleTarget.scrollIntoView({{behavior: "smooth", block: "center"}});
            }}
        }}, 120);
        </script>
        """,
        height=0,
    )
    st.session_state.scroll_to_first_error = False

# 补充 Streamlit 原生控件缺失的中文可访问名称。
accessibility_version = (
    f"{len(st.session_state.app_data['rows'])}:"
    f"{st.session_state.restore_uploader_version}"
)
components.html(
    f"<!-- accessibility:{accessibility_version} -->" + """
    <script>
    const chineseInterfaceText = {
        January: '1月', February: '2月', March: '3月',
        April: '4月', May: '5月', June: '6月',
        July: '7月', August: '8月', September: '9月',
        October: '10月', November: '11月', December: '12月',
        'Choose an option': '请选择',
        'No result': '未找到匹配项',
        'No results': '未找到匹配项',
        'No results found': '未找到匹配项'
    };
    const translateInterfaceText = (root) => {
        const walker = root.createTreeWalker(
            root.body,
            parent.NodeFilter.SHOW_TEXT
        );
        const textNodes = [];
        while (walker.nextNode()) textNodes.push(walker.currentNode);
        textNodes.forEach((node) => {
            const sourceText = node.nodeValue.trim();
            if (chineseInterfaceText[sourceText]) {
                node.nodeValue = chineseInterfaceText[sourceText];
            }
        });
        root.querySelectorAll(
            'input[placeholder="Choose an option"]'
        ).forEach((input) => {
            input.setAttribute('placeholder', '请输入关键词检索');
        });
    };
    const applyChineseLabels = () => {
        const root = parent.document;
        root.querySelectorAll('[data-testid="stDateInput"]').forEach((widget) => {
            const label = widget.querySelector('[data-testid="stWidgetLabel"] p');
            const input = widget.querySelector('input');
            if (input) {
                input.setAttribute('placeholder', '年/月/日');
                if (label) {
                    input.setAttribute('aria-label', label.textContent.trim());
                }
            }
        });
        root.querySelectorAll('[data-testid="stFileUploader"]').forEach((widget) => {
            const button = widget.querySelector(
                '[data-testid="stFileUploaderDropzone"] button, ' +
                '[data-testid="stFileUploadDropzone"] > button'
            );
            if (button) {
                button.setAttribute('aria-label', '恢复填报进度');
                button.setAttribute('title', '恢复填报进度');
                if (button.textContent.trim() !== '📂 恢复填报进度') {
                    button.textContent = '📂 恢复填报进度';
                }
            }
        });
        translateInterfaceText(root);
    };
    applyChineseLabels();
    window.setTimeout(applyChineseLabels, 200);
    if (parent.__workloadCalendarObserver) {
        parent.__workloadCalendarObserver.disconnect();
    }
    let labelUpdateScheduled = false;
    const calendarObserver = new parent.MutationObserver((mutations) => {
        /*
         * Streamlit 会频繁更新状态节点。只在新增日期选择器、上传控件或
         * 弹层时处理中文标签，避免每次细小 DOM 变化都扫描整张页面。
         */
        const needsUpdate = mutations.some((mutation) =>
            mutation.target.parentElement?.closest?.(
                '[role="listbox"], [data-baseweb="popover"]'
            ) ||
            Array.from(mutation.addedNodes).some((node) => {
                if (node.nodeType !== parent.Node.ELEMENT_NODE) return false;
                return (
                    node.matches?.(
                        '[data-testid="stDateInput"], ' +
                        '[data-testid="stFileUploader"], ' +
                        '[role="dialog"], [role="listbox"], ' +
                        '[data-baseweb="calendar"], [data-baseweb="popover"]'
                    ) ||
                    node.querySelector?.(
                        '[data-testid="stDateInput"], ' +
                        '[data-testid="stFileUploader"], ' +
                        '[role="dialog"], [role="listbox"], ' +
                        '[data-baseweb="calendar"], [data-baseweb="popover"]'
                    )
                );
            })
        );
        if (!needsUpdate || labelUpdateScheduled) return;
        labelUpdateScheduled = true;
        parent.requestAnimationFrame(() => {
            labelUpdateScheduled = false;
            applyChineseLabels();
        });
    });
    calendarObserver.observe(parent.document.body, {
        childList: true,
        subtree: true
    });
    parent.__workloadCalendarObserver = calendarObserver;
    </script>
    """,
    height=0,
)

# ================= 6. 底部版权信息 =================
st.markdown(
    f"""
    <div class='app-footer'>
        <div>© 2026 西京学院科技商学院 · 系统版本 v{APP_VERSION}</div>
        <div class='footer-address'>地址：陕西省西安市长安区西京路1号</div>
    </div>
    """,
    unsafe_allow_html=True
)

run_elapsed_ms = (time.perf_counter() - RUN_STARTED_AT) * 1000
if run_elapsed_ms >= 750:
    PERFORMANCE_LOGGER.warning(
        "slow_rerun elapsed_ms=%.1f rows=%d validation_requested=%s",
        run_elapsed_ms,
        len(st.session_state.app_data.get("rows", [])),
        st.session_state.validation_requested,
    )
