"""
Excel 数据清洗网页助手
=======================
用 Streamlit + openpyxl 构建的本地数据清洗工具。
上传培训报名 Excel → 勾选清洗规则 → 一键清洗 → 下载结果。
"""

# ============================================================
# SECTION 1: 导入库和页面配置
# ============================================================
import io

import streamlit as st

# 清洗能力模块
import cleaner

# 页面配置（必须是第一个 Streamlit 命令）
st.set_page_config(
    page_title="Excel 数据清洗网页助手",
    page_icon="🧹",
    layout="wide",
)

# ============================================================
# SECTION 2: 常量与配置
# ============================================================

# —— 文件相关 ——
MAX_FILE_SIZE_MB = 50  # 文件大小上限（MB）

# —— 工作表相关 ——
DEFAULT_SHEET_NAME = "报名明细"  # 默认读取的工作表名称

# —— 必要列清单 ——
REQUIRED_COLUMNS = [
    "姓名",
    "手机号",
    "部门",
    "城市",
    "报名课程",
    "报名日期",
    "报名状态",
    "应缴金额",
    "实缴金额",
    "备注",
]

# —— 6 条清洗规则 ——
CLEANING_RULES = [
    ("rule_trim_names", "清理姓名前后空格", "去除姓名中多余的前后空格，让数据更整洁"),
    ("rule_dept", "统一部门名称", "将「技术部门」「技术中心」等变体统一为「技术部」等标准名称"),
    ("rule_course", "统一课程名称", "将「python 入门」「AI 实战」等变体统一为标准课程名称"),
    ("rule_status", "补全空报名状态", "报名状态为空白时，自动填入「待确认」"),
    ("rule_phone", "检查手机号异常", "检测手机号格式是否正确，异常号码用红色标记提醒"),
    ("rule_duplicate", "识别重复报名", "按手机号查找重复报名记录，用黄色标记提示"),
]

# 下载文件名
OUTPUT_FILENAME = "数据清洗结果-网页版.xlsx"


# ============================================================
# SECTION 3: Session State 初始化
# ============================================================

def _init_session_state():
    """初始化所有 session_state 变量。"""
    defaults = {
        "uploaded_file_name": None,
        "uploaded_file_size": None,
        "result_wb": None,          # 清洗后的 openpyxl Workbook
        "metrics": None,            # 5 个摘要指标
        "cleaning_error": None,     # 清洗失败时的错误信息
        "cleaning_done": False,     # 是否已完成清洗
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


_init_session_state()


def _format_file_size(size_bytes: int) -> str:
    """将字节数转为可读的文件大小字符串。"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    else:
        return f"{size_bytes / (1024 * 1024):.2f} MB"


# ============================================================
# SECTION 4: 主界面
# ============================================================

# ---- 标题区 ----
st.title("🧹 Excel 数据清洗网页助手")

st.markdown(
    "自动检查并修正培训报名 Excel 中的常见数据问题，"
    "让数据更规范、更整洁。",
)

st.divider()

# ---- 上传区 + 规则区（左右两栏） ----
col_upload, col_rules = st.columns([1, 1])

# ========== 左栏：上传区域 ==========
with col_upload:
    st.subheader("📎 上传 Excel 文件")

    uploaded_file = st.file_uploader(
        "上传Excel文件（.xlsx）",
        type=["xlsx"],
        help="请上传 .xlsx 格式的培训报名 Excel 文件，文件大小不超过 50MB",
        label_visibility="visible",
    )

    # 上传新文件时清除旧的清洗结果
    if uploaded_file is not None:
        if st.session_state.uploaded_file_name != uploaded_file.name:
            st.session_state.result_wb = None
            st.session_state.metrics = None
            st.session_state.cleaning_error = None
            st.session_state.cleaning_done = False
            st.session_state.uploaded_file_name = uploaded_file.name
            st.session_state.uploaded_file_size = uploaded_file.size

        file_size_mb = uploaded_file.size / (1024 * 1024)

        # 检查文件大小
        if file_size_mb > MAX_FILE_SIZE_MB:
            st.error(
                f"文件大小为 {_format_file_size(uploaded_file.size)}，"
                f"超过了 {MAX_FILE_SIZE_MB}MB 的上限。"
                f"请压缩文件后重新上传。"
            )
        else:
            st.success("文件已上传成功！")
            st.markdown(
                f"**文件名：** {uploaded_file.name}  \n"
                f"**文件大小：** {_format_file_size(uploaded_file.size)}"
            )
    else:
        # 未上传时清除旧结果
        if st.session_state.uploaded_file_name is not None:
            st.session_state.result_wb = None
            st.session_state.metrics = None
            st.session_state.cleaning_error = None
            st.session_state.cleaning_done = False
            st.session_state.uploaded_file_name = None
            st.session_state.uploaded_file_size = None

        st.info(
            "请点击上方按钮选择 Excel 文件，或直接将文件拖拽到此区域。\n\n"
            "支持 `.xlsx` 格式，文件大小不超过 50MB。"
        )

# ========== 右栏：选择清洗规则 + 开始清洗 ==========
with col_rules:
    st.subheader("🔧 选择清洗规则")

    st.caption("请勾选需要执行的清洗规则（默认全部启用）：")

    # 6 个默认勾选的规则
    rule_states = {}
    for rule_id, rule_name, rule_desc in CLEANING_RULES:
        rule_states[rule_id] = st.checkbox(
            f"**{rule_name}**",
            value=True,
            help=rule_desc,
        )
        st.caption(f"　{rule_desc}")

    st.divider()

    # ---- 开始清洗按钮 ----
    start_clicked = st.button(
        "🚀 开始清洗",
        type="primary",
        use_container_width=True,
    )

    if start_clicked:
        # 检查是否已上传文件
        if uploaded_file is None:
            st.warning("请先上传Excel文件。")
        else:
            # 检查是否有文件大小问题
            file_size_mb = uploaded_file.size / (1024 * 1024)
            if file_size_mb > MAX_FILE_SIZE_MB:
                st.error("文件大小超过限制，请重新上传。")
            else:
                with st.spinner("正在清洗数据，请稍候..."):
                    try:
                        # 读取文件字节并调用清洗模块
                        file_bytes = uploaded_file.read()
                        result_wb, metrics = cleaner.clean(file_bytes, enabled_rules=rule_states)

                        # 保存到 session_state
                        st.session_state.result_wb = result_wb
                        st.session_state.metrics = metrics
                        st.session_state.cleaning_error = None
                        st.session_state.cleaning_done = True

                    except ValueError as e:
                        st.session_state.cleaning_error = str(e)
                        st.session_state.cleaning_done = False
                        st.session_state.result_wb = None
                        st.session_state.metrics = None

                    except Exception as e:
                        st.session_state.cleaning_error = (
                            f"处理失败，请检查Excel文件格式。\n\n"
                            f"错误详情：{e}"
                        )
                        st.session_state.cleaning_done = False
                        st.session_state.result_wb = None
                        st.session_state.metrics = None

                # 刷新页面以显示结果
                st.rerun()

# ---- 清洗结果展示区 ----
if st.session_state.cleaning_done and st.session_state.metrics is not None:
    st.divider()
    st.success("处理完成，可以下载结果文件。")

    # 5 个摘要指标
    metrics = st.session_state.metrics
    m_col1, m_col2, m_col3, m_col4, m_col5 = st.columns(5)

    with m_col1:
        st.metric("总记录数", metrics["总记录数"])
    with m_col2:
        st.metric(
            "异常记录数",
            metrics["异常记录数"],
            delta=None if metrics["异常记录数"] == 0 else f"{metrics['异常记录数']} 条",
        )
    with m_col3:
        st.metric(
            "重复手机号数",
            metrics["重复手机号数"],
            delta=None if metrics["重复手机号数"] == 0 else f"{metrics['重复手机号数']} 个",
        )
    with m_col4:
        st.metric("应缴金额合计", f"{metrics['应缴金额合计']:,.2f}")
    with m_col5:
        st.metric("实缴金额合计", f"{metrics['实缴金额合计']:,.2f}")

    # 下载按钮
    st.divider()

    # 将结果 workbook 转为字节流
    output_bytes = io.BytesIO()
    st.session_state.result_wb.save(output_bytes)
    output_bytes.seek(0)

    st.download_button(
        label="📥 下载数据清洗结果.xlsx",
        data=output_bytes,
        file_name=OUTPUT_FILENAME,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

# ---- 清洗失败提示 ----
if st.session_state.cleaning_error is not None and not st.session_state.cleaning_done:
    st.divider()
    st.error(st.session_state.cleaning_error)

# ---- 文件格式要求（折叠区，供参考） ----
st.divider()

with st.expander("📝 点击查看：Excel 文件的格式要求", expanded=False):
    st.markdown(
        "上传的 Excel 文件需要满足以下要求，否则系统会给出提示。"
        "如果你不确定自己的文件是否符合要求，可以对照检查。"
    )

    fmt_col1, fmt_col2 = st.columns([1, 2])

    with fmt_col1:
        st.markdown("#### 工作表名称")
        st.info(f"**`{DEFAULT_SHEET_NAME}`**")
        st.caption("文件中必须包含此名称的工作表。")

    with fmt_col2:
        st.markdown("#### 必要列（表头，共 10 列）")
        st.caption("表头文字必须一致，顺序不限。")

        col_specs = [
            ("姓名", "报名人姓名，不能为空"),
            ("手机号", "中国大陆手机号，11 位数字"),
            ("部门", "所属部门，系统会自动统一名称变体"),
            ("城市", "所在城市"),
            ("报名课程", "课程名称，系统会自动统一名称变体"),
            ("报名日期", "报名日期，支持多种日期格式"),
            ("报名状态", "如「已缴费」「未缴费」，空值会自动补为「待确认」"),
            ("应缴金额", "应缴纳的报名费用（数字）"),
            ("实缴金额", "实际缴纳的报名费用（数字）"),
            ("备注", "附加信息，可以为空"),
        ]

        for col_name, col_desc in col_specs:
            st.markdown(f"`{col_name}` — {col_desc}")
