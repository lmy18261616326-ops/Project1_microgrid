from __future__ import annotations

from pathlib import Path

from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Image,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


PROJECT_ROOT = Path(__file__).resolve().parents[1]
ASSET_DIR = PROJECT_ROOT / "docs" / "report_assets"
RESULT_DIR = PROJECT_ROOT / "results"
OUTPUT_PATH = (
    PROJECT_ROOT
    / "output"
    / "pdf"
    / "PV_Battery_MPC_Model_Audit_and_System_Report.pdf"
)

PAGE_SIZE = landscape(A4)
PAGE_WIDTH, PAGE_HEIGHT = PAGE_SIZE
MARGIN_X = 14 * mm
MARGIN_TOP = 15 * mm
MARGIN_BOTTOM = 14 * mm
CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN_X
CONTENT_HEIGHT = PAGE_HEIGHT - MARGIN_TOP - MARGIN_BOTTOM

NAVY = colors.HexColor("#17324D")
BLUE = colors.HexColor("#1F6AA5")
TEAL = colors.HexColor("#12827B")
GREEN = colors.HexColor("#2E7D32")
LIGHT_GREEN = colors.HexColor("#E9F5EA")
LIGHT_BLUE = colors.HexColor("#EAF2F8")
LIGHT_GRAY = colors.HexColor("#F3F5F7")
MID_GRAY = colors.HexColor("#65727E")
DARK = colors.HexColor("#1E2933")
AMBER = colors.HexColor("#A96500")
LIGHT_AMBER = colors.HexColor("#FFF4DD")
RED = colors.HexColor("#B42318")


def register_fonts() -> None:
    pdfmetrics.registerFont(
        TTFont(
            "MicrosoftYaHei",
            r"C:\Windows\Fonts\msyh.ttc",
            subfontIndex=0,
        )
    )
    pdfmetrics.registerFont(
        TTFont(
            "MicrosoftYaHeiBold",
            r"C:\Windows\Fonts\msyhbd.ttc",
            subfontIndex=0,
        )
    )


register_fonts()

BASE_STYLES = getSampleStyleSheet()
styles = {
    "cover_title": ParagraphStyle(
        "cover_title",
        fontName="MicrosoftYaHeiBold",
        fontSize=25,
        leading=34,
        textColor=colors.white,
        alignment=TA_CENTER,
        spaceAfter=8,
    ),
    "cover_subtitle": ParagraphStyle(
        "cover_subtitle",
        fontName="MicrosoftYaHei",
        fontSize=11,
        leading=18,
        textColor=colors.HexColor("#DCE8F1"),
        alignment=TA_CENTER,
    ),
    "h1": ParagraphStyle(
        "h1",
        fontName="MicrosoftYaHeiBold",
        fontSize=17,
        leading=23,
        textColor=NAVY,
        spaceBefore=4,
        spaceAfter=8,
    ),
    "h2": ParagraphStyle(
        "h2",
        fontName="MicrosoftYaHeiBold",
        fontSize=12,
        leading=17,
        textColor=BLUE,
        spaceBefore=7,
        spaceAfter=5,
    ),
    "h3": ParagraphStyle(
        "h3",
        fontName="MicrosoftYaHeiBold",
        fontSize=10,
        leading=14,
        textColor=TEAL,
        spaceBefore=5,
        spaceAfter=3,
    ),
    "body": ParagraphStyle(
        "body",
        fontName="MicrosoftYaHei",
        fontSize=8.7,
        leading=14,
        textColor=DARK,
        spaceAfter=4,
    ),
    "small": ParagraphStyle(
        "small",
        fontName="MicrosoftYaHei",
        fontSize=7.4,
        leading=11,
        textColor=DARK,
    ),
    "tiny": ParagraphStyle(
        "tiny",
        fontName="MicrosoftYaHei",
        fontSize=6.5,
        leading=9,
        textColor=DARK,
    ),
    "caption": ParagraphStyle(
        "caption",
        fontName="MicrosoftYaHei",
        fontSize=7.2,
        leading=10,
        textColor=MID_GRAY,
        alignment=TA_CENTER,
        spaceBefore=3,
        spaceAfter=5,
    ),
    "note": ParagraphStyle(
        "note",
        fontName="MicrosoftYaHei",
        fontSize=8,
        leading=13,
        textColor=AMBER,
        leftIndent=8,
        rightIndent=8,
    ),
    "code": ParagraphStyle(
        "code",
        fontName="MicrosoftYaHei",
        fontSize=7.6,
        leading=12,
        textColor=colors.HexColor("#243447"),
        backColor=LIGHT_GRAY,
        borderPadding=5,
    ),
}


def para(text: str, style: str = "body") -> Paragraph:
    return Paragraph(text, styles[style])


def cell(text: object, style: str = "small") -> Paragraph:
    return Paragraph(str(text), styles[style])


def table(
    rows: list[list[object]],
    widths: list[float] | None = None,
    header: bool = True,
    font_style: str = "small",
    repeat_rows: int = 1,
) -> Table:
    formatted = [
        [value if hasattr(value, "wrap") else cell(value, font_style) for value in row]
        for row in rows
    ]
    result = Table(
        formatted,
        colWidths=widths,
        repeatRows=repeat_rows if header else 0,
        hAlign="LEFT",
    )
    commands: list[tuple] = [
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#BCC8D1")),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1 if header else 0), (-1, -1), [colors.white, LIGHT_GRAY]),
    ]
    if header:
        commands.extend(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "MicrosoftYaHeiBold"),
            ]
        )
    result.setStyle(TableStyle(commands))
    return result


def badge(text: str, color: colors.Color = GREEN) -> Table:
    t = Table([[cell(text, "body")]], colWidths=[70 * mm])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), color),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("BOX", (0, 0), (-1, -1), 0.8, colors.white),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return t


def scaled_image(path: Path, max_width: float, max_height: float) -> Image:
    with PILImage.open(path) as source:
        width, height = source.size
    scale = min(max_width / width, max_height / height)
    return Image(str(path), width=width * scale, height=height * scale)


def figure(path: Path, caption: str, max_height: float = 108 * mm) -> list:
    return [
        scaled_image(path, CONTENT_WIDTH, max_height),
        para(caption, "caption"),
    ]


def section_page(title: str, subtitle: str) -> list:
    return [
        PageBreak(),
        Spacer(1, 16 * mm),
        para(title, "h1"),
        Spacer(1, 3 * mm),
        para(subtitle, "body"),
        Spacer(1, 5 * mm),
    ]


def on_page(canvas, doc) -> None:
    page_number = canvas.getPageNumber()
    canvas.saveState()
    if page_number == 1:
        canvas.setFillColor(NAVY)
        canvas.rect(0, 0, PAGE_WIDTH, PAGE_HEIGHT, fill=1, stroke=0)
    else:
        canvas.setStrokeColor(colors.HexColor("#D4DCE2"))
        canvas.setLineWidth(0.5)
        canvas.line(MARGIN_X, PAGE_HEIGHT - 10 * mm, PAGE_WIDTH - MARGIN_X, PAGE_HEIGHT - 10 * mm)
        canvas.setFont("MicrosoftYaHei", 7)
        canvas.setFillColor(MID_GRAY)
        canvas.drawString(MARGIN_X, PAGE_HEIGHT - 7.2 * mm, "PV-Battery MPC 模型检查与系统说明报告")
        canvas.drawRightString(
            PAGE_WIDTH - MARGIN_X,
            PAGE_HEIGHT - 7.2 * mm,
            "Predictive Control of PV/Battery System under Load and Environmental Uncertainty",
        )
        canvas.line(MARGIN_X, 9.2 * mm, PAGE_WIDTH - MARGIN_X, 9.2 * mm)
        canvas.drawString(MARGIN_X, 5.8 * mm, "生成日期：2026-07-31  |  MATLAB/Simulink R2025a")
        canvas.drawRightString(PAGE_WIDTH - MARGIN_X, 5.8 * mm, f"第 {page_number} 页")
    canvas.restoreState()


story: list = []

# Cover
story.extend(
    [
        Spacer(1, 37 * mm),
        para("PV/电池系统预测控制模型", "cover_title"),
        para("故障检查、系统连接、模块构建与参数说明报告", "cover_title"),
        Spacer(1, 7 * mm),
        para(
            "对应论文：Predictive Control of PV/Battery System under Load and Environmental Uncertainty",
            "cover_subtitle",
        ),
        Spacer(1, 13 * mm),
        Table(
            [
                [badge("结构检查：通过"), badge("闭环仿真：通过"), badge("模式切换修复：通过")],
            ],
            colWidths=[82 * mm, 82 * mm, 82 * mm],
            hAlign="CENTER",
        ),
        Spacer(1, 17 * mm),
        para(
            "检查对象：PV_Battery_MPC.slx<br/>"
            "工程目录：D:/PV_MPPT/PV_Battery_MPC_Project<br/>"
            "检查环境：MATLAB、Simulink、Model Predictive Control Toolbox R2025a",
            "cover_subtitle",
        ),
        Spacer(1, 13 * mm),
        para(
            "结论：在已执行的结构、编译、短时闭环、论文阶跃及 Mode III 限功率测试范围内，"
            "模型未发现未连接端口、悬空信号、非有限状态、占空比越界或闭环失稳。"
            "检查过程中发现并修复了 Mode III 进入瞬间一个控制周期的模式编码空档。",
            "cover_subtitle",
        ),
        PageBreak(),
    ]
)

# Executive summary
story.extend(
    [
        para("执行摘要", "h1"),
        para(
            "本报告对完整 MATLAB Project 中的主模型、七个核心子系统、初始化及非线性 MPC "
            "回调进行了逐层复核。模型采用论文给出的六状态平均值方程，并尽量用普通 Simulink "
            "模块实现；主模型内部没有 MATLAB Function 块。非线性优化由标准 "
            "Nonlinear MPC Controller 模块调用外部状态预测函数完成。",
        ),
        para("检查结论", "h2"),
        table(
            [
                ["检查轮次", "检查方法", "覆盖内容", "结果"],
                [
                    "第 1 次",
                    "Simulink 全层级结构检查",
                    "未连接端口、悬空信号线、Stateflow 编辑期问题",
                    "<font color='#2E7D32'><b>通过</b></font>",
                ],
                [
                    "第 2 次",
                    "MPC/PI 独立编译及 0.1 s 闭环仿真",
                    "有限数、SoC 边界、占空比边界、母线安全范围",
                    "<font color='#2E7D32'><b>通过</b></font>",
                ],
                [
                    "第 3 次",
                    "3 s 论文阶跃 + 1 s Mode III",
                    "环境/负载扰动、MPC/PI 指标、满电限功率、模式序列",
                    "<font color='#2E7D32'><b>通过，发现并修复 1 项轻微问题</b></font>",
                ],
                [
                    "修复后复核",
                    "全模型结构检查 + 闭环回归",
                    "修改后三条逻辑线、模式切换、正常模式回归",
                    "<font color='#2E7D32'><b>通过</b></font>",
                ],
            ],
            [27 * mm, 50 * mm, 128 * mm, 48 * mm],
        ),
        Spacer(1, 4 * mm),
        para("检查中发现并修复的问题", "h2"),
        table(
            [
                ["问题", "原因", "修改", "复核证据"],
                [
                    "SoC 达到 90% 时，模式编号曾短暂出现约一个控制周期的 Mode 5",
                    "Mode III 的模式输出使用了 Unit Delay 后的锁存值，当前周期置位信号要到下一采样才生效",
                    "将 Set_or_Hold_Mode_III 的当前有效信号直接送往 alpha、Any_Valid_Mode 和 Mode_III_Code；Unit Delay 仅保留锁存功能",
                    "修复后 0.28-0.36 s 模式集合仅为 {1,3}，Mode 5 样本数为 0；进入 Mode III 约 0.306 s",
                ],
            ],
            [49 * mm, 68 * mm, 79 * mm, 57 * mm],
        ),
        Spacer(1, 3 * mm),
        para(
            "<b>重要边界：</b>“检查通过”表示模型在报告列出的结构与工况范围内运行正常，"
            "并不等同于形式化证明覆盖所有参数组合。论文未公开的电池参数、ARIMA 阶次/系数、"
            "PI 完整实现及原始环境序列仍属于复现假设，已在第 11 节单独列出。",
            "note",
        ),
        PageBreak(),
    ]
)

# Contents and project map
story.extend(
    [
        para("报告内容", "h1"),
        table(
            [
                ["章节", "内容"],
                ["1", "模型总体结构与系统间连接"],
                ["2", "信号、状态、功率方向与数据记录"],
                ["3", "场景输入与 ARIMA 一步预测"],
                ["4", "六状态 PV/电池平均值对象"],
                ["5", "P&O 最大功率点跟踪"],
                ["6", "四模式功率管理与修复后的 Mode III 锁存"],
                ["7", "非线性 MPC 监督控制器"],
                ["8", "PI 基线控制器与根层控制器选择"],
                ["9", "参数、频率、采样时间及其确定依据"],
                ["10", "三次检查、仿真结果与判据"],
                ["11", "已知提示、复现假设、风险与改进建议"],
                ["附录", "模块类型词典、文件结构与复现实验命令"],
            ],
            [24 * mm, 225 * mm],
        ),
        Spacer(1, 5 * mm),
        para("项目关键文件", "h2"),
        table(
            [
                ["文件", "作用"],
                ["PV_Battery_MPC_Project.prj", "MATLAB Project 入口，注册模型、脚本、文档和结果文件"],
                ["models/PV_Battery_MPC.slx", "完整闭环 Simulink 主模型"],
                ["scripts/pvbatt_parameters.m", "论文参数、派生参数和复现假设的唯一集中入口"],
                ["scripts/pvbatt_initialize.m", "生成 PV 查表、场景、稳态初值和 nlmpc 对象"],
                ["scripts/pvbatt_state_transition.m", "MPC 内部 10 ms 离散非线性预测器"],
                ["scripts/run_paper_scenario.m", "通过 SimulationInput 运行指定场景及控制器"],
                ["results/validation_results.mat", "修复后 MPC、PI 和 Mode III 的压缩仿真结果"],
                ["docs/reproduction_notes.md", "论文参数映射、公式修正和复现边界"],
            ],
            [72 * mm, 177 * mm],
        ),
        PageBreak(),
    ]
)

# Architecture
story.extend(
    [
        para("1. 模型总体结构与系统间连接", "h1"),
        para(
            "根层共有 19 个块，包含 7 个功能子系统。全部层级合计约 209 个块。"
            "主闭环以实际扰动 w_actual、六状态 x_state、辅助量 aux、参考电压 Vpv_ref、"
            "模式总线 mode_bus 和两路占空比 u_duty 为主干。ControllerSelect 选择 MPC 或 PI，"
            "最终命令统一经过 0.02-0.95 的占空比保护后送入对象。",
        ),
        *figure(
            PROJECT_ROOT / "docs" / "model_overview.png",
            "图 1  主模型顶层结构及系统间信号连接",
            118 * mm,
        ),
        PageBreak(),
        para("1.1 系统连接关系", "h2"),
        table(
            [
                ["源系统/信号", "连接到", "作用"],
                [
                    "Scenario_Profiles / w_actual",
                    "ARIMA、Plant、Power Management、日志",
                    "提供实际温度 T、辐照度 G 和负载电阻 Rload",
                ],
                [
                    "ARIMA / w_forecast",
                    "MPC Supervisory Controller",
                    "将下一控制周期扰动预测作为 nlmpc 的 measured disturbance",
                ],
                [
                    "Plant / x_state",
                    "MPC、PI、P&O、Power Management、Scope/日志",
                    "提供 Vpv、iLpv、Vb、iLb、SoC、Vdc 六个状态",
                ],
                [
                    "Plant / aux",
                    "P&O、Power Management、Scope/日志",
                    "提供 Ipv、Ib、Ebat、Ppv、Pbat、Pload",
                ],
                [
                    "P&O / Vpv_ref",
                    "MPC、PI",
                    "正常模式下的 PV 电压参考值",
                ],
                [
                    "Power Management / mode_bus",
                    "MPC、日志",
                    "mode_bus=[alpha, mode_id]；alpha=1 表示 Mode III 限功率",
                ],
                [
                    "MPC / u_mpc 与 PI / u_pi",
                    "Controller Selector",
                    "依据 ControllerSelect 选择科研控制器或基线控制器",
                ],
                [
                    "Duty Limits / u_duty",
                    "Plant、Scope/日志",
                    "将实际施加占空比限制在 0.02-0.95",
                ],
            ],
            [58 * mm, 84 * mm, 107 * mm],
        ),
        Spacer(1, 5 * mm),
        para("1.2 总体闭环数据流", "h2"),
        para(
            "<b>扰动链：</b>场景数据 -&gt; 实际扰动 -&gt; ARIMA 一步预测 -&gt; MPC。<br/>"
            "<b>能量链：</b>PV/电池平均值对象 -&gt; DC 母线与负载。<br/>"
            "<b>参考链：</b>对象状态和 PV 电流 -&gt; P&O -&gt; Vpv_ref。<br/>"
            "<b>模式链：</b>SoC、可用 PV 功率、负载功率、辐照度 -&gt; Mode I-IV -&gt; MPC 在线权重和限功率层。<br/>"
            "<b>控制链：</b>MPC 或 PI -&gt; 控制器选择 -&gt; 占空比限幅 -&gt; 两路 Boost 平均值对象。",
            "code",
        ),
        PageBreak(),
    ]
)

# Signals/equations
story.extend(
    [
        para("2. 信号、状态、功率方向与数据记录", "h1"),
        table(
            [
                ["向量", "定义", "单位/符号"],
                ["x_state", "[Vpv, iLpv, Vb, iLb, SoC, Vdc]", "V, A, V, A, 1, V"],
                ["u_duty", "[dpv, db]", "两路 Boost 占空比，0-1"],
                ["w_actual / w_forecast", "[T, G, Rload]", "摄氏度, W/m², ohm"],
                ["aux", "[Ipv, Ib, Ebat, Ppv, Pbat, Pload]", "A, A, V, W, W, W"],
                ["mode_bus", "[alpha, mode_id]", "alpha 为 0/1；mode_id 为 1-5"],
                ["MPC 输出 y", "[Vdc, Vpv, iLb]", "对应模式相关跟踪目标"],
            ],
            [42 * mm, 112 * mm, 95 * mm],
        ),
        Spacer(1, 5 * mm),
        para("功率和电流正方向", "h2"),
        para(
            "Ppv=Vpv*Ipv 为 PV 输出功率；Pbat=(1-db)*Vdc*iLb，Pbat&gt;0 表示电池向母线放电，"
            "Pbat&lt;0 表示电池充电；Pload=Vdc²/Rload。SoC 微分采用 -Ib/(3600Q)，"
            "因此电池放电电流为正时 SoC 下降。",
        ),
        para("数据记录", "h2"),
        table(
            [
                ["块", "工作区变量", "内容"],
                ["State_Log", "sim_x", "六状态 x_state"],
                ["Aux_Log", "sim_aux", "电流、电动势和三类功率"],
                ["Duty_Log", "sim_u", "最终施加的 dpv、db"],
                ["Mode_Log", "sim_mode", "alpha 和 mode_id"],
                ["Disturbance_Log", "sim_w", "实际环境和负载输入"],
                ["Scope", "-", "模型运行时快速观察状态、功率和占空比"],
            ],
            [45 * mm, 45 * mm, 159 * mm],
        ),
        Spacer(1, 4 * mm),
        para(
            "To Workspace 块采用 Structure With Time，日志每 100 个 10 us 步长保存一次，"
            "即结果时间分辨率约 1 ms；对象仍按 10 us 固定步长积分。",
            "note",
        ),
        PageBreak(),
    ]
)

# Scenario and ARIMA
story.extend(
    [
        para("3. 场景输入与 ARIMA 一步预测", "h1"),
        para("3.1 Scenario_Profiles", "h2"),
        *figure(
            ASSET_DIR / "scenario_profiles.png",
            "图 2  场景输入子系统：From Workspace 输出 w_actual",
            75 * mm,
        ),
        table(
            [
                ["模块", "设置/输入", "作用"],
                ["From Workspace", "scenario_w，采样时间 1e-5 s", "输出 [T,G,Rload]；允许论文阶跃、四模式和合成实测场景复用同一接口"],
                ["Outport", "w_actual", "将扰动广播到对象、预测器和功率管理"],
            ],
            [50 * mm, 70 * mm, 129 * mm],
        ),
        Spacer(1, 4 * mm),
        *figure(
            ASSET_DIR / "comparison_disturbance_profile.png",
            "图 3  论文第 5.5 节公开阶跃：1 s 时温度/辐照度变化，2 s 时负载加倍",
            94 * mm,
        ),
        PageBreak(),
        para("3.2 ARIMA_OneStep_Forecast", "h2"),
        *figure(
            ASSET_DIR / "arima_forecast.png",
            "图 4  温度、辐照度和负载电阻的并行 ARIMA 一步预测结构",
            105 * mm,
        ),
        table(
            [
                ["模块", "作用"],
                ["Demux", "把 w_actual 分成 T、G、Rload 三条并行通道"],
                ["Zero-Order Hold", "以 0.01 s 控制周期采样并保持实际扰动"],
                ["Unit Delay", "保存上一采样值 w(k-1)"],
                ["Sum", "计算增量 delta_w=w(k)-w(k-1)，并构造预测值"],
                ["Gain", "施加 AR 系数 phi"],
                ["Mux", "重新组合 w_forecast=[T_hat,G_hat,Rload_hat]"],
            ],
            [62 * mm, 187 * mm],
        ),
        para(
            "<b>实现公式：</b>w_hat(k+1)=w(k)+phi*[w(k)-w(k-1)]。论文说明采用 ARIMA，"
            "但未公开阶次、系数或训练序列，因此 phi_T=phi_G=phi_R=0，"
            "当前等效为稳健的 ARIMA(0,1,0) 持久性预测。以后获得实测历史后，可把 phi 改为标定值，"
            "无需改变模型结构。",
        ),
        PageBreak(),
    ]
)

# Plant
story.extend(
    [
        para("4. 六状态 PV/电池平均值对象", "h1"),
        para(
            "PV_Battery_Averaged_Plant 是科研对象核心，包含 42 个普通 Simulink 块。"
            "它不模拟 10 kHz 开关器件的每次导通/关断，而使用占空比平均值方程，"
            "适合 MPC 长时间闭环研究并与论文建模层级一致。",
        ),
        *figure(
            ASSET_DIR / "averaged_plant.png",
            "图 5  六状态平均值对象内部结构",
            125 * mm,
        ),
        PageBreak(),
        para("4.1 状态方程与模块对应关系", "h2"),
        table(
            [
                ["状态方程", "Simulink 构建", "物理作用"],
                [
                    "dVpv/dt=(Ipv-iLpv)/Cpv",
                    "PV_Current_Lookup - PV_Capacitor_KCL - Inv_Cpv - PV_Voltage_State",
                    "PV 输入电容电荷平衡",
                ],
                [
                    "diLpv/dt=[Vpv-rpv*iLpv-(1-dpv)Vdc]/Lpv",
                    "PV_Inductor_R_Drop、PV_Bus_Voltage_Term、PV_Inductor_KVL、Inv_Lpv、Integrator",
                    "PV Boost 电感伏秒平衡",
                ],
                [
                    "dVb/dt=(Ib-iLb)/Cb",
                    "Battery_OCV_Lookup、Battery_Resistance、Battery_Capacitor_KCL、Inv_Cb、Integrator",
                    "电池端口电容动态",
                ],
                [
                    "diLb/dt=[Vb-rb*iLb-(1-db)Vdc]/Lb",
                    "Battery_Inductor_R_Drop、Battery_Bus_Voltage_Term、KVL、Inv_Lb、Integrator",
                    "双向电池 Boost 电感动态",
                ],
                [
                    "dSoC/dt=-Ib/(3600Q)",
                    "SoC_Coulomb_Counting + Battery_SoC_State",
                    "库仑计数；Q 使用 Ah，3600 完成秒到小时换算",
                ],
                [
                    "dVdc/dt=[(1-dpv)iLpv+(1-db)iLb-Vdc/Rload]/Cdc",
                    "PV_Bus_Current、Battery_Bus_Current、Load_Current、DC_Bus_KCL、Inv_Cdc、Integrator",
                    "DC 母线电流平衡",
                ],
            ],
            [76 * mm, 112 * mm, 61 * mm],
            font_style="tiny",
        ),
        Spacer(1, 4 * mm),
        para(
            "<b>论文公式修正：</b>论文式 (12) 的 SoC 行写成与 x4 相减，但根据论文式 (6)、式 (9) "
            "和量纲，电池电流应为 Ib=(Ebat-Vb)/Rbat，因此项目采用 Vb=x3。"
            "这是公式一致性修正，不是任意改变。",
            "note",
        ),
        para("4.2 查表和辅助功率", "h2"),
        table(
            [
                ["模块", "输入", "输出/作用"],
                ["PV_Current_Lookup (3-D)", "Vpv、T、G", "由离线单二极管模型生成 Ipv；仿真中只做标准查表"],
                ["Battery_OCV_Lookup (1-D)", "SoC", "得到 Ebat，并通过 Rbat 计算端口电流"],
                ["Product", "电压、电流、(1-duty)", "计算母线电流、Ppv、Pbat、Pload"],
                ["Mux", "六个状态或六个辅助量", "形成 x_state 和 aux 固定接口"],
                ["Integrator", "六个状态导数", "连续状态积分，初值来自稳态工作点求解"],
            ],
            [62 * mm, 68 * mm, 119 * mm],
        ),
        PageBreak(),
        para("4.3 PV 与电池特性假设", "h2"),
        Table(
            [
                [
                    scaled_image(
                        ASSET_DIR / "pv_lookup_curves.png",
                        120 * mm,
                        96 * mm,
                    ),
                    scaled_image(
                        ASSET_DIR / "battery_ocv_curve.png",
                        120 * mm,
                        96 * mm,
                    ),
                ]
            ],
            colWidths=[124 * mm, 124 * mm],
        ),
        Table(
            [
                [
                    para("图 6  PV 三维查表切片：I-V 与 P-V 曲线", "caption"),
                    para("图 7  可替换 OCV-SoC 查表及 20%/90% 阈值", "caption"),
                ]
            ],
            colWidths=[124 * mm, 124 * mm],
        ),
        Spacer(1, 3 * mm),
        para(
            "PV 曲线由论文公开的 Isc、Voc、Imp、Vmp 以及 MPP 斜率条件拟合单二极管参数后离线生成。"
            "电池 E0、K、A、B、Rbat 未在论文中公开，因此使用 270-312 V 的可替换 OCV 表和 0.08 ohm 内阻；"
            "这部分是与作者原始结果不能逐点一致的主要原因。",
        ),
        PageBreak(),
    ]
)

# P&O
story.extend(
    [
        para("5. P&O 最大功率点跟踪", "h1"),
        *figure(
            ASSET_DIR / "po_mppt.png",
            "图 8  P_and_O_MPPT 内部普通模块结构",
            117 * mm,
        ),
        table(
            [
                ["处理步骤", "模块", "作用"],
                ["采样", "Demux + Zero-Order Hold", "每 10 ms 取得 Vpv 和 Ipv"],
                ["功率", "Product", "P(k)=Vpv(k)*Ipv(k)"],
                ["差分", "Unit Delay + Sum", "delta_P=P(k)-P(k-1)，delta_V=V(k)-V(k-1)"],
                ["方向判定", "Product + Relational Operator", "delta_P*delta_V>=0 时增加参考电压，否则减少"],
                ["死区", "Abs + Relational Operator", "|delta_P|<1 W 时保持，减少 MPP 附近抖动"],
                ["扰动", "Constants + Switch", "选择 +0.25 V 或 -0.25 V"],
                ["记忆与边界", "Unit Delay + Sum + Saturate", "累加 Vpv_ref，并限制在 180-315 V"],
            ],
            [38 * mm, 78 * mm, 133 * mm],
        ),
        Spacer(1, 4 * mm),
        para(
            "0.25 V 步长用于兼顾跟踪速度和稳态波动；范围 180-315 V 包含 9 串组件的标称 MPP "
            "261 V，并留出温度与辐照度变化余量。P&O 只生成参考值，不直接输出占空比；"
            "实际占空比由 MPC 或 PI 根据对象动态计算。",
        ),
        PageBreak(),
    ]
)

# PMS
story.extend(
    [
        para("6. 四模式功率管理与 Mode III 锁存", "h1"),
        *figure(
            ASSET_DIR / "power_management.png",
            "图 9  修复后的 Power_Management_Logic 内部结构",
            120 * mm,
        ),
        PageBreak(),
        para("6.1 模式判据", "h2"),
        table(
            [
                ["模式", "逻辑条件", "控制目的"],
                ["Mode I", "Pavailable>=Pload，SoC<90%，白天", "PV 按 MPPT 运行，剩余功率给电池充电"],
                ["Mode II", "Pavailable<Pload，SoC>20%，白天", "PV 按 MPPT 运行，电池补足负载缺口"],
                ["Mode III", "Pavailable>=Pload，SoC>=90%，白天；条件成立后锁存", "电池电流趋零，PV 离开 MPP 并限功率"],
                ["Mode IV", "G<=0.001 W/m²，SoC>20%", "夜间或无光时由电池供电"],
                ["Mode 5", "以上模式均不成立", "表示低 SoC 且无足够 PV 时的负载切除请求；论文未展开负载执行器"],
            ],
            [34 * mm, 106 * mm, 109 * mm],
        ),
        Spacer(1, 4 * mm),
        para("6.2 内部模块及作用", "h2"),
        table(
            [
                ["模块组", "作用"],
                ["Demux", "从 x_state 取 SoC，从 aux 取 Pload，从 w_actual 取 G"],
                ["Available_PV_Power_Estimate (Gain)", "Pavailable=G*Parray,rated/1000；用可用功率而不是已限功率后的 Ppv 判断是否剩余"],
                ["Relational Operator", "比较 PV 可用功率/负载、SoC 上下限、昼夜阈值"],
                ["Logic", "组合 Mode I-IV 条件，并形成 Set_or_Hold_Mode_III"],
                ["Unit Delay", "保存 Mode III 锁存状态，避免满电边界附近反复切换"],
                ["Gain + Sum", "把布尔模式乘以 1-5 后相加形成 mode_id"],
                ["Data Type Conversion + Mux", "输出 double 类型 mode_bus=[alpha,mode_id]"],
            ],
            [82 * mm, 167 * mm],
        ),
        para(
            "<b>本次修复：</b>Set_or_Hold_Mode_III 当前周期有效信号直接用于 alpha、有效模式判断和 Mode III "
            "编码；Unit Delay 只负责下一周期保持。这样既立即进入 Mode III，又保留锁存。",
            "note",
        ),
        PageBreak(),
    ]
)

# MPC
story.extend(
    [
        para("7. 非线性 MPC 监督控制器", "h1"),
        *figure(
            ASSET_DIR / "mpc_controller.png",
            "图 10  MPC_Supervisory_Controller 内部结构",
            119 * mm,
        ),
        PageBreak(),
        para("7.1 标准 Nonlinear MPC Controller", "h2"),
        table(
            [
                ["项目", "配置", "作用"],
                ["状态数 / 输出数", "6 / 3", "预测六状态；目标输出 y=[Vdc,Vpv,iLb]"],
                ["操纵量 MV", "dpv、db", "两路 Boost 占空比"],
                ["测量扰动 MD", "T、G、Rload", "把环境和负载预测送入状态转移"],
                ["采样周期", "0.01 s", "100 Hz 在线优化"],
                ["预测/控制时域", "1 / 1", "严格对应论文的单步预测"],
                ["求解器", "SQP，最多 20 次迭代", "求解非线性受约束优化；允许使用可行次优解"],
                ["MV 范围", "0-1；变化率 +/-0.05/周期", "限制占空比及控制动作突变"],
                ["SoC 预测约束", "0.1995-0.9005", "围绕论文 20%-90% 阈值留 0.05% 数值保护带"],
            ],
            [52 * mm, 78 * mm, 119 * mm],
        ),
        Spacer(1, 4 * mm),
        para("7.2 目标函数与模式相关权重", "h2"),
        table(
            [
                ["模式", "输出权重 [Vdc,Vpv,iLb]", "MV 变化率权重", "含义"],
                ["正常 Mode I/II/IV", "[0.75,0.15,0]", "[0.10,0.10]", "优先稳定 600 V 母线，同时跟踪 P&O 电压"],
                ["Mode III", "[0.75,0,0.15]", "[0.10,0.10]", "母线稳定且电池电流归零，不再强制跟踪 MPP"],
            ],
            [46 * mm, 68 * mm, 55 * mm, 80 * mm],
        ),
        para(
            "权重 0.75/0.15/0.10/0.15 来自论文。采用单位 ScaleFactor，使这些权重直接作用于 V 和 A，"
            "而不是被自动工程量纲归一化。",
        ),
        PageBreak(),
        para("7.3 内部预测器和 Mode III 约束执行层", "h2"),
        para(
            "MPC 状态函数每 10 ms 调用一次，内部使用 10 个 1 ms RK4 子步。"
            "对象中的电池端口电容时间常数 Rbat*Cb=24 us，远小于 10 ms 控制周期；"
            "若在单步预测器内直接显式积分该快状态会数值发散，因此预测器令 "
            "Vb=Eoc(SoC)-Rbat*iLb 为准稳态值，而实际对象仍保留全部六个连续状态并按 10 us 积分。",
        ),
        table(
            [
                ["Mode III 标准模块", "作用"],
                ["Curtailment_State_Split", "提取 iLb、SoC 和 Vdc"],
                ["Battery_Zero_Current_Increment", "delta_db=-0.0002*iLb，使电池电流趋零"],
                ["Curtailment_DC_Error", "计算 600-Vdc"],
                ["PV_Curtailment_Increment", "delta_dpv=0.0002*(600-Vdc)，通过 PV 侧占空比调节母线能量"],
                ["Previous_Duty + Sum", "在上一控制量基础上增量更新"],
                ["Saturate", "将限功率层输出限制到 0.02-0.95"],
                ["Switch", "alpha=1 使用限功率执行层，否则使用 nlmpc 输出"],
            ],
            [76 * mm, 173 * mm],
        ),
        PageBreak(),
    ]
)

# PI
story.extend(
    [
        para("8. PI 基线控制器与根层控制器选择", "h1"),
        *figure(
            ASSET_DIR / "pi_controller.png",
            "图 11  PI_Baseline_Controller 内部结构",
            116 * mm,
        ),
        table(
            [
                ["回路", "误差", "参数", "输出"],
                ["PV 电压 PI", "Vpv-Vpv_ref", "Kp=0.002，Ki=0.20 s^-1", "加到稳态 dpv 偏置"],
                ["DC 母线 PI", "600-Vdc", "Kp=0.001，Ki=0.05 s^-1", "加到稳态 db 偏置"],
            ],
            [50 * mm, 70 * mm, 70 * mm, 59 * mm],
        ),
        Spacer(1, 4 * mm),
        para(
            "PI 积分器由 Unit Delay、误差乘 Ki*0.01 和 Sum 构成离散累加器。"
            "ControllerSelect=1 时使能 MPC；其逻辑非值使能 PI。Switch 在两组占空比之间选择，"
            "未选控制器停止执行以节省仿真时间。最终 Duty Limits 对两种控制器执行同一保护。",
        ),
        para(
            "论文给出了控制对照思路但没有公开完整 PI 实现，因此本 PI 结构是透明、可复核的重构基线，"
            "不应被解释为作者原始控制器的逐块复制。",
            "note",
        ),
        PageBreak(),
    ]
)

# Parameters
story.extend(
    [
        para("9. 参数、频率、采样时间及其确定依据", "h1"),
        para("9.1 电力级与仿真参数", "h2"),
        table(
            [
                ["参数", "数值", "来源/确定方法", "作用"],
                ["Cpv, Cb", "300 uF", "论文公开", "平滑 PV/电池端口电压；决定端口快动态"],
                ["Lpv, Lb", "10 mH", "论文公开", "平滑两路 Boost 电流；决定电感动态"],
                ["rpv, rb", "10 mOhm", "论文公开", "电感串联损耗和阻尼"],
                ["Cdc", "1500 uF", "论文公开", "储存母线能量并抑制负载阶跃电压波动"],
                ["Vdc_ref", "600 V", "论文公开", "MPC/PI 的首要母线目标"],
                ["fsw", "10 kHz", "论文公开", "物理变换器开关频率；平均值模型不逐开关仿真"],
                ["固定步长", "10 us", "论文公开", "100 kHz 对象积分；每开关周期 10 步"],
                ["求解器", "ode4", "复现实现", "固定步长四阶 Runge-Kutta，兼顾精度和确定性"],
                ["占空比", "优化 0-1；施加 0.02-0.95", "论文范围 + 数值保护", "避免平均值方程在极限占空比附近退化"],
            ],
            [44 * mm, 42 * mm, 67 * mm, 96 * mm],
            font_style="tiny",
        ),
        Spacer(1, 4 * mm),
        para("9.2 PV 参数", "h2"),
        table(
            [
                ["参数", "数值", "来源/作用"],
                ["单模块 Pmp / Vmp / Imp", "213.15 W / 29 V / 7.35 A", "论文公开的 MPP 额定点"],
                ["单模块 Voc / Isc", "36.3 V / 7.84 A", "论文公开的开路/短路边界"],
                ["阵列", "9 串 x 5 并，45 块，9.59175 kW", "论文只给 9.5 kW；选择最接近且母线/MPP 合理的整数布局"],
                ["电池片数", "60/模块", "晶硅组件典型假设，用于单二极管温度关系"],
                ["二极管理想因子", "1.1395082016", "由 Isc/Voc/Imp/Vmp 与 MPP 斜率条件拟合"],
                ["Rs / Rsh", "0.327685 ohm / 4972.0696 ohm", "单二极管拟合参数"],
                ["Iph_STC / I0_STC", "7.84000085 A / 8.3047e-9 A", "单二极管拟合参数"],
                ["短路电流温度系数", "0.0005*Isc / C", "论文未给，采用典型晶硅假设"],
                ["带隙", "1.12 eV", "硅材料典型值"],
            ],
            [62 * mm, 70 * mm, 117 * mm],
            font_style="tiny",
        ),
        PageBreak(),
        para("9.3 电池、MPPT、预测与基线参数", "h2"),
        table(
            [
                ["类别/参数", "数值", "来源", "作用"],
                ["电池额定电压", "300 V", "论文公开", "电池侧电压等级"],
                ["容量", "20 Ah，约 6 kWh", "论文公开", "决定 SoC 变化速度和储能规模"],
                ["SoC 范围", "20%-90%", "论文公开", "防止过放/过充并决定模式切换"],
                ["Rbat", "0.08 ohm", "复现假设", "端口压降和电池电流动态；论文未公开"],
                ["OCV 表", "270-312 V", "复现假设", "以可替换查表代替未公开 E0/K/A/B 参数"],
                ["P&O 周期", "10 ms / 100 Hz", "等于 MPC 控制周期", "参考更新与控制同步"],
                ["P&O 步长", "0.25 V", "复现调谐", "跟踪速度与 MPP 附近波动折中"],
                ["P&O 范围", "180-315 V", "阵列 9 串额定点派生", "覆盖正常 MPP 并防止不合理参考"],
                ["功率死区", "1 W", "复现调谐", "抑制数值噪声导致的方向反转"],
                ["ARIMA phi", "T/G/R 均为 0", "论文未公开", "当前为一步持久性预测，接口保留待标定"],
                ["PV PI", "Kp=0.002, Ki=0.20", "透明基线重构", "控制 Vpv 跟踪 P&O 参考"],
                ["DC PI", "Kp=0.001, Ki=0.05", "透明基线重构", "控制 Vdc 跟踪 600 V"],
            ],
            [58 * mm, 60 * mm, 54 * mm, 77 * mm],
            font_style="tiny",
        ),
        Spacer(1, 5 * mm),
        para("9.4 多时间尺度关系", "h2"),
        table(
            [
                ["层级", "周期/频率", "相互关系", "为什么这样设置"],
                ["对象积分", "10 us / 100 kHz", "每 10 kHz 开关周期 10 个积分步", "论文给定；解析平均值对象的快速 LC/RC 动态"],
                ["物理开关", "100 us / 10 kHz", "每个 10 ms 控制周期含 100 个开关周期", "论文给定的变换器频率"],
                ["MPC / PI / P&O / ARIMA", "10 ms / 100 Hz", "每控制周期含 1000 个对象积分步", "论文给定控制周期；降低在线优化负担"],
                ["MPC 内部 RK4", "1 ms x 10 子步", "覆盖一个 10 ms 预测步", "避免单个大步长预测造成非线性数值误差"],
                ["日志", "1 ms / 1 kHz", "每 100 个对象步记录一次", "在保留动态细节的同时控制结果文件体积"],
                ["电池端口 RC", "24 us", "远小于 10 ms 控制周期", "解释 MPC 预测器采用 Vb 准稳态约束的必要性"],
            ],
            [48 * mm, 52 * mm, 72 * mm, 77 * mm],
        ),
        Spacer(1, 4 * mm),
        para(
            "<b>频率说明：</b>模型采用平均值变换器，因此 fsw=10 kHz 是被复现系统的物理参数，"
            "不是模型里 PWM 脉冲块的运行频率。10 us 步长仍按照论文保留，以正确覆盖对象的电容、电感快动态。",
            "note",
        ),
        PageBreak(),
    ]
)

# Audits
story.extend(
    [
        para("10. 三次检查、仿真结果与判据", "h1"),
        para("10.1 第一次：结构连接检查", "h2"),
        table(
            [
                ["检查项", "范围", "结果"],
                ["Unconnected Ports", "根层及全部子系统", "0 个错误"],
                ["Unconnected Lines", "根层及全部子系统", "0 个错误"],
                ["Stateflow Lint", "模型内状态图检查", "无问题；模型实际未使用 Stateflow"],
                ["MATLAB Function 块", "逐层模型读取", "0 个"],
                ["模型状态", "Simulink model_check", "healthy"],
            ],
            [66 * mm, 96 * mm, 87 * mm],
        ),
        Spacer(1, 5 * mm),
        para("10.2 第二次：MPC/PI 编译与 0.1 s 闭环检查", "h2"),
        table(
            [
                ["控制器", "Vdc 范围", "SoC 范围", "dpv 范围", "db 范围", "结果"],
                ["MPC", "599.990241-600.009367 V", "0.79998780-0.80000000", "0.586775-0.586781", "0.501285-0.501349", "通过"],
                ["PI", "599.981459-600.015222 V", "0.79998780-0.80000000", "0.586771-0.586784", "0.501308-0.501334", "通过"],
            ],
            [30 * mm, 61 * mm, 58 * mm, 42 * mm, 42 * mm, 24 * mm],
        ),
        para(
            "自动判据：所有状态和占空比必须为有限数；SoC 必须在 0-1；施加占空比必须在 "
            "0.02-0.95；Vdc 必须处于 500-700 V 的短时安全包络。两种控制器全部满足。",
        ),
        PageBreak(),
        para("10.3 第三次：论文 3 s 环境/负载阶跃", "h2"),
        *figure(
            RESULT_DIR / "comparison_mpc_vs_pi.png",
            "图 12  论文对照工况下 MPC 与 PI 的 DC 母线响应",
            112 * mm,
        ),
        table(
            [
                ["控制器", "平均 VRI", "最大 VRI", "Vdc 范围", "结论"],
                ["MPC", "0.121544%", "3.945106%", "576.329-618.158 V", "峰值偏差较小"],
                ["PI", "0.110473%", "5.989661%", "564.062-625.628 V", "平均偏差略小，峰值较大"],
            ],
            [38 * mm, 42 * mm, 42 * mm, 74 * mm, 53 * mm],
        ),
        para(
            "MPC 占空比：dpv=0.554957-0.604957，db=0.461831-0.548815；"
            "PI 占空比：dpv=0.551236-0.672050，db=0.475692-0.534753。"
            "两者均未触及 0.02/0.95 保护边界。当前假设下 MPC 降低了最大 VRI，但平均 VRI 不优于 PI，"
            "因此报告不宣称 MPC 在所有指标上均胜出。",
        ),
        PageBreak(),
        para("10.4 Mode III 满电限功率与修复后复核", "h2"),
        *figure(
            RESULT_DIR / "mode3_curtailment.png",
            "图 13  修复后 Mode III：无 Mode 5 瞬时空档，PV 功率向负载需求收敛",
            105 * mm,
        ),
        table(
            [
                ["指标", "结果", "判据"],
                ["Mode III 首次进入", "约 0.306 s", "SoC 达到 90% 后应立即从 Mode I 进入 Mode III"],
                ["切换窗口模式集合", "{1,3}", "不得出现瞬时 Mode 5"],
                ["1 s 末 Ppv", "5.916 kW", "与约 6.07 kW 负载接近，误差小于 0.5 kW"],
                ["1 s 末 Pbat", "0.156 kW", "绝对值小于 0.5 kW，向零收敛"],
                ["1 s 末 Vdc", "603.602 V", "保持稳定，仍存在阻尼暂态"],
                ["占空比/状态", "有限且在边界内", "无 NaN、Inf 或占空比越界"],
            ],
            [62 * mm, 64 * mm, 123 * mm],
            font_style="tiny",
        ),
        PageBreak(),
    ]
)

# Limitations and recommendations
story.extend(
    [
        para("11. 已知提示、复现假设、风险与改进建议", "h1"),
        para("11.1 已知但非故障的提示", "h2"),
        table(
            [
                ["提示", "原因", "判断"],
                [
                    "Zero weights are applied to one or more OVs because there are fewer MVs than OVs.",
                    "模型有三个 OV 和两个 MV；正常模式将 iLb 权重置零，Mode III 将 Vpv 权重置零，以便在线切换目标。",
                    "预期提示。StateFcn、OutputFcn 和用户函数验证均通过。",
                ],
            ],
            [88 * mm, 105 * mm, 56 * mm],
        ),
        Spacer(1, 4 * mm),
        para("11.2 不能视为论文原值的假设", "h2"),
        table(
            [
                ["未公开内容", "项目处理", "对结果的影响"],
                ["PV 阵列串并联布局", "采用 9 串 x 5 并，额定 9.59175 kW", "额定功率略高于论文名义 9.5 kW"],
                ["单二极管 Rs/Rsh/n 与温度系数", "由公开四点拟合并补充晶硅典型值", "PV 曲线不能保证与作者查表逐点一致"],
                ["电池 E0/K/A/B/Rbat", "OCV-SoC 查表 + 0.08 ohm", "电池瞬态、SoC 与功率分配会影响 VRI"],
                ["ARIMA 阶次、系数和训练数据", "phi=0 的一步持久性预测", "当前不体现数据训练带来的预测改进"],
                ["PI 完整实现", "重构两个透明离散 PI 回路", "PI 对照指标仅对本基线有效"],
                ["Mode I-IV 原始数据序列", "按论文文字构造；comparison 严格采用公开阶跃", "四模式长期曲线不等于作者原始数据"],
            ],
            [67 * mm, 91 * mm, 91 * mm],
        ),
        Spacer(1, 4 * mm),
        para("11.3 模型适用范围和风险", "h2"),
        para(
            "本模型用于平均值级控制算法研究，不包含 PWM、半导体开关损耗、二极管反向恢复、"
            "电磁干扰、采样量化、传感器噪声、通信延迟、接触器、BMS 保护、温升和老化。"
            "因此不能直接作为硬件设计或安全认证依据。Mode 5 只输出负载切除请求，没有实际负载执行器。",
        ),
        para("11.4 建议的下一步", "h2"),
        table(
            [
                ["优先级", "建议", "价值"],
                ["高", "用论文作者数据或实测 T/G/Rload 历史标定 ARIMA 阶次和 phi", "验证论文核心“扰动预测”创新的真实增益"],
                ["高", "用目标电芯/电池包 HPPC 或脉冲数据辨识 OCV、R0、RC 支路", "缩小 VRI、SoC 和功率分配误差"],
                ["中", "把关键判据固化为 Simulink Test 回归用例", "今后修改模型时自动检测连接、边界和模式退化"],
                ["中", "加入开关级子模型进行短时验证", "检查电流纹波、PWM 饱和和器件应力"],
                ["低", "加入传感器噪声、预测误差和参数蒙特卡洛", "评估不确定性下的鲁棒性和统计置信度"],
            ],
            [28 * mm, 139 * mm, 82 * mm],
        ),
        PageBreak(),
    ]
)

# Appendix glossary
story.extend(
    [
        para("附录 A：Simulink 模块类型与作用词典", "h1"),
        table(
            [
                ["模块类型", "在本模型中的作用"],
                ["Inport / Outport", "定义每个子系统的固定输入输出接口"],
                ["From Workspace", "注入环境与负载场景矩阵"],
                ["To Workspace / Scope", "保存与显示状态、功率、模式和占空比"],
                ["Demux / Mux", "拆分和重新组合状态、扰动、输出及模式向量"],
                ["Sum", "实现 KCL、KVL、误差、差分、增量和模式编码"],
                ["Gain", "实现 1/C、1/L、电阻压降、PI 系数、AR 系数和模式编号"],
                ["Product", "乘法、除法、功率、(1-duty) 电流和方向指标"],
                ["Integrator", "积分六个连续状态导数"],
                ["Lookup Table n-D", "PV 的 V/T/G 三维电流表和电池 OCV-SoC 一维表"],
                ["Zero-Order Hold", "以 10 ms 周期采样状态、PV 量和扰动"],
                ["Unit Delay", "保存上一采样、积分状态、上一占空比和 Mode III 锁存"],
                ["Relational Operator", "判断功率方向、SoC 阈值、昼夜和 P&O 方向"],
                ["Logic", "组合 Mode I-IV、控制器使能和锁存条件"],
                ["Switch", "选择 P&O 方向、保持/更新、在线权重、MPC/PI 和 Mode III 输出"],
                ["Saturate", "限制 Vpv_ref 和最终占空比"],
                ["Abs", "形成 P&O 功率变化死区"],
                ["Reshape", "把参考、上一操纵量、扰动、在线权重转换为 nlmpc 所需行向量"],
                ["Data Type Conversion", "把布尔模式转换为 double 总线"],
                ["Enable Port", "只执行当前选中的 MPC 或 PI 子系统"],
                ["Terminator", "明确终止当前算法不使用的向量分量或状态输出"],
                ["Nonlinear MPC Controller", "标准 MPC Toolbox 在线非线性优化块"],
                ["Constant", "参考值、阈值、步长、权重、偏置和单位常数"],
            ],
            [68 * mm, 181 * mm],
            font_style="tiny",
        ),
        PageBreak(),
        para("附录 B：复现实验命令与检查记录", "h1"),
        para(
            "在 MATLAB 中双击 PV_Battery_MPC_Project.prj，随后运行：",
        ),
        para(
            "START_HERE<br/>"
            "resultMPC = run_paper_scenario(&quot;comparison&quot;,&quot;MPC&quot;);<br/>"
            "resultPI = run_paper_scenario(&quot;comparison&quot;,&quot;PI&quot;);<br/>"
            "resultMode3 = run_paper_scenario(&quot;mode3_quick&quot;,&quot;MPC&quot;);<br/>"
            "plot_paper_scenario(resultMPC);",
            "code",
        ),
        Spacer(1, 5 * mm),
        table(
            [
                ["检查证据文件", "内容"],
                ["results/model_audit_check2.mat", "修复前独立 MPC/PI 0.1 s 编译与边界检查数据"],
                ["results/model_audit_check3.mat", "3 s MPC/PI 与 Mode III 工况检查数据"],
                ["results/mode3_transition_fix_audit.mat", "Mode III 瞬时模式空档修复验证"],
                ["results/post_fix_regression.mat", "修复后正常模式短时回归"],
                ["results/validation_results.mat", "修复后最终验证曲线数据"],
            ],
            [88 * mm, 161 * mm],
        ),
        Spacer(1, 5 * mm),
        para("参考文献", "h2"),
        para(
            "S. Batiyah, R. Sharma, S. Abdelwahed, W. Alhosaini, and O. Aldosari, "
            "\"Predictive Control of PV/Battery System under Load and Environmental Uncertainty,\" "
            "Energies, vol. 15, no. 11, 4100, 2022. DOI: 10.3390/en15114100.",
        ),
        para(
            "报告生成依据：PV_Battery_MPC.slx 修复后保存版本、pvbatt_parameters.m、"
            "pvbatt_create_nlmpc.m、pvbatt_build_scenario.m、逐层 model_read/model_check "
            "结果及本次重新运行的 SimulationInput 仿真。",
            "small",
        ),
    ]
)


OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
document = SimpleDocTemplate(
    str(OUTPUT_PATH),
    pagesize=PAGE_SIZE,
    rightMargin=MARGIN_X,
    leftMargin=MARGIN_X,
    topMargin=MARGIN_TOP,
    bottomMargin=MARGIN_BOTTOM,
    title="PV/电池系统预测控制模型检查与系统说明报告",
    author="Codex",
    subject="PV_Battery_MPC Simulink model audit and system documentation",
)
document.build(story, onFirstPage=on_page, onLaterPages=on_page)
print(OUTPUT_PATH)
