#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""食品标签审核报告生成器（GB 7718-2025 / GB 28050-2025 / 食品标识监督管理办法）。

用法：
    python build_audit_report.py 审核数据.json -o 报告.xlsx

设计约束（WPS 兼容）：
  * 全部单元格写入静态值，不使用任何公式，避免 WPS 重算错位；
  * 不使用条件格式、表格对象、数据验证、透视表；
  * 仅用基础字体、填充、细边框、自动换行、列宽、行高、冻结首行。
"""

import argparse
import json
import math
import os
import sys
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------- 样式常量
FONT_NAME = "宋体"
THIN = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="E7E6E6")
TITLE_FONT = Font(name=FONT_NAME, size=14, bold=True)
SUB_FONT = Font(name=FONT_NAME, size=10, color="595959")
HDR_FONT = Font(name=FONT_NAME, size=10.5, bold=True)
BODY_FONT = Font(name=FONT_NAME, size=10)
SEV_FILL = {
    "严重": PatternFill("solid", fgColor="FBD5D5"),
    "一般": PatternFill("solid", fgColor="FDE9C8"),
    "瑕疵": PatternFill("solid", fgColor="DCE9F7"),
    "无法判定": PatternFill("solid", fgColor="EDEDED"),
}
BODY_ALIGN = Alignment(horizontal="left", vertical="top", wrap_text=True)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)

# ---------------------------------------------------------------- 表1 / 表A.1
# name, aliases, unit, rounding, zero_limit, nrv( None 表示无NRV规定 )
TABLE1 = [
    ("能量", ["能量", "热量"], "kJ", 1, 17, 8400),
    ("蛋白质", [], "g", 0.1, 0.5, 60),
    ("脂肪", [], "g", 0.1, 0.5, 60),
    ("饱和脂肪", ["饱和脂肪", "饱和脂肪酸"], "g", 0.1, 0.1, 20),
    ("反式脂肪酸", ["反式脂肪酸", "反式脂肪"], "g", 0.1, 0.3, None),
    ("单不饱和脂肪", ["单不饱和脂肪", "单不饱和脂肪酸"], "g", 0.1, 0.1, None),
    ("多不饱和脂肪", ["多不饱和脂肪", "多不饱和脂肪酸"], "g", 0.1, 0.1, None),
    ("n-3多不饱和脂肪酸", ["n-3多不饱和脂肪酸", "ω-3多不饱和脂肪酸", "n-3"], "mg", 1, 20, None),
    ("α-亚麻酸", ["α-亚麻酸", "a-亚麻酸"], "mg", 1, 5, None),
    ("二十碳五烯酸(EPA)", ["二十碳五烯酸", "EPA"], "mg", 1, 5, None),
    ("二十二碳六烯酸(DHA)", ["二十二碳六烯酸", "DHA"], "mg", 1, 5, None),
    ("胆固醇", [], "mg", 1, 5, None),
    ("碳水化合物", ["碳水化合物"], "g", 0.1, 0.5, 300),
    ("糖", ["糖", "糖类"], "g", 0.1, 0.5, None),
    ("乳糖", [], "g", 0.1, 0.5, None),
    ("膳食纤维", ["膳食纤维", "可溶性膳食纤维", "不溶性膳食纤维"], "g", 0.1, 0.5, 25),
    ("钠", [], "mg", 1, 5, 2000),
    ("维生素A", ["维生素A", "维生素 A", "视黄醇"], "μg RE", 1, 10, 800),
    ("维生素D", ["维生素D", "维生素 D"], "μg", 0.1, 0.1, 10),
    ("维生素E", ["维生素E", "维生素 E"], "mg α-TE", 0.01, 0.20, 14),
    ("维生素K", ["维生素K", "维生素 K"], "μg", 0.1, 1.6, 80),
    ("维生素B1", ["维生素B1", "维生素 B1", "硫胺素"], "mg", 0.01, 0.03, 1.4),
    ("维生素B2", ["维生素B2", "维生素 B2", "核黄素"], "mg", 0.01, 0.03, 1.4),
    ("维生素B6", ["维生素B6", "维生素 B6"], "mg", 0.01, 0.03, 1.4),
    ("维生素B12", ["维生素B12", "维生素 B12"], "μg", 0.01, 0.10, 2.4),
    ("维生素C", ["维生素C", "维生素 C", "抗坏血酸"], "mg", 0.1, 2.0, 100),
    ("烟酸", ["烟酸", "烟酰胺"], "mg", 0.1, 0.3, 14),
    ("叶酸", [], "μg", 1, 8, 400),
    ("泛酸", [], "mg", 0.01, 0.10, 5),
    ("生物素", [], "μg", 0.1, 0.6, 40),
    ("胆碱", [], "mg", 0.1, 9.0, 500),
    ("钙", [], "mg", 1, 8, 800),
    ("磷", [], "mg", 1, 14, 700),
    ("钾", [], "mg", 1, 20, 2000),
    ("镁", [], "mg", 1, 3, 300),
    ("铁", [], "mg", 0.1, 0.3, 15),
    ("锌", [], "mg", 0.01, 0.10, 11),
    ("碘", [], "μg", 0.1, 3.0, 120),
    ("硒", [], "μg", 0.1, 1.0, 60),
    ("铜", [], "mg", 0.01, 0.03, 0.8),
    ("氟", [], "mg", 0.01, 0.02, 1),
    ("锰", [], "mg", 0.01, 0.06, 3),
]

ALIAS_INDEX = {}
for row in TABLE1:
    ALIAS_INDEX[row[0].replace(" ", "")] = row
    for a in row[1]:
        ALIAS_INDEX[a.replace(" ", "")] = row


def find_spec(name):
    key = str(name).replace(" ", "").replace("\u3000", "")
    if key in ALIAS_INDEX:
        return ALIAS_INDEX[key]
    for alias, row in ALIAS_INDEX.items():
        if alias and (key.startswith(alias) or alias in key):
            return row
    return None


def num(v):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(",", ""))
    except ValueError:
        return None


def fmt_num(value):
    """数值显示：去掉无意义的尾随零，返回 int 或 float。"""
    if value is None:
        return "—"
    f = float(value)
    if abs(f - round(f)) < 1e-9:
        return int(round(f))
    return float(f"{f:.4f}".rstrip("0").rstrip("."))


def round_half_up(value, step=1):
    """按 四舍五入（ROUND_HALF_UP）修约，避免 Python round 的银行家舍入把 X.5 判错。"""
    d = Decimal(str(value))
    q = Decimal(str(step))
    return float((d / q).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * q)


def is_multiple(value, step):
    """用十进制精确判断是否为修约间隔的整数倍，避免 0.3/0.1 这类浮点误差。"""
    if value is None:
        return False
    try:
        d = Decimal(str(value))
        s = Decimal(str(step))
    except Exception:
        return False
    if s <= 0:
        return True
    return d % s == 0


def disp_width(text):
    return sum(2 if ord(ch) > 127 else 1 for ch in str(text))


def set_row_height(ws, row_idx, texts, widths, base=15.0, cap=409):
    lines = 1
    for text, width in zip(texts, widths):
        if text is None:
            continue
        usable = max(width - 1.5, 4)
        for seg in str(text).split("\n"):
            lines = max(lines, max(1, math.ceil(disp_width(seg) / usable)))
    ws.row_dimensions[row_idx].height = min(cap, base * lines + 4)


def write_header(ws, row_idx, headers, widths, height=30):
    for col, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=row_idx, column=col, value=h)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.border = BORDER
        cell.alignment = CENTER_ALIGN
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[row_idx].height = height


def page_setup(ws, landscape=True):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


# ---------------------------------------------------------------- 营养核对
def audit_nutrition(nut_block):
    """返回 (rows, issues)。rows 供表格，issues 为提示列表。"""
    rows, issues = [], []
    items = nut_block.get("items") or []
    values = {}
    declared_pct = {}
    for it in items:
        spec = find_spec(it.get("name", ""))
        value = num(it.get("value"))
        if spec is None:
            rows.append({
                "name": it.get("name", ""), "value": it.get("value"), "unit": it.get("unit", ""),
                "rounding": "-", "zero": "-", "round_ok": "未知成分", "zero_ok": "-",
                "nrv": "-", "pct_declared": it.get("nrv_pct", "-"), "pct_expect": "-",
                "verdict": "需人工确认", "note": "表中未列出的成分，须属 GB 14880 允许强化并在表1成分之后",
            })
            continue
        canonical, _, unit, rounding, zero, nrv = spec
        if canonical in values:
            issues.append(f"「{canonical}」重复标示")
        values[canonical] = value
        declared_pct[canonical] = num(it.get("nrv_pct"))

        round_ok, note = "—", []
        if value is None:
            round_ok = "无法判定"
            note.append("未填数值")
        elif value == 0:
            # 标 0 合法，不涉及小数位与修约间隔问题
            round_ok = "符合"
        else:
            if not is_multiple(value, rounding):
                round_ok = "不符"
                note.append(f"须为修约间隔 {rounding} 的整数倍")
            else:
                round_ok = "符合"

        zero_ok = "—"
        if value is None:
            zero_ok = "无法判定"
        elif zero <= 0:
            zero_ok = "—"
        elif value == 0:
            zero_ok = "符合"
        elif value <= zero:
            zero_ok = "不符"
            note.append(f"≤“0”界限值 {zero}，应标示为 0")

        if nrv is None:
            pct_expect, pct_verdict = "—", "—"
            if declared_pct[canonical] not in (None, 0, ""):
                pct_verdict = "不符"
                note.append("该成分无 NRV，NRV% 应空白/横线/斜线")
        elif value is None:
            pct_expect, pct_verdict = "无法计算", "无法判定"
        else:
            raw = value / nrv * 100
            base = int(round_half_up(raw, 1))
            pct_expect = base
            # 落在 X.5 边界时，四舍五入与 GB/T 8170“五留双”结论不同，两种都合规
            at_half = abs((raw - math.floor(raw)) - 0.5) < 1e-9
            allowed = {base, math.floor(raw), math.ceil(raw)} if at_half else {base}
            declared = declared_pct[canonical]
            if declared is None:
                pct_verdict = "未标示"
                note.append("强制/声称成分须标 NRV%")
            elif declared in allowed:
                pct_verdict = "符合"
            else:
                pct_verdict = "不符"
                note.append(f"应为 {base}%（{fmt_num(value)}/{nrv}×100，修约间隔1）")

        verdicts = [round_ok, zero_ok, pct_verdict]
        if "不符" in verdicts:
            verdict = "不符"
        elif "无法判定" in verdicts or "未标示" in verdicts:
            verdict = "需复核"
        else:
            verdict = "符合"

        rows.append({
            "name": canonical, "value": fmt_num(value) if value is not None else "—", "unit": unit, "rounding": rounding,
            "zero": f"≤{zero} {unit}", "round_ok": round_ok, "zero_ok": zero_ok,
            "nrv": f"{nrv} {unit}" if nrv is not None else "无NRV规定", "pct_declared": declared_pct[canonical] if declared_pct[canonical] is not None else "—",
            "pct_expect": pct_expect, "verdict": verdict, "note": "；".join(note),
        })
        if verdict == "不符":
            issues.append(f"{canonical}：{'；'.join(note)}")

    # 能量加和复核
    p, f, c, fib = values.get("蛋白质"), values.get("脂肪"), values.get("碳水化合物"), values.get("膳食纤维")
    energy = values.get("能量")
    if None not in (p, f, c) and energy is not None:
        expect = round_half_up(17 * p + 37 * f + 17 * c + (8 * fib if fib else 0), 1)
        diff = abs(expect - energy)
        ratio = diff / energy if energy else 1.0
        formula = f"17/37/17{'/8' if fib else ''}"
        if ratio <= 0.05:
            verdict = "符合"
            note = (f"按 {formula} 加和得 {expect} kJ，与标示 {fmt_num(energy)} kJ 差异 {ratio:.1%}，"
                    "在合理范围（可能按检测值标示或含糖醇、有机酸等其他供能成分）")
        elif ratio <= 0.20:
            verdict = "需复核"
            note = (f"按 {formula} 加和得 {expect} kJ，与标示 {fmt_num(energy)} kJ 差异 {ratio:.1%}；"
                    "属内部一致性提示，请核实碳水/膳食纤维取值，或确认是否含糖醇、有机酸、乙醇等供能成分")
            issues.append(f"【提示】能量：与加和值差异 {ratio:.1%}，非直接判罚项，需说明原因")
        else:
            verdict = "不符"
            note = (f"按 {formula} 加和得 {expect} kJ，与标示 {fmt_num(energy)} kJ 相差 {expect - energy:+.0f} kJ"
                    f"（{ratio:.1%}），难以由常规供能成分解释，须复核")
            issues.append(f"能量：{note}")
        rows.append({
            "name": "能量（加和复核）", "value": expect, "unit": "kJ", "rounding": "-",
            "zero": "-", "round_ok": "-", "zero_ok": "-", "nrv": "8400 kJ",
            "pct_declared": "—", "pct_expect": int(round_half_up(expect / 8400 * 100, 1)),
            "verdict": verdict, "note": note,
        })
    elif energy is not None:
        rows.append({
            "name": "能量（加和复核）", "value": "—", "unit": "kJ", "rounding": "-", "zero": "-",
            "round_ok": "-", "zero_ok": "-", "nrv": "8400 kJ", "pct_declared": "—",
            "pct_expect": "—", "verdict": "无法判定",
            "note": "蛋白质/脂肪/碳水化合物未全部提供，无法复算能量",
        })
    return rows, issues


# ---------------------------------------------------------------- 主流程
def build(data, out_path):
    meta = data.get("meta") or {}
    findings = data.get("findings") or []
    nutrition = data.get("nutrition") or {}
    checklist = data.get("checklist") or []
    tasks = data.get("tasks") or []

    # 营养复算前移：结论需要把营养标签问题一并计入，避免“有 7 处营养问题却写未发现不符合项”
    nut_rows, nut_issues = audit_nutrition(nutrition)
    hard_nut = [x for x in nut_issues if not x.startswith("【提示】")]

    order = {"严重": 0, "一般": 1, "瑕疵": 2, "无法判定": 3}
    findings = sorted(findings, key=lambda x: order.get(str(x.get("severity", "")).strip(), 9))

    counts = {k: 0 for k in order}
    for f in findings:
        sev = str(f.get("severity", "")).strip()
        if sev in counts:
            counts[sev] += 1

    conclusion = meta.get("conclusion")
    if not conclusion:
        if counts["严重"]:
            conclusion = "不合格"
        elif counts["一般"]:
            conclusion = "存在不符合项，需整改"
        elif counts["瑕疵"]:
            conclusion = "存在标签瑕疵"
        elif findings or hard_nut:
            conclusion = "存在需关注事项"
        else:
            conclusion = "未发现不符合项"

    wb = Workbook()

    # ---------------- 表1 审核概览
    ws = wb.active
    ws.title = "审核概览"
    widths = [18, 26, 18, 26]
    for col, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.merge_cells("A1:D1")
    t = ws.cell(row=1, column=1, value="预包装食品标签合规审核报告")
    t.font = TITLE_FONT
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 34

    ws.merge_cells("A2:D2")
    s = ws.cell(row=2, column=1, value="依据 GB 7718-2025、GB 28050-2025、《食品标识监督管理办法》（2027-03-16 施行）")
    s.font = SUB_FONT
    s.alignment = Alignment(horizontal="center", vertical="center")

    def kv_rows(start, pairs):
        r = start
        for i in range(0, len(pairs), 2):
            for j in range(2):
                if i + j >= len(pairs):
                    break
                label, value = pairs[i + j]
                lc = ws.cell(row=r, column=1 + j * 2, value=label)
                lc.font = HDR_FONT
                lc.fill = HDR_FILL
                lc.border = BORDER
                lc.alignment = CENTER_ALIGN
                vc = ws.cell(row=r, column=2 + j * 2, value=value)
                vc.font = BODY_FONT
                vc.border = BORDER
                vc.alignment = BODY_ALIGN
            ws.row_dimensions[r].height = 22
            r += 1
        return r

    pairs = [
        ("产品名称", meta.get("product_name", "—")),
        ("规格/净含量", meta.get("spec", "—")),
        ("标签类型", meta.get("label_type", "直接提供给消费者的预包装食品")),
        ("产品类别", meta.get("category", "—")),
        ("审核日期", meta.get("audit_date") or date.today().isoformat()),
        ("审核人", meta.get("auditor", "—")),
        ("图片/资料", "、".join(meta.get("source_images") or []) or "—"),
        ("最大表面面积", f"{meta.get('max_surface_area_cm2')} cm²" if meta.get("max_surface_area_cm2") else "未提供（需实测）"),
    ]
    r = kv_rows(3, pairs)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1)
    c = ws.cell(row=r, column=1, value="审核结论")
    c.font = HDR_FONT
    c.fill = HDR_FILL
    c.border = BORDER
    c.alignment = CENTER_ALIGN
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    v = ws.cell(row=r, column=2, value=conclusion)
    v.font = Font(name=FONT_NAME, size=12, bold=True)
    v.border = BORDER
    v.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    fill = {"不合格": SEV_FILL["严重"], "存在不符合项，需整改": SEV_FILL["一般"],
            "存在标签瑕疵": SEV_FILL["瑕疵"]}.get(conclusion, PatternFill("solid", fgColor="E2EFDA"))
    v.fill = fill
    ws.row_dimensions[r].height = 26
    r += 1

    stat_pairs = [
        ("严重（可能按食安法 125/124 条处罚）", counts["严重"]),
        ("一般（责令改正类）", counts["一般"]),
        ("瑕疵（办法第 39 条情形）", counts["瑕疵"]),
        ("无法判定（需实测/检测）", counts["无法判定"]),
        ("合计问题项", len(findings)),
        ("核查底稿条目", len(checklist)),
    ]
    r = kv_rows(r, stat_pairs)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1)
    c = ws.cell(row=r, column=1, value="审核摘要")
    c.font = HDR_FONT
    c.fill = HDR_FILL
    c.border = BORDER
    c.alignment = CENTER_ALIGN
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
    v = ws.cell(row=r, column=2, value=meta.get("summary", "—"))
    v.font = BODY_FONT
    v.border = BORDER
    v.alignment = BODY_ALIGN
    set_row_height(ws, r, [meta.get("summary", "—")], [26 + 18 + 26], base=15)
    r += 1

    notes = meta.get("confidence_notes")
    if notes:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1)
        c = ws.cell(row=r, column=1, value="识别说明")
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.border = BORDER
        c.alignment = CENTER_ALIGN
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        v = ws.cell(row=r, column=2, value=notes)
        v.font = BODY_FONT
        v.border = BORDER
        v.alignment = BODY_ALIGN
        set_row_height(ws, r, [notes], [70], base=15)
        r += 1

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    d = ws.cell(row=r, column=1, value=(
        "免责边界：本报告基于上传图片与委托方提供信息作合规性文件审查，不替代营养成分检测、标签字符高度实测、"
        "产品执行标准与生产许可真实性核查、商标/专利/广告用语合法性审查。图片无法辨识之处已列入“无法判定”，未计入不合格项。"))
    d.font = Font(name=FONT_NAME, size=9, color="7F7F7F")
    d.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[r].height = 42
    page_setup(ws, landscape=False)

    # ---------------- 表2 不合格项清单
    ws2 = wb.create_sheet("不合格项清单")
    headers = ["序号", "编号", "严重程度", "项目类别", "所在位置", "标签现状（原文）",
               "问题描述", "违反条款", "判定依据", "整改建议", "复核方式", "优先级"]
    widths2 = [6, 8, 10, 14, 14, 26, 28, 22, 32, 32, 18, 9]
    write_header(ws2, 1, headers, widths2)
    ws2.freeze_panes = "A2"

    row_i = 2
    if not findings:
        ws2.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
        empty_text = "本次未发现不符合项。"
        if hard_nut:
            empty_text = (f"标签文字项未发现不符合项；营养成分表另有 {len(hard_nut)} 项问题，"
                          "逐项判定见「营养成分核对」表。")
        c = ws2.cell(row=2, column=1, value=empty_text)
        c.font = BODY_FONT
        c.border = BORDER
        c.alignment = CENTER_ALIGN
        ws2.row_dimensions[2].height = 24
    for idx, f in enumerate(findings, start=1):
        vals = [
            idx,
            f.get("id", f"F{idx:02d}"),
            str(f.get("severity", "")).strip() or "一般",
            f.get("category", "—"),
            f.get("location", "—"),
            f.get("current", "—"),
            f.get("issue", "—"),
            f.get("clause", "—"),
            f.get("basis", "—"),
            f.get("rectification", "—"),
            f.get("verify", "—"),
            f.get("priority", "—"),
        ]
        for col, val in enumerate(vals, start=1):
            cell = ws2.cell(row=row_i, column=col, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = CENTER_ALIGN if col in (1, 2, 3, 12) else BODY_ALIGN
        sev = vals[2]
        if sev in SEV_FILL:
            ws2.cell(row=row_i, column=3).fill = SEV_FILL[sev]
        set_row_height(ws2, row_i, vals, widths2)
        row_i += 1
    page_setup(ws2)

    # ---------------- 表3 整改任务表
    ws3 = wb.create_sheet("整改任务表")
    headers3 = ["序号", "问题编号", "严重程度", "整改事项", "建议修改文案/动作", "责任岗位", "完成时限", "复核方式", "状态"]
    widths3 = [6, 9, 10, 30, 44, 14, 14, 22, 10]
    write_header(ws3, 1, headers3, widths3)
    ws3.freeze_panes = "A2"

    task_source = tasks or [
        {
            "ref": f.get("id", f"F{i:02d}"),
            "severity": str(f.get("severity", "")).strip() or "一般",
            "action": f.get("issue", ""),
            "text": f.get("suggested_text") or f.get("rectification", ""),
            "owner": f.get("owner", ""),
            "deadline": f.get("deadline", "") or ("立即" if str(f.get("severity")) == "严重" else "改版前"),
            "verify": f.get("verify", ""),
        }
        for i, f in enumerate(findings, start=1)
    ]

    row_i = 2
    if not task_source:
        ws3.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers3))
        c = ws3.cell(row=2, column=1, value="无待整改事项。")
        c.font = BODY_FONT
        c.border = BORDER
        c.alignment = CENTER_ALIGN
        ws3.row_dimensions[2].height = 24
    for idx, t in enumerate(task_source, start=1):
        vals = [
            idx, t.get("ref", f"F{idx:02d}"), t.get("severity", "一般"),
            t.get("action", ""), t.get("text", ""), t.get("owner") or "待指定",
            t.get("deadline", ""), t.get("verify", ""), t.get("status", "未开始"),
        ]
        for col, val in enumerate(vals, start=1):
            cell = ws3.cell(row=row_i, column=col, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = CENTER_ALIGN if col in (1, 2, 3, 6, 7, 9) else BODY_ALIGN
        sev = vals[2]
        if sev in SEV_FILL:
            ws3.cell(row=row_i, column=3).fill = SEV_FILL[sev]
        set_row_height(ws3, row_i, vals, widths3)
        row_i += 1
    page_setup(ws3)

    # ---------------- 表4 营养成分核对
    ws4 = wb.create_sheet("营养成分核对")
    headers4 = ["项目", "标示含量", "单位", "修约间隔", "“0”界限值", "修约", "“0”界限",
                "NRV", "标示 NRV%", "应标 NRV%", "结论", "备注"]
    widths4 = [18, 12, 12, 10, 12, 10, 12, 12, 12, 12, 14, 36]
    write_header(ws4, 1, headers4, widths4)
    ws4.freeze_panes = "A2"

    row_i = 2
    if not nut_rows:
        ws4.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers4))
        c = ws4.cell(row=2, column=1, value="未提供营养成分表数据（可能豁免标示或未上传营养成分表区域）。")
        c.font = BODY_FONT
        c.border = BORDER
        c.alignment = CENTER_ALIGN
        ws4.row_dimensions[2].height = 24
    for nr in nut_rows:
        vals = [
            nr["name"], nr["value"] if nr["value"] is not None else "—", nr["unit"], nr["rounding"],
            nr["zero"], nr["round_ok"], nr["zero_ok"], nr["nrv"], nr["pct_declared"],
            nr["pct_expect"], nr["verdict"], nr["note"] or "—",
        ]
        for col, val in enumerate(vals, start=1):
            cell = ws4.cell(row=row_i, column=col, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = CENTER_ALIGN if col <= 11 else BODY_ALIGN
        if nr["verdict"] == "不符":
            ws4.cell(row=row_i, column=11).fill = SEV_FILL["严重"]
        elif nr["verdict"] == "需复核":
            ws4.cell(row=row_i, column=11).fill = SEV_FILL["一般"]
        set_row_height(ws4, row_i, vals, widths4)
        row_i += 1

    if nutrition:
        row_i += 1
        ws4.merge_cells(start_row=row_i, start_column=1, end_row=row_i, end_column=len(headers4))
        c = ws4.cell(row=row_i, column=1, value=(
            f"标示基准：{nutrition.get('basis', '每100g')}　"
            f"盐油糖提示语：{'已标示' if nutrition.get('salt_sugar_warning') else '未见'}　"
            f"反式脂肪酸：" + ('应标（使用氢化/部分氢化油脂）' if nutrition.get("trans_fat_required") else "不强制") +
            "　｜　允许误差仅用于判定标示值是否可靠，不能替代检测判定产品真实含量。"))
        c.font = Font(name=FONT_NAME, size=9, color="595959")
        c.alignment = BODY_ALIGN
        ws4.row_dimensions[row_i].height = 30
    page_setup(ws4)

    # ---------------- 表5 核查底稿
    ws5 = wb.create_sheet("核查底稿")
    headers5 = ["序号", "核查项", "依据条款", "是否适用", "核查结论", "备注/证据"]
    widths5 = [6, 34, 22, 10, 12, 46]
    write_header(ws5, 1, headers5, widths5)
    ws5.freeze_panes = "A2"

    row_i = 2
    if not checklist:
        ws5.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers5))
        c = ws5.cell(row=2, column=1, value="未提供逐项核查记录。")
        c.font = BODY_FONT
        c.border = BORDER
        c.alignment = CENTER_ALIGN
        ws5.row_dimensions[2].height = 24
    for idx, item in enumerate(checklist, start=1):
        applicable = item.get("applicable", True)
        vals = [
            idx, item.get("item", ""), item.get("clause", ""),
            "适用" if applicable else "不适用",
            item.get("result", "—"), item.get("note", "") or "—",
        ]
        for col, val in enumerate(vals, start=1):
            cell = ws5.cell(row=row_i, column=col, value=val)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = CENTER_ALIGN if col in (1, 4, 5) else BODY_ALIGN
        if vals[4] == "不符合":
            ws5.cell(row=row_i, column=5).fill = SEV_FILL["严重"]
        elif vals[4] == "无法判定":
            ws5.cell(row=row_i, column=5).fill = SEV_FILL["无法判定"]
        set_row_height(ws5, row_i, vals, widths5)
        row_i += 1
    page_setup(ws5)

    wb.save(out_path)
    return {
        "path": os.path.abspath(out_path),
        "conclusion": conclusion,
        "counts": counts,
        "nutrition_issues": nut_issues,
        "nutrition_rows": len(nut_rows),
    }


def main():
    parser = argparse.ArgumentParser(description="生成食品标签审核报告（Excel）")
    parser.add_argument("data", help="审核数据 JSON 文件路径")
    parser.add_argument("-o", "--output", default="食品标签审核报告.xlsx", help="输出 xlsx 路径")
    args = parser.parse_args()

    with open(args.data, "r", encoding="utf-8-sig") as fh:
        data = json.load(fh)

    try:
        result = build(data, args.output)
    except KeyError as exc:
        print(f"数据结构缺少字段：{exc}", file=sys.stderr)
        return 1

    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
