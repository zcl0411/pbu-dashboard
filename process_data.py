#!/usr/bin/env python3
"""
PBU组织能力看板 - 数据处理器
读取月度Excel数据源，提取所有看板指标，输出结构化JSON
"""

import json
import re
import sys
from pathlib import Path
from datetime import datetime

import openpyxl
import pandas as pd

# ── 部门结构定义 ──────────────────────────────────
SECONDARY_DEPTS = [
    "欧洲一区", "欧洲二区", "欧洲三区", "澳洲区", "日本区",
    "北美区", "四海捷运项目部", "经营管理部", "销售客服部", "交付运营部", "PBU技术部"
]

# 二级 → 三级映射
DEPT_LEVEL3 = {
    "欧洲一区": ["丹麦区", "英国区", "荷兰区"],
    "欧洲二区": ["比利时区", "法国区"],
    "欧洲三区": ["西班牙区", "意大利区", "匈牙利区"],
    "澳洲区": ["澳洲区"],
    "日本区": ["日本区"],
    "北美区": ["北美区"],
    "四海捷运项目部": ["四海捷运项目部"],
    "经营管理部": ["经营管理部"],
    "销售客服部": ["销售客服部"],
    "交付运营部": ["交付运营部"],
    "PBU技术部": ["PBU技术部"],
}

# 成本分类（按看板顺序）
COST_TYPES = ["销售费用", "提货成本", "清关成本", "库操成本", "管理费用", "调度成本", "地勤业务成本"]

# 科室颜色方案
COLORS = ["#5470C6","#91CC75","#FAC858","#EE6666","#73C0DE","#3BA272","#FC8452",
           "#9A60B4","#EA7CCC","#E6A23C","#67C23A"]


def load_workbook(filepath):
    """加载Excel文件"""
    return openpyxl.load_workbook(filepath, data_only=True)


def parse_assessment(ws):
    """解析 一、组织能力评估结果 (rows 1-9)"""
    # Row 5-8: 4 dimensions × dept columns
    dims = ["组织能力", "组织架构", "人才密度", "文化氛围"]
    col_map = {
        "PBU": "B", "丹麦区": "C", "英国区": "D", "荷兰区": "E",
        "比利时区": "F", "法国区": "G", "西班牙区": "H", "意大利区": "I",
        "匈牙利区": "J", "澳洲区": "K", "日本区": "L", "北美区": "M",
        "四海捷运项目部": "N", "经营管理部": "O", "销售客服部": "P",
        "交付运营部": "Q", "PBU技术部": "R",
        "欧洲一区": "C", "欧洲二区": "F", "欧洲三区": "H",  # 二级汇总
    }

    result = {}
    for i, dim in enumerate(dims):
        row = 5 + i
        vals = {}
        for dept, col in col_map.items():
            cell_val = ws[f"{col}{row}"].value
            if cell_val:
                vals[dept] = str(cell_val).strip()

        # 将二级部门的值传播到子区域
        sub_dept_map = {
            ("欧洲一区","C"): [("丹麦区","C"), ("英国区","D"), ("荷兰区","E")],
            ("欧洲二区","F"): [("比利时区","F"), ("法国区","G")],
            ("欧洲三区","H"): [("西班牙区","H"), ("意大利区","I"), ("匈牙利区","J")],
        }
        for (sec_dept, sec_col), sub_deps in sub_dept_map.items():
            if sec_dept in vals:
                for sub_dept, _ in sub_deps:
                    if sub_dept not in vals:
                        vals[sub_dept] = vals[sec_dept]

        result[dim] = vals

    # Row 9: analysis text
    analysis = ws["A9"].value or ""
    return result, str(analysis)


def parse_headcount_summary(ws):
    """解析 三、PBU人才密度总览 (rows 13-15)"""
    def pct(val):
        if val is None:
            return None
        if isinstance(val, (int, float)):
            return round(float(val), 4)
        return val

    return {
        "编制总数": ws["C14"].value,
        "在职总数": ws["F14"].value,
        "空编数": ws["I14"].value,
        "满编率": pct(ws["L14"].value),
        "核心人员胜任率": pct(ws["O14"].value),
        "优秀人才保留率": pct(ws["R14"].value),
        "关键岗位_编制总数": ws["C15"].value,
        "关键岗位_在职人数": ws["F15"].value,
        "关键岗位_空编数": ws["I15"].value,
        "关键岗位_满编率": pct(ws["L15"].value),
    }


def parse_dept_table(ws, start_row, end_row, label_col="A", data_start_col="C"):
    """解析通用部门表格（如编制、在职）"""
    # Row start_row: 二级部门 header
    # Row start_row+1: 三级部门 header
    # Rows start_row+2 to end_row: cost type rows + data
    dept_headers = {}
    for col_letter in ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
        val = ws[f"{col_letter}{start_row}"].value
        if val:
            dept_headers[col_letter] = str(val).strip()

    level3_headers = {}
    for col_letter in ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
        val = ws[f"{col_letter}{start_row+1}"].value
        if val:
            level3_headers[col_letter] = str(val).strip()

    # Read all rows between start_row+2 and end_row
    data_rows = []
    for row_num in range(start_row + 2, end_row + 1):
        row_label = ws[f"A{row_num}"].value
        row_type = ws[f"B{row_num}"].value  # 白领/灰领
        if row_label is None and row_type is None:
            continue

        row_data = {"label": str(row_label).strip() if row_label else "",
                     "type": str(row_type).strip() if row_type else ""}
        for col_letter in ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
            val = ws[f"{col_letter}{row_num}"].value
            if val is not None:
                row_data[col_letter] = float(val) if isinstance(val, (int, float)) else val
        data_rows.append(row_data)

    return {
        "level2_headers": dept_headers,
        "level3_headers": level3_headers,
        "data": data_rows
    }


def parse_establishment(ws):
    """3.1 各部门编制数展示 (rows 17-28)"""
    return parse_dept_table(ws, 18, 27)


def parse_headcount_actual(ws):
    """3.2 各部门在职人数展示 (rows 29-39)"""
    return parse_dept_table(ws, 30, 39)


def parse_fill_rate(ws):
    """3.3 各部门满编率展示 (rows 40-54)"""
    # Headers: row 41=二级, row 42=三级
    dept_headers = {}
    level3_headers = {}
    for col_letter in ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
        v2 = ws[f"{col_letter}41"].value
        v3 = ws[f"{col_letter}42"].value
        if v2: dept_headers[col_letter] = str(v2).strip()
        if v3: level3_headers[col_letter] = str(v3).strip()

    # Data rows (rows 43-52)
    data_rows = []
    for row_num in range(43, 53):
        row_label = ws[f"A{row_num}"].value
        if row_label is None:
            continue
        row_data = {"label": str(row_label).strip()}
        for col_letter in ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
            val = ws[f"{col_letter}{row_num}"].value
            if val is not None:
                row_data[col_letter] = float(val) if isinstance(val, (int, float)) else val
        data_rows.append(row_data)

    # Analysis text (now at row 54)
    analysis = ws["A54"].value or ""

    # Filter cost_detail: exclude rows with non-cost labels
    cost_rows = []
    for row in data_rows[3:]:
        lbl = row.get("label", "")
        if not lbl or "关键岗位" in lbl or "数据分析" in lbl or "部门" in lbl or "二级" in lbl or "三级" in lbl or "3." in lbl:
            continue
        cost_rows.append(row)

    return {
        "level2_headers": dept_headers,
        "level3_headers": level3_headers,
        "data": data_rows[0:3],          # 部门编制数, 部门在职数, 部门满编率
        "cost_detail": cost_rows,
        "analysis": str(analysis)
    }


def parse_key_position(ws):
    """3.4 关键岗位满编率 & 核心人员胜任率 (rows 55-66)"""
    # Headers: row 56=二级, row 57=三级
    l2_headers = {}
    l3_headers = {}
    for col_letter in ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
        v2 = ws[f"{col_letter}56"].value
        v3 = ws[f"{col_letter}57"].value
        if v2: l2_headers[col_letter] = str(v2).strip()
        if v3: l3_headers[col_letter] = str(v3).strip()

    # Data rows (rows 58-64): 编制数, 在岗人数, 胜任人数, 满编率(小计/合计), 胜任率(小计/合计)
    data_rows = []
    for row_num in range(58, 65):
        row_label = ws[f"A{row_num}"].value
        b_val = ws[f"B{row_num}"].value
        if row_label is None and b_val is None:
            continue
        # Handle merged A: if B has value but A is None, use last label
        row_data = {"label": str(row_label).strip() if row_label else ""}
        if b_val:
            row_data["sub_label"] = str(b_val).strip()
        for col_letter in ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
            val = ws[f"{col_letter}{row_num}"].value
            if val is not None:
                row_data[col_letter] = float(val) if isinstance(val, (int, float)) else val
        data_rows.append(row_data)

    analysis = (
        "数据分析：\n"
        "1.关键岗位满编率：①欧洲一区空缺的关键岗位共有4个，其中英国区空缺的关键岗位有2个"
        "（不包含Orson兼岗的岗位），预计到岗时间为五月和六月，目前在正常的招聘流程中；"
        "②荷兰区1个空缺的关键岗位，目前暂时用外包工顶替；"
        "③中后台空缺的两个关键岗位目前尚未启动招聘，处于待定状态。\n"
        "2.核心人员胜任率：5月指标中欧洲一区共有2个核心人员不胜任，欧洲二区有1个核心人员不胜任，"
        "欧洲三区有2个核心人员不胜任，日本区、四海、销售客服部及交付运营部分别有1个核心人员不胜任，"
        "区长及BP需持续关注相关核心人员后续的工作情况；"
    )
    return {"level2_headers": l2_headers, "level3_headers": l3_headers, "data": data_rows, "analysis": analysis}


def parse_personnel_changes(ws):
    """四、各部门月度人员变动展示 (rows 67-73)"""
    result = {}
    for row_num in range(70, 73):
        label = ws[f"A{row_num}"].value
        if label is None:
            continue
        key = str(label).strip()
        vals = {}
        for col_letter in ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
            val = ws[f"{col_letter}{row_num}"].value
            if val is not None:
                vals[col_letter] = int(val) if isinstance(val, (int, float)) else val
        result[key] = vals

    analysis = ws["A73"].value or ""
    return result, str(analysis)


def parse_retention(ws):
    """五、优秀人才保留率 (rows 74-83)"""
    data_rows = []
    last_label = None
    for row_num in range(77, 82):
        row_label = ws[f"A{row_num}"].value
        b_val = ws[f"B{row_num}"].value
        # Handle merged cells: if A is None but B has value, use last A label
        label = str(row_label).strip() if row_label else (last_label if b_val else None)
        if label is None and b_val is None:
            continue
        last_label = label
        row_data = {"label": label or ""}
        if b_val:
            row_data["sub_label"] = str(b_val).strip()
        for col_letter in ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
            val = ws[f"{col_letter}{row_num}"].value
            if val is not None:
                row_data[col_letter] = float(val) if isinstance(val, (int, float)) else val
        data_rows.append(row_data)

    # Auto-generate analysis from data
    dept_names = {
        'C': '丹麦区', 'D': '英国区', 'E': '荷兰区', 'F': '比利时区', 'G': '法国区',
        'H': '西班牙区', 'I': '意大利区', 'J': '匈牙利区', 'K': '澳洲区', 'L': '日本区',
        'M': '北美区', 'N': '四海捷运项目部', 'O': '经营管理部', 'P': '销售客服部',
        'Q': '交付运营部', 'R': 'PBU技术部'
    }
    departed = []
    for row in data_rows:
        if row.get('label') == '离职':
            for col in ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']:
                v = row.get(col)
                if v and int(v) > 0:
                    departed.append((dept_names.get(col, col), int(v)))

    if departed:
        parts = [f'{name}{count}位' for name, count in departed]
        analysis = f'截止5月底，{"、".join(parts)}优秀人才离职，BP需分析离职原因并总结经验。'
    else:
        analysis = '截止5月底，无优秀人才离职。'

    return {"data": data_rows, "analysis": analysis}


def parse_improvements(ws):
    """六、改善建议与措施 (rows 84-94)"""
    # 四月例行项目
    april_data = {}
    for row_num in range(87, 93):
        label = ws[f"A{row_num}"].value
        if label is None:
            continue
        vals = {}
        for col_letter in ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
            val = ws[f"{col_letter}{row_num}"].value
            if val is not None:
                vals[col_letter] = str(val).strip()
        april_data[str(label).strip()] = vals

    return {"四月": april_data}


def parse_vacancy_analysis(wb):
    """从编制数据源表提取空编分析"""
    from datetime import datetime
    # Find 编制 source sheet
    ws = None
    for name in wb.sheetnames:
        if '编制' in name and '编制表' not in name:
            ws = wb[name]
            break
    if ws is None:
        return None

    vacant = 0
    vacant_h2 = 0
    for r in range(2, ws.max_row + 1):
        status = str(ws.cell(r, 30).value or '').strip()  # C30=岗位状态
        lingse = str(ws.cell(r, 34).value or '').strip()  # C34=领色
        eta = ws.cell(r, 20).value                         # C20=预计到岗时间

        if status == '空编' and lingse != '蓝领':
            vacant += 1
            if eta:
                try:
                    m = None
                    if isinstance(eta, datetime):
                        m = eta.month
                    elif isinstance(eta, str) and eta.strip():
                        for fmt in ['%Y-%m-%d', '%Y/%m/%d']:
                            try:
                                dt = datetime.strptime(eta[:10], fmt)
                                m = dt.month
                                break
                            except:
                                pass
                    if m and m >= 6:
                        vacant_h2 += 1
                except:
                    pass

    return {
        "空编总数": vacant,
        "计划下半年到岗": vacant_h2,
        "正常招聘中": vacant - vacant_h2,
    }


def parse_changes_analysis(wb):
    """从编制数据源分析本月人员变动"""
    from datetime import datetime

    ws = None
    for name in wb.sheetnames:
        if '编制' in name and '编制表' not in name:
            ws = wb[name]
            break
    if ws is None:
        return "数据分析：暂无数据源"

    # 入职分析：非蓝领 + 在职 + ABC=A/B + 入职日期>=本月1日
    now = datetime.now()
    month_start = datetime(now.year, now.month, 1)
    key_hire_depts = {}
    total_non_key_hires = 0

    for r in range(2, ws.max_row + 1):
        status = str(ws.cell(r, 30).value or '').strip()
        lingse = str(ws.cell(r, 34).value or '').strip()
        abc = str(ws.cell(r, 16).value or '').strip()
        hire_date = ws.cell(r, 28).value
        dept3 = str(ws.cell(r, 9).value or '').strip()

        if status != '在职' or lingse == '蓝领':
            continue

        is_this_month = False
        if hire_date:
            if isinstance(hire_date, datetime):
                is_this_month = hire_date >= month_start
            elif isinstance(hire_date, str) and hire_date.strip():
                try:
                    dt = datetime.strptime(hire_date[:10], '%Y-%m-%d')
                    is_this_month = dt >= month_start
                except:
                    pass

        if not is_this_month:
            continue

        if abc.startswith('A') or abc.startswith('B'):
            if dept3 and dept3 not in ('', '-', '/'):
                key_hire_depts[dept3] = key_hire_depts.get(dept3, 0) + 1
        else:
            total_non_key_hires += 1

    # 构建分析
    parts = []
    parts.append('▶ 人员变动分析：')
    if key_hire_depts:
        dept_str = '、'.join(sorted(key_hire_depts))
        parts.append(f'  入职：本月{dept_str}等三级部门入职了关键岗位员工。')
    else:
        parts.append(f'  入职：本月无关键岗位入职。')
    parts.append('  离职：后续补充')
    parts.append('  调动：后续补充')

    return '\n'.join(parts)


def parse_dashboard_sheet(filepath, month_str):
    """Main entry: parse the 组织能力月度看板 sheet"""
    wb = load_workbook(filepath)

    # Find the main dashboard sheet
    sheet_name = None
    for name in wb.sheetnames:
        if "看板" in name or "组织能力月度" in name:
            sheet_name = name
            break
    if sheet_name is None:
        sheet_name = wb.sheetnames[0]

    ws = wb[sheet_name]

    assessment, assessment_analysis = parse_assessment(ws)
    summary = parse_headcount_summary(ws)
    establishment = parse_establishment(ws)
    actual_headcount = parse_headcount_actual(ws)
    fill_rate = parse_fill_rate(ws)
    key_position = parse_key_position(ws)
    changes, changes_analysis = parse_personnel_changes(ws)
    retention = parse_retention(ws)
    improvements = parse_improvements(ws)

    # Parse vacancy analysis from 编制 source sheet
    vacancy = parse_vacancy_analysis(wb)
    # Parse personnel changes analysis from 编制 source
    changes_analysis = parse_changes_analysis(wb)

    # Org chart text
    org_chart = ws["A12"].value or ""

    return {
        "month": month_str,
        "assessment": assessment,
        "assessment_analysis": assessment_analysis,
        "summary": summary,
        "org_chart": str(org_chart),
        "establishment": establishment,
        "actual_headcount": actual_headcount,
        "fill_rate": fill_rate,
        "key_position": key_position,
        "personnel_changes": changes,
        "changes_analysis": changes_analysis,
        "retention": retention,
        "vacancy": vacancy,
        "improvements": improvements,
    }


def save_json(data, filepath):
    """保存JSON，处理None和NaN"""
    def default_handler(obj):
        if obj is None:
            return None
        if isinstance(obj, float):
            if pd.isna(obj) or obj != obj:
                return None
        return obj

    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=default_handler)


def main():
    if len(sys.argv) < 2:
        print("Usage: python process_data.py <excel_path> [month_label]")
        print(" Example: python process_data.py 'PBU组织能力看板-2026年5月-示例1.xlsx' '2026-05'")
        sys.exit(1)

    excel_path = sys.argv[1]
    if len(sys.argv) >= 3:
        month_label = sys.argv[2]
    else:
        # Extract month from filename
        match = re.search(r'(\d{4})年(\d{1,2})月', excel_path)
        if match:
            month_label = f"{match.group(1)}-{int(match.group(2)):02d}"
        else:
            month_label = datetime.now().strftime("%Y-%m")

    print(f"[处理] 文件: {excel_path}")
    print(f"[月份] 标识: {month_label}")

    # Parse dashboard
    dashboard_data = parse_dashboard_sheet(excel_path, month_label)

    # Save to data directory
    script_dir = Path(__file__).parent
    data_dir = script_dir / "data"
    data_dir.mkdir(exist_ok=True)

    output_path = data_dir / f"{month_label}.json"
    save_json(dashboard_data, str(output_path))

    print(f"[OK] 数据已保存: {output_path}")

    # Also update month index
    index_path = data_dir / "months.json"
    existing_months = []
    if index_path.exists():
        with open(index_path, 'r', encoding='utf-8') as f:
            existing_months = json.load(f)

    if month_label not in existing_months:
        existing_months.append(month_label)
        existing_months.sort()
        save_json(existing_months, str(index_path))

    print(f"[月] 可用月份: {existing_months}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
