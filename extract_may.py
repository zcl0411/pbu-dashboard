#!/usr/bin/env python3
"""Extract June 2026 data from Excel with updated row numbers matching current layout."""
import json, re, sys, openpyxl
from datetime import datetime

EXCEL = "D:/工作内容/数据分析/PBU数据需求/26年新增维度数据分析/5月/PBU组织能力看板-2026年5月.xlsx"

# Column mapping - includes both L2 and L3 departments
COL_MAP = {
    "PBU": "B",
    "欧洲一区": "C", "丹麦区": "C", "英国区": "D", "荷兰区": "E",
    "欧洲二区": "F", "比利时区": "F", "法国区": "G",
    "欧洲三区": "H", "西班牙区": "H", "意大利区": "I", "匈牙利区": "J",
    "澳洲区": "K", "日本区": "L", "北美区": "M",
    "四海捷运项目部": "N", "经营管理部": "O", "销售客服部": "P",
    "交付运营部": "Q", "PBU技术部": "R",
}
# Only L3 department names for iteration (not the L2 parent names)
DEPT_NAMES = ["PBU","丹麦区","英国区","荷兰区","比利时区","法国区",
    "西班牙区","意大利区","匈牙利区","澳洲区","日本区","北美区",
    "四海捷运项目部","经营管理部","销售客服部","交付运营部","PBU技术部"]
COLS = ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']

wb = openpyxl.load_workbook(EXCEL, data_only=True)
ws = wb['组织能力月度看板-5月']

def pct(v):
    if v is None: return None
    if isinstance(v, (int,float)): return round(float(v),4)
    return v

# ── 1. Assessment (Rows 3-8) ──
dims = ["组织能力","组织架构","人才密度","文化氛围"]
assessment = {}
for i, dim in enumerate(dims):
    row = 5 + i  # rows 5-8
    vals = {}
    # Read both L2 and L3 values
    for dept, col in COL_MAP.items():
        cv = ws[f"{col}{row}"].value
        if cv: vals[dept] = str(cv).strip()
    # Ensure L3 departments get values from L2 parent
    for sec, subs in [("欧洲一区",["丹麦区","英国区","荷兰区"]),
                       ("欧洲二区",["比利时区","法国区"]),
                       ("欧洲三区",["西班牙区","意大利区","匈牙利区"])]:
        if sec in vals:
            for s in subs:
                if s not in vals: vals[s] = vals[sec]
    assessment[dim] = vals

# ── 2. Summary (Rows 13-14) ──
summary = {
    "编制总数": ws["C13"].value, "在职总数": ws["F13"].value,
    "空编数": ws["I13"].value, "满编率": pct(ws["L13"].value),
    "核心人员胜任率": pct(ws["O13"].value), "优秀人才保留率": pct(ws["R13"].value),
    "关键岗位_编制总数": ws["C14"].value, "关键岗位_在职人数": ws["F14"].value,
    "关键岗位_空编数": ws["I14"].value, "关键岗位_满编率": pct(ws["L14"].value),
}

# ── 3.1 Establishment (Rows 17-27) ──
def parse_dept_table(start_row, end_row, label_col="A"):
    l2h = {}; l3h = {}
    for c in ['B']+COLS[1:]:
        v2 = ws[f"{c}{start_row}"].value; v3 = ws[f"{c}{start_row+1}"].value
        if v2: l2h[c] = str(v2).strip()
        if v3: l3h[c] = str(v3).strip()
    data = []
    for r in range(start_row+2, end_row+1):
        lbl = ws[f"A{r}"].value
        if lbl is None: continue
        lbl = str(lbl).strip()
        if lbl=='二级部门' or lbl=='三级部门': continue
        row_data = {"label": lbl}
        bv = ws[f"B{r}"].value
        if bv is not None:
            row_data['B'] = float(bv) if isinstance(bv,(int,float)) else str(bv).strip()
        for c in COLS[1:]:
            v = ws[f"{c}{r}"].value
            if v is not None:
                row_data[c] = float(v) if isinstance(v,(int,float)) else v
        data.append(row_data)
    return {"level2_headers": l2h, "level3_headers": l3h, "data": data}

establishment = parse_dept_table(17, 26)  # rows 17-26
estab_analysis = str(ws["A27"].value or "")

# ── 3.2 Actual headcount (Rows 28-38, analysis row 39) ──
actual_headcount = parse_dept_table(29, 38)  # headers rows 29-30, data 31-38
actual_headcount_analysis = str(ws["A39"].value or "")

# ── 3.3 Fill rate (Rows 39-53, analysis row 56) ──
fill_rate_raw = parse_dept_table(40, 53)  # headers rows 40-41, data 42-53
fill_rate_analysis = str(ws["A56"].value or "")
# Split cost_detail from main data
cost_types = ["销售费用","提货成本","清关成本","库操成本","管理费用","调度成本","地勤业务成本"]
main_data = []; cost_data = []; extras = []
for row in fill_rate_raw["data"]:
    lbl = row.get("label","")
    if lbl in ("在招岗位数","未启动岗位数"):
        extras.append(row)
    elif lbl in cost_types:
        cost_data.append(row)
    else:
        main_data.append(row)
fill_rate = {
    "level2_headers": fill_rate_raw["level2_headers"],
    "level3_headers": fill_rate_raw["level3_headers"],
    "data": main_data + extras,
    "cost_detail": cost_data,
    "analysis": fill_rate_analysis,
}

# ── 3.4 Key position (Rows 57-68) ──
l2h_kp = {}
for c in ['B']+COLS[1:]:
    v = ws[f"{c}58"].value
    if v: l2h_kp[c] = str(v).strip()
kp_data = []
for r in range(59, 67):
    lbl = ws[f"A{r}"].value
    if lbl is None: continue
    lbl = str(lbl).strip()
    if '部门' in lbl: continue
    row_data = {"label": lbl}
    for c in ['B','C','F','H','K','L','M','N','O','P','Q','R']:
        v = ws[f"{c}{r}"].value
        if v is not None:
            row_data[c] = float(v) if isinstance(v,(int,float)) else v
    kp_data.append(row_data)
kp_analysis = str(ws["A68"].value or "")
key_position = {"level2_headers": l2h_kp, "data": kp_data, "analysis": kp_analysis}

# ── 4. Personnel changes (Rows 69-75) ──
changes = {}
for r in range(72, 75):
    lbl = str(ws[f"A{r}"].value or "").strip()
    vals = {}
    for c in COLS:
        v = ws[f"{c}{r}"].value
        if v is not None:
            vals[c] = int(v) if isinstance(v,(int,float)) else v
    changes[lbl] = vals
changes_analysis = str(ws["A75"].value or "").replace("▶ ","")

# ── 5. Retention (Rows 76-84) ──
ret_data = []
for r in range(79, 84):
    lbl = ws[f"A{r}"].value
    if lbl is None: continue
    lbl = str(lbl).strip()
    if '部门' in lbl or '实际留任' in lbl: continue
    row_data = {"label": lbl}
    bv = ws[f"B{r}"].value
    if bv is not None:
        if isinstance(bv,(int,float)): row_data['B'] = float(bv)
        else: row_data['B'] = str(bv).strip()
    for c in COLS[1:]:
        v = ws[f"{c}{r}"].value
        if v is not None:
            row_data[c] = float(v) if isinstance(v,(int,float)) else v
    ret_data.append(row_data)
ret_analysis = str(ws["A84"].value or "")
retention = {"data": ret_data, "analysis": ret_analysis}

# ── 6. Improvements (Row 85+) ──
april = {}
for r in range(86, 96):
    lbl = ws[f"A{r}"].value
    if lbl is None: continue
    lbl = str(lbl).strip()
    if len(lbl) < 3: continue
    vals = {}
    for c in ['C','G','K','O']:  # week columns
        v = ws[f"{c}{r}"].value
        if v is not None: vals[c] = str(v).strip()
    if vals: april[lbl] = vals
improvements = {"四月": april}

# Build full data
data = {
    "month": "2026-05",
    "assessment": assessment,
    "assessment_analysis": "",
    "summary": summary,
    "org_chart": str(ws["A11"].value or "PBU组织架构调整：本月暂无"),
    "establishment": establishment,
    "estab_analysis": estab_analysis,
    "actual_headcount": actual_headcount,
    "actual_headcount_analysis": actual_headcount_analysis,
    "fill_rate": fill_rate,
    "key_position": key_position,
    "personnel_changes": changes,
    "changes_analysis": changes_analysis,
    "retention": retention,
    "vacancy": None,
    "improvements": improvements,
}

# Write JSON
with open("data/2026-05.json", "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2, default=lambda x: None)

# Print summary
for k in ['assessment','summary','establishment','actual_headcount','fill_rate','key_position','personnel_changes','retention']:
    v = data[k]
    if isinstance(v, dict) and 'data' in v:
        print(f"  {k}: {len(v['data'])} rows")
    elif isinstance(v, dict):
        print(f"  {k}: {json.dumps(v, ensure_ascii=False)[:100]}")
    else:
        print(f"  {k}: {str(v)[:100]}")

print(f"\n  changes: {json.dumps(changes, ensure_ascii=False)}")
print(f"  changes_analysis: {changes_analysis[:200]}")
print(f"  fill_rate_analysis: {fill_rate_analysis[:200]}")
print(f"  retention data: {len(ret_data)} rows")
print(f"\nDone! Data saved to data/2026-05.json")
