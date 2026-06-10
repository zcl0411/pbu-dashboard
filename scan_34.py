import sys; sys.path.insert(0, '.')
import process_data as pd

wb = pd.load_workbook(r'D:\工作内容\数据分析\PBU数据需求\26年新增维度数据分析\5月\PBU组织能力看板-2026年5月-示例2.xlsx')
for name in wb.sheetnames:
    if '看板' in name or '组织能力月度' in name:
        ws = wb[name]; break

# Scan rows around 3.4 section (55-75)
print("=== 3.4 section rows ===")
for r in range(54, 75):
    a = ws.cell(r, 1).value
    b = ws.cell(r, 2).value
    if a:
        vals = {}
        for col_letter in 'CFHKLMNOPQR':
            col_idx = ord(col_letter) - ord('A') + 1
            v = ws.cell(r, col_idx).value
            if v is not None:
                vals[col_letter] = v
        print(f"Row {r}: A={repr(a)}, B={repr(b)}")
        if vals: print(f"  Data: {vals}")
