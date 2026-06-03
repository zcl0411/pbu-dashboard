#!/usr/bin/env python3
"""
一键更新：解析Excel → 嵌入数据 → 推送GitHub
Usage: python update.py "Excel路径" "月份标识"
Example: python update.py "D:/工作内容/数据分析/PBU数据需求/26年新增维度数据分析/6月/xxx.xlsx" "2026-06"
"""
import json
import re
import sys
import subprocess
from pathlib import Path

# Import existing processor
import process_data

def embed_data(month_label):
    """将 data/{month}.json 嵌入到 index.html 中"""
    html_path = Path(__file__).parent / "index.html"
    data_path = Path(__file__).parent / "data" / f"{month_label}.json"
    months_path = Path(__file__).parent / "data" / "months.json"

    with open(data_path, 'r', encoding='utf-8') as f:
        fresh = json.load(f)

    with open(months_path, 'r', encoding='utf-8') as f:
        months = json.load(f)

    with open(html_path, 'r', encoding='utf-8') as f:
        html = f.read()

    # Remove old EMBED_DATA blocks
    old_len = len(html)
    for _ in range(10):
        html = re.sub(
            r'// === EMBEDDED DATA ===[^\n]*\nconst EMBED_MONTHS\s*=\s*\[[^\]]*\];?\s*\nconst EMBED_DATA\s*=\s*\{[\s\S]*?\n\};',
            '', html)
        if len(html) == old_len:
            break
        old_len = len(html)
    html = re.sub(r'\n{4,}', '\n\n\n', html)

    # Build full embedded data
    all_data = {}
    for m in months:
        if m == month_label:
            all_data[m] = fresh
        else:
            # Load historical data
            hist_path = Path(__file__).parent / "data" / f"{m}.json"
            if hist_path.exists():
                with open(hist_path, 'r', encoding='utf-8') as f:
                    all_data[m] = json.load(f)

    embed = '// === EMBEDDED DATA ===\nconst EMBED_MONTHS = ' + \
        json.dumps(months, ensure_ascii=False) + ';\nconst EMBED_DATA = ' + \
        json.dumps(all_data, ensure_ascii=False) + ';\n'

    html = html.replace(
        "let allData={}, months=[], currentMonth='';",
        "let allData={}, months=[], currentMonth='';\n" + embed
    )

    cnt = html.count('const EMBED_DATA')
    if cnt != 1:
        print(f"[警告] EMBED_DATA 声明了 {cnt} 次，可能有残留数据")
    else:
        print(f"[OK] 数据已嵌入 index.html")

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(html)


def git_push(month_label):
    """提交并推送到 GitHub"""
    import os
    script_dir = Path(__file__).parent
    os.chdir(str(script_dir))

    subprocess.run(['git', 'add', 'index.html', 'data/'], check=False)
    subprocess.run(['git', 'commit', '-m',
                   f'Update dashboard data: {month_label}'], check=False)
    subprocess.run(['git', 'push'], check=False)
    print(f"[OK] 已推送到 GitHub Pages")


def main():
    if len(sys.argv) < 2:
        print("Usage: python update.py <excel_path> [month_label]")
        sys.exit(1)

    excel_path = sys.argv[1]
    month_label = sys.argv[2] if len(sys.argv) >= 3 else None

    # Step 1: Parse Excel
    process_data.main()

    # Determine month label
    if month_label is None:
        match = re.search(r'(\d{4})年(\d{1,2})月', excel_path)
        if match:
            month_label = f"{match.group(1)}-{int(match.group(2)):02d}"
        else:
            from datetime import datetime
            month_label = datetime.now().strftime("%Y-%m")

    # Step 2: Embed into HTML
    embed_data(month_label)

    # Step 3: Push to GitHub
    git_push(month_label)

    print(f"\n[完成] 看板已更新: https://zcl0411.github.io/pbu-dashboard/")


if __name__ == "__main__":
    main()
