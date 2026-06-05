#!/usr/bin/env python3
"""合并示例1和示例2的数据：S3.4从示例2取，S4/S5从示例1恢复"""
import json
from pathlib import Path

data_dir = Path(__file__).parent / "data"

with open(data_dir / "2026-05-s2.json", "r", encoding="utf-8") as f:
    merged = json.load(f)

with open(data_dir / "2026-05-s1.json", "r", encoding="utf-8") as f:
    s1 = json.load(f)

# 从示例1恢复 S4 人员变动和 S5 保留率
merged["personnel_changes"] = s1["personnel_changes"]
merged["changes_analysis"] = s1["changes_analysis"]
merged["retention"] = s1["retention"]

# 保存最终合并文件
output_path = data_dir / "2026-05.json"
with open(output_path, "w", encoding="utf-8") as f:
    json.dump(merged, f, ensure_ascii=False, indent=2)

print("[OK] 合并完成: 2026-05.json")

# 验证
with open(output_path, "r", encoding="utf-8") as f:
    verify = json.load(f)

pc = verify["personnel_changes"]
total = sum(int(v) for v in pc.get("入职", {}).values())
print(f"[验证] 入职总数: {total}")
print(f"[验证] 离职总数: {sum(int(v) for v in pc.get('离职', {}).values())}")
print(f"[验证] changes_analysis 入职数: {verify['changes_analysis'][:80]}")
print(f"[验证] 3.4 key_position labels: {[r['label'] for r in verify['key_position']['data']]}")
print(f"[验证] S5 retention labels: {[(r.get('label'), r.get('sub_label','')) for r in verify['retention']['data']]}")
