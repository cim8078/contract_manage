# -*- coding: utf-8 -*-
"""生成「数据清洗清单」（只读、基于 test_data 基线副本；不触碰生产库）。
输出：
  test_data/数据清洗清单_明细.csv   —— 逐条明细（含名称/原值，供内部清洗使用；该目录已 gitignore）
  docs/test_reports/数据清洗清单_汇总.md —— 汇总统计 + 脱敏样例（进仓库，不含姓名与金额原文）
"""
import csv
import os
import re
import sys
import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _common as C  # noqa: E402
import _oracle as O  # noqa: E402

AMOUNT_COLS = ["决策的招标控制价", "合同金额（万元）", "结算金额", "已支付", "未支付"]
DATE_COLS = ["采购时间", "开工时间", "完工时间", "竣工验收时间", "下次招采时间"]
ENUM_COLS = {"项目分类": ["服务类", "维护维修类", "货物类", "建设项目"],
             "采购方式": ["公开招标", "竞争性磋商", "竞争性谈判", "自行采购"],
             "资金来源": ["年初部门预算", "年中追加预算", "统筹预算", "上级专项转移支付", "上年结转结余", "其他"],
             "是否结算": ["是", "否", "部分结算"],
             "是否有签证或变更": ["是", "否"]}


def main():
    C.setup_stdout()
    C.ensure_dirs()
    C.prepare_runtime()
    C.prepare_base()
    db = C.BASE_DB
    rows = C.query(db, "SELECT * FROM project_contracts")
    users = set((x["姓名"] or "").strip() for x in C.query(db, "SELECT 姓名 FROM users"))
    depts = set(x["d"] for x in C.query(db, "SELECT 单位名称||'-'||科室名称 d FROM departments"))
    issues = []          # (类别, 编号, 项目名称, 字段, 原值, 建议, 所属部门)

    def add(kind, r, field, raw, suggest=""):
        issues.append([kind, str(r.get("编号", "")), str(r.get("项目名称", "")), field,
                       "" if raw is None else str(raw), suggest, str(r.get("所属部门", ""))])

    # 1) 金额非数字 / 负值
    neg = 0
    for r in rows:
        for col in AMOUNT_COLS:
            v = r.get(col)
            if v in (None, ""):
                continue
            try:
                f = float(v)
            except (TypeError, ValueError):
                m = re.search(r"-?\d+(?:\.\d+)?", str(v))
                add("金额-非数字", r, col, v, ("可提取数字 %s（待业务确认）" % m.group(0)) if m else "无法自动提取，需人工确认")
                continue
            if f < 0:
                neg += 1
                add("金额-负值", r, col, v, "确认是否为录入符号错误")

    # 2) Q9 勾稽：已支付+未支付 ≠ 结算金额（容差 0.01）
    q9 = 0
    for r in rows:
        s, p, u = O.num(r.get("结算金额")), O.num(r.get("已支付")), O.num(r.get("未支付"))
        if s is None or p is None or u is None:
            continue
        if abs(p + u - s) > 0.01:
            q9 += 1
            add("勾稽-已支付+未支付≠结算金额", r, "已支付/未支付/结算金额", "%s/%s/%s" % (p, u, s),
                "差额 %.2f" % (p + u - s))

    # 3) Q10：结算金额 > 合同金额（容差 1%）
    q10 = 0
    for r in rows:
        a, s = O.num(r.get("合同金额（万元）")), O.num(r.get("结算金额"))
        if a is None or s is None:
            continue
        if s > a * 1.01:
            q10 += 1
            add("勾稽-结算金额>合同金额", r, "结算金额/合同金额", "%s/%s" % (s, a), "超出 %.2f%%" % ((s / a - 1) * 100 if a else 0))

    # 4) 招标控制价单位疑似错位（控制价/合同额 ≥ 100）
    unit = 0
    for r in rows:
        b, a = O.num(r.get("决策的招标控制价")), O.num(r.get("合同金额（万元）"))
        if b is None or a is None or a <= 0:
            continue
        if b / a >= 100:
            unit += 1
            add("金额-控制价单位疑似错位", r, "决策的招标控制价/合同金额（万元）", "%s/%s" % (b, a),
                "比值 %.0f 倍，疑似以「元」填写（应换算为万元）" % (b / a))

    # 5) 经办人不在用户主数据
    handler = 0
    for r in rows:
        h = str(r.get("经办人") or "").strip()
        if h and h not in users:
            handler += 1
            add("经办人-不在用户表", r, "经办人", h, "补录用户或确认白名单（Q8：中级别，不强制）")

    # 6) 日期不可解析 / 枚举非标准 / 所属部门不在部门表
    bad_date = bad_enum = bad_dept = 0
    for r in rows:
        for col in DATE_COLS:
            v = str(r.get(col) or "").strip()
            if v and O.parse_date(v) is None:
                bad_date += 1
                add("日期-不可解析", r, col, v, "统一为 yyyy-mm-dd")
        for col, allowed in ENUM_COLS.items():
            v = str(r.get(col) or "").strip()
            if v and v not in allowed:
                bad_enum += 1
                add("枚举-非标准取值", r, col, v, "标准取值：" + "、".join(allowed))
        d = str(r.get("所属部门") or "").strip()
        if d and d not in depts:
            bad_dept += 1
            add("部门-不在部门表", r, "所属部门", d, "补录部门或修正合同归属")

    # 明细 CSV（内部使用，test_data/ 已 gitignore）
    csv_path = os.path.join(C.TEST_DATA, "数据清洗清单_明细.csv")
    with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["类别", "合同编号", "项目名称", "字段", "原值", "建议", "所属部门"])
        w.writerows(issues)

    # 汇总 MD（进仓库：仅计数与编号，不含姓名与金额原文）
    kinds = {}
    for x in issues:
        kinds[x[0]] = kinds.get(x[0], 0) + 1
    md = ["# 数据清洗清单（汇总）", "",
          "- **生成时间**：%s" % datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
          "- **数据来源**：`test_data/base_合同台账.db`（由 `backups/` 只读复制；生产库未被读写）",
          "- **明细文件**：`test_data/数据清洗清单_明细.csv`（含合同名称/原值，目录已 gitignore，**请勿提交公开仓库**）",
          "- **合同总量**：%d 条" % len(rows), "",
          "## 一、问题汇总", "", "| 类别 | 条数 |", "|---|---|"]
    for k in sorted(kinds, key=lambda x: -kinds[x]):
        md.append("| %s | %d |" % (k, kinds[k]))
    md += ["| **合计** | **%d** |" % len(issues), "",
           "> 注：金额「非数字」为**字段级**计数（同一合同可能有多个金额字段异常），与按“条”统计的口径不同。", "",
           "## 二、按合同金额口径的影响", "",
           "| 项 | 数量 | 说明 |", "|---|---|---|",
           "| 金额非数字（字段级） | %d | 该字段在统计/首页中按 0 参与 |" % kinds.get("金额-非数字", 0),
           "| 金额负值 | %d | 违反 Q14 |" % neg,
           "| 勾稽 Q9 违背 | %d | 已支付+未支付 ≠ 结算金额（容差 0.01） |" % q9,
           "| 勾稽 Q10 违背 | %d | 结算金额 > 合同金额（容差 1%%） |" % q10,
           "| 控制价单位疑似错位 | %d | 控制价/合同额 ≥ 100 倍 |" % unit,
           "| 经办人不在用户表 | %d | Q8 中级别，不强制；影响短信收件人解析 |" % handler,
           "| 日期不可解析 | %d | 会导致该合同永不触发到期提醒 |" % bad_date,
           "| 枚举非标准 | %d | 影响分类/采购方式/资金来源统计维度 |" % bad_enum,
           "| 所属部门不在部门表 | %d | 违反 Q7 |" % bad_dept, ""]
    samp = [x for x in issues if x[0] == "金额-非数字"][:8]
    if samp:
        md += ["## 三、样例（已脱敏：仅编号与字段名）", "", "| 合同编号 | 字段 |", "|---|---|"]
        for x in samp:
            md.append("| %s | %s |" % (x[1], x[3]))
        md.append("")
    md += ["## 四、清洗建议顺序", "",
           "1. **金额单位/文本**（影响首页总合同额、统计与导出）：先处理「控制价单位错位」，再处理「金额非数字」；",
           "2. **勾稽 Q9/Q10**（影响月审计）：与业务逐条确认差额合理性；",
           "3. **经办人归属**（影响短信收件人解析与责任追溯）；",
           "4. **部门/枚举**（影响权限可见范围与统计维度，数量少可批量修正）。", "",
           "> 清洗方式建议：先导出本清单 → 业务逐条确认 → 生成 UPSERT 脚本 → 在**副本库**演练 → 再对生产库执行（执行前备份）。", ""]
    md_path = os.path.join(C.REPORT_DIR, "数据清洗清单_汇总.md")
    with open(md_path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(md))
    print("[OK] 明细 %d 条 → %s" % (len(issues), csv_path))
    print("[OK] 汇总 → %s" % md_path)
    for k in sorted(kinds, key=lambda x: -kinds[x]):
        print("   %-28s %d" % (k, kinds[k]))


if __name__ == "__main__":
    main()
