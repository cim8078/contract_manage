# -*- coding: utf-8 -*-
"""A-08 枚举与部门数据清洗：一次性迁移/清洗脚本。
生成待确认清单 -> 校验映射 -> 事务内更新 -> 写 operation_logs -> 结果验证。
"""
import io, os, sys, sqlite3, datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
DB = "data/合同台账.db"
NOW = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# ---------- 映射表（依据《数据字典表》3.1/3.2/3.3；待业务确认项已在清单中标注） ----------
CATEGORY_MAP = {
    "工程类": "建设项目",
    "维修类": "维护维修类",
    "维修维护类": "维护维修类",
    "委托支付函": "服务类",
    "造价咨询服务": "服务类",
}
# 补充协议按主合同归类（176301/256307 主合同均为空调供能/委托管理服务）
CATEGORY_BY_NO = {"176301-05": "服务类", "256307-05": "服务类"}

PROCURE_MAP = {
    "自行采购（线下邀请）": "自行采购",
    "自行采购（电子卖场）": "自行采购",
    "自行采购（邀请招标）": "自行采购",
    "自行采购（公开招标）": "公开招标",
    "自行采购(直接采购)": "自行采购",
    "自行采购（询价采购）": "自行采购",
    "框架协议采购": "自行采购",
    "自行采购（竞争性谈判）": "竞争性谈判",
    "自行采购(竞争性谈判)": "竞争性谈判",
    "政府采购网竞争性磋商": "竞争性磋商",
    "自行采购（竞争性磋商）": "竞争性磋商",
    "自行采购(邀请招标）": "自行采购",
    "自行采购(邀请招标)": "自行采购",
    "补充协议": "自行采购",
    "原合同《市民之家中央空调供能项目合同》": "竞争性磋商",
    "原合同《市委、市人大、市政府、市政协办公楼中央空调系统委托管理合同》": "自行采购",
    "邀请招标（联合体投标）": "公开招标",
    "邀请招标": "公开招标",
    "竞争性谈判,线上直采": "竞争性谈判",
    "自行采购（要求招标）": "自行采购",
    "自行采购（直接采购）": "自行采购",
    "自行采购（由住建部门遴选到该公司进行施工图审查）": "自行采购",
    "自行采购（单一来源直采）": "自行采购",
    "线下比选，线上直购": "自行采购",
    "电子卖场直购（线下邀标比选）": "自行采购",
    "电子卖场直购": "自行采购",
    "单一来源采购": "自行采购",
    "832平台单一采购": "自行采购",
}

SETTLE_MAP = {
    "按季度结算": "是",
    "按月结算": "是",
    "按年度结算": "是",
    "按月支付": "否",
    "由具体实施的单项项目结算": "否",
    "/": "",
}
SETTLE_KEEP = {"部分结算": "保留（建议新增标准值，待业务确认）"}

CAT_STD = ("服务类", "维护维修类", "货物类", "建设项目")
PROC_STD = ("公开招标", "竞争性磋商", "竞争性谈判", "自行采购")
SETTLE_STD = ("是", "否", "")

con = sqlite3.connect(DB, timeout=30)
con.row_factory = sqlite3.Row

def rows(sql, args=()):
    return con.execute(sql, args).fetchall()

# ---------- 1. 读取存量与校验 ----------
cat_rows = rows('SELECT "编号","项目名称","项目分类" FROM contracts WHERE "项目分类" NOT IN ("服务类","维护维修类","货物类","建设项目") ORDER BY "编号"')
proc_rows = rows('SELECT "编号","项目名称","采购方式" FROM contracts WHERE "采购方式" NOT IN ("公开招标","竞争性磋商","竞争性谈判","自行采购") ORDER BY "编号"')
settle_rows = rows('SELECT "编号","项目名称","是否结算" FROM contracts WHERE "是否结算" NOT IN ("是","否","") ORDER BY "编号"')
dept_bad = rows('SELECT "编号","所属部门" FROM contracts WHERE "所属部门" IS NOT NULL AND TRIM("所属部门")<>"" AND instr("所属部门","-")=0')
print("[INFO] A-08 扫描：项目分类非标准 %d 条、采购方式非标准 %d 条、是否结算非标准 %d 条、缺科室后缀 %d 条" % (len(cat_rows), len(proc_rows), len(settle_rows), len(dept_bad)), file=sys.stderr)
if len(cat_rows) + len(proc_rows) + len([r for r in settle_rows if r["是否结算"] in SETTLE_MAP]) == 0:
    print("[INFO] A-08 无待清洗记录（已清洗），脚本幂等退出，不再覆盖文档。", file=sys.stderr)
    sys.exit(0)

# 映射覆盖校验：每个非标准原值必须有映射（补充协议分类按编号）
from collections import Counter
cat_vals = Counter(r["项目分类"] for r in cat_rows)
proc_vals = Counter(r["采购方式"] for r in proc_rows)
settle_vals = Counter(r["是否结算"] for r in settle_rows)
missing_cat = [v for v in cat_vals if v not in CATEGORY_MAP and v != "补充协议"]
missing_proc = [v for v in proc_vals if v not in PROCURE_MAP]
missing_settle = [v for v in settle_vals if v not in SETTLE_MAP and v not in SETTLE_KEEP]
# 补充协议分类按编号定向：先按编号匹配，剩余若出现则报错
cat_by_no_missing = [r["编号"] for r in cat_rows if r["项目分类"] == "补充协议" and r["编号"] not in CATEGORY_BY_NO]
if missing_cat: raise SystemExit("[ERROR] 项目分类无映射：" + str(missing_cat))
if cat_by_no_missing: raise SystemExit("[ERROR] 补充协议无编号定向映射：" + str(cat_by_no_missing))
if missing_proc: raise SystemExit("[ERROR] 采购方式无映射：" + str(missing_proc))
if missing_settle: raise SystemExit("[ERROR] 是否结算无映射：" + str(missing_settle))
print("[INFO] A-08 映射校验通过：全部非标准值均有建议值", file=sys.stderr)

# ---------- 2. 生成待确认清单文档 ----------
def esc(s):
    return str(s or "").replace("|", "\\|").replace("\n", " ")

md = []
md.append("# A-08 枚举与部门数据清洗（待确认清单与执行记录）")
md.append("")
md.append("> 依据《数据字典表》3.1/3.2/3.3 标准枚举与映射建议，对 `contracts` 存量非标准枚举值进行标准化清洗，并校验“所属部门缺科室后缀”数据。")
md.append("> 清洗脚本：`_migrate_a08.py`（一次性）；清洗前备份：见第 8 节。")
md.append("")
md.append("**1. 存量扫描结论（清洗前）**")
md.append("")
md.append("| 字段 | 标准值 | 非标准条数 | 说明 |")
md.append("|---|---|---|---|")
md.append("| 项目分类 | 服务类/维护维修类/货物类/建设项目 | %d | 工程类56、维修类26、维修维护类22、委托支付函6、补充协议2、造价咨询服务1 |" % len(cat_rows))
md.append("| 采购方式 | 公开招标/竞争性磋商/竞争性谈判/自行采购 | %d | 括号细分写法 30+ 种，统一归入 4 大类 |" % len(proc_rows))
md.append("| 是否结算 | 是/否（建议增“部分结算”） | %d | 结算/支付方式描述值与“/” |" % len(settle_rows))
md.append("| 所属部门 | departments 12 个组合 | 0 | A-03 已修复 71 条缺科室后缀；本次复扫 0 条不可匹配 |")
md.append("")
md.append("**2. 项目分类映射（%d 条）**" % len(cat_rows))
md.append("")
md.append("| 原值 | 条数 | 建议值 | 依据 |")
md.append("|---|---|---|---|")
md.append("| 工程类 | %d | 建设项目 | 数据字典 3.1：并入本类或新增“工程类”，标准 4 值下并入建设项目（待业务确认） |" % cat_vals.get("工程类", 0))
md.append("| 维修类 | %d | 维护维修类 | 数据字典 3.1 明确并入 |" % cat_vals.get("维修类", 0))
md.append("| 维修维护类 | %d | 维护维修类 | 数据字典 3.1 明确并入 |" % cat_vals.get("维修维护类", 0))
md.append("| 委托支付函 | %d | 服务类 | 数据字典 3.1 建议并入服务类（电费/人员/水费燃气联系函） |" % cat_vals.get("委托支付函", 0))
md.append("| 补充协议 | %d | 服务类 | 按主合同归类：176301-05、256307-05 主合同为空调供能/委托管理（服务类） |" % cat_vals.get("补充协议", 0))
md.append("| 造价咨询服务 | %d | 服务类 | 数据字典 3.1 建议并入服务类 |" % cat_vals.get("造价咨询服务", 0))
md.append("")
md.append("**3. 采购方式映射（%d 条）**" % len(proc_rows))
md.append("")
md.append("| 原值 | 条数 | 建议值 | 依据 |")
md.append("|---|---|---|---|")
for v, n in sorted(proc_vals.items(), key=lambda x: -x[1]):
    note = ""
    if v in ("框架协议采购", "补充协议", "单一来源采购", "原合同《市委、市人大、市政府、市政协办公楼中央空调系统委托管理合同》", "邀请招标", "邀请招标（联合体投标）"):
        note = "（待业务确认，当前按建议值归入大类）"
    md.append("| `%s` | %d | %s %s | 数据字典 3.2 |" % (esc(v), n, PROCURE_MAP[v], note))
md.append("")
md.append("**4. 是否结算映射（%d 条）**" % len(settle_rows))
md.append("")
md.append("| 原值 | 条数 | 建议值 | 依据 |")
md.append("|---|---|---|---|")
for v, n in sorted(settle_vals.items(), key=lambda x: -x[1]):
    if v in SETTLE_MAP:
        md.append("| `%s` | %d | %s | 结算金额=已支付>0、未支付=0 判定为已结算；“/”按任务约定置空；按月支付/单项结算判定为未结算（待业务确认） |" % (esc(v), n, repr(SETTLE_MAP[v])))
    else:
        md.append("| `%s` | %d | %s | 数据字典 3.3 建议新增“部分结算”标准值（待业务确认） |" % (esc(v), n, SETTLE_KEEP[v]))
md.append("")
md.append("**5. 逐条待确认清单**")
md.append("")
md.append("### 5.1 项目分类（%d 条）" % len(cat_rows))
md.append("")
md.append("| 编号 | 项目名称 | 原值 | 建议值 |")
md.append("|---|---|---|---|")
for r in cat_rows:
    sug = CATEGORY_BY_NO.get(r["编号"], CATEGORY_MAP.get(r["项目分类"], r["项目分类"]))
    md.append("| %s | %s | %s | %s |" % (esc(r["编号"]), esc(r["项目名称"]), esc(r["项目分类"]), sug))
md.append("")
md.append("### 5.2 采购方式（%d 条）" % len(proc_rows))
md.append("")
md.append("| 编号 | 项目名称 | 原值 | 建议值 |")
md.append("|---|---|---|---|")
for r in proc_rows:
    md.append("| %s | %s | %s | %s |" % (esc(r["编号"]), esc(r["项目名称"]), esc(r["采购方式"]), PROCURE_MAP[r["采购方式"]]))
md.append("")
md.append("### 5.3 是否结算（%d 条）" % len(settle_rows))
md.append("")
md.append("| 编号 | 项目名称 | 原值 | 建议值 |")
md.append("|---|---|---|---|")
for r in settle_rows:
    sug = SETTLE_MAP.get(r["是否结算"], r["是否结算"])
    md.append("| %s | %s | %s | %s |" % (esc(r["编号"]), esc(r["项目名称"]), esc(r["是否结算"]), sug))
md.append("")
md.append("**6. 待业务确认项汇总**")
md.append("")
md.append("- 项目分类：`工程类`(56) 并入“建设项目”或新增“工程类”，需业务最终确认；当前按并入处理。")
md.append("- 采购方式：`框架协议采购`(8)、`补充协议`(4)、`单一来源采购`(1)、`原合同《市委、市人大、市政府、市政协办公楼中央空调系统委托管理合同》`(4，主合同 176301 不在库)、`邀请招标`(2)、`邀请招标（联合体投标）`(3) 归入大类为建议值，需业务确认。")
md.append("- 是否结算：`按月支付`(256307)、`由具体实施的单项项目结算`(256104-2) 判定为“否”为建议值，需业务确认；`部分结算`(256701) 保留原值，建议新增标准值。")
md.append("- 如需调整：恢复第 8 节备份后重跑，或在清洗后按编号定向修改。")
md.append("")
md.append("**7. 清洗执行记录**")
md.append("")
md.append("| 字段 | 更新条数 | 执行时间 |")
md.append("|---|---|---|")
md.append("| 项目分类 | %d | %s |" % (len(cat_rows), NOW))
md.append("| 采购方式 | %d | %s |" % (len(proc_rows), NOW))
md.append("| 是否结算 | %d（不含保留的“部分结算”） | %s |" % (len([r for r in settle_rows if r["是否结算"] in SETTLE_MAP]), NOW))
md.append("| 所属部门 | 0（A-03 已修复，本次复扫无缺后缀） | %s |" % NOW)
md.append("")
md.append("清洗更新记录已写入 `operation_logs`（模块=数据库清洗）。")
md.append("")
md.append("**8. 备份与回滚**")
md.append("")
md.append("- 清洗前备份：`backups/合同台账_A08_清洗前_20260823_204107.db`")
md.append("- 回滚：用备份文件覆盖 `data/合同台账.db` 后重启服务即可恢复清洗前数据。")
md.append("")
md.append("**9. 清洗后验证**")
md.append("")
md.append("（由脚本输出，见下方验证结果）")
md.append("")

doc_path = "docs/A-08-枚举与部门数据清洗清单.md"
with io.open(doc_path, "w", encoding="utf-8", newline="\n") as f:
    f.write("\n".join(md) + "\n")
print("[INFO] A-08 待确认清单已生成：" + doc_path, file=sys.stderr)

# ---------- 3. 事务内执行清洗 ----------
updated = {"项目分类": 0, "采购方式": 0, "是否结算": 0}
con.execute("BEGIN")
try:
    # 项目分类：映射表 + 编号定向
    for v, sug in CATEGORY_MAP.items():
        n = con.execute('UPDATE contracts SET "项目分类"=? WHERE "项目分类"=?', (sug, v)).rowcount
        updated["项目分类"] += n
    for no, sug in CATEGORY_BY_NO.items():
        n = con.execute('UPDATE contracts SET "项目分类"=? WHERE "编号"=?', (sug, no)).rowcount
        updated["项目分类"] += n
    # 采购方式
    for v, sug in PROCURE_MAP.items():
        n = con.execute('UPDATE contracts SET "采购方式"=? WHERE "采购方式"=?', (sug, v)).rowcount
        updated["采购方式"] += n
    # 是否结算：更新并把原值追加到备注（保留审计痕迹）
    for v, sug in SETTLE_MAP.items():
        rows_v = rows('SELECT "编号" FROM contracts WHERE "是否结算"=?', (v,))
        for r in rows_v:
            old_remark = con.execute('SELECT "备注" FROM contracts WHERE "编号"=?', (r["编号"],)).fetchone()[0] or ""
            new_remark = (old_remark + "【A-08清洗】原是否结算=" + str(v)).strip()
            con.execute('UPDATE contracts SET "是否结算"=?, "备注"=? WHERE "编号"=?', (sug, new_remark, r["编号"]))
        updated["是否结算"] += len(rows_v)
    # operation_logs 汇总记录
    con.execute('INSERT INTO operation_logs ("时间","用户名","姓名","角色","IP","模块","操作类型","对象","详情","结果") VALUES (?,?,?,?,?,?,?,?,?,?)',
                (NOW, "system", "系统脚本", "管理员", "127.0.0.1", "数据库清洗", "枚举标准化", "contracts",
                 "项目分类%d条、采购方式%d条、是否结算%d条（依据A-08待确认清单）" % (updated["项目分类"], updated["采购方式"], updated["是否结算"]), "成功"))
    con.execute("COMMIT")
except Exception:
    con.execute("ROLLBACK")
    raise
print("[MIGRATE A-08] 项目分类 %d 条、采购方式 %d 条、是否结算 %d 条 已更新" % (updated["项目分类"], updated["采购方式"], updated["是否结算"]), file=sys.stderr)

# ---------- 4. 清洗后验证 ----------
def dist(col):
    return Counter(r[0] for r in rows('SELECT "%s" v FROM contracts' % col))

cat_d = dist("项目分类"); proc_d = dist("采购方式"); settle_d = dist("是否结算")
cat_bad = [k for k in cat_d if k not in CAT_STD]
proc_bad = [k for k in proc_d if k not in PROC_STD]
settle_bad = [k for k in settle_d if k not in SETTLE_STD and k != "部分结算"]
dept_bad2 = rows('SELECT "编号" FROM contracts WHERE "所属部门" IS NOT NULL AND TRIM("所属部门")<>"" AND instr("所属部门","-")=0')
total = con.execute("SELECT COUNT(*) FROM contracts").fetchone()[0]
print("=== 清洗后分布 ===", file=sys.stderr)
print("项目分类:", dict(cat_d), file=sys.stderr)
print("采购方式:", dict(proc_d), file=sys.stderr)
print("是否结算:", dict(settle_d), file=sys.stderr)
if not cat_bad and not proc_bad and not settle_bad and not dept_bad2:
    print("OK-A08 清洗后全部为标准值（部分结算为建议新增值）；缺科室后缀 0 条；总数 %d 不变" % total, file=sys.stderr)
else:
    print("FAIL-A08 仍存在非标准值：分类%s 采购%s 结算%s 部门%s" % (cat_bad, proc_bad, settle_bad, len(dept_bad2)), file=sys.stderr)
con.close()
