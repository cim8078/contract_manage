import os,sys,shutil,hashlib,datetime,uuid,sqlite3,threading,time,json
import urllib.request
import contract_no
from functools import wraps
from io import BytesIO
from flask import Flask,jsonify,request,session,send_from_directory,send_file,Response
import openpyxl
from openpyxl.styles import Font,Alignment,PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

app = Flask(__name__)
app.secret_key = "xtsgj_contract_manager_secret_key_2026"
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DATA_DIR = os.path.join(BASE_DIR,"data")
BACKUP_DIR = os.path.join(BASE_DIR,"backups")
STATIC_DIR = os.path.join(BASE_DIR,"static")
EXCEL_PATH = os.path.join(DATA_DIR,"\u5408\u540c\u53f0\u8d26.xlsx")
DB_PATH = os.path.join(DATA_DIR,"合同台账.db")
ATTACH_DIR = os.path.join(DATA_DIR,"attachments")
ATTACH_SHEET = "\u5408\u540c\u9644\u4ef6"
ATTACH_HEADERS = ["\u9644\u4ef6ID","\u5408\u540c\u7f16\u53f7","\u9644\u4ef6\u7c7b\u578b","\u663e\u793a\u540d\u79f0","\u5b58\u50a8\u6587\u4ef6\u540d","\u6587\u4ef6\u5927\u5c0f","\u4e0a\u4f20\u4eba","\u4e0a\u4f20\u65f6\u95f4"]
ATTACH_TYPES = {"decision":"\u51b3\u7b56\u4f9d\u636e\u9644\u4ef6","contract":"\u5408\u540c\u9644\u4ef6","acceptance":"\u9a8c\u6536\u7ed3\u7b97\u9644\u4ef6"}
ATTACH_ALLOWED_EXT = {".jpg",".jpeg",".png",".gif",".webp",".bmp",".pdf"}
DEFAULT_ATTACH_MAX_MB = 20
SALT = "xtsgj_2024_contract"

for d in [DATA_DIR,BACKUP_DIR,STATIC_DIR,ATTACH_DIR,os.path.join(BASE_DIR,"templates")]:
    os.makedirs(d,exist_ok=True)

CONTRACT_COLUMNS = [
    "\u7f16\u53f7","\u9879\u76ee\u5206\u7c7b","\u9879\u76ee\u540d\u79f0","\u51b3\u7b56\u4f9d\u636e",
    "\u51b3\u7b56\u7684\u62db\u6807\u63a7\u5236\u4ef7","\u62db\u6807\u63a7\u5236\u4ef7\u51b3\u7b56\u7684\u4f9d\u636e",
    "\u91c7\u8d2d\u65b9\u5f0f","\u4e2d\u6807\u5355\u4f4d","\u5b9e\u9645\u5b9e\u65bd\u5355\u4f4d",
    "\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09","\u8d44\u91d1\u6765\u6e90","\u5408\u540c\u4e3b\u8981\u5185\u5bb9",
    "\u91c7\u8d2d\u65f6\u95f4","\u5f00\u5de5\u65f6\u95f4","\u5b8c\u5de5\u65f6\u95f4",
    "\u7ae3\u5de5\u9a8c\u6536\u65f6\u95f4","\u73b0\u573a\u8d1f\u8d23\u4eba",
    "\u662f\u5426\u6709\u7b7e\u8bc1\u6216\u53d8\u66f4","\u7b7e\u8bc1\u6216\u53d8\u66f4\u60c5\u51b5",
    "\u662f\u5426\u7ed3\u7b97","\u7ed3\u7b97\u91d1\u989d","\u5df2\u652f\u4ed8","\u672a\u652f\u4ed8",
    "\u671f\u9650","\u7ecf\u529e\u4eba","\u5907\u6ce8",
    "\u6240\u5c5e\u90e8\u95e8","\u4e0b\u6b21\u62db\u91c7\u65f6\u95f4","\u63d0\u9192\u63d0\u524d\u5929\u6570",
]

COL_IDX = {n:i for i,n in enumerate(CONTRACT_COLUMNS)}

# ──── 资产合同台账（asset_contracts）────
# 依据《资产合同台账.xlsx》（湘潭市国有资产租金缴纳情况明细表，2026 年度）建表；
# 列名/类型/约束参照 project_contracts 表约定：金额 REAL CHECK(>=0)、日期 TEXT CHECK(yyyy-mm-dd)、必填列 NOT NULL。
# 说明：原表“已缴租金”区域共 10 列已于 v13 删除；合同级“缴款方式”列已于 v15 删除，
#       缴款方式改为租金收取明细（asset_rent_collections）上的枚举字段（直缴非税/现金/微信/其他）；
#       租金收缴明细改由 asset_rent_collections 按“收款期次”一对多记录。
ASSET_CONTRACTS_DDL = """
CREATE TABLE IF NOT EXISTS asset_contracts (
  id TEXT PRIMARY KEY NOT NULL,   -- 6 位合同编号（年度后两位+科室序号+中心组室序号+顺序号），见 contract_no；v26 起不再自增
  "位置" TEXT NOT NULL, "序号" INTEGER, "资产名称" TEXT NOT NULL, "承租人" TEXT NOT NULL, "联系电话" TEXT,
  "管理部门" TEXT, "责任人" TEXT,
  "面积（㎡）" REAL CHECK ("面积（㎡）" IS NULL OR CAST("面积（㎡）" AS REAL) >= 0),
  "租期开始日" TEXT CHECK ("租期开始日" IS NULL OR "租期开始日"='' OR "租期开始日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "租期结束日" TEXT CHECK ("租期结束日" IS NULL OR "租期结束日"='' OR "租期结束日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "押金（元）" REAL CHECK ("押金（元）" IS NULL OR CAST("押金（元）" AS REAL) >= 0),
  "年租金（元）" REAL CHECK ("年租金（元）" IS NULL OR CAST("年租金（元）" AS REAL) >= 0),
  "季租（元）" REAL CHECK ("季租（元）" IS NULL OR CAST("季租（元）" AS REAL) >= 0),
  "备注" TEXT,
  "是否归档" TEXT, "归档人" TEXT, "归档时间" TEXT,
  "下次出租时间" DATE, "总合同额" DECIMAL(12,2) DEFAULT 0,
  "提醒提前天数" INTEGER
);
"""

# asset_contracts 常用查询列索引（参照 A-04 对 project_contracts 的处理，幂等）
ASSET_CONTRACTS_INDEXES = [
    'CREATE INDEX IF NOT EXISTS idx_asset_contracts_location ON asset_contracts("位置")',
    'CREATE INDEX IF NOT EXISTS idx_asset_contracts_tenant ON asset_contracts("承租人")',
    'CREATE INDEX IF NOT EXISTS idx_asset_contracts_end_date ON asset_contracts("租期结束日")',
    'CREATE INDEX IF NOT EXISTS idx_asset_contracts_dept ON asset_contracts("管理部门")',
    'CREATE INDEX IF NOT EXISTS idx_asset_contracts_person ON asset_contracts("责任人")',
    'CREATE INDEX IF NOT EXISTS idx_asset_contracts_archived ON asset_contracts("是否归档")',
]

# ──── 资产合同租金收缴（asset_rent_collections：一个资产合同对应多条收款记录，一期一行）────
# 接口字段（英文）↔ 数据库列名（中文）映射（供后续接口/页面接线使用；本步骤只建表）
ASSET_RENT_FIELDS = [
    ("asset_id", "资产合同ID"), ("period_label", "收款期次"),
    ("period_start", "所属租期开始日"), ("period_end", "所属租期结束日"),
    ("due_amount", "应缴金额（元）"), ("paid_amount", "实缴金额（元）"),
    ("pay_date", "缴款日期"), ("pay_method", "缴款方式"), ("payee", "收款人"),
]
ASSET_RENT_KEY_TO_COL = {k: cn for k, cn in ASSET_RENT_FIELDS}
ASSET_RENT_DDL = """
CREATE TABLE IF NOT EXISTS asset_rent_collections (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  "资产合同ID" INTEGER NOT NULL,
  "收款期次" TEXT,
  "所属租期开始日" TEXT CHECK ("所属租期开始日" IS NULL OR "所属租期开始日"='' OR "所属租期开始日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "所属租期结束日" TEXT CHECK ("所属租期结束日" IS NULL OR "所属租期结束日"='' OR "所属租期结束日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "应缴金额（元）" REAL NOT NULL DEFAULT 0 CHECK ("应缴金额（元）" IS NULL OR CAST("应缴金额（元）" AS REAL) >= 0),
  "实缴金额（元）" REAL NOT NULL DEFAULT 0 CHECK ("实缴金额（元）" IS NULL OR CAST("实缴金额（元）" AS REAL) >= 0),
  "缴款日期" TEXT CHECK ("缴款日期" IS NULL OR "缴款日期"='' OR "缴款日期" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "缴款方式" TEXT DEFAULT '直缴非税' CHECK ("缴款方式" IS NULL OR "缴款方式"='' OR "缴款方式" IN ('直缴非税','现金','微信','其他')),
  "收款人" TEXT,
  FOREIGN KEY ("资产合同ID") REFERENCES asset_contracts(id) ON UPDATE CASCADE ON DELETE RESTRICT
);
"""
ASSET_RENT_INDEXES = [
    'CREATE INDEX IF NOT EXISTS idx_asset_rent_collections_contract ON asset_rent_collections("资产合同ID")',
    'CREATE INDEX IF NOT EXISTS idx_asset_rent_collections_pay_date ON asset_rent_collections("缴款日期")',
    'CREATE INDEX IF NOT EXISTS idx_asset_rent_collections_period_start ON asset_rent_collections("所属租期开始日")',
    'CREATE INDEX IF NOT EXISTS idx_asset_rent_collections_period_label ON asset_rent_collections("收款期次")',
]
# 租金收取明细：校验规则（列名 = 接口字段 → 数据库列）
ASSET_RENT_PAY_METHODS = ("直缴非税", "现金", "微信", "其他")   # 缴款方式枚举（v16 起由 CHECK 约束保证）
ASSET_RENT_REQUIRED = ("period_label", "due_amount")          # 收款期次、应缴金额（元）必填
ASSET_RENT_NUM_KEYS = ("due_amount", "paid_amount")           # 金额：数字、≥0、保留两位
ASSET_RENT_DATE_KEYS = ("period_start", "period_end", "pay_date")   # 日期：yyyy-mm-dd，可空
ASSET_RENT_MAX_ROWS = 200                                     # 单个合同最多保存的收款记录数

# ──── v13：删除 asset_contracts 旧租金收缴列（重建表用常量）────
ASSET_CONTRACT_KEEP_COLS = (
    "id", "位置", "序号", "资产名称", "承租人", "联系电话", "管理部门", "责任人",
    "面积（㎡）", "租期开始日", "租期结束日", "押金（元）", "缴款方式", "年租（元）", "季租（元）",
    "备注", "是否归档", "归档人", "归档时间", "下次出租时间", "总合同额", "提醒提前天数",
)
ASSET_CONTRACT_LEGACY_RENT_COLS = (
    "已缴租金1季度（元）", "已缴租金1季度交款时间",
    "已缴租金2季度（元）", "已缴租金2季度交款时间",
    "已缴租金3季度（元）", "已缴租金3季度交款时间",
    "已缴租金4季度（元）", "已缴租金4季度交款时间",
    "已缴租金开票对应租期", "已缴租金合计（元）",
)
ASSET_CONTRACTS_NEW_TABLE = "asset_contracts_new"
# 重建用建表语句：列定义与 ASSET_CONTRACTS_DDL 保留列逐字一致（主键 AUTOINCREMENT、金额 CHECK ≥ 0、
# 日期 CHECK yyyy-mm-dd、总合同额 DEFAULT 0），仅表名换为临时表名；属 v13 历史版本冻结内容。
ASSET_CONTRACTS_REBUILD_DDL = """
CREATE TABLE asset_contracts_new (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  "位置" TEXT NOT NULL, "序号" INTEGER, "资产名称" TEXT NOT NULL, "承租人" TEXT NOT NULL, "联系电话" TEXT,
  "管理部门" TEXT, "责任人" TEXT,
  "面积（㎡）" REAL CHECK ("面积（㎡）" IS NULL OR CAST("面积（㎡）" AS REAL) >= 0),
  "租期开始日" TEXT CHECK ("租期开始日" IS NULL OR "租期开始日"='' OR "租期开始日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "租期结束日" TEXT CHECK ("租期结束日" IS NULL OR "租期结束日"='' OR "租期结束日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "押金（元）" REAL CHECK ("押金（元）" IS NULL OR CAST("押金（元）" AS REAL) >= 0),
  "缴款方式" TEXT,
  "年租（元）" REAL CHECK ("年租（元）" IS NULL OR CAST("年租（元）" AS REAL) >= 0),
  "季租（元）" REAL CHECK ("季租（元）" IS NULL OR CAST("季租（元）" AS REAL) >= 0),
  "备注" TEXT,
  "是否归档" TEXT, "归档人" TEXT, "归档时间" TEXT,
  "下次出租时间" DATE, "总合同额" DECIMAL(12,2) DEFAULT 0,
  "提醒提前天数" INTEGER
);
"""

# ──── v15：删除 asset_contracts 的“缴款方式”列（重建表用常量，冻结内容）────
ASSET_CONTRACTS_V15_KEEP_COLS = (
    "id", "位置", "序号", "资产名称", "承租人", "联系电话", "管理部门", "责任人",
    "面积（㎡）", "租期开始日", "租期结束日", "押金（元）", "年租金（元）", "季租（元）",
    "备注", "是否归档", "归档人", "归档时间", "下次出租时间", "总合同额", "提醒提前天数",
)
ASSET_CONTRACTS_V15_DDL = """
CREATE TABLE asset_contracts_new (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  "位置" TEXT NOT NULL, "序号" INTEGER, "资产名称" TEXT NOT NULL, "承租人" TEXT NOT NULL, "联系电话" TEXT,
  "管理部门" TEXT, "责任人" TEXT,
  "面积（㎡）" REAL CHECK ("面积（㎡）" IS NULL OR CAST("面积（㎡）" AS REAL) >= 0),
  "租期开始日" TEXT CHECK ("租期开始日" IS NULL OR "租期开始日"='' OR "租期开始日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "租期结束日" TEXT CHECK ("租期结束日" IS NULL OR "租期结束日"='' OR "租期结束日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "押金（元）" REAL CHECK ("押金（元）" IS NULL OR CAST("押金（元）" AS REAL) >= 0),
  "年租金（元）" REAL CHECK ("年租金（元）" IS NULL OR CAST("年租金（元）" AS REAL) >= 0),
  "季租（元）" REAL CHECK ("季租（元）" IS NULL OR CAST("季租（元）" AS REAL) >= 0),
  "备注" TEXT,
  "是否归档" TEXT, "归档人" TEXT, "归档时间" TEXT,
  "下次出租时间" DATE, "总合同额" DECIMAL(12,2) DEFAULT 0,
  "提醒提前天数" INTEGER
);
"""
# ──── v16：asset_rent_collections.“缴款方式”枚举化（重建表用常量，冻结内容）────
ASSET_RENT_V16_DDL = """
CREATE TABLE asset_rent_collections_new (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  "资产合同ID" INTEGER NOT NULL,
  "收款期次" TEXT,
  "所属租期开始日" TEXT CHECK ("所属租期开始日" IS NULL OR "所属租期开始日"='' OR "所属租期开始日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "所属租期结束日" TEXT CHECK ("所属租期结束日" IS NULL OR "所属租期结束日"='' OR "所属租期结束日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "应缴金额（元）" REAL NOT NULL DEFAULT 0 CHECK ("应缴金额（元）" IS NULL OR CAST("应缴金额（元）" AS REAL) >= 0),
  "实缴金额（元）" REAL NOT NULL DEFAULT 0 CHECK ("实缴金额（元）" IS NULL OR CAST("实缴金额（元）" AS REAL) >= 0),
  "缴款日期" TEXT CHECK ("缴款日期" IS NULL OR "缴款日期"='' OR "缴款日期" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "缴款方式" TEXT CHECK ("缴款方式" IS NULL OR "缴款方式"='' OR "缴款方式" IN ('直缴非税','现金','微信','其他')),
  "收款人" TEXT,
  FOREIGN KEY ("资产合同ID") REFERENCES asset_contracts(id) ON UPDATE CASCADE ON DELETE RESTRICT
);
"""

# ──── v17：asset_rent_collections.“缴款方式”加默认值 '直缴非税'（重建表用常量，冻结内容）────
ASSET_RENT_V17_DDL = """
CREATE TABLE asset_rent_collections_new (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  "资产合同ID" INTEGER NOT NULL,
  "收款期次" TEXT,
  "所属租期开始日" TEXT CHECK ("所属租期开始日" IS NULL OR "所属租期开始日"='' OR "所属租期开始日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "所属租期结束日" TEXT CHECK ("所属租期结束日" IS NULL OR "所属租期结束日"='' OR "所属租期结束日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "应缴金额（元）" REAL NOT NULL DEFAULT 0 CHECK ("应缴金额（元）" IS NULL OR CAST("应缴金额（元）" AS REAL) >= 0),
  "实缴金额（元）" REAL NOT NULL DEFAULT 0 CHECK ("实缴金额（元）" IS NULL OR CAST("实缴金额（元）" AS REAL) >= 0),
  "缴款日期" TEXT CHECK ("缴款日期" IS NULL OR "缴款日期"='' OR "缴款日期" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "缴款方式" TEXT DEFAULT '直缴非税' CHECK ("缴款方式" IS NULL OR "缴款方式"='' OR "缴款方式" IN ('直缴非税','现金','微信','其他')),
  "收款人" TEXT,
  FOREIGN KEY ("资产合同ID") REFERENCES asset_contracts(id) ON UPDATE CASCADE ON DELETE RESTRICT
);
"""

# ──── 资产合同台账字段映射（接口字段名 ↔ 数据库列名；顺序即页面/导出的列顺序）────

# ──── v26：asset_contracts.id 由“自增整数”改为“6 位合同编号”（TEXT；重建表用常量，冻结内容）────
# 说明：6 位编号为字符串，含年度后两位 + 科室（中心）序号 + 中心组室序号 + 顺序号；
#       历史数据的 id 保持原值（重建时按列逐一搬运，不重编、不修改）。
ASSET_CONTRACTS_V26_KEEP_COLS = (
    "id", "位置", "序号", "资产名称", "承租人", "联系电话", "管理部门", "责任人",
    "面积（㎡）", "租期开始日", "租期结束日", "押金（元）", "年租金（元）", "季租（元）",
    "备注", "是否归档", "归档人", "归档时间", "下次出租时间", "总合同额", "提醒提前天数", "短信提醒",
)
ASSET_CONTRACTS_V26_DDL = """
CREATE TABLE asset_contracts_new (
  id TEXT PRIMARY KEY NOT NULL,
  "位置" TEXT NOT NULL, "序号" INTEGER, "资产名称" TEXT NOT NULL, "承租人" TEXT NOT NULL, "联系电话" TEXT,
  "管理部门" TEXT, "责任人" TEXT,
  "面积（㎡）" REAL CHECK ("面积（㎡）" IS NULL OR CAST("面积（㎡）" AS REAL) >= 0),
  "租期开始日" TEXT CHECK ("租期开始日" IS NULL OR "租期开始日"='' OR "租期开始日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "租期结束日" TEXT CHECK ("租期结束日" IS NULL OR "租期结束日"='' OR "租期结束日" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "押金（元）" REAL CHECK ("押金（元）" IS NULL OR CAST("押金（元）" AS REAL) >= 0),
  "年租金（元）" REAL CHECK ("年租金（元）" IS NULL OR CAST("年租金（元）" AS REAL) >= 0),
  "季租（元）" REAL CHECK ("季租（元）" IS NULL OR CAST("季租（元）" AS REAL) >= 0),
  "备注" TEXT,
  "是否归档" TEXT, "归档人" TEXT, "归档时间" TEXT,
  "下次出租时间" DATE, "总合同额" DECIMAL(12,2) DEFAULT 0,
  "提醒提前天数" INTEGER,
  "短信提醒" TEXT
);
"""

ASSET_FIELDS = [
    ("place", "位置"), ("seq", "序号"), ("asset_name", "资产名称"), ("tenant", "承租人"), ("phone", "联系电话"),
    ("dept", "管理部门"), ("person", "责任人"),
    ("area", "面积（㎡）"), ("lease_start", "租期开始日"), ("lease_end", "租期结束日"),
    ("deposit", "押金（元）"), ("annual_rent", "年租金（元）"), ("quarter_rent", "季租（元）"),
    ("remark", "备注"),
    ("next_lease_date", "下次出租时间"), ("total_amount", "总合同额"),
    ("remind_days", "提醒提前天数"),
]
ASSET_KEY_TO_COL = {k: cn for k, cn in ASSET_FIELDS}
# 说明：原“已缴租金1~4季度（元）/交款时间/已缴租金开票对应租期/已缴租金合计（元）”10 列已于 v13 删除，
#       租金收取明细改由子表 asset_rent_collections 一对多管理（见 ASSET_RENT_* 常量与 /rent_collections 接口）；
#       列表与统计中的“已缴租金合计”为子表实缴金额汇总（见 read_assets 计算列 paid_total），不再是主表列。
ASSET_NUM_KEYS = ("area", "deposit", "annual_rent", "quarter_rent", "total_amount")
ASSET_DATE_KEYS = ("lease_start", "lease_end", "next_lease_date")
ASSET_REQUIRED = ("place", "asset_name", "tenant")
ASSET_QUERY_KEYS = ("place", "tenant", "asset_name", "seq", "lease_end_start", "lease_end_end",
                    "dept", "person", "archived", "unit", "lease_start_start", "lease_start_end")
# 归档相关列：单独维护，不进入 ASSET_FIELDS（导入/模板/导出与字段校验均不含归档状态）
ASSET_ARCHIVE_COLS = ("是否归档", "归档人", "归档时间")

# ──── 资产合同附件（asset_attachments：按资产合同 id 关联，附件类型仅 决策依据/合同）────
ASSET_ATTACH_TYPES = {"condition": "资产现状图片", "decision": "决策依据附件", "contract": "合同附件"}
ASSET_ATTACH_HEADERS = ["附件ID", "资产合同ID", "附件类型", "显示名称", "存储文件名", "文件大小", "上传人", "上传时间"]
ASSET_ATTACHMENTS_DDL = """
CREATE TABLE IF NOT EXISTS asset_attachments (
  "附件ID" TEXT PRIMARY KEY,
  "资产合同ID" INTEGER NOT NULL,
  "附件类型" TEXT, "显示名称" TEXT, "存储文件名" TEXT, "文件大小" TEXT, "上传人" TEXT, "上传时间" TEXT,
  FOREIGN KEY ("资产合同ID") REFERENCES asset_contracts(id) ON DELETE RESTRICT ON UPDATE RESTRICT
);
"""
ASSET_ATTACHMENTS_INDEXES = [
    'CREATE INDEX IF NOT EXISTS idx_asset_attachments_contract ON asset_attachments("资产合同ID")',
]

# ──── 字段级变更日志（field_change_logs：记录“谁在何时把某字段从 A 改成 B”）────
FIELD_LOG_HEADERS = ["时间", "用户名", "姓名", "角色", "IP", "模块", "对象编号", "对象名称", "字段", "原值", "新值"]
FIELD_LOG_DDL = """
CREATE TABLE IF NOT EXISTS field_change_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  "时间" TEXT, "用户名" TEXT, "姓名" TEXT, "角色" TEXT, "IP" TEXT,
  "模块" TEXT, "对象编号" TEXT, "对象名称" TEXT, "字段" TEXT, "原值" TEXT, "新值" TEXT
);
"""
FIELD_LOG_INDEXES = [
    'CREATE INDEX IF NOT EXISTS idx_field_logs_object ON field_change_logs("对象编号")',
    'CREATE INDEX IF NOT EXISTS idx_field_logs_time ON field_change_logs("时间")',
]

# ──── 短信提醒（v20~v23）：配置项 / 模板 / 发送流水 / 合同开关 ────
# 说明：短信功能是“只读现有台账数据 + 写发送流水”的叠加层；本段仅建表、写默认配置与模板，
#       不含发送接口 / 定时任务 / 前端页面。
#       平台：湘潭市政务短信服务平台（对接文档 2024-11-08），接口前缀 https://sms.xtdzzw.cn/api。
#       收件人口径：项目合同=经办人+现场负责人；资产合同=责任人+承租人；
#       经办人/现场负责人/责任人手机号按姓名精确匹配 users.电话，承租人取 asset_contracts.联系电话。
SMS_SETTINGS_DEFAULTS = [
    ("sms_enabled", "0"),          # 短信功能总开关（默认关，联调通过后再开）
    ("sms_base_url", "https://sms.xtdzzw.cn/api"),
    ("sms_access_key", "1621c79cb95147598662"),
    ("sms_secret_key", "1621c79cb951475986621ad4baec2fdf"),
    ("sms_weekly_enabled", "1"),   # 每周自动提醒开关
    ("sms_weekly_weekday", "1"),   # 1=周一
    ("sms_weekly_hour", "9"),
    ("sms_weekly_minute", "0"),
    ("sms_max_per_run", "200"),    # 单次定时任务发送上限
    ("sms_retry_max", "1"),        # 失败重试次数
    ("sms_last_run_at", ""),       # 上次定时执行时间（防重复触发）
    ("sms_receipt_token", ""),     # 回执/回复回调校验口令（可选）
    ("sms_timeout_sec", "10"),     # 连接/读取超时（秒）
    ("sms_include_sign", "1"),     # 是否在正文末尾追加机构署名
    ("sms_sign_suffix", "（湘潭市机关事务管理局）"),  # 正文末尾署名文本（全角括号；平台禁止【】）
    ("sms_segment_len", "70"),     # 单条计费长度，超出自动拆分（平台注意事项 1）
    ("sms_max_segments", "3"),     # 单条内容最多拆分段数
    ("sms_block_check", "1"),      # 敏感词校验开关（平台注意事项 3）
    ("sms_block_words", "楼价,打折,特价,优惠,促销,秒杀,返现,中奖,免费领,贷款,发票,代开,加微信,扫码领取"),  # 敏感词（可维护）
    ("sms_daily_per_phone", "100"),# 同一号码每日发送上限（平台硬性要求 ≤100，注意事项 5）
    ("sms_min_interval_sec", "30"),# 同一号码两次发送最小间隔（秒，0=不限；注意事项 6）
]

# 短信模板：场景 expire=到期 / receivable=待收 / payable=待支 / urge=催办 / manual=点对点
SMS_TEMPLATES_DDL = """
CREATE TABLE IF NOT EXISTS sms_templates (
  "场景" TEXT PRIMARY KEY,
  "场景名称" TEXT NOT NULL,
  "模板内容" TEXT NOT NULL,
  "启用" TEXT DEFAULT '是',
  "更新时间" TEXT, "更新人" TEXT
);
"""
SMS_TEMPLATE_DEFAULTS = [
    ("expire", "合同到期提醒", "合同「{name}」（编号{number}）将于{due}到期，剩余{days}天，请及时办理。"),
    ("receivable", "待收提醒", "资产合同「{name}」（承租方{tenant}）尚有租金{amount}元未收讫，请及时处理。"),
    ("payable", "待支提醒", "项目合同「{name}」（编号{number}）尚有未支付金额{amount}万元，请及时办理支付。"),
    ("urge", "短信催办", "合同「{name}」（编号{number}）提醒事项请尽快办理：{note}"),
    ("manual", "点对点短信", "{content}"),
]

# 短信发送流水：一次发送一行；"去重键" 部分唯一索引保证同一提醒同一周同一号码仅发一条（每人一条）
SMS_LOGS_DDL = """
CREATE TABLE IF NOT EXISTS sms_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  "时间" TEXT, "场景" TEXT, "关联类型" TEXT, "关联ID" TEXT, "关联编号" TEXT, "关联名称" TEXT,
  "接收人类型" TEXT, "接收人" TEXT, "手机号" TEXT, "短信内容" TEXT, "去重键" TEXT,
  "触发方式" TEXT, "发送人" TEXT, "状态" TEXT,
  "外部guId" TEXT, "回执状态码" TEXT, "回执描述" TEXT, "回复时间" TEXT, "回复内容" TEXT,
  "错误信息" TEXT, "请求时间戳" INTEGER
);
"""
SMS_LOGS_INDEXES = [
    'CREATE INDEX IF NOT EXISTS idx_sms_logs_time ON sms_logs("时间")',
    'CREATE INDEX IF NOT EXISTS idx_sms_logs_object ON sms_logs("关联类型","关联ID")',
    'CREATE INDEX IF NOT EXISTS idx_sms_logs_phone ON sms_logs("手机号")',
    'CREATE INDEX IF NOT EXISTS idx_sms_logs_status ON sms_logs("状态")',
    'CREATE INDEX IF NOT EXISTS idx_sms_logs_guid ON sms_logs("外部guId")',
    "CREATE UNIQUE INDEX IF NOT EXISTS idx_sms_logs_dedup ON sms_logs(\"去重键\") "
    "WHERE \"去重键\" IS NOT NULL AND \"去重键\" <> ''",
]

# ──── SQLite I/O ────
db_lock = threading.Lock()

def _conn():
    con = sqlite3.connect(DB_PATH, timeout=15)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA busy_timeout=15000")
    con.execute("PRAGMA foreign_keys=ON")
    return con

def db_query(sql, params=()):
    con = _conn()
    try:
        cur = con.execute(sql, params)
        return [{k: r[k] for k in r.keys()} for r in cur.fetchall()]
    finally:
        con.close()

def db_exec(sql, params=()):
    with db_lock:
        _backup()
        con = _conn()
        try:
            cur = con.execute(sql, params)
            con.commit()
            _invalidate_caches()      # 写操作后清空只读缓存，避免首页/列表读到旧数据
            return cur
        except Exception:
            con.rollback()
            raise
        finally:
            con.close()

# ─── 只读结果缓存（首页/列表高频只读路径共用；TTL 秒级 + 写操作即时失效）───
# 背景：一次首页加载会分别调用 dashboard / reminders / pending_payments 三个接口，
# 每个接口都会 read_contracts()/read_assets() 各读一次全表；缓存后同一时间窗内只读一次库。
# 语义保证：缓存只保存“查库结果”，每次返回浅拷贝（调用方可自由往 items 上挂 attachments 等字段，
#          不会污染缓存）；任何写库操作都会立即失效，另有 TTL 兜底。
READ_CACHE_TTL = 2.0
_read_cache = {"contracts": (0.0, None), "assets": (0.0, None)}
_cache_lock = threading.Lock()

def _invalidate_caches():
    with _cache_lock:
        _read_cache["contracts"] = (0.0, None)
        _read_cache["assets"] = (0.0, None)

def _cache_get(key):
    with _cache_lock:
        ts, val = _read_cache[key]
    if val is not None and (time.time() - ts) < READ_CACHE_TTL:
        return val
    return None

def _cache_put(key, val):
    with _cache_lock:
        _read_cache[key] = (time.time(), val)
def _backup():
    if not os.path.exists(DB_PATH): return
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest_path = os.path.join(BACKUP_DIR, f"合同台账_{ts}.db")
    src = sqlite3.connect(DB_PATH, timeout=15)
    dst = sqlite3.connect(dest_path)
    try:
        src.backup(dst)
    finally:
        dst.close(); src.close()
    cutoff = datetime.datetime.now() - datetime.timedelta(days=30)
    for fn in os.listdir(BACKUP_DIR):
        fp = os.path.join(BACKUP_DIR, fn)
        if fn.endswith(".db") and datetime.datetime.fromtimestamp(os.path.getmtime(fp)) < cutoff:
            try: os.remove(fp)
            except: pass

_SCHEMA = """
CREATE TABLE IF NOT EXISTS project_contracts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  "编号" TEXT NOT NULL, "项目分类" TEXT NOT NULL, "项目名称" TEXT NOT NULL, "决策依据" TEXT,
  "决策的招标控制价" REAL CHECK ("决策的招标控制价" IS NULL OR CAST("决策的招标控制价" AS REAL) >= 0),
  "招标控制价决策的依据" TEXT, "采购方式" TEXT NOT NULL, "中标单位" TEXT, "实际实施单位" TEXT,
  "合同金额（万元）" REAL NOT NULL CHECK ("合同金额（万元）" IS NULL OR CAST("合同金额（万元）" AS REAL) >= 0),
  "资金来源" TEXT NOT NULL, "合同主要内容" TEXT NOT NULL,
  "采购时间" TEXT CHECK ("采购时间" IS NULL OR "采购时间" = '' OR "采购时间" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "开工时间" TEXT CHECK ("开工时间" IS NULL OR "开工时间" = '' OR "开工时间" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "完工时间" TEXT CHECK ("完工时间" IS NULL OR "完工时间" = '' OR "完工时间" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "竣工验收时间" TEXT CHECK ("竣工验收时间" IS NULL OR "竣工验收时间" = '' OR "竣工验收时间" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "现场负责人" TEXT,
  "是否有签证或变更" TEXT, "签证或变更情况" TEXT,
  "是否结算" TEXT, "结算金额" REAL CHECK ("结算金额" IS NULL OR CAST("结算金额" AS REAL) >= 0),
  "已支付" REAL CHECK ("已支付" IS NULL OR CAST("已支付" AS REAL) >= 0),
  "未支付" REAL CHECK ("未支付" IS NULL OR CAST("未支付" AS REAL) >= 0),
  "期限" TEXT, "经办人" TEXT NOT NULL, "备注" TEXT,
  "所属部门" TEXT NOT NULL,
  "下次招采时间" TEXT CHECK ("下次招采时间" IS NULL OR "下次招采时间" = '' OR "下次招采时间" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'),
  "提醒提前天数" INTEGER,
  "是否归档" TEXT, "归档人" TEXT, "归档时间" TEXT,
  "所属单位" TEXT, "所属科室" TEXT,
  -- A-05 枚举 CHECK（项目分类/采购方式/是否结算/资金来源/是否归档）延后至 A-08 清洗完成后启用
  FOREIGN KEY ("所属单位","所属科室") REFERENCES departments("单位名称","科室名称") ON DELETE RESTRICT ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS users (
  "用户名" TEXT PRIMARY KEY, "密码" TEXT, "姓名" TEXT, "所属部门" TEXT, "角色" TEXT, "分管部门" TEXT,
  "所属单位" TEXT, "所属科室" TEXT, "电话" TEXT,
  FOREIGN KEY ("所属单位","所属科室") REFERENCES departments("单位名称","科室名称") ON DELETE RESTRICT ON UPDATE CASCADE
);
CREATE TABLE IF NOT EXISTS departments (
  "单位名称" TEXT, "科室名称" TEXT
);
CREATE TABLE IF NOT EXISTS settings (
  "配置项" TEXT PRIMARY KEY, "配置值" TEXT
);
CREATE TABLE IF NOT EXISTS project_attachments (
  "附件ID" TEXT PRIMARY KEY, "合同编号" TEXT, "附件类型" TEXT, "显示名称" TEXT,
  "存储文件名" TEXT, "文件大小" TEXT, "上传人" TEXT, "上传时间" TEXT,
  FOREIGN KEY ("合同编号") REFERENCES project_contracts("编号") ON DELETE RESTRICT ON UPDATE RESTRICT
);
CREATE TABLE IF NOT EXISTS operation_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  "时间" TEXT, "用户名" TEXT, "姓名" TEXT, "角色" TEXT, "IP" TEXT,
  "模块" TEXT, "操作类型" TEXT, "对象" TEXT, "详情" TEXT, "结果" TEXT
);
""" + ASSET_CONTRACTS_DDL + ASSET_ATTACHMENTS_DDL + FIELD_LOG_DDL + ASSET_RENT_DDL \
    + SMS_TEMPLATES_DDL + SMS_LOGS_DDL

# ─── 数据库结构版本化迁移（schema_version 记录已应用版本；迁移函数须保持幂等）───
def _contracts_table(con):
    """合同主表当前表名：改名后为 project_contracts；仍为旧名 contracts 的历史库按旧名处理。"""
    names = set(r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall())
    return "project_contracts" if "project_contracts" in names else "contracts"

def _mig_v1_contract_extra_cols(con):
    """v1：合同主表追加历史字段列（资金来源/是否归档/归档人/归档时间），幂等（存在即跳过）。"""
    tbl = _contracts_table(con)
    cols = [r[1] for r in con.execute("PRAGMA table_info(" + tbl + ")").fetchall()]
    for cname, ctype in (("资金来源","TEXT"),("是否归档","TEXT"),("归档人","TEXT"),("归档时间","TEXT")):
        if cname not in cols:
            con.execute('ALTER TABLE ' + tbl + ' ADD COLUMN "' + cname + '" ' + ctype)

def _mig_v2_contract_rename_cols(con):
    """v2：合同主表历史列改名（经办人/分管领导→经办人、甲方现场负责人→现场负责人），幂等。"""
    tbl = _contracts_table(con)
    cols = [r[1] for r in con.execute("PRAGMA table_info(" + tbl + ")").fetchall()]
    if "经办人/分管领导" in cols and "经办人" not in cols:
        con.execute('ALTER TABLE ' + tbl + ' RENAME COLUMN "经办人/分管领导" TO "经办人"')
    if "甲方现场负责人" in cols and "现场负责人" not in cols:
        con.execute('ALTER TABLE ' + tbl + ' RENAME COLUMN "甲方现场负责人" TO "现场负责人"')

def _mig_v3_users_supervise_dept(con):
    """v3：users 追加 分管部门 字段，幂等。"""
    ucols = [r[1] for r in con.execute("PRAGMA table_info(users)").fetchall()]
    if "分管部门" not in ucols:
        con.execute('ALTER TABLE users ADD COLUMN "分管部门" TEXT')

def _ensure_asset_contracts(con):
    """建立资产合同台账表 asset_contracts 及其索引（幂等：已存在则跳过）。"""
    con.execute(ASSET_CONTRACTS_DDL)
    for sql in ASSET_CONTRACTS_INDEXES:
        con.execute(sql)

def _mig_v4_asset_contracts(con):
    """v4：新增资产合同台账表 asset_contracts（依据《资产合同台账.xlsx》），幂等。"""
    _ensure_asset_contracts(con)

def _mig_v5_rename_contracts_table(con):
    """v5：合同主表改名 contracts → project_contracts，并同步索引/触发器名（幂等）。

    - SQLite 的 ALTER TABLE ... RENAME TO 会自动改写外键子句（项目附件表的外键随之
      指向新表名）以及索引/触发器中对该表的引用，故外键列与触发器体无需手工维护。
    - 索引/触发器名含 contracts 前缀，SQLite 不支持改名，按其原定义重建后删除旧名。
    - 旧库首次启动时 _SCHEMA 会先按新表名建出空表（IF NOT EXISTS），故此处识别
      “旧表可读、新表为空且无自定义索引”的情形：删除该空表后再改名，旧库数据随之落到新表名下。
    """
    names = set(r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall())
    if "contracts" in names and "project_contracts" in names:
        n_old = con.execute("SELECT COUNT(*) FROM contracts").fetchone()[0]
        n_new = con.execute("SELECT COUNT(*) FROM project_contracts").fetchone()[0]
        new_idx = [r[0] for r in con.execute(
            "SELECT name FROM sqlite_master WHERE type='index' AND tbl_name='project_contracts' AND sql IS NOT NULL")]
        if n_new == 0 and not new_idx:
            con.execute("DROP TABLE project_contracts")
            names.discard("project_contracts")
            print("[INFO] v5：已删除 _SCHEMA 预建的空表 project_contracts，随后将 contracts（%d 行）改名为 project_contracts"
                  % n_old, file=sys.stderr)
        else:
            print("[WARN] 库中同时存在 contracts（%d 行）与 project_contracts（%d 行），v5 未改名，请人工确认"
                  % (n_old, n_new), file=sys.stderr)
    if "contracts" in names and "project_contracts" not in names:
        con.execute("ALTER TABLE contracts RENAME TO project_contracts")
    for kind, old_pfx, new_pfx in (("index", "idx_contracts_", "idx_project_contracts_"),
                                   ("trigger", "trg_contracts_", "trg_project_contracts_")):
        rows = con.execute("SELECT name, sql FROM sqlite_master WHERE type=? AND name LIKE ?",
                           (kind, old_pfx + "%")).fetchall()
        for name, sql in rows:
            if not sql or name not in sql:
                continue
            new_name = new_pfx + name[len(old_pfx):]
            con.execute("DROP " + kind.upper() + " " + name)
            con.execute(sql.replace(name, new_name, 1))

def _mig_v6_asset_dept_person_cols(con):
    """v6：资产合同台账追加 管理部门/责任人 两列，并建对应查询索引（幂等）。"""
    cols = [r[1] for r in con.execute("PRAGMA table_info(asset_contracts)").fetchall()]
    if not cols:
        return
    for cname in ("管理部门", "责任人"):
        if cname not in cols:
            con.execute('ALTER TABLE asset_contracts ADD COLUMN "%s" TEXT' % cname)
    for sql in ASSET_CONTRACTS_INDEXES:
        con.execute(sql)

def _mig_v7_asset_archive_cols(con):
    """v7：资产合同台账追加 归档列（是否归档/归档人/归档时间），并建归档索引（幂等）。"""
    cols = [r[1] for r in con.execute("PRAGMA table_info(asset_contracts)").fetchall()]
    if not cols:
        return
    for cname in ASSET_ARCHIVE_COLS:
        if cname not in cols:
            con.execute('ALTER TABLE asset_contracts ADD COLUMN "%s" TEXT' % cname)
    for sql in ASSET_CONTRACTS_INDEXES:
        con.execute(sql)

def _mig_v8_asset_next_date_total_amount(con):
    """v8：资产合同台账追加 下次出租时间（DATE，可空）/ 总合同额（DECIMAL(12,2)，默认 0），幂等。"""
    cols = [r[1] for r in con.execute("PRAGMA table_info(asset_contracts)").fetchall()]
    if not cols:
        return
    if "下次出租时间" not in cols:
        con.execute('ALTER TABLE asset_contracts ADD COLUMN "下次出租时间" DATE')
    if "总合同额" not in cols:
        con.execute('ALTER TABLE asset_contracts ADD COLUMN "总合同额" DECIMAL(12,2) DEFAULT 0')

def _ensure_asset_attachments(con):
    """建立资产合同附件表 asset_attachments 及其索引（幂等：已存在则跳过）。"""
    con.execute(ASSET_ATTACHMENTS_DDL)
    for sql in ASSET_ATTACHMENTS_INDEXES:
        con.execute(sql)

def _mig_v9_asset_attachments(con):
    """v9：新增资产合同附件表 asset_attachments（按资产合同 id 关联），幂等。"""
    _ensure_asset_attachments(con)

def _mig_v10_asset_remind_days(con):
    """v10：资产合同台账追加 提醒提前天数（INTEGER，可空），幂等。"""
    cols = [r[1] for r in con.execute("PRAGMA table_info(asset_contracts)").fetchall()]
    if not cols:
        return
    if "提醒提前天数" not in cols:
        con.execute('ALTER TABLE asset_contracts ADD COLUMN "提醒提前天数" INTEGER')

def _ensure_field_logs(con):
    """建立字段级变更日志表 field_change_logs 及其索引（幂等）。"""
    con.execute(FIELD_LOG_DDL)
    for sql in FIELD_LOG_INDEXES:
        con.execute(sql)

def _mig_v11_field_change_logs(con):
    """v11：新增字段级变更日志表 field_change_logs，幂等。"""
    _ensure_field_logs(con)

def _ensure_asset_rent_collections(con):
    """建立资产合同租金收缴表 asset_rent_collections 及其索引（幂等：已存在则跳过）。"""
    con.execute(ASSET_RENT_DDL)
    for sql in ASSET_RENT_INDEXES:
        con.execute(sql)

def _mig_v12_asset_rent_collections(con):
    """v12：新增资产合同租金收缴表 asset_rent_collections（资产合同 1 : N 收款记录），幂等。

    - 仅建表 + 4 个普通索引（"资产合同ID" 上不建唯一索引/唯一约束，允许一个合同多条收款记录）。
    - 外键：资产合同ID → asset_contracts(id)，ON UPDATE CASCADE（id 为自增主键，实际不变更）
      ON DELETE RESTRICT——租金收款记录属财务留痕，删除资产合同不得连带删除，故不采用 CASCADE；
      资产附件的删除逻辑为“先删附件再删主记录”，口径一致。
    - 不迁移任何历史已缴租金数据（旧季度列/已缴租金合计仍在 asset_contracts，由 v13 删除）。
    """
    _ensure_asset_rent_collections(con)

def _mig_v13_drop_asset_rent_cols(con):
    """v13：删除 asset_contracts 的 10 个旧租金收缴列（重建表，保留列数据不丢），幂等。

    - 采用“建临时表 → INSERT…SELECT 保留列（显式含 id）→ DROP 旧表 → RENAME → 按原名重建索引/触发器”
      的官方重建流程；旧列带 CHECK/GLOB 约束，且 asset_attachments、asset_rent_collections 均外键引用
      asset_contracts(id)，直接删列/删表都不安全（外键开启时 DROP TABLE 会因 RESTRICT 报错）。
    - PRAGMA foreign_keys 仅在事务外生效（SQLite 规范），故本迁移在 MIGRATIONS 中标记 own_transaction，
      自行 BEGIN/COMMIT，并在提交前用 PRAGMA foreign_key_check 校验引用完整性；结束后恢复 foreign_keys=ON。
    - 幂等：若表中已无旧列，仅补齐 6 个查询索引后返回，不执行重建。
    """
    cols = [r[1] for r in con.execute("PRAGMA table_info(asset_contracts)").fetchall()]
    if not cols:
        return
    legacy = [c for c in ASSET_CONTRACT_LEGACY_RENT_COLS if c in cols]
    if not legacy:
        for sql in ASSET_CONTRACTS_INDEXES:
            con.execute(sql)
        print("[INFO] v13：asset_contracts 已无旧租金列，跳过重建", file=sys.stderr)
        return
    missing = [c for c in ASSET_CONTRACT_KEEP_COLS if c not in cols]
    if missing:
        raise RuntimeError("v13：asset_contracts 缺少预期保留列，请人工确认：" + "、".join(missing))
    keep_objs = con.execute(
        "SELECT type, name, sql FROM sqlite_master WHERE tbl_name='asset_contracts' "
        "AND ((type='index' AND sql IS NOT NULL) OR type='trigger') ORDER BY type DESC, name").fetchall()
    cols_sql = ",".join('"%s"' % c for c in ASSET_CONTRACT_KEEP_COLS)
    prev_fk = con.execute("PRAGMA foreign_keys").fetchone()[0]
    con.execute("PRAGMA foreign_keys=OFF")
    try:
        con.execute("BEGIN")
        try:
            con.execute("DROP TABLE IF EXISTS " + ASSET_CONTRACTS_NEW_TABLE)
            con.execute(ASSET_CONTRACTS_REBUILD_DDL)
            con.execute("INSERT INTO %s (%s) SELECT %s FROM asset_contracts"
                        % (ASSET_CONTRACTS_NEW_TABLE, cols_sql, cols_sql))
            con.execute("DROP TABLE asset_contracts")
            con.execute("ALTER TABLE %s RENAME TO asset_contracts" % ASSET_CONTRACTS_NEW_TABLE)
            for _type, _name, _sql in keep_objs:
                con.execute(_sql)
            for sql in ASSET_CONTRACTS_INDEXES:
                con.execute(sql)
            bad = con.execute("PRAGMA foreign_key_check").fetchall()
            if bad:
                raise RuntimeError("v13：重建后外键校验失败：" + repr(bad))
            con.execute("COMMIT")
        except Exception:
            con.execute("ROLLBACK")
            raise
    finally:
        con.execute("PRAGMA foreign_keys=%d" % (1 if prev_fk else 0))

def _mig_v14_rename_year_rent_col(con):
    """v14：资产合同台账列改名 年租（元） → 年租金（元）（只改名，数据/约束/默认值随列保留），幂等。

    - 用 ALTER TABLE … RENAME COLUMN（SQLite ≥ 3.25）：SQLite 会自动改写该列在 CHECK/索引/触发器等
      结构定义中的引用；本列带 CHECK、未参与索引、也未被任何子表外键引用，故无需重建表。
    - 接口字段名保持不变（annual_rent），仅数据库列名与界面/统计标签变化，前端 v-model 与取数不受影响。
    - 幂等：列名已是 年租金（元） 则跳过；两列都不存在时只告警不报错。
    """
    cols = [r[1] for r in con.execute("PRAGMA table_info(asset_contracts)").fetchall()]
    if not cols:
        return
    if "年租金（元）" in cols:
        return
    if "年租（元）" not in cols:
        print("[WARN] v14：asset_contracts 既无 年租（元） 也无 年租金（元），跳过改名", file=sys.stderr)
        return
    if sqlite3.sqlite_version_info < (3, 25, 0):
        raise RuntimeError("v14：当前 SQLite %s 不支持 RENAME COLUMN（需 ≥ 3.25），请升级后重试" % sqlite3.sqlite_version)
    con.execute('ALTER TABLE asset_contracts RENAME COLUMN "年租（元）" TO "年租金（元）"')

def _mig_v15_drop_contract_pay_method(con):
    """v15：删除 asset_contracts 的“缴款方式”列（含数据，重建表，其余列数据不丢），幂等。

    - 走官方重建表流程（建临时表 → INSERT…SELECT 保留列 → DROP → RENAME → 重建索引/触发器），
      版本无关；asset_attachments / asset_rent_collections 仍外键引用 asset_contracts(id)，
      故在事务外关闭外键后重建（MIGRATIONS 中标记 own_transaction）。
    - 缴款方式改由 asset_rent_collections 按收款记录枚举维护（见 v16）。
    - 幂等：列已不存在则仅补齐索引后返回。
    """
    cols = [r[1] for r in con.execute("PRAGMA table_info(asset_contracts)").fetchall()]
    if not cols:
        return
    if "缴款方式" not in cols:
        for sql in ASSET_CONTRACTS_INDEXES:
            con.execute(sql)
        print("[INFO] v15：asset_contracts 已无“缴款方式”列，跳过重建", file=sys.stderr)
        return
    keep_objs = con.execute(
        "SELECT type, name, sql FROM sqlite_master WHERE tbl_name='asset_contracts' "
        "AND ((type='index' AND sql IS NOT NULL) OR type='trigger') ORDER BY type DESC, name").fetchall()
    cols_sql = ",".join('"%s"' % c for c in ASSET_CONTRACTS_V15_KEEP_COLS)
    prev_fk = con.execute("PRAGMA foreign_keys").fetchone()[0]
    con.execute("PRAGMA foreign_keys=OFF")
    try:
        con.execute("BEGIN")
        try:
            con.execute("DROP TABLE IF EXISTS " + ASSET_CONTRACTS_NEW_TABLE)
            con.execute(ASSET_CONTRACTS_V15_DDL)
            con.execute("INSERT INTO %s (%s) SELECT %s FROM asset_contracts"
                        % (ASSET_CONTRACTS_NEW_TABLE, cols_sql, cols_sql))
            con.execute("DROP TABLE asset_contracts")
            con.execute("ALTER TABLE %s RENAME TO asset_contracts" % ASSET_CONTRACTS_NEW_TABLE)
            for _type, _name, _sql in keep_objs:
                con.execute(_sql)
            for sql in ASSET_CONTRACTS_INDEXES:
                con.execute(sql)
            bad = con.execute("PRAGMA foreign_key_check").fetchall()
            if bad:
                raise RuntimeError("v15：重建后外键校验失败：" + repr(bad))
            con.execute("COMMIT")
        except Exception:
            con.execute("ROLLBACK")
            raise
    finally:
        con.execute("PRAGMA foreign_keys=%d" % (1 if prev_fk else 0))

def _mig_v16_rent_pay_method_enum(con):
    """v16：asset_rent_collections.“缴款方式”改为枚举（直缴非税/现金/微信/其他），并规范化存量取值，幂等。

    - SQLite 无法给已有列追加 CHECK，故重建该子表（子表不被其它表引用，重建不影响外键关系）；
      同时按规则归并历史取值：缴非税/直缴非税→直缴非税，微信/微信转账→微信，现金→现金，其他→其他，
      其余旧值（如“年交”等缴款频次）与空值一律置空，避免误标为“其他”。
    - 幂等：建表语句中已含枚举约束（含 '直缴非税'）则跳过。
    """
    row = con.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='asset_rent_collections'").fetchone()
    if not row or not row[0]:
        return
    if "直缴非税" in row[0]:
        for sql in ASSET_RENT_INDEXES:
            con.execute(sql)
        print("[INFO] v16：asset_rent_collections.“缴款方式”已是枚举约束，跳过重建", file=sys.stderr)
        return
    con.execute("DROP TABLE IF EXISTS asset_rent_collections_new")
    con.execute(ASSET_RENT_V16_DDL)
    con.execute('''
        INSERT INTO asset_rent_collections_new
          (id,"资产合同ID","收款期次","所属租期开始日","所属租期结束日","应缴金额（元）","实缴金额（元）","缴款日期","缴款方式","收款人")
        SELECT id,"资产合同ID","收款期次","所属租期开始日","所属租期结束日","应缴金额（元）","实缴金额（元）","缴款日期",
               CASE
                 WHEN "缴款方式" IN ('直缴非税','缴非税') THEN '直缴非税'
                 WHEN "缴款方式" IN ('微信','微信转账') THEN '微信'
                 WHEN "缴款方式" = '现金' THEN '现金'
                 WHEN "缴款方式" = '其他' THEN '其他'
                 ELSE ''
               END,
               "收款人"
        FROM asset_rent_collections''')
    con.execute("DROP TABLE asset_rent_collections")
    con.execute("ALTER TABLE asset_rent_collections_new RENAME TO asset_rent_collections")
    for sql in ASSET_RENT_INDEXES:
        con.execute(sql)

def _mig_v17_rent_pay_method_default(con):
    """v17：asset_rent_collections.“缴款方式”默认值设为 '直缴非税'（重建表，数据原样保留），幂等。

    - SQLite 无法直接修改列默认值，故重建该子表（子表不被其它表引用，重建不影响外键关系）。
    - 只改默认值：已有记录取值原样保留（空值仍为空）；新增/脚本导入未指定该列时默认落在“直缴非税”。
    - 幂等：建表语句中已含 DEFAULT '直缴非税' 则跳过。
    """
    row = con.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='asset_rent_collections'").fetchone()
    if not row or not row[0]:
        return
    if "DEFAULT '直缴非税'" in row[0]:
        for sql in ASSET_RENT_INDEXES:
            con.execute(sql)
        print("[INFO] v17：asset_rent_collections.“缴款方式”已有默认值，跳过重建", file=sys.stderr)
        return
    con.execute("DROP TABLE IF EXISTS asset_rent_collections_new")
    con.execute(ASSET_RENT_V17_DDL)
    con.execute('''
        INSERT INTO asset_rent_collections_new
          (id,"资产合同ID","收款期次","所属租期开始日","所属租期结束日","应缴金额（元）","实缴金额（元）","缴款日期","缴款方式","收款人")
        SELECT id,"资产合同ID","收款期次","所属租期开始日","所属租期结束日","应缴金额（元）","实缴金额（元）","缴款日期",
               COALESCE("缴款方式", ''),"收款人"
        FROM asset_rent_collections''')
    con.execute("DROP TABLE asset_rent_collections")
    con.execute("ALTER TABLE asset_rent_collections_new RENAME TO asset_rent_collections")
    for sql in ASSET_RENT_INDEXES:
        con.execute(sql)

def _mig_v18_stats_indexes(con):
    """v18：为首页统计/提醒/待收待支相关谓词补索引（幂等，纯新增索引，不改数据与约束）。

    - 资产侧到期提醒按“下次出租时间”判定，此前无任何索引；
    - 项目侧提醒/待支走“下次招采时间 + 提醒提前天数 + 所属部门”“是否归档 + 所属部门 + 金额列”；
    - 子表按“资产合同ID + 收款期次”分组/过滤。
    说明：当前数据量（639/5/16 行）下这些索引基本不会被优化器选中（全表扫描更便宜），
    属于面向数据增长的预置；索引写在库文件里，写放大可忽略。
    """
    for sql in STATS_INDEXES:
        con.execute(sql)

STATS_INDEXES = [
    'CREATE INDEX IF NOT EXISTS idx_asset_contracts_next_lease_date ON asset_contracts("下次出租时间")',
    'CREATE INDEX IF NOT EXISTS idx_project_contracts_reminder ON project_contracts("下次招采时间","提醒提前天数","所属部门")',
    'CREATE INDEX IF NOT EXISTS idx_project_contracts_pending ON project_contracts("是否归档","所属部门","未支付","合同金额（万元）","已支付")',
    'CREATE INDEX IF NOT EXISTS idx_arc_contract_period ON asset_rent_collections("资产合同ID","收款期次")',
]

def _mig_v19_users_phone(con):
    """v19：users 追加 电话 字段（TEXT，可空，无默认值），幂等。"""
    ucols = [r[1] for r in con.execute("PRAGMA table_info(users)").fetchall()]
    if "电话" not in ucols:
        con.execute('ALTER TABLE users ADD COLUMN "电话" TEXT')

def _mig_v20_sms_settings(con):
    """v20：settings 追加短信平台配置项（AK/SK、开关、每周提醒时间等），幂等。

    - 仅首次写入默认值（INSERT OR IGNORE），不覆盖用户已修改的线上配置。
    - 总开关 sms_enabled 默认 '0'（关闭），联调通过后在设置页开启。
    - 发送时段可配置：sms_weekly_weekday/hour/minute（默认 周一 09:00）。
    """
    for k, v in SMS_SETTINGS_DEFAULTS:
        con.execute('INSERT OR IGNORE INTO settings ("配置项","配置值") VALUES (?,?)', (k, v))

def _mig_v21_sms_templates(con):
    """v21：新增短信模板表 sms_templates 并写入默认模板，幂等。

    - 场景：expire 到期 / receivable 待收 / payable 待支 / urge 催办 / manual 点对点。
    - 模板内容支持占位符 {name}{number}{due}{days}{tenant}{amount}{note}，发送时替换并自动拼接签名。
    - 默认模板仅首次写入（INSERT OR IGNORE），不覆盖用户已修改内容。
    """
    con.execute(SMS_TEMPLATES_DDL)
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for scene, name, tpl in SMS_TEMPLATE_DEFAULTS:
        con.execute('INSERT OR IGNORE INTO sms_templates ("场景","场景名称","模板内容","启用","更新时间") '
                    'VALUES (?,?,?,?,?)', (scene, name, tpl, "是", ts))

def _mig_v22_sms_logs(con):
    """v22：新增短信发送流水表 sms_logs 及其索引（含周去重唯一索引），幂等。

    - "去重键" 部分唯一索引：同一提醒同一周同一手机号只发一条（每周“每人一条”）；
      服务重启或定时任务重复触发时，靠该唯一索引天然去重。
    - 回执 / 回复回调按 "外部guId" 就近写入对应流水。
    """
    con.execute(SMS_LOGS_DDL)
    for sql in SMS_LOGS_INDEXES:
        con.execute(sql)

def _mig_v23_sms_switch_col(con):
    """v23：project_contracts / asset_contracts 追加“短信提醒”开关列（TEXT，可空），幂等。

    - 取值：'开启' / '关闭'；空值与 '开启' 等价（向后兼容）。
    - 关闭后该合同不参与自动提醒与催办短信；点对点手动发送仍允许。
    - 仅 ADD COLUMN，不重建表、不影响外键与既有数据。
    """
    for tbl in ("project_contracts", "asset_contracts"):
        cols = [r[1] for r in con.execute("PRAGMA table_info(%s)" % tbl).fetchall()]
        if cols and "短信提醒" not in cols:
            con.execute('ALTER TABLE %s ADD COLUMN "短信提醒" TEXT' % tbl)

def _mig_v24_sms_timeout_setting(con):
    """v24：settings 追加短信连接/读取超时项 sms_timeout_sec（默认 10 秒），幂等。

    - 政务内网链路较慢时可调大；代码侧读取缺省值同样为 10，故本项缺失也不影响运行。
    """
    con.execute('INSERT OR IGNORE INTO settings ("配置项","配置值") VALUES (?,?)', ("sms_timeout_sec", "10"))

def _mig_v25_sms_rules_settings(con):
    """v25：settings 追加“短信平台使用规范”相关配置项（幂等 INSERT OR IGNORE）。

    对应《政务短信平台使用注意事项》：签名拼接开关、单条分段长度、最大段数、
    敏感词校验开关与词表、同一号码每日上限、同一号码最小发送间隔。
    """
    for k in ("sms_include_sign", "sms_segment_len", "sms_max_segments", "sms_block_check",
              "sms_block_words", "sms_daily_per_phone", "sms_min_interval_sec"):
        v = next((val for key, val in SMS_SETTINGS_DEFAULTS if key == k), "")
        con.execute('INSERT OR IGNORE INTO settings ("配置项","配置值") VALUES (?,?)', (k, v))

def _mig_v26_asset_id_text(con):
    """v26：asset_contracts.id 由“自增整数”改为“6 位合同编号”（TEXT），历史数据原值保留，幂等。

    - 统一编号规则后，id 不再自增，新增资产合同时按《合同编号规则》生成 6 位编号写入 id；
    - 采用“建临时表 → INSERT…SELECT（显式列出全部列，id 原值搬运）→ DROP 旧表 → RENAME →
      重建索引/触发器”的官方重建流程；asset_attachments、asset_rent_collections 外键引用
      asset_contracts(id)，重建被引用的表必须在外键关闭状态下进行（PRAGMA foreign_keys 仅在事务外生效）。
    - 幂等：若 id 列已是 TEXT 类型，仅补齐查询索引后返回，不执行重建。
    """
    info = con.execute("PRAGMA table_info(asset_contracts)").fetchall()
    if not info:
        return
    cols = [r[1] for r in info]
    id_type = next((str(r[2] or "").upper() for r in info if r[1] == "id"), "")
    if id_type.startswith("TEXT"):
        for sql in ASSET_CONTRACTS_INDEXES:
            con.execute(sql)
        print("[INFO] v26：asset_contracts.id 已是 TEXT，跳过重建", file=sys.stderr)
        return
    missing = [c for c in ASSET_CONTRACTS_V26_KEEP_COLS if c not in cols]
    if missing:
        raise RuntimeError("v26：asset_contracts 缺少预期保留列，请人工确认：" + "、".join(missing))
    keep_objs = con.execute(
        "SELECT type, name, sql FROM sqlite_master WHERE tbl_name='asset_contracts' "
        "AND ((type='index' AND sql IS NOT NULL) OR type='trigger') ORDER BY type DESC, name").fetchall()
    cols_sql = ",".join('"%s"' % c for c in ASSET_CONTRACTS_V26_KEEP_COLS)
    prev_fk = con.execute("PRAGMA foreign_keys").fetchone()[0]
    con.execute("PRAGMA foreign_keys=OFF")
    try:
        con.execute("BEGIN")
        try:
            con.execute("DROP TABLE IF EXISTS asset_contracts_new")
            con.execute(ASSET_CONTRACTS_V26_DDL)
            con.execute("INSERT INTO asset_contracts_new (%s) SELECT %s FROM asset_contracts"
                        % (cols_sql, cols_sql))
            con.execute("DROP TABLE asset_contracts")
            con.execute("ALTER TABLE asset_contracts_new RENAME TO asset_contracts")
            for _type, _name, _sql in keep_objs:
                con.execute(_sql)
            for sql in ASSET_CONTRACTS_INDEXES:
                con.execute(sql)
            bad = con.execute("PRAGMA foreign_key_check").fetchall()
            if bad:
                raise RuntimeError("v26：重建后外键校验失败：" + repr(bad))
            con.execute("COMMIT")
        except Exception:
            con.execute("ROLLBACK")
            raise
    finally:
        con.execute("PRAGMA foreign_keys=%d" % (1 if prev_fk else 0))

# 部门表历史残留行：原“人事与政策法规科”已拆分为“人事科”(第3位 4) 与“政策法规科”(第3位 9)
LEGACY_DEPT_ROW = ("湘潭市机关事务管理局本级", "人事与政策法规科")

def _mig_v27_drop_legacy_dept_row(con):
    """v27：按业务确认，删除部门表历史残留行「本级-人事与政策法规科」，幂等。

    - 仅在该行确实存在、且没有被 users / project_contracts 通过复合外键引用时删除；
    - 若仍被引用则跳过并打印告警（不静默删除业务数据，不猜测归属）；
    - 历史合同/用户数据一律不改动、不回填。
    """
    unit, section = LEGACY_DEPT_ROW
    exists = con.execute('SELECT 1 FROM departments WHERE "单位名称"=? AND "科室名称"=?',
                         (unit, section)).fetchone()
    if not exists:
        return
    ref_u = con.execute('SELECT COUNT(*) FROM users WHERE "所属单位"=? AND "所属科室"=?', (unit, section)).fetchone()[0]
    ref_c = con.execute('SELECT COUNT(*) FROM project_contracts WHERE "所属单位"=? AND "所属科室"=?', (unit, section)).fetchone()[0]
    if ref_u or ref_c:
        print("[WARN] v27：部门「%s-%s」仍被引用（users=%d、project_contracts=%d），跳过删除"
              % (unit, section, ref_u, ref_c), file=sys.stderr)
        return
    con.execute('DELETE FROM departments WHERE "单位名称"=? AND "科室名称"=?', (unit, section))

def _mig_v28_rename_attachments_table(con):
    """v28：项目合同附件表改名 attachments → project_attachments（幂等）。

    - SQLite 的 ALTER TABLE ... RENAME TO 会自动改写表定义中的外键子句，
      主键自动索引 sqlite_autoindex_attachments_1 也随之改名为
      sqlite_autoindex_project_attachments_1；该表无自定义索引与触发器，故无需重建对象。
    - 旧库首次启动时 _SCHEMA 会先按新表名建出空表（IF NOT EXISTS），故此处识别
      “旧表存在、新表为空”的情形：删除该空表后再改名，旧库数据随之落到新表名下。
    - 两表都有数据时打印告警并跳过，交人工确认（不静默丢数据）。
    """
    names = set(r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall())
    if "attachments" in names and "project_attachments" in names:
        n_old = con.execute("SELECT COUNT(*) FROM attachments").fetchone()[0]
        n_new = con.execute("SELECT COUNT(*) FROM project_attachments").fetchone()[0]
        if n_new == 0:
            con.execute("DROP TABLE project_attachments")
            names.discard("project_attachments")
            print("[INFO] v28：已删除 _SCHEMA 预建的空表 project_attachments，随后将 attachments（%d 行）改名为 project_attachments"
                  % n_old, file=sys.stderr)
        else:
            print("[WARN] 库中同时存在 attachments（%d 行）与 project_attachments（%d 行），v28 未改名，请人工确认"
                  % (n_old, n_new), file=sys.stderr)
    if "attachments" in names and "project_attachments" not in names:
        con.execute("ALTER TABLE attachments RENAME TO project_attachments")

def _mig_v29_sms_sign_suffix(con):
    """v29：短信正文末尾追加机构署名。

    - 新增配置项 `sms_sign_suffix`（默认「（湘潭市机关事务管理局）」）：发送时自动追加到正文末尾；
    - 确保署名开关 `sms_include_sign` 为 1（对应“每条短信后面加（湘潭市机关事务管理局）”的要求）；
    - 使用**全角括号**而非「【】」：平台注意事项 2 规定内容含中括号会被判多签名并阻断；
    - 幂等：配置项存在则保留现值（不覆盖管理员后续修改）。
    """
    con.execute('INSERT OR IGNORE INTO settings ("配置项","配置值") VALUES (?,?)',
                ("sms_sign_suffix", "（湘潭市机关事务管理局）"))
    con.execute('UPDATE settings SET "配置值"=? WHERE "配置项"=?', ("1", "sms_include_sign"))

def _mig_v30_drop_sms_sign_name(con):
    """v30：删除已废弃的「短信签名」配置项 `sms_sign_name`（幂等）。

    背景：原先用 `sms_sign_name` 在正文前拼接 `【签名】正文`，平台注意事项 2 明令内容不得含「【】」；
    现改为在正文**末尾**追加全角括号署名（`sms_sign_suffix`），`sms_sign_name` 已无任何业务引用，
    且设置页已移除该输入框。其值等价保留在 `sms_sign_suffix`（默认「（湘潭市机关事务管理局）」）中。
    """
    con.execute('DELETE FROM settings WHERE "配置项"=?', ("sms_sign_name",))

MIGRATIONS = [
    (1, "合同主表追加历史字段列：资金来源/是否归档/归档人/归档时间", _mig_v1_contract_extra_cols),
    (2, "合同主表列改名：经办人/分管领导→经办人、甲方现场负责人→现场负责人", _mig_v2_contract_rename_cols),
    (3, "users 追加 分管部门 字段", _mig_v3_users_supervise_dept),
    (4, "新增资产合同台账表 asset_contracts（依据资产合同台账.xlsx，字段参照 project_contracts 表约定）", _mig_v4_asset_contracts),
    (5, "合同主表改名：contracts → project_contracts（同步索引/触发器名与外键引用）", _mig_v5_rename_contracts_table),
    (6, "资产合同台账追加字段：管理部门/责任人（同步新增查询索引）", _mig_v6_asset_dept_person_cols),
    (7, "资产合同台账追加归档列：是否归档/归档人/归档时间（同步新增归档索引）", _mig_v7_asset_archive_cols),
    (8, "资产合同台账追加字段：下次出租时间（DATE，可空）/ 总合同额（DECIMAL(12,2)，默认 0）", _mig_v8_asset_next_date_total_amount),
    (9, "新增资产合同附件表 asset_attachments（决策依据附件/合同附件）", _mig_v9_asset_attachments),
    (10, "资产合同台账追加字段：提醒提前天数（INTEGER，可空）", _mig_v10_asset_remind_days),
    (11, "新增字段级变更日志表 field_change_logs（记录字段修改前后值）", _mig_v11_field_change_logs),
    (12, "新增资产合同租金收缴表 asset_rent_collections（资产合同 1:N 收款记录，不迁移历史已缴租金）", _mig_v12_asset_rent_collections),
    (13, "删除资产合同台账旧租金收缴列（已缴租金1~4季度/交款时间/开票对应租期/已缴租金合计，重建表，旧数据不迁移）",
     _mig_v13_drop_asset_rent_cols, {"own_transaction": True}),
    (14, "资产合同台账列改名：年租（元） → 年租金（元）", _mig_v14_rename_year_rent_col),
    (15, "删除资产合同台账“缴款方式”列（含数据；缴款方式改由租金收取明细按枚举维护）",
     _mig_v15_drop_contract_pay_method, {"own_transaction": True}),
    (16, "租金收取明细“缴款方式”枚举化（直缴非税/现金/微信/其他）并规范化存量取值", _mig_v16_rent_pay_method_enum),
    (17, "租金收取明细“缴款方式”默认值设为“直缴非税”", _mig_v17_rent_pay_method_default),
    (18, "统计/提醒查询索引优化（资产下次出租时间、项目提醒与待支覆盖索引、租赁期次复合索引）", _mig_v18_stats_indexes),
    (19, "users 追加 电话 字段", _mig_v19_users_phone),
    (20, "settings 追加短信平台配置项（AK/SK、开关、每周提醒时间等）", _mig_v20_sms_settings),
    (21, "新增短信模板表 sms_templates 并写入默认模板（到期/待收/待支/催办/点对点）", _mig_v21_sms_templates),
    (22, "新增短信发送流水表 sms_logs（含回执/回复、每周每人一条去重唯一索引）", _mig_v22_sms_logs),
    (23, "project_contracts/asset_contracts 追加“短信提醒”开关列", _mig_v23_sms_switch_col),
    (24, "settings 追加短信连接超时项 sms_timeout_sec（默认 10 秒）", _mig_v24_sms_timeout_setting),
    (25, "settings 追加短信平台使用规范配置项（签名拼接/分段长度/敏感词/每日上限/最小间隔）", _mig_v25_sms_rules_settings),
    (26, "asset_contracts.id 改为 6 位合同编号（TEXT，按《合同编号规则》生成；历史 id 原值保留）",
     _mig_v26_asset_id_text, {"own_transaction": True}),
    (27, "按业务确认删除部门表历史残留行：本级-人事与政策法规科（已拆分为人事科/政策法规科）",
     _mig_v27_drop_legacy_dept_row),
    (28, "项目合同附件表改名：attachments → project_attachments（外键与自动索引随表名更新）",
    _mig_v28_rename_attachments_table),
    (29, "短信正文末尾追加机构署名（settings.sms_sign_suffix 默认「（湘潭市机关事务管理局）」，并开启署名开关）", _mig_v29_sms_sign_suffix),
    (30, "删除已废弃的短信签名配置项 sms_sign_name（改用正文末尾署名 sms_sign_suffix）", _mig_v30_drop_sms_sign_name),
]

def _run_versioned_migrations(con):
    """按版本号顺序执行未应用的数据库结构迁移。

    - 建表：schema_version(version INTEGER PRIMARY KEY, applied_at TEXT, description TEXT)。
    - 仅执行 version > 当前最大已应用版本的迁移；每个迁移独立事务，成功后写入版本记录。
    - 失败：回滚事务、不记录版本、向上抛出错误（下次启动重试）。
    - 迁移函数本身保持幂等（含存在性判断），删除版本行可强制重跑对应版本。
    - 新增结构变更须向 MIGRATIONS 追加新版本号，禁止修改已执行的历史版本。
    - 条目格式：(版本, 说明, 函数) 为默认写法；(版本, 说明, 函数, {"own_transaction": True}) 表示该迁移
      自行管理事务——用于必须临时关闭外键的重建表迁移（SQLite 规范下 PRAGMA foreign_keys 在事务内设置无效，
      而重建被外键引用的表必须在外键关闭状态下 DROP/RENAME）。此类迁移运行器只负责调用与记录版本。
    """
    con.execute("CREATE TABLE IF NOT EXISTS schema_version (version INTEGER PRIMARY KEY, applied_at TEXT, description TEXT)")
    applied = set(r[0] for r in con.execute("SELECT version FROM schema_version").fetchall())
    for entry in MIGRATIONS:
        ver, desc, fn = entry[0], entry[1], entry[2]
        opts = entry[3] if len(entry) > 3 else {}
        if ver in applied:
            continue
        if opts.get("own_transaction"):
            con.commit()          # 先结束可能存在的隐式事务，保证迁移内 PRAGMA foreign_keys 生效
            try:
                fn(con)
                ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                con.execute("BEGIN")
                con.execute("INSERT INTO schema_version (version, applied_at, description) VALUES (?,?,?)", (ver, ts, desc))
                con.execute("COMMIT")
            except Exception:
                try: con.execute("ROLLBACK")
                except sqlite3.OperationalError: pass
                raise
            print(f"[MIGRATE] schema_version {ver}: {desc}", file=sys.stderr)
            continue
        con.execute("BEGIN")
        try:
            fn(con)
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            con.execute("INSERT INTO schema_version (version, applied_at, description) VALUES (?,?,?)", (ver, ts, desc))
            con.execute("COMMIT")
        except Exception:
            con.execute("ROLLBACK")
            raise
        print(f"[MIGRATE] schema_version {ver}: {desc}", file=sys.stderr)

def init_db():
    fresh = not os.path.exists(DB_PATH)
    with db_lock:
        con = _conn()
        try:
            con.executescript(_SCHEMA)
            _run_versioned_migrations(con)
            # A-01 合同编号唯一约束（数据库层兜底）：先扫描存量重复编号，存在则终止并输出清单
            dup_nos = con.execute('SELECT "编号", COUNT(*) c FROM project_contracts GROUP BY "编号" HAVING c>1').fetchall()
            if dup_nos:
                lst = "、".join(f"{r[0]}（{r[1]}条）" for r in dup_nos)
                print(f"[ERROR] project_contracts 存在重复编号，无法建立唯一索引，请先处理：{lst}", file=sys.stderr)
                raise RuntimeError("project_contracts 存在重复编号，无法建立唯一索引：" + lst)
            con.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_project_contracts_no ON project_contracts("编号")')
            # A-02 departments 复合主键（单位名称, 科室名称）数据库层兜底：重建表，先校验存量无重复组合
            dept_pk = [r[5] for r in con.execute("PRAGMA table_info(departments)").fetchall()]
            if not all(dept_pk):
                dup_depts = con.execute('SELECT "单位名称","科室名称",COUNT(*) c FROM departments GROUP BY "单位名称","科室名称" HAVING c>1').fetchall()
                if dup_depts:
                    lst = "、".join(f"{r[0]}-{r[1]}（{r[2]}条）" for r in dup_depts)
                    print(f"[ERROR] departments 存在重复部门组合，无法建立复合主键，请先处理：{lst}", file=sys.stderr)
                    raise RuntimeError("departments 存在重复部门组合，无法建立复合主键：" + lst)
                old_ddl = con.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='departments'").fetchone()[0]
                old_rows = con.execute('SELECT "单位名称","科室名称" FROM departments').fetchall()
                backup_path = os.path.join(os.path.dirname(DB_PATH), "departments_ddl_backup.sql")
                with open(backup_path, "w", encoding="utf-8", newline="\n") as f:
                    f.write(f"-- departments 迁移前原表结构与数据备份：{datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n")
                    f.write(old_ddl.rstrip().rstrip(";") + ";\n")
                    for r in old_rows:
                        f.write('INSERT INTO departments ("单位名称","科室名称") VALUES (\'{}\',\'{}\');\n'.format(
                            str(r[0] or "").replace("'", "''"), str(r[1] or "").replace("'", "''")))
                con.execute("BEGIN")
                try:
                    # 修复后全量校验：project_contracts/users/project_attachments 任一悬空引用即终止（禁止静默跳过）
                    con.execute('CREATE TABLE departments_new ("单位名称" TEXT NOT NULL, "科室名称" TEXT NOT NULL, PRIMARY KEY ("单位名称","科室名称"))')
                    con.execute('INSERT INTO departments_new ("单位名称","科室名称") SELECT "单位名称","科室名称" FROM departments')
                    con.execute("DROP TABLE departments")
                    con.execute("ALTER TABLE departments_new RENAME TO departments")
                    con.execute("COMMIT")
                except Exception:
                    con.execute("ROLLBACK")
                    raise

            # A-03 外键约束（数据库层兜底）：重建 project_contracts/users/project_attachments 三张表并声明外键；先修复缺科室后缀数据
            ccols3 = [r[1] for r in con.execute("PRAGMA table_info(project_contracts)").fetchall()]
            if "所属单位" not in ccols3:
                dept_pairs = set((r[0], r[1]) for r in con.execute('SELECT "单位名称","科室名称" FROM departments').fetchall())
                def _split_dept(v):
                    v = str(v or "").strip()
                    if "-" in v:
                        u, k = v.split("-", 1)
                        return (u.strip(), k.strip())
                    return None
                broken = []
                for r in con.execute('SELECT id,"编号","项目名称","经办人","所属部门" FROM project_contracts ORDER BY "编号"').fetchall():
                    sp = _split_dept(r[4])
                    if sp is None:
                        if str(r[4] or "").strip(): broken.append(r)
                    elif sp not in dept_pairs:
                        broken.append(r)
                _DEPT_RULES = [
                    (("公车","车辆","平台委托","平台服务","新能源"), "公车组"),
                    (("周转房","周转住房"), "周转住房专班"),
                    (("施工图","预算","结算","审核","设计","幕墙","桥梁","钢结构","给水","钢筋","建设工程","工程质量","项目检测","检测合同","施工"), "工程组"),
                    (("物业","食堂","电梯","直饮水","消防","空调","窗帘","消杀","保洁","环境治理","标识","标牌","车库","会议室","弱电","网络","监控","配电","排污","管道","维保","维修","维护","运维","灭火器","卷帘门","限高","照明","下水道","水质","监测","采购","物资"), "物业组"),
                    (("市民之家",), "市民之家专班"),
                ]
                def _suggest_dept(name):
                    for kws, dept in _DEPT_RULES:
                        if any(k in (name or "") for k in kws): return dept, kws
                    return "后勤组", ("兜底",)
                repair_rows = []
                if broken:
                    md = ["# A-03 部门数据修复清单（缺科室后缀 → 建议科室）", "",
                          "> 依据关键字规则自动生成，**待业务逐条复核确认**；如需调整，按编号修改“所属部门”后重新执行迁移（回滚见 backups 重建前备份）。",
                          "", "| 编号 | 项目名称 | 原所属部门 | 修复后所属部门 | 修复依据（命中关键字） |", "|---|---|---|---|---|"]
                    for r in broken:
                        dept, kws = _suggest_dept(r[2])
                        newval = (str(r[4] or "").strip()) + "-" + dept
                        repair_rows.append((r[0], newval))
                        md.append("| {} | {} | {} | {} | {} |".format(
                            str(r[1]).replace("|", "\\|"), str(r[2] or "").replace("|", "\\|").replace("\n", " "),
                            str(r[4] or "").replace("|", "\\|"), newval.replace("|", "\\|"), "、".join(kws)))
                    md_path = os.path.join(BASE_DIR, "docs", "A-03-部门数据修复清单.md")
                    os.makedirs(os.path.dirname(md_path), exist_ok=True)
                    with open(md_path, "w", encoding="utf-8", newline="\n") as f:
                        f.write("\n".join(md) + "\n")
                    print(f"[INFO] A-03 缺科室后缀部门数据 {len(broken)} 条，已按建议修复，清单：{md_path}", file=sys.stderr)
                # 重建前整库备份
                os.makedirs(BACKUP_DIR, exist_ok=True)
                ts1 = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                bak_before = os.path.join(BACKUP_DIR, f"合同台账_A03_重建前_{ts1}.db")
                if os.path.exists(bak_before): os.remove(bak_before)
                con.execute("VACUUM INTO '" + bak_before.replace("'", "''") + "'")
                # 事务内：先修复缺科室后缀数据并全量校验（任一悬空引用即终止），再原子重建三张表
                con.execute("BEGIN")
                try:
                    for rid, newval in repair_rows:
                        con.execute('UPDATE project_contracts SET "所属部门"=? WHERE id=?', (newval, rid))
                    bad_dept = []
                    for r in con.execute('SELECT id,"编号","所属部门" FROM project_contracts').fetchall():
                        sp = _split_dept(r[2])
                        if sp is None:
                            if str(r[2] or "").strip(): bad_dept.append(r[1])
                        elif sp not in dept_pairs:
                            bad_dept.append(r[1])
                    if bad_dept:
                        raise RuntimeError("project_contracts 存在无法匹配部门的所属部门，请先修复：" + "、".join(str(x) for x in bad_dept[:50]))
                    bad_user = []
                    for r in con.execute('SELECT "用户名","所属部门" FROM users').fetchall():
                        sp = _split_dept(r[1])
                        if sp is None or sp not in dept_pairs:
                            bad_user.append(str(r[0]))
                    if bad_user:
                        raise RuntimeError("users 存在无法匹配部门的所属部门，请先修复：" + "、".join(bad_user))
                    cno_set = set(x[0] for x in con.execute('SELECT "编号" FROM project_contracts').fetchall())
                    bad_att = []
                    for r in con.execute('SELECT "合同编号" FROM project_attachments').fetchall():
                        if str(r[0] or "").strip() and str(r[0]) not in cno_set:
                            bad_att.append(str(r[0]))
                    if bad_att:
                        raise RuntimeError("project_attachments 存在不存在的合同编号，请先修复：" + "、".join(bad_att))
                    def _mk_new_ddl(con, tbl, new_name, extra_cols, fk_clause):
                        info = con.execute("PRAGMA table_info(" + tbl + ")").fetchall()
                        segs = []
                        for cid, cname, ctype, notnull, dflt, pk in info:
                            seg = '"%s" %s' % (cname, ctype or "TEXT")
                            if notnull: seg += " NOT NULL"
                            if dflt is not None: seg += " DEFAULT " + str(dflt)
                            if pk == 1:
                                if cname == "id" and (ctype or "").upper() == "INTEGER":
                                    seg = '"id" INTEGER PRIMARY KEY AUTOINCREMENT'
                                else:
                                    seg += " PRIMARY KEY"
                            segs.append(seg)
                        if extra_cols: segs.append(extra_cols)
                        if fk_clause: segs.append(fk_clause)
                        return "CREATE TABLE " + new_name + " (\n  " + ",\n  ".join(segs) + "\n)"
                    # project_contracts：追加 所属单位/所属科室 + 复合外键 → departments
                    old_ccols = [r[1] for r in con.execute("PRAGMA table_info(project_contracts)").fetchall()]
                    c_ddl = _mk_new_ddl(con, "project_contracts", "project_contracts_new",
                                        '"所属单位" TEXT, "所属科室" TEXT',
                                        'FOREIGN KEY ("所属单位","所属科室") REFERENCES departments("单位名称","科室名称") ON DELETE RESTRICT ON UPDATE CASCADE')
                    con.execute(c_ddl)
                    csel = ",".join('"%s"' % c for c in old_ccols)
                    con.execute("INSERT INTO project_contracts_new (" + csel + ") SELECT " + csel + " FROM project_contracts")
                    con.execute('''UPDATE project_contracts_new SET
  "所属单位" = CASE WHEN instr("所属部门",'-')>0 THEN substr("所属部门",1,instr("所属部门",'-')-1) ELSE '' END,
  "所属科室" = CASE WHEN instr("所属部门",'-')>0 THEN substr("所属部门",instr("所属部门",'-')+1) ELSE '' END
  WHERE "所属部门" IS NOT NULL AND TRIM("所属部门") <> '' ''')
                    # DROP/重命名父表前，先移除引用 project_contracts 的触发器（否则 RENAME 时触发解析报错），
                    # 事务提交后由下方 A-03 触发器块 CREATE TRIGGER IF NOT EXISTS 自动重建
                    con.execute("DROP TRIGGER IF EXISTS trg_project_contracts_dept_sync_i")
                    con.execute("DROP TRIGGER IF EXISTS trg_project_contracts_dept_sync_u")
                    con.execute("DROP TRIGGER IF EXISTS trg_departments_rename_sync")
                    con.execute("DROP TABLE project_contracts")
                    con.execute("ALTER TABLE project_contracts_new RENAME TO project_contracts")
                    con.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_project_contracts_no ON project_contracts("编号")')
                    # users：追加 所属单位/所属科室 + 复合外键 → departments
                    old_ucols = [r[1] for r in con.execute("PRAGMA table_info(users)").fetchall()]
                    u_ddl = _mk_new_ddl(con, "users", "users_new",
                                        '"所属单位" TEXT, "所属科室" TEXT',
                                        'FOREIGN KEY ("所属单位","所属科室") REFERENCES departments("单位名称","科室名称") ON DELETE RESTRICT ON UPDATE CASCADE')
                    con.execute(u_ddl)
                    usel = ",".join('"%s"' % c for c in old_ucols)
                    con.execute("INSERT INTO users_new (" + usel + ") SELECT " + usel + " FROM users")
                    con.execute('''UPDATE users_new SET
  "所属单位" = CASE WHEN instr("所属部门",'-')>0 THEN substr("所属部门",1,instr("所属部门",'-')-1) ELSE '' END,
  "所属科室" = CASE WHEN instr("所属部门",'-')>0 THEN substr("所属部门",instr("所属部门",'-')+1) ELSE '' END
  WHERE "所属部门" IS NOT NULL AND TRIM("所属部门") <> '' ''')
                    con.execute("DROP TABLE users")
                    con.execute("ALTER TABLE users_new RENAME TO users")
                    # project_attachments：追加外键 → project_contracts(编号)
                    old_acols = [r[1] for r in con.execute("PRAGMA table_info(project_attachments)").fetchall()]
                    a_ddl = _mk_new_ddl(con, "project_attachments", "project_attachments_new",
                                        None,
                                        'FOREIGN KEY ("合同编号") REFERENCES project_contracts("编号") ON DELETE RESTRICT ON UPDATE RESTRICT')
                    con.execute(a_ddl)
                    asel = ",".join('"%s"' % c for c in old_acols)
                    con.execute("INSERT INTO project_attachments_new (" + asel + ") SELECT " + asel + " FROM project_attachments")
                    con.execute("DROP TABLE project_attachments")
                    con.execute("ALTER TABLE project_attachments_new RENAME TO project_attachments")
                    con.execute("COMMIT")
                except Exception:
                    con.execute("ROLLBACK")
                    raise
                finally:
                    con.execute("PRAGMA foreign_keys=ON")
                # 重建后整库备份
                ts2 = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                bak_after = os.path.join(BACKUP_DIR, f"合同台账_A03_重建后_{ts2}.db")
                if os.path.exists(bak_after): os.remove(bak_after)
                con.execute("VACUUM INTO '" + bak_after.replace("'", "''") + "'")
            # A-05 CHECK 与 NOT NULL 约束（金额非负/日期格式/必填字段）：重建 project_contracts；枚举 CHECK 延后至 A-08 清洗后启用
            cinfo5 = con.execute("PRAGMA table_info(project_contracts)").fetchall()
            cnames5 = [r[1] for r in cinfo5]
            cnot5 = [r[3] for r in cinfo5]
            if "编号" in cnames5 and not cnot5[cnames5.index("编号")]:
                req5 = ["编号","项目分类","项目名称","采购方式","合同金额（万元）","资金来源","合同主要内容","经办人","所属部门"]
                amt5 = ["合同金额（万元）","结算金额","已支付","未支付","决策的招标控制价"]
                date5 = ["采购时间","开工时间","完工时间","竣工验收时间","下次招采时间"]
                import re as _re5
                def _norm_date5(v):
                    s = str(v or "").strip()
                    if not s: return ""
                    m = _re5.search(r"(20\d{2})[.\-\s年]+(\d{1,2})[.\-\s月]+(\d{1,2})", s)
                    if m: return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
                    return ""
                def _extract_num5(v):
                    # 金额文本解析：纯数字直接转换；否则提取首个数字作为“建议值”参考；无法解析返回 None
                    if v is None: return None
                    s = str(v).strip()
                    if not s: return None
                    try:
                        return float(s)
                    except ValueError:
                        m = _re5.search(r"-?\d+(?:\.\d+)?", s)
                        return float(m.group(0)) if m else None
                def _sib_handler5(no):
                    base = str(no).split("-")[0]
                    for pfx in (base, base[:5]):
                        row = con.execute('SELECT "经办人" FROM project_contracts WHERE "编号" LIKE ? AND "经办人" IS NOT NULL AND TRIM(CAST("经办人" AS TEXT))<>"" ORDER BY "编号" LIMIT 1', (pfx + "%",)).fetchone()
                        if row: return str(row[0])
                    return ""
                repair5 = []  # (id, 编号, 列, 原值, 新值, 依据)：本次自动应用
                for r in con.execute('SELECT id,"编号","项目分类","项目名称","采购方式","合同金额（万元）","资金来源","合同主要内容","经办人" FROM project_contracts ORDER BY id').fetchall():
                    rid, no = r[0], str(r[1])
                    def g5(v): return "" if v is None else str(v).strip()
                    if not g5(r[2]): repair5.append((rid, no, "项目分类", "", "待补充", "必填字段为空，回填占位值"))
                    if not g5(r[3]): repair5.append((rid, no, "项目名称", "", "待补充", "必填字段为空，回填占位值"))
                    if not g5(r[4]): repair5.append((rid, no, "采购方式", "", "自行采购", "必填字段为空，按默认枚举回填"))
                    if r[5] is None or g5(r[5])=="": repair5.append((rid, no, "合同金额（万元）", "", 0, "必填字段为空，按 0 回填"))
                    if not g5(r[6]): repair5.append((rid, no, "资金来源", "", "其他", "必填字段为空，按兜底枚举“其他”回填"))
                    if not g5(r[7]): repair5.append((rid, no, "合同主要内容", "", "待补充", "必填字段为空，回填占位值"))
                    if not g5(r[8]):
                        sh = _sib_handler5(no)
                        repair5.append((rid, no, "经办人", "", sh if sh else "待补充", "必填字段为空，优先取同前缀兄弟合同经办人，否则回填占位值"))
                for r in con.execute('SELECT id,"编号","采购时间","开工时间","完工时间","竣工验收时间","下次招采时间" FROM project_contracts ORDER BY id').fetchall():
                    rid, no = r[0], str(r[1])
                    for ci, col in enumerate(date5):
                        v = r[2 + ci]
                        if v is None or str(v).strip() == "": continue
                        sv = str(v).strip()
                        if _re5.match(r"^\d{4}-\d{2}-\d{2}$", sv): continue
                        nv = _norm_date5(sv)
                        repair5.append((rid, no, col, sv, nv, "日期格式非 YYYY-MM-DD，提取日期或清空"))
                # 金额列非数字文本：按金额 CHECK（CAST(字段 AS REAL)>=0）语义，文本 CAST 后为 0、满足约束，
                # 故迁移不阻塞、不自动改写（避免将文号“潭财办发[2014]15号”、年份“2025年”、单位说明等误判为金额），仅输出清单供业务复核。
                amt_warn5 = []  # (id, 编号, 列, 原值, 建议值)
                for r in con.execute('SELECT id,"编号","合同金额（万元）","结算金额","已支付","未支付","决策的招标控制价" FROM project_contracts ORDER BY id').fetchall():
                    rid, no = r[0], str(r[1])
                    for ci, col in enumerate(amt5):
                        v = r[2 + ci]
                        if v is None or str(v).strip() == "": continue
                        try:
                            float(str(v).strip())
                            continue
                        except ValueError:
                            pass
                        sug = _extract_num5(v)
                        amt_warn5.append((rid, no, col, str(v).strip(), sug if sug is not None else "无法解析"))
                if repair5 or amt_warn5:
                    md5 = ["# A-05 数据修复清单（CHECK/NOT NULL 前置：存量违规逐条回填）", "",
                           "> 依据明确规则自动修复，**待业务复核确认**；回滚见 backups A-05 重建前备份。",
                           "> 枚举 CHECK（项目分类/采购方式/是否结算/资金来源/是否归档）延后至 A-08 清洗后启用，本次未启用。",
                           "", "## 一、自动回填（本次迁移已应用）", "",
                           "| 编号 | 字段 | 原值 | 修复值 | 修复依据 |", "|---|---|---|---|---|"]
                    for rid, no, col, old, new, why in repair5:
                        md5.append("| {} | {} | {} | {} | {} |".format(str(no).replace("|", "\\|").replace("\n", " "), str(col).replace("|", "\\|").replace("\n", " "), str(old).replace("|", "\\|").replace("\n", " "), str(new).replace("|", "\\|").replace("\n", " "), str(why).replace("|", "\\|").replace("\n", " ")))
                    md5 += ["", "## 二、金额列非数字文本（建议人工复核，本次未自动修改）", "",
                            "> 按 A-05 金额 CHECK（CAST(字段 AS REAL)>=0）语义，以下文本 CAST 后为 0，满足约束，故迁移不阻塞、不改写。",
                            "> 建议业务逐条核实后手动修正；“建议值”仅为正则提取的首个数字参考，不构成金额换算。",
                            "", "| id | 编号 | 字段 | 原值 | 建议值 |", "|---|---|---|---|---|"]
                    for rid, no, col, old, sug in amt_warn5:
                        md5.append("| {} | {} | {} | {} | {} |".format(rid, str(no).replace("|", "\\|").replace("\n", " "), str(col).replace("|", "\\|").replace("\n", " "), str(old).replace("|", "\\|").replace("\n", " "), str(sug).replace("|", "\\|").replace("\n", " ")))
                    md5_path = os.path.join(BASE_DIR, "docs", "A-05-数据修复清单.md")
                    os.makedirs(os.path.dirname(md5_path), exist_ok=True)
                    with open(md5_path, "w", encoding="utf-8", newline="\n") as f:
                        f.write("\n".join(md5) + "\n")
                    print(f"[INFO] A-05 自动回填 {len(repair5)} 条；金额文本待复核 {len(amt_warn5)} 条，清单：{md5_path}", file=sys.stderr)
                # 重建前整库备份
                os.makedirs(BACKUP_DIR, exist_ok=True)
                ts5 = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                bak5_before = os.path.join(BACKUP_DIR, f"合同台账_A05_重建前_{ts5}.db")
                if os.path.exists(bak5_before): os.remove(bak5_before)
                con.execute("VACUUM INTO '" + bak5_before.replace("'", "''") + "'")
                # 事务内：回填违规数据 → 复验零违规 → 重建 project_contracts（含 CHECK/NOT NULL）
                # 重建需 DROP 被 project_attachments 外键引用的 project_contracts 父表；外键开关仅对当前连接生效，
                # 临时关闭（PRAGMA 须在事务外执行），事务结束后恢复；API 连接(_conn)仍强制外键。
                con.execute("PRAGMA foreign_keys=OFF")
                con.execute("BEGIN")
                try:
                    for rid, no, col, old, new, why in repair5:
                        con.execute('UPDATE project_contracts SET "' + col + '"=? WHERE id=?', (new, rid))
                    bad5 = []
                    for r in con.execute('SELECT id,"编号","项目分类","项目名称","采购方式","合同金额（万元）","资金来源","合同主要内容","经办人","所属部门" FROM project_contracts').fetchall():
                        if any((r[i] is None or str(r[i]).strip()=="") for i in range(2, 10)): bad5.append((str(r[1]), "必填为空"))
                    for r in con.execute('SELECT id,"编号","合同金额（万元）","结算金额","已支付","未支付","决策的招标控制价" FROM project_contracts').fetchall():
                        for i, c in enumerate(amt5):
                            v = r[2 + i]
                            if v is None or str(v).strip() == "": continue
                            nf = _extract_num5(v)
                            # 非数字文本按 CHECK 语义（CAST 为 0）通过，仅负值视为违规
                            if nf is not None and nf < 0: bad5.append((str(r[1]), c + "为负"))
                    for r in con.execute('SELECT id,"编号","采购时间","开工时间","完工时间","竣工验收时间","下次招采时间" FROM project_contracts').fetchall():
                        for i, c in enumerate(date5):
                            sv = str(r[2 + i] or "").strip()
                            if sv and not _re5.match(r"^\d{4}-\d{2}-\d{2}$", sv): bad5.append((str(r[1]), c + "日期格式"))
                    if bad5:
                        raise RuntimeError("A-05 存量仍有违规，无法启用约束：" + "、".join(str(x) for x in bad5[:50]))
                    info5 = con.execute("PRAGMA table_info(project_contracts)").fetchall()
                    segs5 = []
                    for cid, cname, ctype, notnull, dflt, pk in info5:
                        seg = '"%s" %s' % (cname, ctype or "TEXT")
                        if cname in req5: seg += " NOT NULL"
                        if cname in amt5: seg += ' CHECK ("%s" IS NULL OR CAST("%s" AS REAL) >= 0)' % (cname, cname)
                        if cname in date5: seg += " CHECK (\"%s\" IS NULL OR \"%s\"='' OR \"%s\" GLOB '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]')" % (cname, cname, cname)
                        if dflt is not None: seg += " DEFAULT " + str(dflt)
                        if pk == 1:
                            if cname == "id" and (ctype or "").upper() == "INTEGER": seg = '"id" INTEGER PRIMARY KEY AUTOINCREMENT'
                            else: seg += " PRIMARY KEY"
                        segs5.append(seg)
                    segs5.append('FOREIGN KEY ("所属单位","所属科室") REFERENCES departments("单位名称","科室名称") ON DELETE RESTRICT ON UPDATE CASCADE')
                    con.execute("CREATE TABLE project_contracts_new (\n  " + ",\n  ".join(segs5) + "\n)")
                    old5 = [r[1] for r in con.execute("PRAGMA table_info(project_contracts)").fetchall()]
                    sel5 = ",".join('"%s"' % c for c in old5)
                    con.execute("INSERT INTO project_contracts_new (" + sel5 + ") SELECT " + sel5 + " FROM project_contracts")
                    # DROP/重命名父表前，先移除引用 project_contracts 的触发器（否则 RENAME 时触发解析报错），
                    # 事务提交后由下方 A-03 触发器块 CREATE TRIGGER IF NOT EXISTS 自动重建
                    con.execute("DROP TRIGGER IF EXISTS trg_project_contracts_dept_sync_i")
                    con.execute("DROP TRIGGER IF EXISTS trg_project_contracts_dept_sync_u")
                    con.execute("DROP TRIGGER IF EXISTS trg_departments_rename_sync")
                    con.execute("DROP TABLE project_contracts")
                    con.execute("ALTER TABLE project_contracts_new RENAME TO project_contracts")
                    con.execute('CREATE UNIQUE INDEX IF NOT EXISTS idx_project_contracts_no ON project_contracts("编号")')
                    con.execute("COMMIT")
                except Exception:
                    con.execute("ROLLBACK")
                    raise
                finally:
                    con.execute("PRAGMA foreign_keys=ON")
                # 重建后整库备份
                ts52 = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                bak5_after = os.path.join(BACKUP_DIR, f"合同台账_A05_重建后_{ts52}.db")
                if os.path.exists(bak5_after): os.remove(bak5_after)
                con.execute("VACUUM INTO '" + bak5_after.replace("'", "''") + "'")

            # A-06 字段类型规范化：project_attachments.文件大小 TEXT→INTEGER；日期列统一 YYYY-MM-DD；
            # 金额列声明类型已为 REAL，存量文本值须人工确认（迁移不猜测转换），输出清单供业务复核
            import re as _re6
            # ── 1) project_attachments.文件大小 → INTEGER：重建子表（保留主键与外键），存量字节数字串 CAST ──
            a_info6 = con.execute("PRAGMA table_info(project_attachments)").fetchall()
            a_types6 = {r[1]: r[2] for r in a_info6}
            if str(a_types6.get("文件大小", "") or "").upper() != "INTEGER":
                os.makedirs(BACKUP_DIR, exist_ok=True)
                ts6 = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                bak6 = os.path.join(BACKUP_DIR, f"合同台账_A06_附件重建前_{ts6}.db")
                if os.path.exists(bak6): os.remove(bak6)
                con.execute("VACUUM INTO '" + bak6.replace("'", "''") + "'")
                con.execute("BEGIN")
                try:
                    segs6 = []
                    for cid6, cn6, ct6, nn6, df6, pk6 in a_info6:
                        seg = '"%s" %s' % (cn6, "INTEGER" if cn6 == "文件大小" else (ct6 or "TEXT"))
                        if df6 is not None: seg += " DEFAULT " + str(df6)
                        if pk6 == 1: seg += " PRIMARY KEY"
                        segs6.append(seg)
                    segs6.append('FOREIGN KEY ("合同编号") REFERENCES project_contracts("编号") ON DELETE RESTRICT ON UPDATE RESTRICT')
                    con.execute("CREATE TABLE project_attachments_new (\n  " + ",\n  ".join(segs6) + "\n)")
                    aold6 = [r[1] for r in a_info6]
                    asel6 = ",".join('"%s"' % c for c in aold6)
                    acast6 = ",".join('CAST("文件大小" AS INTEGER)' if c == "文件大小" else '"%s"' % c for c in aold6)
                    con.execute("INSERT INTO project_attachments_new (" + asel6 + ") SELECT " + acast6 + " FROM project_attachments")
                    con.execute("DROP TABLE project_attachments")
                    con.execute("ALTER TABLE project_attachments_new RENAME TO project_attachments")
                    con.execute("COMMIT")
                except Exception:
                    con.execute("ROLLBACK")
                    raise
                ts62 = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                bak6b = os.path.join(BACKUP_DIR, f"合同台账_A06_附件重建后_{ts62}.db")
                if os.path.exists(bak6b): os.remove(bak6b)
                con.execute("VACUUM INTO '" + bak6b.replace("'", "''") + "'")
                print("[INFO] A-06 project_attachments.文件大小 已重建为 INTEGER", file=sys.stderr)
            # ── 2) 日期列统一 YYYY-MM-DD（含归档时间）；无法解析的仅列清单，禁止静默置空 ──
            date6 = ["采购时间","开工时间","完工时间","竣工验收时间","下次招采时间","归档时间"]
            def _norm_date6(v):
                s = str(v or "").strip()
                if not s: return None
                m = _re6.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})([ T].*)?$", s)
                if m: return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
                m = _re6.search(r"(20\d{2})[.\-\s年]+(\d{1,2})[.\-\s月]+(\d{1,2})", s)
                if m: return "%04d-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
                return None
            def _extract_num6(v):
                if v is None: return None
                s = str(v).strip()
                if not s: return None
                try:
                    return float(s)
                except ValueError:
                    m = _re6.search(r"-?\d+(?:\.\d+)?", s)
                    return float(m.group(0)) if m else None
            date_fix6 = []   # (编号, 列, 原值, 新值)：自动规范化
            date_bad6 = []   # (编号, 列, 原值)：无法解析，待业务确认
            for r in con.execute('SELECT id,"编号","采购时间","开工时间","完工时间","竣工验收时间","下次招采时间","归档时间" FROM project_contracts').fetchall():
                rid, no = r[0], str(r[1])
                for ci, col in enumerate(date6):
                    v = r[2 + ci]
                    if v is None or str(v).strip() == "": continue
                    sv = str(v).strip()
                    if _re6.match(r"^\d{4}-\d{2}-\d{2}$", sv): continue
                    nv = _norm_date6(sv)
                    if nv is None: date_bad6.append((no, col, sv))
                    else: date_fix6.append((no, col, sv, nv))
            if date_fix6:
                con.execute("BEGIN")
                try:
                    for no, col, old, nv in date_fix6:
                        con.execute('UPDATE project_contracts SET "' + col + '"=? WHERE "编号"=?', (nv, no))
                    con.execute("COMMIT")
                except Exception:
                    con.execute("ROLLBACK")
                    raise
            md6 = ["# A-06 日期规范化清单（统一 YYYY-MM-DD）", "",
                   "> 自动规范化已应用；无法解析的日期仅列出，**待业务确认后回填**，未静默置空。",
                   "", "## 一、自动规范化（本次迁移已应用）", "",
                   "| 编号 | 字段 | 原值 | 新值 |", "|---|---|---|---|"]
            for no, col, old, nv in date_fix6:
                md6.append("| {} | {} | {} | {} |".format(str(no).replace("|", "\\|").replace("\n", " "), col, old.replace("|", "\\|").replace("\n", " "), nv))
            md6 += ["", "## 二、无法解析（待业务确认回填，未修改）", "",
                    "| 编号 | 字段 | 原值 |", "|---|---|---|"]
            for no, col, old in date_bad6:
                md6.append("| {} | {} | {} |".format(str(no).replace("|", "\\|").replace("\n", " "), col, old.replace("|", "\\|").replace("\n", " ")))
            md6_path = os.path.join(BASE_DIR, "docs", "A-06-日期规范化清单.md")
            os.makedirs(os.path.dirname(md6_path), exist_ok=True)
            with open(md6_path, "w", encoding="utf-8", newline="\n") as f:
                f.write("\n".join(md6) + "\n")
            print(f"[INFO] A-06 日期自动规范化 {len(date_fix6)} 条；无法解析待确认 {len(date_bad6)} 条，清单：{md6_path}", file=sys.stderr)
            # ── 3) 金额列文本：声明类型已为 REAL；存量文本值迁移不猜测转换，输出清单待业务确认 ──
            amt6 = ["合同金额（万元）","结算金额","已支付","未支付","决策的招标控制价"]
            amt_warn6 = []  # (id, 编号, 列, 原值, 建议值)
            for r in con.execute('SELECT id,"编号","合同金额（万元）","结算金额","已支付","未支付","决策的招标控制价" FROM project_contracts ORDER BY id').fetchall():
                rid, no = r[0], str(r[1])
                for ci, col in enumerate(amt6):
                    v = r[2 + ci]
                    if v is None or str(v).strip() == "": continue
                    try:
                        float(str(v).strip())
                        continue
                    except ValueError:
                        pass
                    sug = _extract_num6(v)
                    amt_warn6.append((rid, no, col, str(v).strip(), sug if sug is not None else "无法解析"))
            if amt_warn6:
                md6a = ["# A-06 金额类型转换确认清单（存量文本值，待业务确认）", "",
                        "> 实况：金额列声明类型已为 REAL（合同金额/结算金额/已支付/未支付/决策的招标控制价），残留问题仅是**存量文本值**。",
                        "> 迁移脚本**不做猜测转换**（避免把文号/年份/单位说明误判为金额）；请业务逐条核对正确金额后，",
                        "> 通过系统“修改合同”表单或 SQL 更新为数字（写入 REAL 列时自动转为数值）。",
                        "> “建议值”仅为正则提取的首个数字参考，不构成金额换算。",
                        "", "| id | 编号 | 字段 | 原值 | 建议值 | 状态 |", "|---|---|---|---|---|---|"]
                for rid, no, col, old, sug in amt_warn6:
                    md6a.append("| {} | {} | {} | {} | {} | 待确认 |".format(rid, str(no).replace("|", "\\|").replace("\n", " "), col, old.replace("|", "\\|").replace("\n", " "), str(sug).replace("|", "\\|").replace("\n", " ")))
                md6a_path = os.path.join(BASE_DIR, "docs", "A-06-金额类型转换确认清单.md")
                os.makedirs(os.path.dirname(md6a_path), exist_ok=True)
                with open(md6a_path, "w", encoding="utf-8", newline="\n") as f:
                    f.write("\n".join(md6a) + "\n")
                print(f"[INFO] A-06 金额列文本值 {len(amt_warn6)} 条待业务确认，清单：{md6a_path}", file=sys.stderr)
            else:
                print("[INFO] A-06 金额列无文本残留，无需转换", file=sys.stderr)
            # A-03 同步触发器：所属部门(字符串) ↔ 所属单位/所属科室(外键列)，保证 API 写入一致；部门改名级联同步
            con.execute('''CREATE TRIGGER IF NOT EXISTS trg_project_contracts_dept_sync_i
AFTER INSERT ON project_contracts
FOR EACH ROW WHEN (NEW."所属部门" IS NOT NULL AND TRIM(NEW."所属部门") <> '')
BEGIN
  UPDATE project_contracts SET
    "所属单位" = CASE WHEN instr(NEW."所属部门",'-')>0 THEN substr(NEW."所属部门",1,instr(NEW."所属部门",'-')-1) ELSE '' END,
    "所属科室" = CASE WHEN instr(NEW."所属部门",'-')>0 THEN substr(NEW."所属部门",instr(NEW."所属部门",'-')+1) ELSE '' END
  WHERE id = NEW.id;
END''')
            con.execute('''CREATE TRIGGER IF NOT EXISTS trg_project_contracts_dept_sync_u
AFTER UPDATE OF "所属部门" ON project_contracts
FOR EACH ROW WHEN (NEW."所属部门" IS NOT NULL AND TRIM(NEW."所属部门") <> '')
BEGIN
  UPDATE project_contracts SET
    "所属单位" = CASE WHEN instr(NEW."所属部门",'-')>0 THEN substr(NEW."所属部门",1,instr(NEW."所属部门",'-')-1) ELSE '' END,
    "所属科室" = CASE WHEN instr(NEW."所属部门",'-')>0 THEN substr(NEW."所属部门",instr(NEW."所属部门",'-')+1) ELSE '' END
  WHERE id = NEW.id;
END''')
            con.execute('''CREATE TRIGGER IF NOT EXISTS trg_users_dept_sync_i
AFTER INSERT ON users
FOR EACH ROW WHEN (NEW."所属部门" IS NOT NULL AND TRIM(NEW."所属部门") <> '')
BEGIN
  UPDATE users SET
    "所属单位" = CASE WHEN instr(NEW."所属部门",'-')>0 THEN substr(NEW."所属部门",1,instr(NEW."所属部门",'-')-1) ELSE '' END,
    "所属科室" = CASE WHEN instr(NEW."所属部门",'-')>0 THEN substr(NEW."所属部门",instr(NEW."所属部门",'-')+1) ELSE '' END
  WHERE "用户名" = NEW."用户名";
END''')
            con.execute('''CREATE TRIGGER IF NOT EXISTS trg_users_dept_sync_u
AFTER UPDATE OF "所属部门" ON users
FOR EACH ROW WHEN (NEW."所属部门" IS NOT NULL AND TRIM(NEW."所属部门") <> '')
BEGIN
  UPDATE users SET
    "所属单位" = CASE WHEN instr(NEW."所属部门",'-')>0 THEN substr(NEW."所属部门",1,instr(NEW."所属部门",'-')-1) ELSE '' END,
    "所属科室" = CASE WHEN instr(NEW."所属部门",'-')>0 THEN substr(NEW."所属部门",instr(NEW."所属部门",'-')+1) ELSE '' END
  WHERE "用户名" = NEW."用户名";
END''')
            con.execute('''CREATE TRIGGER IF NOT EXISTS trg_departments_rename_sync
AFTER UPDATE OF "单位名称","科室名称" ON departments
FOR EACH ROW
BEGIN
  UPDATE project_contracts SET "所属部门" = NEW."单位名称" || '-' || NEW."科室名称"
    WHERE "所属部门" = OLD."单位名称" || '-' || OLD."科室名称";
  UPDATE users SET "所属部门" = NEW."单位名称" || '-' || NEW."科室名称"
    WHERE "所属部门" = OLD."单位名称" || '-' || OLD."科室名称";
END''')

            # A-04 project_contracts 常用查询列索引（幂等；编号唯一索引 A-01 已建，直接复用不重复创建）
            # 经办人为多值列（顿号/斜杠分隔多人），索引价值有限，作为过渡索引；档位 B 拆分多值列后将废弃本索引
            con.execute('CREATE INDEX IF NOT EXISTS idx_project_contracts_department ON project_contracts("所属部门")')
            con.execute('CREATE INDEX IF NOT EXISTS idx_project_contracts_procurement_date ON project_contracts("采购时间")')
            con.execute('CREATE INDEX IF NOT EXISTS idx_project_contracts_next_bid_date ON project_contracts("下次招采时间")')
            con.execute('CREATE INDEX IF NOT EXISTS idx_project_contracts_archived ON project_contracts("是否归档")')
            con.execute('CREATE INDEX IF NOT EXISTS idx_project_contracts_handler ON project_contracts("经办人")')

            con.execute('INSERT OR IGNORE INTO settings ("配置项","配置值") VALUES (?,?)', ("log_keep_days", "90"))


            con.execute('INSERT OR IGNORE INTO settings ("配置项","配置值") VALUES (?,?)', ("default_attach_max_mb", "20"))
            con.commit()
        finally:
            con.close()
    if fresh:
        if os.path.exists(EXCEL_PATH):
            _migrate_from_excel()
        _seed_defaults()

def _norm(v):
    if v is None: return ""
    if isinstance(v, datetime.datetime): return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date): return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)): return v
    return str(v).strip()

def _insert_dict(con, table, colnames, rowdict):
    cols = list(colnames)
    sql = "INSERT INTO " + table + " (" + ",".join('"'+c+'"' for c in cols) + ") VALUES (" + ",".join("?"*len(cols)) + ")"
    con.execute(sql, [rowdict.get(c, "") for c in cols])

def _migrate_from_excel():
    if not os.path.exists(EXCEL_PATH): return
    print("[INFO] 从 Excel 迁移数据到 SQLite...")
    wb = openpyxl.load_workbook(EXCEL_PATH)
    with db_lock:
        con = _conn()
        try:
            ws = wb["合同台账"]
            rows = list(ws.iter_rows(values_only=True))
            if rows and len(rows) > 1:
                hdr = [str(x).strip() if x else "" for x in rows[0]]
                for ri in range(1, len(rows)):
                    vals = {cn: "" for cn in CONTRACT_COLUMNS}
                    src = rows[ri]
                    for ci, cn in enumerate(hdr):
                        if cn == "\u7aef\u5de5\u9a8c\u6536\u65f6\u95f4": cn = "\u7ae3\u5de5\u9a8c\u6536\u65f6\u95f4"
                        if cn in ("经办人/分管领导","甲方现场负责人"):
                            cn = {"经办人/分管领导":"经办人","甲方现场负责人":"现场负责人"}[cn]
                        if cn in vals:
                            vals[cn] = _norm(src[ci]) if ci < len(src) else ""
                    cols = ["id"] + list(CONTRACT_COLUMNS)
                    sql = "INSERT INTO project_contracts (" + ",".join('"'+c+'"' for c in cols) + ") VALUES (" + ",".join("?"*len(cols)) + ")"
                    con.execute(sql, [ri+1] + [vals[cn] for cn in CONTRACT_COLUMNS])
            ws = wb["用户表"]
            rows = list(ws.iter_rows(values_only=True))
            if rows and len(rows) > 1:
                hdr = [str(x).strip() if x else "" for x in rows[0]]
                umap = {"用户名":"用户名","密码哈希":"密码","密码":"密码","姓名":"姓名","所属部门":"所属部门","角色":"角色"}
                for r in rows[1:]:
                    d = {}
                    for ci, hh in enumerate(hdr):
                        if hh in umap:
                            d[umap[hh]] = _norm(r[ci]) if ci < len(r) else ""
                    if not d.get("用户名"): continue
                    _insert_dict(con, "users", ["用户名","密码","姓名","所属部门","角色"], d)
            ws = wb["部门列表"]
            rows = list(ws.iter_rows(values_only=True))
            if rows and len(rows) > 1:
                hdr = [str(x).strip() if x else "" for x in rows[0]]
                for r in rows[1:]:
                    d = {hh: (_norm(r[i]) if i < len(r) else "") for i, hh in enumerate(hdr) if hh in ("单位名称","科室名称")}
                    if not d.get("单位名称") and not d.get("科室名称"): continue
                    _insert_dict(con, "departments", ["单位名称","科室名称"], d)
            ws = wb["系统配置"]
            rows = list(ws.iter_rows(values_only=True))
            if rows and len(rows) > 1:
                hdr = [str(x).strip() if x else "" for x in rows[0]]
                kcol = "配置项" if "配置项" in hdr else ("键" if "键" in hdr else None)
                vcol = "配置值" if "配置值" in hdr else ("值" if "值" in hdr else None)
                if kcol and vcol:
                    ki = hdr.index(kcol); vi = hdr.index(vcol)
                    for r in rows[1:]:
                        k = _norm(r[ki]) if ki < len(r) else ""
                        v = _norm(r[vi]) if vi < len(r) else ""
                        if k:
                            con.execute('INSERT INTO settings ("配置项","配置值") VALUES (?,?)', (k, v))
            ws = wb[ATTACH_SHEET]
            rows = list(ws.iter_rows(values_only=True))
            if rows and len(rows) > 1:
                hdr = [str(x).strip() if x else "" for x in rows[0]]
                for r in rows[1:]:
                    d = {hh: (_norm(r[i]) if i < len(r) else "") for i, hh in enumerate(hdr) if hh in ATTACH_HEADERS}
                    if not d.get("附件ID"): continue
                    _insert_dict(con, "project_attachments", ATTACH_HEADERS, d)
            con.commit()
        except Exception:
            con.rollback(); raise
        finally:
            con.close()
    print("[INFO] 迁移完成。")

def _seed_defaults():
    with db_lock:
        con = _conn()
        try:
            # 注意：必须先插入部门再插入用户——用户表含外键(所属单位,所属科室)→departments，
            # 且 trg_users_dept_sync_i 触发器会回填外键列，先存在部门表中才能通过外键校验。
            if con.execute('SELECT COUNT(*) FROM departments').fetchone()[0] == 0:
                con.execute('INSERT INTO departments ("单位名称","科室名称") VALUES (?,?)',
                            ("湘潭市机关事务管理局本级", "办公室"))
            if con.execute('SELECT COUNT(*) FROM users').fetchone()[0] == 0:
                con.execute('INSERT INTO users ("用户名","密码","姓名","所属部门","角色") VALUES (?,?,?,?,?)',
                            ("admin", hash_pw("123456"), "系统管理员", "湘潭市机关事务管理局本级-办公室", "管理员"))
            if con.execute('SELECT COUNT(*) FROM settings').fetchone()[0] == 0:
                con.execute('INSERT INTO settings ("配置项","配置值") VALUES (?,?)', ("default_remind_days", "60"))
            con.commit()
        finally:
            con.close()

# ---- Attachments ----
def read_attachments():
    return db_query("SELECT * FROM project_attachments ORDER BY rowid")

def write_attachments(rows):
    with db_lock:
        _backup()
        con = _conn()
        try:
            con.execute("DELETE FROM project_attachments")
            for r in rows:
                _insert_dict(con, "project_attachments", ATTACH_HEADERS, r)
            con.commit()
        except Exception:
            con.rollback(); raise
        finally:
            con.close()

def attachment_json(row):
    tlabel = row.get("附件类型","")
    tkey = next((k for k,v in ATTACH_TYPES.items() if v==tlabel), "")
    return {
        "id":row.get("附件ID",""),"contract_number":row.get("合同编号",""),
        "type":tkey,"type_label":tlabel,"display_name":row.get("显示名称",""),
        "size":row.get("文件大小",""),"uploader":row.get("上传人",""),
        "uploaded_at":row.get("上传时间",""),
    }

def attach_map_by_number():
    m={}
    for r in read_attachments():
        m.setdefault(str(r.get("合同编号","")),[]).append(attachment_json(r))
    return m

def remove_contract_attachments(number):
    rows = read_attachments()
    keep=[]
    for r in rows:
        if str(r.get("合同编号",""))==str(number):
            fp = os.path.join(ATTACH_DIR, r.get("存储文件名",""))
            try:
                if os.path.exists(fp): os.remove(fp)
            except Exception: pass
        else: keep.append(r)
    if len(keep)!=len(rows): write_attachments(keep)

def read_contracts():
    cached = _cache_get("contracts")
    if cached is None:
        rows = db_query("SELECT * FROM project_contracts ORDER BY id")
        result = []
        for r in rows:
            item = {"_row": r.get("id")}
            for cn in CONTRACT_COLUMNS:
                v = r.get(cn, "")
                item[cn] = "" if v is None else v
            for xc in ("是否归档","归档人","归档时间","短信提醒"):
                v2 = r.get(xc, "")
                item[xc] = "" if v2 is None else v2
            result.append(item)
        _cache_put("contracts", result)
        cached = result
    return [dict(x) for x in cached]      # 浅拷贝：调用方可安全挂载 attachments 等附加字段

def c2r(c):
    return {
        "id":f"row_{c['_row']}","number":str(c.get("编号","")),
        "category":c.get("项目分类",""),"name":c.get("项目名称",""),
        "decision_basis":c.get("决策依据",""),"bid_control_price":str(c.get("决策的招标控制价","")),
        "price_basis":c.get("招标控制价决策的依据",""),"procurement_method":c.get("采购方式",""),
        "winner":c.get("中标单位",""),"actual_implementor":c.get("实际实施单位",""),
        "amount":_num(c.get("合同金额（万元）")),
        "fund_source":c.get("资金来源",""),
        "main_content":c.get("合同主要内容",""),"procurement_date":c.get("采购时间",""),
        "start_date":c.get("开工时间",""),"end_date":c.get("完工时间",""),
        "acceptance_date":c.get("竣工验收时间",""),"site_manager":c.get("现场负责人",""),
        "has_change":c.get("是否有签证或变更",""),"change_detail":c.get("签证或变更情况",""),
        "is_settled":c.get("是否结算",""),"settled_amount":_num(c.get("结算金额")),
        "paid_amount":_num(c.get("已支付")),"unpaid_amount":_num(c.get("未支付")),
        "term":c.get("期限",""),"handler":c.get("经办人",""),
        "remark":c.get("备注",""),"department":c.get("所属部门",""),
        "next_bid_date":c.get("下次招采时间",""),"remind_days":_int(c.get("提醒提前天数")),
        "archived":c.get("是否归档","")=="是","archive_by":c.get("归档人",""),"archive_at":c.get("归档时间",""),
        "sms_switch":c.get("短信提醒","") or "开启",
    }

def _num(v):
    if v=="" or v is None: return None
    try: return round(float(v),2)
    except: return None

def _int(v):
    if v=="" or v is None: return None
    try: return int(float(v))
    except: return None

# ──── 合同编号（6 位）生成：年度后两位 + 科室（中心）序号 + 中心组室序号 + 顺序号 ────
# 规则实现见 contract_no 模块；此处只负责取“编号归属部门”并查询对应表的既有编号。
def _number_rule_dept(user=None, dept=""):
    """编号归属部门：显式传入优先，否则取当前用户所属部门。"""
    d = str(dept or "").strip()
    if d:
        return d
    u = user or get_cur_user()
    return str((u or {}).get("\u6240\u5c5e\u90e8\u95e8", "") or "").strip()

def _existing_numbers(table, column):
    """读取某表某列的全部既有编号（用于计算同前缀最大顺序号）。"""
    return [r.get("n") for r in db_query('SELECT "%s" AS n FROM %s' % (column, table))]

def next_contract_number(table, column, dept, year=None):
    """按《合同编号规则》计算下一个 6 位编号；返回 (编号, 错误信息)。

    - 未配置编号规则的科室（中心）/中心组室：返回 (None, 原因)，禁止自动生成或猜测编号；
    - 同一“年度 + 科室（中心）+ 中心组室”前缀下取既有最大顺序号 +1，无记录时从 01 开始；
    - 顺序号超过 99：返回 (None, 原因)，需扩展编号规则。
    """
    try:
        p = contract_no.plan(dept=dept, year=year)
    except contract_no.NoRuleError as e:
        return None, str(e)
    try:
        return contract_no.next_number(_existing_numbers(table, column), p["prefix"]), None
    except OverflowError as e:
        return None, str(e)

def _pending_unpaid(c):
    # 待支付金额：优先取台账"未支付"（含0）；为空时按 合同金额-已支付 计算（空按0）
    up = _num(c.get("未支付"))
    if up is not None:
        return up
    amt = _num(c.get("合同金额（万元）")) or 0
    paid = _num(c.get("已支付")) or 0
    return round(amt - paid, 2)

def _pending_count(cts):
    # 待支付项目数：未归档 且 未支付金额>0
    n = 0
    for c in cts:
        if c.get("是否归档", "") == "是":
            continue
        if _pending_unpaid(c) > 0:
            n += 1
    return n

def _split_names(s):
    # 经办人姓名列表：兼容中文顿号、顿号、斜杠、分号等分隔符
    if s is None: return []
    return [p.strip() for p in str(s).replace("\u3001", ",").replace("/", ",").replace("\uff1b", ";").split(",") if p.strip()]

def upd_cell(rn,cn,val):
    if cn not in COL_IDX: return
    db_exec('UPDATE project_contracts SET "' + cn + '"=? WHERE id=?', (val, rn))

def upd_row(rn,d):
    sets = [cn for cn in CONTRACT_COLUMNS if cn in d]
    if not sets: return
    params = [d[cn] for cn in sets] + [rn]
    sql = "UPDATE project_contracts SET " + ",".join('"'+cn+'"=?' for cn in sets) + " WHERE id=?"
    db_exec(sql, params)

def del_row(rn):
    db_exec("DELETE FROM project_contracts WHERE id=?", (rn,))

def app_row(d):
    cols = list(CONTRACT_COLUMNS)
    with db_lock:
        _backup()
        con = _conn()
        try:
            sql = "INSERT INTO project_contracts (" + ",".join('"'+c+'"' for c in cols) + ") VALUES (" + ",".join("?"*len(cols)) + ")"
            cur = con.execute(sql, [d.get(cn, "") for cn in cols])
            con.commit()
            _invalidate_caches()
            return cur.lastrowid
        except Exception:
            con.rollback(); raise
        finally:
            con.close()

# ──── Helpers ────
def get_all_users(): return db_query('SELECT * FROM users')
def get_user(un):
    for u in get_all_users():
        if u["\u7528\u6237\u540d"] == un: return u
    return None

LEADER_DEPT_NAMES = ("\u5c40\u9886\u5bfc", "\u4e3b\u4efb\u73ed\u5b50", "\u4e2d\u5fc3\u73ed\u5b50")
def is_leadership_dept(dept):
    return str(dept or "").strip().split("-")[-1] in LEADER_DEPT_NAMES

def norm_managed(departments, allowed=None):
    if allowed is None: allowed = get_dept_strings()
    if isinstance(departments, (list, tuple)):
        parts = [str(x).strip() for x in departments if str(x).strip()]
    else:
        parts = [p.strip() for p in str(departments or "").replace("\u3001", ",").replace("\uff1b", ";").split(",") if p.strip()]
    bad = [p for p in parts if p not in allowed]
    if bad: raise ValueError("\u5206\u7ba1\u90e8\u95e8\u4e0d\u5b58\u5728\uff1a" + "\u3001".join(bad))
    return "\u3001".join(parts)
def get_all_depts(): return db_query('SELECT * FROM departments')

# ─────────────────────────────────────────────────────────────────────────────
# 部门显示名工具（兼容 Python 3.8 ~ 3.12，勿删本注释）
#
# 背景：Python 3.12 起（PEP 701）才允许 f-string 的「表达式部分」出现反斜杠（含 \uXXXX 转义），
#       3.11 及更早版本在编译源码阶段即报语法错误：
#           SyntaxError: f-string expression part cannot include a backslash
#       本项目源码中的中文列名以 \uXXXX 转义书写，故统一约定：
#           ✗ 禁止：f"{d['\u5355\u4f4d\u540d\u79f0']}-{d['\u79d1\u5ba4\u540d\u79f0']}"
#           ✓ 允许：先把取值语句放到 f-string 外面，f-string 内只引用变量（表达式内无反斜杠）
#       说明：该限制只针对「反斜杠」，与中文/非 ASCII 字符无关。
#       本函数不使用 3.9+/3.10+/3.12+ 语法，Python 3.8 与 3.12 均可运行。
# ─────────────────────────────────────────────────────────────────────────────
def _dept_display(d):
    """返回部门的显示名「单位名称-科室名称」。

    参数：d —— get_all_depts() 返回的一行（dict，键为数据库中文列名）
    返回：形如 "湘潭市机关事务管理局后勤中心-综合组" 的字符串
    """
    unit = d["\u5355\u4f4d\u540d\u79f0"]      # 单位名称
    room = d["\u79d1\u5ba4\u540d\u79f0"]      # 科室名称
    return f"{unit}-{room}"                   # f-string 内只引用变量：无反斜杠、无同类引号嵌套

def get_dept_strings():
    """返回全部部门的「单位名称-科室名称」字符串列表（部门下拉选项、部门字符串校验用）。

    兼容性：改为调用 _dept_display() 组装，避免在 f-string 表达式内书写 Unicode 转义（形如 \\uXXXX），
            从而同时兼容 Python 3.8（服务器）与 3.12（开发机）。
    """
    # 原实现（仅 Python 3.12+ 可用：f-string 表达式内含 \uXXXX 反斜杠，3.11 及以前报 SyntaxError）
    # 为便于回溯保留在此，禁止启用：
    # def get_dept_strings(): return [f"{d['\u5355\u4f4d\u540d\u79f0']}-{d['\u79d1\u5ba4\u540d\u79f0']}" for d in get_all_depts()]
    return [_dept_display(d) for d in get_all_depts()]

PHONE_MAX_LEN = 20
_PHONE_CHARS = set("0123456789+-() ")
def norm_phone(value):
    """用户“电话”字段规范化：可空；非空时仅允许数字、+、-、括号与空格，最长 20 位且至少含 1 位数字。"""
    s = str(value if value is not None else "").strip()
    if not s: return ""
    if len(s) > PHONE_MAX_LEN: raise ValueError("电话长度不能超过 %d 位" % PHONE_MAX_LEN)
    if any(ch not in _PHONE_CHARS for ch in s) or not any(ch.isdigit() for ch in s):
        raise ValueError("电话格式不正确：仅支持数字、+、-、括号和空格")
    return s
def get_cfg(k):
    rows = db_query('SELECT "配置值" FROM settings WHERE "配置项"=?', (k,))
    return rows[0]["配置值"] if rows else None
def set_cfg(k,v):
    db_exec('INSERT INTO settings ("配置项","配置值") VALUES (?,?) '
            'ON CONFLICT("配置项") DO UPDATE SET "配置值"=excluded."配置值"', (k, str(v)))
def hash_pw(p): return hashlib.sha256((p+SALT).encode()).hexdigest()
def parse_date(s):
    if not s or str(s).strip()=="": return None
    ds = str(s).strip().replace("/","-")
    try:
        p=ds.split("-"); return datetime.date(int(p[0]),int(p[1]),int(p[2]))
    except: return None
def get_default_days():
    v = get_cfg("default_remind_days")
    try: return int(v)
    except: return 60

def get_attach_max_mb():
    v = get_cfg("default_attach_max_mb")
    try:
        n = int(float(v))
        if n < 1: return DEFAULT_ATTACH_MAX_MB
        return n
    except Exception:
        return DEFAULT_ATTACH_MAX_MB

def get_log_keep_days():
    v = get_cfg("log_keep_days")
    try:
        n = int(v)
        if n < 1: return LOG_KEEP_DAYS
        return n
    except Exception:
        return LOG_KEEP_DAYS
def get_cur_user():
    un = session.get("username")
    if not un: return None
    return get_user(un)

# ---- 操作日志（审计） ----
LOG_KEEP_DAYS = 90
_last_log_cleanup = {"date": ""}

def _log_exec(sql, params=()):
    con = _conn()
    try:
        con.execute(sql, params)
        con.commit()
    finally:
        con.close()

def _client_ip():
    ip = request.headers.get("X-Forwarded-For", "") or ""
    if ip and "," in ip: ip = ip.split(",")[0].strip()
    return ip or (request.remote_addr or "")

def _cleanup_logs(days=None):
    if days is None: days = get_log_keep_days()
    cutoff = (datetime.datetime.now() - datetime.timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
    try:
        _log_exec('DELETE FROM operation_logs WHERE "时间" < ?', (cutoff,))
    except Exception:
        pass
    try:
        _log_exec('DELETE FROM field_change_logs WHERE "时间" < ?', (cutoff,))
    except Exception:
        pass

def _maybe_cleanup_logs():
    today = datetime.date.today().isoformat()
    if _last_log_cleanup.get("date") == today: return
    _last_log_cleanup["date"] = today
    _cleanup_logs()

def write_log(module, action, obj="", detail="", result="成功"):
    try:
        u = get_cur_user()
        un = u["用户名"] if u else (session.get("username", "") or "")
        nm = u["姓名"] if u else ""
        rl = u["角色"] if u else ""
    except Exception:
        un = session.get("username", "") or ""; nm = ""; rl = ""
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        _log_exec('INSERT INTO operation_logs ("时间","用户名","姓名","角色","IP","模块","操作类型","对象","详情","结果") VALUES (?,?,?,?,?,?,?,?,?,?)',
                  (ts, un, nm, rl, _client_ip(), module, action, str(obj), str(detail), result))
        _maybe_cleanup_logs()
    except Exception:
        pass

def _same_number(a, b):
    """两侧是否都是纯数字且数值相等（日志去噪：'0' 与 '0.0'、'150000' 与 '150000.0' 视为未变更）。"""
    import re as _re
    pa = _re.compile(r"^-?\d+(\.\d+)?$")
    sa, sb = str(a or "").strip(), str(b or "").strip()
    if not pa.match(sa) or not pa.match(sb):
        return False
    try:
        return float(sa) == float(sb)
    except Exception:
        return False

def log_field_changes(module, obj_no, obj_name, old_vals, new_vals, fields=None):
    """字段级变更日志：逐字段比较 old_vals/new_vals（{数据库列名: 值}），仅记录发生变化的字段。

    返回记录条数；任何异常都不影响主流程。
    数值比较：两侧都是纯数字（含 "0" 与 "0.0"、"150000" 与 "150000.0" 这类同值不同写法）时按数值判定，
    避免产生"值未变却记一条变更"的噪声日志。
    """
    try:
        changes = []
        for cn, nv in (new_vals or {}).items():
            if fields is not None and cn not in fields:
                continue
            ov = (old_vals or {}).get(cn, "")
            a = "" if ov is None else str(ov)
            b = "" if nv is None else str(nv)
            if a == b:
                continue
            if _same_number(a, b):
                continue                      # 数值等价（如 0 / 0.0），不视为变更
            changes.append((cn, a, b))
        if not changes:
            return 0
        try:
            u = get_cur_user()
            un = u["用户名"] if u else (session.get("username", "") or "")
            nm = u["姓名"] if u else ""
            rl = u["角色"] if u else ""
        except Exception:
            un = session.get("username", "") or ""; nm = ""; rl = ""
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ip = _client_ip()
        with db_lock:
            con = _conn()
            try:
                for cn, a, b in changes:
                    con.execute(
                        'INSERT INTO field_change_logs ("时间","用户名","姓名","角色","IP","模块","对象编号","对象名称","字段","原值","新值")'
                        ' VALUES (?,?,?,?,?,?,?,?,?,?,?)',
                        (ts, un, nm, rl, ip, module, str(obj_no), str(obj_name), cn, a, b))
                con.commit()
            except Exception:
                con.rollback(); raise
            finally:
                con.close()
        _maybe_cleanup_logs()
        return len(changes)
    except Exception:
        return 0

def _visible_depts(u=None):
    # 可见部门集合：本部门 + 分管部门；管理员返回 None 表示全部可见
    if u is None: u = get_cur_user()
    if not u or u.get("\u89d2\u8272") == "\u7ba1\u7406\u5458": return None
    depts = [str(u.get("\u6240\u5c5e\u90e8\u95e8", "")).strip()]
    md = str(u.get("\u5206\u7ba1\u90e8\u95e8", "") or "").strip()
    depts += [p.strip() for p in md.replace("\u3001", ",").replace("\uff1b", ";").split(",") if p.strip()]
    return set(depts)

def _scope_contracts(cts):
    u = get_cur_user()
    if u:
        vis = _visible_depts(u)
        if vis is not None:
            cts = [c for c in cts if str(c.get("\u6240\u5c5e\u90e8\u95e8", "")).strip() in vis]
    return cts

def _scope_assets(items):
    """资产合同数据范围：管理员全部；普通/查询用户按“管理部门”限本部门+分管部门（与 _scope_contracts 口径一致）。"""
    u = get_cur_user()
    if u:
        vis = _visible_depts(u)
        if vis is not None:
            items = [a for a in items if str(a.get("dept", "")).strip() in vis]
    return items

def _can_write(c):
    # 写权限：管理员全部；普通用户限可见部门（本部门+分管部门）
    u = get_cur_user()
    if not u or u["\u89d2\u8272"]=="\u7ba1\u7406\u5458": return True
    vis = _visible_depts(u)
    return vis is not None and str(c.get("\u6240\u5c5e\u90e8\u95e8", "")).strip() in vis

# ──── Auth Decorators ────
def login_required(f):
    @wraps(f)
    def wrapper(*a,**kw):
        if "username" not in session: return jsonify({"error":"\u672a\u767b\u5f55"}),401
        return f(*a,**kw)
    return wrapper

def admin_required(f):
    @wraps(f)
    def wrapper(*a,**kw):
        if "username" not in session: return jsonify({"error":"\u672a\u767b\u5f55"}),401
        u = get_user(session["username"])
        if not u or u.get("\u89d2\u8272")!="\u7ba1\u7406\u5458": return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
        return f(*a,**kw)
    return wrapper

# ═══════════════════ API Routes ═══════════════════

# ─── Static ───
@app.route("/static/<path:fn>")
def serve_static(fn): return send_from_directory(STATIC_DIR,fn)

# ─── Auth ───
@app.route("/api/login",methods=["POST"])
def api_login():
    d = request.get_json(silent=True) or {}
    un = d.get("username","").strip()
    pw = d.get("password","").strip()
    if not un or not pw: return jsonify({"error":"\u7528\u6237\u540d\u548c\u5bc6\u7801\u4e0d\u80fd\u4e3a\u7a7a"}),400
    u = get_user(un)
    if not u:
        write_log("用户登录", "登录", obj=un, result="失败")
        return jsonify({"error":"用户名或密码错误"}),401
    stored_pw = u.get("密码哈希", u.get("密码", ""))
    if stored_pw != hash_pw(pw):
        write_log("用户登录", "登录", obj=un, result="失败")
        return jsonify({"error":"用户名或密码错误"}),401
    session["username"]=un; session.permanent=True
    write_log("用户登录", "登录", obj=un, result="成功")
    return jsonify({"username":u["\u7528\u6237\u540d"],"name":u["\u59d3\u540d"],"department":u["\u6240\u5c5e\u90e8\u95e8"],"role":u["\u89d2\u8272"]})

@app.route("/api/logout")
def api_logout():
    un = session.get("username","")
    write_log("用户登录", "登出", obj=un)
    session.clear()
    return jsonify({"message":"已退出"})

@app.route("/api/current_user")
@login_required
def api_cur_user():
    u = get_cur_user()
    if not u: session.clear(); return jsonify({"error":"\u7528\u6237\u4e0d\u5b58\u5728"}),401
    return jsonify({"username":u["\u7528\u6237\u540d"],"name":u["\u59d3\u540d"],"department":u["\u6240\u5c5e\u90e8\u95e8"],"role":u["\u89d2\u8272"]})

# ─── Dashboard & Reminders ───
@app.route("/api/dashboard")
@login_required
def api_dashboard():
    u = get_cur_user(); cts = _scope_contracts(read_contracts()); total = len(cts)
    assets = _scope_assets(read_assets())
    # 到期提醒：只统计条数（不构造明细对象），资产侧复用上面已过滤的 assets，避免重复读表
    exp_project = _project_reminder_count(cts, u)
    exp_asset = _asset_reminder_count(assets, u)
    exp_total = exp_project + exp_asset
    # 总合同额 = 项目合同总额（万元，沿用原口径）+ 资产合同总额（asset_contracts.总合同额 求和，元→万元）
    project_amt = sum((_num(c.get("合同金额（万元）")) or 0) for c in cts)
    asset_amt_yuan = sum((_num(a.get("total_amount")) or 0) for a in assets)
    asset_amt = asset_amt_yuan / 10000.0
    total_amt = project_amt + asset_amt
    unpaid_total = sum(_pending_unpaid(c) for c in cts)
    pending_count = _pending_count(cts)
    asset_count = len(assets)
    # 待收（资产租金）：按资产合同汇总“应缴合计 − 实缴合计”（read_assets 已挂计算列 due_total/paid_total，
    # 底层即 asset_rent_collections 按 资产合同ID 分组的 SUM 口径）。
    # 待收合同数 = 差额 > 0 的合同数（去重计数）；待收总金额 = 全部合同差额之和（含多缴的负数差额），元 → 万元。
    recv_diffs = [round((_num(a.get("due_total")) or 0) - (_num(a.get("paid_total")) or 0), 2) for a in assets]
    recv_count = sum(1 for d in recv_diffs if d > 0)
    recv_amount = round(sum(recv_diffs) / 10000.0, 2)
    return jsonify({"total_contracts":total + asset_count,
                    "project_contract_count":total, "asset_contract_count":asset_count,
                    "project_amount":round(project_amt,2), "asset_amount":round(asset_amt,2),
                    "total_amount":round(total_amt,2),"dept_contracts":total,"expiring_contracts":exp_total,
                    "expiring_project_count":exp_project, "expiring_asset_count":exp_asset,
                    "pending_project_count":pending_count,"unpaid_amount_total":round(unpaid_total,2),
                    # 待收/支：拆分项 + 合并值（卡片显示合并值，悬停气泡显示拆分项）
                    "pending_receivable_count":recv_count, "pending_receivable_amount":recv_amount,
                    "pending_total_count":recv_count + pending_count,
                    "pending_total_amount":round(recv_amount + unpaid_total, 2)})

def _reminders(u,cts=None):
    if cts is None: cts = read_contracts()
    today = datetime.date.today(); dd = get_default_days(); res = []
    for c in cts:
        nb = parse_date(c.get("\u4e0b\u6b21\u62db\u91c7\u65f6\u95f4",""))
        if not nb: continue
        rd = _int(c.get("\u63d0\u9192\u63d0\u524d\u5929\u6570")); 
        if rd is None: rd = dd
        if today < nb - datetime.timedelta(days=rd): continue
        if u["\u89d2\u8272"]!="\u7ba1\u7406\u5458":
            vis = _visible_depts(u)
            if vis is not None and str(c.get("\u6240\u5c5e\u90e8\u95e8","")).strip() not in vis: continue
        item = c2r(c); item["remaining_days"] = (nb-today).days; res.append(item)
    return res

def _asset_reminders(items=None, today=None, dd=None, u=None):
    """资产合同提醒：与项目合同同一口径 ——
    按“下次出租时间 − 提醒提前天数（空则取全局默认）”判定，并按数据范围过滤：
    管理员可见全部；普通/查询用户仅可见“本部门 + 分管部门”（按“管理部门”精确匹配，与项目提醒一致）。
    """
    if items is None: items = read_assets()
    if u is None: u = get_cur_user()
    vis = _visible_depts(u)          # None 表示全部可见（管理员）
    today = today or datetime.date.today()
    dd = dd if dd is not None else get_default_days()
    res = []
    for a in items:
        if vis is not None and str(a.get("dept", "")).strip() not in vis: continue
        nb = parse_date(a.get("next_lease_date", ""))
        if not nb: continue
        rd = _int(a.get("remind_days"))
        if rd is None: rd = dd
        if today < nb - datetime.timedelta(days=rd): continue
        res.append({
            "id": a.get("id", ""), "kind": "asset", "next_label": "下次出租",
            "number": str(a.get("_row", "")), "name": a.get("asset_name", ""),
            "next_bid_date": a.get("next_lease_date", ""), "remaining_days": (nb-today).days,
            "handler": a.get("person", ""), "department": a.get("dept", ""),
        })
    return res

def _all_reminders(u, cts=None):
    """首页到期提醒合并口径：项目合同（下次招采时间）+ 资产合同（下次出租时间），按剩余天数升序（最紧急在前）。"""
    items = _reminders(u, cts) + _asset_reminders(u=u)
    for it in items:
        it.setdefault("kind", "project")
        it.setdefault("next_label", "下次招采")
    items.sort(key=lambda x: (x.get("remaining_days", 0), str(x.get("next_bid_date", "")), str(x.get("name", ""))))
    return items

def _project_reminder_count(cts, u=None):
    """项目合同到期提醒“条数”（与 _reminders 完全同口径，但不构造明细对象，供首页卡片使用）。

    判定：下次招采时间非空 且 今天 ≥ 下次招采时间 − 提醒提前天数（空则取全局默认）；
    数据范围：非管理员仅统计“本部门 + 分管部门”。
    """
    if u is None: u = get_cur_user()
    vis = _visible_depts(u)
    today = datetime.date.today()
    dd = get_default_days()
    n = 0
    for c in cts:
        nb = parse_date(c.get("下次招采时间", ""))
        if not nb: continue
        rd = _int(c.get("提醒提前天数"))
        if rd is None: rd = dd
        if today < nb - datetime.timedelta(days=rd): continue
        if vis is not None and str(c.get("所属部门", "")).strip() not in vis: continue
        n += 1
    return n

def _asset_reminder_count(items, u=None):
    """资产合同到期提醒“条数”（与 _asset_reminders 完全同口径，复用已读取/已过滤的资产列表）。"""
    if u is None: u = get_cur_user()
    vis = _visible_depts(u)
    today = datetime.date.today()
    dd = get_default_days()
    n = 0
    for a in items:
        if vis is not None and str(a.get("dept", "")).strip() not in vis: continue
        nb = parse_date(a.get("next_lease_date", ""))
        if not nb: continue
        rd = _int(a.get("remind_days"))
        if rd is None: rd = dd
        if today < nb - datetime.timedelta(days=rd): continue
        n += 1
    return n

@app.route("/api/reminders")
@login_required
def api_reminders():
    items=_all_reminders(get_cur_user()); return jsonify({"count":len(items),"items":items})
@app.route("/api/pending_payments")
@login_required
def api_pending_payments():
    """待收/支提醒数据：待收（资产合同租金未收齐差额）+ 待支（项目合同待支付，原口径）。

    - 待收：asset_rent_collections 按 资产合同ID 分组计算 SUM(应缴金额)-SUM(实缴金额) > 0 的资产合同
      （金额元 → 万元），口径与首页“待收/支合同数/总金额”卡片、资产统计页一致。
    - 待支：沿用原待支付口径（未归档 且 待支付金额 > 0）。
    - 排序：待收在前、待支在后（与卡片“待收/支”的表述顺序一致）。
    """
    items = []
    for a in _scope_assets(read_assets()):
        diff_yuan = round((_num(a.get("due_total")) or 0) - (_num(a.get("paid_total")) or 0), 2)
        if diff_yuan <= 0:
            continue
        items.append({
            "id": a.get("id"), "kind": "receivable",
            "number": str(a.get("_row", "")), "name": a.get("asset_name", ""),
            "department": a.get("dept", ""), "handler": a.get("person", ""),
            "due_amount": round((_num(a.get("due_total")) or 0) / 10000.0, 2),
            "received_amount": round((_num(a.get("paid_total")) or 0) / 10000.0, 2),
            "next_lease_date": a.get("next_lease_date", ""),
            "pending_amount": round(diff_yuan / 10000.0, 2),
        })
    cts = _scope_contracts(read_contracts())
    for c in cts:
        # 与“待支合同数（个）”口径一致：未归档 且 待支付金额>0
        if c.get("是否归档", "") == "是":
            continue
        amt = _pending_unpaid(c)
        if amt <= 0:
            continue
        it = c2r(c)
        it["kind"] = "payable"
        it["pending_amount"] = amt
        items.append(it)
    return jsonify({"count": len(items),
                    "receivable_count": sum(1 for x in items if x.get("kind") == "receivable"),
                    "payable_count": sum(1 for x in items if x.get("kind") == "payable"),
                    "items": items})

# ═══════════════════ 短信提醒：发送服务 / 定时任务 / 回执接口 ═══════════════════
# 平台：湘潭市政务短信服务平台（前缀 https://sms.xtdzzw.cn/api）
#   POST /message/smsBatch   点对点批量（一号码一内容一 guId）——本系统定时/催办/点对点均按此发送
#   POST /message/smsSingle  单号码单内容（保留可用，本实现统一走 smsBatch，一条也走批量）
# 签名：参数按 ASCII 升序拼 k=v&…，末尾追加 key=SK，MD5 后转大写；sign 与数组字段（sendData）不参与签名。
# 收件人口径：项目=经办人(+现场负责人，仅到期)；资产=承租人(联系电话)+责任人；均按手机号去重、每人一条。
_sms_lock = threading.Lock()
_sms_sched_started = False
SMS_MAX_PHONES = 50          # 单次点对点最多接收号码数
SMS_MAX_CONTENT_LEN = 500    # 单条短信正文最大字数（不含签名）
SMS_RESEND_MAX = 50          # 单次最多重发条数

def _sms_int(v, d):
    try:
        return int(float(v))
    except Exception:
        return d

def _sms_bool(v, d=False):
    if v is None: return d
    return str(v).strip().lower() in ("1", "true", "yes", "on", "是", "开启")

def _sms_cfg():
    def g(k, d=""):
        v = get_cfg(k)
        return d if v is None else v
    return {
        "enabled": _sms_bool(g("sms_enabled", "0")),
        "base_url": str(g("sms_base_url", "https://sms.xtdzzw.cn/api")).strip().rstrip("/"),
        "ak": str(g("sms_access_key", "")).strip(),
        "sk": str(g("sms_secret_key", "")).strip(),
        "weekly_enabled": _sms_bool(g("sms_weekly_enabled", "1"), True),
        "weekday": min(max(_sms_int(g("sms_weekly_weekday", 1), 1), 1), 7),   # ISO：1=周一 … 7=周日
        "hour": min(max(_sms_int(g("sms_weekly_hour", 9), 9), 0), 23),
        "minute": min(max(_sms_int(g("sms_weekly_minute", 0), 0), 0), 59),
        "max_per_run": min(max(_sms_int(g("sms_max_per_run", 200), 200), 1), 1000),
        "retry_max": max(_sms_int(g("sms_retry_max", 1), 1), 0),
        "receipt_token": str(g("sms_receipt_token", "")).strip(),
        "timeout": min(max(_sms_int(g("sms_timeout_sec", 10), 10), 1), 120),
        "include_sign": _sms_bool(g("sms_include_sign", "0")),
        "sign_suffix": str(g("sms_sign_suffix", "（湘潭市机关事务管理局）")).strip(),
        "segment_len": min(max(_sms_int(g("sms_segment_len", 70), 70), 20), 500),
        "max_segments": min(max(_sms_int(g("sms_max_segments", 3), 3), 1), 10),
        "block_check": _sms_bool(g("sms_block_check", "1"), True),
        "block_words": [w.strip() for w in str(g("sms_block_words", "")).replace("、", ",").replace("|", ",").replace("；", ",").split(",") if w.strip()],
        "daily_per_phone": min(max(_sms_int(g("sms_daily_per_phone", 100), 100), 0), 100),
        "min_interval_sec": min(max(_sms_int(g("sms_min_interval_sec", 30), 30), 0), 3600),
    }

def _sms_mobile(v):
    """规范化手机号：仅保留数字，容忍 86 等前缀；返回 11 位手机号或 ''。"""
    import re as _re
    s = str(v or "").strip()
    if not s: return ""
    d = _re.sub(r"\D", "", s)
    if len(d) == 13 and d.startswith("86"): d = d[2:]
    return d if _re.fullmatch(r"1[3-9]\d{9}", d) else ""

def _sms_name_phone_map():
    """姓名 → 手机号（取 users.姓名/电话，仅收录有效手机号；重名取首个）。"""
    m = {}
    for u in get_all_users():
        n = str(u.get("姓名", "") or "").strip()
        p = _sms_mobile(u.get("电话", ""))
        if n and p and n not in m: m[n] = p
    return m

def _sms_sign(params, sk):
    """平台签名：ascii 升序 k=v&… + key=SK，MD5 大写；跳过 sign 与数组/对象字段。"""
    parts = []
    for k in sorted(params.keys()):
        v = params[k]
        if k == "sign" or v is None: continue
        if isinstance(v, (list, tuple, dict)): continue
        parts.append("%s=%s" % (k, v))
    parts.append("key=" + sk)
    return hashlib.md5("&".join(parts).encode("utf-8")).hexdigest().upper()

SMS_NET_TIP = ("请确认本机/服务器在政务网内，并已开通 sms.xtdzzw.cn:443、10.129.17.93:443 网络策略"
               "及调用方 IP 白名单（见对接文档第一章）。")

def _sms_err_text(e, url=""):
    """把底层网络异常翻译成可操作的中文提示（保留原始错误便于排查）。"""
    import socket, ssl, urllib.error
    raw = str(e)
    reason = getattr(e, "reason", None)
    inner = reason if reason is not None else e            # urllib 会把底层异常包进 URLError.reason
    low = (raw + " " + str(inner)).lower()
    winerr = getattr(inner, "winerror", None)
    if winerr is None: winerr = getattr(e, "winerror", None)
    if isinstance(e, urllib.error.HTTPError):
        return "短信平台返回 HTTP %s %s（%s）" % (e.code, e.reason, url)
    # Windows 连接超时为 WinError 10060（WSAETIMEDOUT），文本为中文提示，故需按错误码判断
    if winerr == 10060 or isinstance(inner, (socket.timeout, TimeoutError)) or "timed out" in low:
        return "连接短信平台超时。%s（原始错误：%s）" % (SMS_NET_TIP, raw)
    if winerr == 11001 or isinstance(inner, socket.gaierror) or "getaddrinfo" in low or "name or service not known" in low:
        return "无法解析短信平台域名。请检查 DNS 或网络策略。（原始错误：%s）" % raw
    if winerr == 10061 or "connection refused" in low or isinstance(inner, ConnectionRefusedError):
        return "短信平台拒绝连接（端口未开放或被策略拦截）。%s（原始错误：%s）" % (SMS_NET_TIP, raw)
    if winerr in (10051, 10065) or "unreachable" in low or "no route" in low or isinstance(inner, OSError) and getattr(inner, "errno", None) in (101, 113):
        return "网络不可达。%s（原始错误：%s）" % (SMS_NET_TIP, raw)
    if isinstance(inner, ssl.SSLError) or "ssl" in low or "certificate" in low:
        return "与短信平台 TLS 握手失败（可能被中间设备拦截）。（原始错误：%s）" % raw
    return "发送失败：%s" % raw

def _sms_post(path, body):
    """POST JSON 到平台；返回解析后的响应 dict（解析失败时给出原文）。超时取自 sms_timeout_sec。"""
    cfg = _sms_cfg()
    data = json.dumps(body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(cfg["base_url"] + path, data=data, method="POST",
                                 headers={"Content-Type": "application/json;charset=utf-8"})
    with urllib.request.urlopen(req, timeout=cfg["timeout"]) as resp:
        txt = resp.read().decode("utf-8", "replace")
    try:
        return json.loads(txt)
    except Exception:
        return {"code": "", "success": False, "msg": txt}

def _sms_render(scene, ctx):
    """按场景取启用中的模板并替换占位符；模板缺失时回退内置默认模板。"""
    tpl = ""
    for r in db_query('SELECT "模板内容","启用" FROM sms_templates WHERE "场景"=?', (scene,)):
        if str(r.get("启用") or "是").strip() != "否":
            tpl = r.get("模板内容") or ""
    if not tpl:
        tpl = next((t for s, _n, t in SMS_TEMPLATE_DEFAULTS if s == scene), "{content}")
    out = tpl
    for k, v in (ctx or {}).items():
        out = out.replace("{" + k + "}", "" if v is None else str(v))
    return out

def _sms_sign_content(body):
    """下发内容 = 正文 + 末尾机构署名（如「……请及时办理。（湘潭市机关事务管理局）」）。

    - 署名文本取自 settings.sms_sign_suffix（默认 `（湘潭市机关事务管理局）`），
      可用 sms_include_sign 开关关闭；改用**全角括号**而非「【】」——
      平台注意事项 2：内容含「【】」会被判多签名并直接阻断。
    - 幂等：正文末尾已带该署名时不再重复追加（避免历史内容/重复渲染出现两遍）。
    """
    cfg = _sms_cfg()
    text = str(body)
    suffix = cfg["sign_suffix"] if cfg["include_sign"] else ""
    if not suffix:
        return text
    # 幂等：正文中已出现该署名（可能写成「（机构名）」，也可能在句末句号之前）就不再追加，
    # 避免模板/历史内容里已手工写过署名时出现两遍。
    if suffix in text:
        return text
    core = suffix.strip("（）() \t")
    if core and (("（%s）" % core) in text or ("(%s)" % core) in text):
        return text
    return text + suffix

def _sms_default_receivers(kind, obj, scene):
    """默认收件人 [(类型,姓名,手机号)]：项目=经办人(+现场负责人,仅到期)；资产=承租人+责任人。"""
    names = _sms_name_phone_map()
    out, seen = [], set()
    def add(rtype, name, phone):
        m = _sms_mobile(phone)
        if not m or m in seen: return
        seen.add(m); out.append((rtype, str(name or "").strip(), m))
    if kind == "project":
        for n in _split_names(obj.get("经办人", "")): add("经办人", n, names.get(n, ""))
        if scene == "expire":
            for n in _split_names(obj.get("现场负责人", "")): add("现场负责人", n, names.get(n, ""))
    else:
        add("承租人", obj.get("tenant", ""), obj.get("phone", ""))
        for n in _split_names(obj.get("person", "")): add("责任人", n, names.get(n, ""))
    return out

def _sms_obj_fields(kind, obj):
    """(关联ID, 关联编号, 关联名称)；项目取编号/项目名称，资产取位置(id)/资产名称。"""
    if kind == "project":
        return (obj.get("_row"), obj.get("编号", ""), obj.get("项目名称", ""))
    return (obj.get("_row"), obj.get("place", "") or str(obj.get("_row", "")), obj.get("asset_name", ""))

def _sms_db_write(fn):
    """单连接单事务执行写操作（整批只备份一次），返回 fn(con) 的结果。"""
    with db_lock:
        _backup()
        con = _conn()
        try:
            r = fn(con); con.commit(); _invalidate_caches(); return r
        except Exception:
            con.rollback(); raise
        finally:
            con.close()

# ── 平台注意事项合规校验（短信平台使用规范）──
def _sms_content_issue(text, cfg=None):
    """内容合规检查：返回问题描述，通过返回 ''。
    注意事项 2：内容不得含「【】」（否则被判多签名、直接阻断）；
    注意事项 3：不得含敏感词（省垃圾短信拦截系统会拦截）。"""
    cfg = cfg or _sms_cfg()
    t = str(text or "")
    if "【" in t or "】" in t:
        return "短信内容不能包含“【】”中括号（运营商会判定为多签名并直接阻断发送）"
    if cfg["block_check"]:
        hit = [w for w in cfg["block_words"] if w and w in t]
        if hit:
            return "短信内容含敏感词（可能被垃圾短信拦截系统拦截）：" + "、".join(hit[:5])
    return ""

def _sms_split_content(content, seg_len, max_parts):
    """注意事项 1：单条内容不宜过长，过长按单条计费长度拆分多条发送（尽量在标点处断句）。
    返回段落列表；超过 max_parts 段时返回 None，表示应拒绝发送。"""
    text = str(content or "")
    if seg_len <= 0 or len(text) <= seg_len:
        return [text] if text else None
    parts, cur, marks = [], "", "，。；！？,.;!?、"
    for ch in text:
        cur += ch
        if len(cur) >= seg_len:
            cut = max([cur.rfind(m) for m in marks] or [-1])
            if cut >= max(1, seg_len // 2):
                parts.append(cur[:cut + 1]); cur = cur[cut + 1:]
            else:
                parts.append(cur); cur = ""
    if cur.strip():
        parts.append(cur)
    parts = [p for p in parts if p.strip()]
    if len(parts) > max_parts:
        return None
    return parts or None

def _sms_precheck(msgs, cfg):
    """发送前合规校验（注意事项 4/5/6）：返回 (通过列表, 被拦截列表)。

    - 注意事项 4：号码必须是有效 11 位手机号（固话 / 虚拟号直接拦截）；
    - 注意事项 5：同一号码每日成功发送不超过 sms_daily_per_phone（平台硬性 ≤100）；
    - 注意事项 6：同一号码两次发送的最小间隔 sms_min_interval_sec；
    - 同时在发送前拦截含「【】」与敏感词的内容。
    说明：同一批次内对同一号码的多条消息按“批次开始前”的状态统一判定，
    因此周一自动提醒（同一人可能同时有到期/待收/待支）不会被间隔规则误拦。
    """
    ok, blocked = [], []
    today0 = datetime.date.today().strftime("%Y-%m-%d") + " 00:00:00"
    per_day, gap = cfg["daily_per_phone"], cfg["min_interval_sec"]
    sent_today, last_time = {}, {}
    for m in msgs:
        phone = m.get("phone") or ""
        issue = ""
        if not _sms_mobile(phone):
            issue = "不是有效手机号（仅支持 11 位手机号；固话、虚拟号无法发送）"
        if not issue:
            issue = _sms_content_issue(m.get("content"), cfg)
        if not issue and per_day > 0:
            if phone not in sent_today:
                sent_today[phone] = list(db_query(
                    'SELECT COUNT(*) c FROM sms_logs WHERE "手机号"=? AND "状态"=? AND "时间">=?',
                    (phone, "成功", today0)))[0]["c"]
            if sent_today[phone] + sum(1 for x in ok if x["phone"] == phone) >= per_day:
                issue = "该号码今日发送已达平台上限（%d 条/天）" % per_day
        if not issue and gap > 0:
            if phone not in last_time:
                rows = db_query('SELECT "时间" t FROM sms_logs WHERE "手机号"=? AND "状态"=? ORDER BY id DESC LIMIT 1',
                                (phone, "成功"))
                last_time[phone] = rows[0]["t"] if rows else ""
            lt = last_time[phone]
            if lt:
                try:
                    sec = (datetime.datetime.now() - datetime.datetime.strptime(lt, "%Y-%m-%d %H:%M:%S")).total_seconds()
                    if sec < gap:
                        issue = "距该号码上次发送不足 %d 秒，已按平台要求限流" % gap
                except Exception:
                    pass
        if issue:
            blocked.append({"phone": phone, "name": m.get("rname", ""), "reason": issue})
        else:
            ok.append(m)
    return ok, blocked

def _sms_insert_logs(msgs):
    """批量为待发送消息写流水（含周去重）；返回 (已受理列表, 跳过条数)。"""
    accepted, skipped = [], [0]
    def _fn(con):
        for m in msgs:
            try:
                cur = con.execute(
                    'INSERT INTO sms_logs ("时间","场景","关联类型","关联ID","关联编号","关联名称",'
                    '"接收人类型","接收人","手机号","短信内容","去重键","触发方式","发送人","状态") '
                    'VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                    (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), m["scene"], m["kind"],
                     str(m["obj_id"]), m["obj_no"], m["obj_name"], m["rtype"], m["rname"], m["phone"],
                     m["content"], m["dedup"] or None, m["trigger"], m["operator"], "待发送"))
                m["log_id"] = cur.lastrowid
                accepted.append(m)
            except sqlite3.IntegrityError:
                skipped[0] += 1
        return None
    _sms_db_write(_fn)
    return accepted, skipped[0]

def _sms_update_status(pairs):
    """pairs: [(log_id, 状态, 外部guId, 错误信息, 请求时间戳)]"""
    if not pairs: return
    def _fn(con):
        for log_id, status, guid, err, ts in pairs:
            con.execute('UPDATE sms_logs SET "状态"=?,"外部guId"=?,"错误信息"=?,"请求时间戳"=? WHERE id=?',
                        (status, guid, err, ts, log_id))
        return None
    _sms_db_write(_fn)

def _sms_dispatch(msgs, trigger, operator):
    """合规校验 + 超长拆分 → 写流水（去重）→ 按 max_per_run 分批调 /message/smsBatch → 回写状态。"""
    cfg = _sms_cfg()
    out = {"total": len(msgs), "sent": 0, "skipped": 0, "failed": 0, "error": "", "blocked": []}
    if not cfg["enabled"]:
        out["error"] = "短信功能未启用"; out["total"] = 0; return out
    if not cfg["ak"] or not cfg["sk"]:
        out["error"] = "未配置 AccessKey/SecretKey"; out["total"] = 0; return out
    valid, blocked = _sms_precheck(msgs, cfg)
    out["blocked"] = list(blocked)
    expanded = []
    for m in valid:
        parts = _sms_split_content(m.get("content"), cfg["segment_len"], cfg["max_segments"])
        if parts is None:
            out["blocked"].append({"phone": m.get("phone", ""), "name": m.get("rname", ""),
                                   "reason": "内容过长（超过 %d 段），请精简内容" % cfg["max_segments"]})
            continue
        for i, part in enumerate(parts):
            x = dict(m); x["content"] = part
            if len(parts) > 1 and x.get("dedup"):
                x["dedup"] = "%s#%d" % (x["dedup"], i + 1)   # 拆分后每段单独去重
            expanded.append(x)
    if not expanded:
        out["error"] = (out["blocked"][0]["reason"] if out["blocked"] else "没有可发送的内容")
        out["total"] = len(msgs)
        return out
    accepted, skipped = _sms_insert_logs(expanded)
    out["skipped"] = skipped
    if not accepted: return out
    chunk = cfg["max_per_run"]
    for i in range(0, len(accepted), chunk):
        part = accepted[i:i + chunk]
        ts = int(time.time() * 1000)
        send_data = []
        for m in part:
            m["guid"] = "SMS%d" % m["log_id"]
            send_data.append({"desAddr": m["phone"], "msgContent": m["content"], "guId": m["guid"]})
        body = {"accessKey": cfg["ak"], "timestamp": ts, "sendData": send_data}
        body["sign"] = _sms_sign(body, cfg["sk"])
        try:
            resp = _sms_post("/message/smsBatch", body)
            ok = str(resp.get("code")) == "200" and bool(resp.get("success"))
            err = "" if ok else str(resp.get("msg") or resp.get("code") or "发送失败")
        except Exception as e:
            ok, err = False, _sms_err_text(e, cfg["base_url"] + "/message/smsBatch")
        _sms_update_status([(m["log_id"], "成功" if ok else "失败", m["guid"] if ok else "", err, ts) for m in part])
        if ok: out["sent"] += len(part)
        else: out["failed"] += len(part); out["error"] = err
    return out

def _sms_due_tasks():
    """本周应发提醒任务（系统级，不按登录人过滤）：到期(expire)/待支(payable)/待收(receivable)。"""
    today = datetime.date.today(); dd = get_default_days(); tasks = []
    for c in read_contracts():
        if str(c.get("短信提醒", "")).strip() == "关闭": continue
        if c.get("是否归档", "") == "是": continue
        nb = parse_date(c.get("下次招采时间", ""))
        if nb:
            rd = _int(c.get("提醒提前天数")); rd = dd if rd is None else rd
            if today >= nb - datetime.timedelta(days=rd):
                tasks.append({"scene": "expire", "kind": "project", "obj": c,
                              "ctx": {"name": c.get("项目名称", ""), "number": c.get("编号", ""),
                                      "due": c.get("下次招采时间", ""), "days": (nb - today).days}})
        amt = _pending_unpaid(c)
        if amt and amt > 0:
            tasks.append({"scene": "payable", "kind": "project", "obj": c,
                          "ctx": {"name": c.get("项目名称", ""), "number": c.get("编号", ""), "amount": amt}})
    for a in read_assets():
        if str(a.get("短信提醒", "")).strip() == "关闭": continue
        if a.get("archived"): continue
        nb = parse_date(a.get("next_lease_date", ""))
        if nb:
            rd = _int(a.get("remind_days")); rd = dd if rd is None else rd
            if today >= nb - datetime.timedelta(days=rd):
                tasks.append({"scene": "expire", "kind": "asset", "obj": a,
                              "ctx": {"name": a.get("asset_name", ""), "number": a.get("place", "") or str(a.get("_row", "")),
                                      "due": a.get("next_lease_date", ""), "days": (nb - today).days}})
        diff = round((_num(a.get("due_total")) or 0) - (_num(a.get("paid_total")) or 0), 2)
        if diff > 0:
            tasks.append({"scene": "receivable", "kind": "asset", "obj": a,
                          "ctx": {"name": a.get("asset_name", ""), "tenant": a.get("tenant", ""), "amount": diff}})
    return tasks

def _sms_tasks_to_msgs(tasks, trigger, operator, week_tag=None):
    msgs = []
    for t in tasks:
        content = _sms_sign_content(_sms_render(t["scene"], t["ctx"]))
        obj_id, obj_no, obj_name = _sms_obj_fields(t["kind"], t["obj"])
        for rtype, rname, phone in _sms_default_receivers(t["kind"], t["obj"], t["scene"]):
            dedup = ""
            if week_tag: dedup = "%s|%s|%s|%s|%s" % (t["scene"], t["kind"], obj_id, phone, week_tag)
            msgs.append({"scene": t["scene"], "kind": t["kind"], "obj_id": obj_id, "obj_no": obj_no,
                         "obj_name": obj_name, "rtype": rtype, "rname": rname, "phone": phone,
                         "content": content, "dedup": dedup, "trigger": trigger, "operator": operator})
    return msgs

def sms_weekly_run(operator="系统", force=False):
    """执行一次“每周提醒”：构建到期/待收/待支短信并按周去重发送。"""
    cfg = _sms_cfg()
    if not cfg["enabled"]:
        return {"ok": False, "error": "短信功能未启用", "total": 0, "sent": 0, "skipped": 0, "failed": 0}
    if not cfg["weekly_enabled"] and not force:
        return {"ok": True, "error": "", "total": 0, "sent": 0, "skipped": 0, "failed": 0, "note": "每周提醒未启用"}
    iso = datetime.date.today().isocalendar()
    week_tag = "%04d-W%02d" % (iso[0], iso[1])
    r = _sms_dispatch(_sms_tasks_to_msgs(_sms_due_tasks(), "定时", operator, week_tag), "定时", operator)
    r["ok"] = (r.get("failed", 0) == 0) and not r.get("error")
    r["week"] = week_tag
    return r

def sms_send_urge(kind, obj, note, operator):
    """催办：按默认收件人发送 urge 模板短信。"""
    obj_id, obj_no, obj_name = _sms_obj_fields(kind, obj)
    content = _sms_sign_content(_sms_render("urge", {"name": obj_name, "number": obj_no, "note": note}))
    msgs = []
    for rtype, rname, phone in _sms_default_receivers(kind, obj, "urge"):
        msgs.append({"scene": "urge", "kind": kind, "obj_id": obj_id, "obj_no": obj_no, "obj_name": obj_name,
                     "rtype": rtype, "rname": rname, "phone": phone, "content": content,
                     "dedup": "", "trigger": "催办", "operator": operator})
    r = _sms_dispatch(msgs, "催办", operator)
    r["ok"] = (r.get("failed", 0) == 0)
    return r

def sms_send_point(kind, obj, phones, content_body, operator):
    """点对点：向指定手机号发送自由内容（自动加签名）。"""
    obj_id, obj_no, obj_name = _sms_obj_fields(kind, obj)
    content = _sms_sign_content(content_body)
    msgs = []
    for p in (phones or []):
        m = _sms_mobile(p)
        if not m: continue
        msgs.append({"scene": "manual", "kind": kind, "obj_id": obj_id, "obj_no": obj_no, "obj_name": obj_name,
                     "rtype": "手动", "rname": "", "phone": m, "content": content,
                     "dedup": "", "trigger": "点对点", "operator": operator})
    if not msgs:
        return {"ok": False, "error": "未提供有效手机号", "total": 0, "sent": 0, "skipped": 0, "failed": 0}
    r = _sms_dispatch(msgs, "点对点", operator)
    r["ok"] = (r.get("failed", 0) == 0)
    return r

def sms_send_free(phones, content_body, operator):
    """独立的点对点发送（无合同上下文，供“短信发送记录”页使用）；关联类型记为“手动”。"""
    content = _sms_sign_content(content_body)
    msgs = []
    for p in (phones or []):
        m = _sms_mobile(p)
        if not m: continue
        msgs.append({"scene": "manual", "kind": "手动", "obj_id": "", "obj_no": "", "obj_name": "",
                     "rtype": "手动", "rname": "", "phone": m, "content": content,
                     "dedup": "", "trigger": "点对点", "operator": operator})
    if not msgs:
        return {"ok": False, "error": "未提供有效手机号", "total": 0, "sent": 0, "skipped": 0, "failed": 0}
    r = _sms_dispatch(msgs, "点对点", operator)
    r["ok"] = (r.get("failed", 0) == 0)
    return r

def _sms_conn_test():
    """平台连通性探测（DNS + TCP，不发送短信）：对应对接文档中的 telnet 测试。"""
    import socket
    from urllib.parse import urlparse
    cfg = _sms_cfg()
    u = urlparse(cfg["base_url"])
    host = u.hostname or "sms.xtdzzw.cn"
    port = u.port or 443
    timeout = cfg["timeout"]
    steps, ok_all = [], True
    t0 = time.time()
    try:
        infos = socket.getaddrinfo(host, port, proto=socket.IPPROTO_TCP)
        ips = sorted(set(i[4][0] for i in infos))
        steps.append({"step": "解析域名 %s" % host, "ok": True, "detail": "、".join(ips),
                      "ms": int((time.time() - t0) * 1000)})
    except Exception as e:
        ok_all = False
        steps.append({"step": "解析域名 %s" % host, "ok": False, "detail": _sms_err_text(e),
                      "ms": int((time.time() - t0) * 1000)})
    for h in (host, "10.129.17.93"):
        t0 = time.time()
        try:
            s = socket.create_connection((h, port), timeout=timeout); s.close()
            ok, detail = True, "已连通"
        except Exception as e:
            ok, detail = False, _sms_err_text(e)
        if h == host: ok_all = ok_all and ok
        steps.append({"step": "TCP 连接 %s:%d" % (h, port), "ok": ok, "detail": detail,
                      "ms": int((time.time() - t0) * 1000)})
    return {"ok": ok_all, "host": host, "port": port, "timeout": timeout, "base_url": cfg["base_url"],
            "enabled": cfg["enabled"], "steps": steps,
            "message": "连通正常，可以进行短信发送。" if ok_all else ("连接失败。" + SMS_NET_TIP)}

def _sms_cb_token_ok():
    """回调鉴权：配置了 sms_receipt_token 时要求 ?token= 或 X-SMS-Token 匹配。"""
    tk = _sms_cfg()["receipt_token"]
    if not tk: return True
    got = request.args.get("token", "") or request.headers.get("X-SMS-Token", "")
    return str(got).strip() == tk

def _sms_find_project(cid):
    if not cid or not str(cid).startswith("row_"): return None
    try: rn = int(str(cid).split("_")[1])
    except Exception: return None
    return next((c for c in read_contracts() if c["_row"] == rn), None)

def _sms_find_asset(aid):
    rid = _asset_id(aid)
    if rid is None: return None
    return next((a for a in read_assets() if a["_row"] == rid), None)

def _sms_send_guard(u, obj, archived=False):
    """发送权限：查询用户禁止；其余沿用写权限（可见部门）。"""
    if u and u.get("角色") == "查询用户":
        return jsonify({"error": "权限不足"}), 403
    if archived:
        return jsonify({"error": "该合同已归档，仅可浏览"}), 403
    if not _can_write(obj):
        return jsonify({"error": "权限不足"}), 403
    return None

def _sms_phones_param(d):
    phones = d.get("phones")
    if isinstance(phones, str):
        phones = [x for x in phones.replace("，", ",").replace(";", ",").split(",") if x.strip()]
    return phones or []

@app.route("/api/sms/test_connection", methods=["POST"])
@admin_required
def api_sms_test_connection():
    """管理员触发的平台连通性测试（DNS + TCP），不发送短信。"""
    r = _sms_conn_test()
    write_log("系统配置", "短信连通性测试", obj="sms", detail=json.dumps(r, ensure_ascii=False),
              result="成功" if r.get("ok") else "失败")
    return jsonify(r)

@app.route("/api/sms/status")
@login_required
def api_sms_status():
    cfg = _sms_cfg()
    return jsonify({"enabled": cfg["enabled"], "weekly_enabled": cfg["weekly_enabled"],
                    "weekday": cfg["weekday"], "hour": cfg["hour"], "minute": cfg["minute"],
                    "configured": bool(cfg["ak"] and cfg["sk"]),
                    "last_run_at": get_cfg("sms_last_run_at") or ""})

@app.route("/api/sms/run_weekly", methods=["POST"])
@admin_required
def api_sms_run_weekly():
    u = get_cur_user()
    r = sms_weekly_run(operator=(u.get("用户名") if u else "系统"), force=True)
    write_log("短信提醒", "定时发送(手动触发)", obj="weekly", detail=json.dumps(r, ensure_ascii=False))
    return jsonify(r)

@app.route("/api/sms_recipients")
@login_required
def api_sms_recipients():
    """预览某合同在某场景下的默认收件人（前端催办前展示）。"""
    kind = (request.args.get("kind", "") or "").strip()
    scene = (request.args.get("scene", "urge") or "urge").strip()
    obj = _sms_find_project(request.args.get("id", "")) if kind == "project" else \
          (_sms_find_asset(request.args.get("id", "")) if kind == "asset" else None)
    if obj is None: return jsonify({"error": "合同不存在"}), 404
    if not _can_write(obj): return jsonify({"error": "权限不足"}), 403
    recv = [{"type": t, "name": n, "phone": p} for t, n, p in _sms_default_receivers(kind, obj, scene)]
    return jsonify({"recipients": recv, "count": len(recv)})

@app.route("/api/contracts/<cid>/sms_urge", methods=["POST"])
@login_required
def api_project_sms_urge(cid):
    u = get_cur_user(); t = _sms_find_project(cid)
    if not t: return jsonify({"error": "合同不存在"}), 404
    g = _sms_send_guard(u, t, archived=(t.get("是否归档", "") == "是"))
    if g: return g
    note = str((request.get_json(silent=True) or {}).get("note", "") or "").strip() or "请尽快办理。"
    r = sms_send_urge("project", t, note, u.get("用户名", ""))
    write_log("合同台账", "短信催办", obj=str(t.get("编号", "")), detail=json.dumps(r, ensure_ascii=False),
              result="成功" if r.get("ok") else "失败")
    return jsonify(r), (200 if (r.get("ok") or r.get("sent")) else 400)

@app.route("/api/contracts/<cid>/sms", methods=["POST"])
@login_required
def api_project_sms_point(cid):
    u = get_cur_user(); t = _sms_find_project(cid)
    if not t: return jsonify({"error": "合同不存在"}), 404
    g = _sms_send_guard(u, t, archived=(t.get("是否归档", "") == "是"))
    if g: return g
    d = request.get_json(silent=True) or {}
    content = str(d.get("content", "") or "").strip()
    if not content: return jsonify({"error": "短信内容不能为空"}), 400
    if len(content) > SMS_MAX_CONTENT_LEN: return jsonify({"error": "短信内容不能超过 %d 字" % SMS_MAX_CONTENT_LEN}), 400
    issue = _sms_content_issue(content)
    if issue: return jsonify({"error": issue}), 400
    r = sms_send_point("project", t, _sms_phones_param(d), content, u.get("用户名", ""))
    write_log("合同台账", "点对点短信", obj=str(t.get("编号", "")), detail=json.dumps(r, ensure_ascii=False),
              result="成功" if r.get("ok") else "失败")
    return jsonify(r), (200 if (r.get("ok") or r.get("sent")) else 400)

@app.route("/api/asset_contracts/<aid>/sms_urge", methods=["POST"])
@login_required
def api_asset_sms_urge(aid):
    u = get_cur_user(); t = _sms_find_asset(aid)
    if not t: return jsonify({"error": "资产合同不存在"}), 404
    g = _sms_send_guard(u, t, archived=bool(t.get("archived")))
    if g: return g
    note = str((request.get_json(silent=True) or {}).get("note", "") or "").strip() or "请尽快办理。"
    r = sms_send_urge("asset", t, note, u.get("用户名", ""))
    write_log("资产合同台账", "短信催办", obj=str(t.get("asset_name", "")), detail=json.dumps(r, ensure_ascii=False),
              result="成功" if r.get("ok") else "失败")
    return jsonify(r), (200 if (r.get("ok") or r.get("sent")) else 400)

@app.route("/api/asset_contracts/<aid>/sms", methods=["POST"])
@login_required
def api_asset_sms_point(aid):
    u = get_cur_user(); t = _sms_find_asset(aid)
    if not t: return jsonify({"error": "资产合同不存在"}), 404
    g = _sms_send_guard(u, t, archived=bool(t.get("archived")))
    if g: return g
    d = request.get_json(silent=True) or {}
    content = str(d.get("content", "") or "").strip()
    if not content: return jsonify({"error": "短信内容不能为空"}), 400
    if len(content) > SMS_MAX_CONTENT_LEN: return jsonify({"error": "短信内容不能超过 %d 字" % SMS_MAX_CONTENT_LEN}), 400
    issue = _sms_content_issue(content)
    if issue: return jsonify({"error": issue}), 400
    r = sms_send_point("asset", t, _sms_phones_param(d), content, u.get("用户名", ""))
    write_log("资产合同台账", "点对点短信", obj=str(t.get("asset_name", "")), detail=json.dumps(r, ensure_ascii=False),
              result="成功" if r.get("ok") else "失败")
    return jsonify(r), (200 if (r.get("ok") or r.get("sent")) else 400)

@app.route("/api/contracts/<cid>/sms_switch", methods=["PUT"])
@login_required
def api_project_sms_switch(cid):
    """单合同短信提醒开关（v23）：开启/关闭，空值视为开启。"""
    u = get_cur_user(); t = _sms_find_project(cid)
    if not t: return jsonify({"error": "合同不存在"}), 404
    g = _sms_send_guard(u, t, archived=(t.get("是否归档", "") == "是"))
    if g: return g
    on = bool((request.get_json(silent=True) or {}).get("enabled", True))
    val = "开启" if on else "关闭"
    db_exec('UPDATE project_contracts SET "短信提醒"=? WHERE id=?', (val, t["_row"]))
    log_field_changes("合同台账", t.get("编号", ""), t.get("项目名称", ""), t, {"短信提醒": val})
    write_log("合同台账", "短信提醒开关", obj=str(t.get("编号", "")), detail=val)
    return jsonify({"message": "已" + val, "sms_switch": val})

@app.route("/api/asset_contracts/<aid>/sms_switch", methods=["PUT"])
@login_required
def api_asset_sms_switch(aid):
    u = get_cur_user(); t = _sms_find_asset(aid)
    if not t: return jsonify({"error": "资产合同不存在"}), 404
    g = _sms_send_guard(u, t, archived=bool(t.get("archived")))
    if g: return g
    on = bool((request.get_json(silent=True) or {}).get("enabled", True))
    val = "开启" if on else "关闭"
    db_exec('UPDATE asset_contracts SET "短信提醒"=? WHERE id=?', (val, t["_row"]))
    log_field_changes("资产合同台账", t["_row"], t.get("asset_name", ""), t, {"短信提醒": val})
    write_log("资产合同台账", "短信提醒开关", obj=str(t.get("asset_name", "")), detail=val)
    return jsonify({"message": "已" + val, "sms_switch": val})

@app.route("/api/sms/send", methods=["POST"])
@login_required
def api_sms_send():
    """独立点对点发送（无合同上下文）：自填手机号与内容，自动加签名并留痕。"""
    u = get_cur_user()
    if u and u.get("角色") == "查询用户":
        return jsonify({"error": "权限不足"}), 403
    d = request.get_json(silent=True) or {}
    content = str(d.get("content", "") or "").strip()
    if not content:
        return jsonify({"error": "短信内容不能为空"}), 400
    if len(content) > SMS_MAX_CONTENT_LEN:
        return jsonify({"error": "短信内容不能超过 %d 字" % SMS_MAX_CONTENT_LEN}), 400
    issue = _sms_content_issue(content)
    if issue:
        return jsonify({"error": issue}), 400
    uniq, seen = [], set()
    for p in _sms_phones_param(d):
        m = _sms_mobile(p)
        if m and m not in seen:
            seen.add(m); uniq.append(m)
    if not uniq:
        return jsonify({"error": "未提供有效手机号"}), 400
    if len(uniq) > SMS_MAX_PHONES:
        return jsonify({"error": "单次最多发送 %d 个手机号" % SMS_MAX_PHONES}), 400
    r = sms_send_free(uniq, content, (u.get("用户名") if u else ""))
    r["phones"] = uniq
    write_log("短信提醒", "点对点发送", obj="手动", detail=json.dumps(r, ensure_ascii=False),
              result="成功" if r.get("ok") else "失败")
    return jsonify(r), (200 if (r.get("ok") or r.get("sent")) else 400)

@app.route("/api/sms/resend", methods=["POST"])
@login_required
def api_sms_resend():
    """失败记录重发：按 sms_logs.id 重发（原记录保留，另写一条 trigger=重发 的新流水）。"""
    u = get_cur_user()
    if u and u.get("角色") == "查询用户":
        return jsonify({"error": "权限不足"}), 403
    d = request.get_json(silent=True) or {}
    raw = d.get("ids")
    if isinstance(raw, (int, str)): raw = [raw]
    ids = []
    for x in (raw or []):
        s = str(x).strip()
        if s.isdigit(): ids.append(int(s))
    if not ids:
        return jsonify({"error": "未选择记录"}), 400
    if len(ids) > SMS_RESEND_MAX:
        return jsonify({"error": "单次最多重发 %d 条" % SMS_RESEND_MAX}), 400
    uname = u.get("用户名", "") if u else ""
    is_admin = bool(u and u.get("角色") == "管理员")
    msgs, picked = [], 0
    for rid in ids:
        r = next(iter(db_query("SELECT * FROM sms_logs WHERE id=?", (rid,))), None)
        if not r: continue
        if str(r.get("状态") or "") != "失败": continue          # 仅失败记录可重发
        if not is_admin and str(r.get("发送人") or "") != uname: continue
        picked += 1
        phone = _sms_mobile(r.get("手机号"))
        if not phone: continue
        # 历史记录内容可能带「【】」（旧版签名格式），平台会判定多签名阻断，
        # 重发时自动清理中括号（签名由平台侧按应用签名自动添加）。
        content = str(r.get("短信内容") or "").replace("【", "").replace("】", "")
        msgs.append({"scene": r.get("场景") or "manual", "kind": r.get("关联类型") or "手动",
                     "obj_id": r.get("关联ID") or "", "obj_no": r.get("关联编号") or "",
                     "obj_name": r.get("关联名称") or "", "rtype": r.get("接收人类型") or "手动",
                     "rname": r.get("接收人") or "", "phone": phone,
                     "content": content, "dedup": "",
                     "trigger": "重发", "operator": uname})
    if not picked:
        return jsonify({"error": "没有可重发的失败记录（仅“失败”且本人触发的记录可重发）"}), 400
    if not msgs:
        return jsonify({"error": "所选记录的手机号无效，无法重发"}), 400
    res = _sms_dispatch(msgs, "重发", uname)
    res["ok"] = (res.get("failed", 0) == 0) and not res.get("error")
    write_log("短信提醒", "失败重发", obj=",".join(str(i) for i in ids), detail=json.dumps(res, ensure_ascii=False),
              result="成功" if res.get("ok") else "失败")
    return jsonify(res), (200 if (res.get("ok") or res.get("sent")) else 400)

@app.route("/api/sms_logs")
@login_required
def api_sms_logs():
    """短信发送记录：管理员看全部，其他用户仅看本人触发的记录。"""
    u = get_cur_user()
    a = request.args
    conds, params = [], []
    if not u or u.get("角色") != "管理员":
        conds.append('"发送人" = ?'); params.append(u.get("用户名", "") if u else "")
    for key, col in (("scene", "场景"), ("status", "状态"), ("kind", "关联类型"), ("phone", "手机号")):
        v = (a.get(key, "") or "").strip()
        if v: conds.append('"%s" = ?' % col); params.append(v)
    if (a.get("start_date", "") or "").strip():
        conds.append('"时间" >= ?'); params.append(a.get("start_date").strip() + " 00:00:00")
    if (a.get("end_date", "") or "").strip():
        conds.append('"时间" <= ?'); params.append(a.get("end_date").strip() + " 23:59:59")
    q = (a.get("q", "") or "").strip()
    if q:
        conds.append('("关联名称" LIKE ? OR "关联编号" LIKE ? OR "接收人" LIKE ? OR "短信内容" LIKE ?)')
        params += ["%" + q + "%"] * 4
    where = (" WHERE " + " AND ".join(conds)) if conds else ""
    # 分页口径与“操作日志（审计）”/api/logs 保持一致：page / page_size，
    # 返回 total / page / page_size / items；仍兼容旧参数 limit（等价于 page_size）。
    try:
        page = max(1, int(a.get("page", 1)))
        ps = min(200, max(1, int(a.get("page_size") or a.get("limit") or 20)))
    except Exception:
        page, ps = 1, 20
    total = db_query("SELECT COUNT(*) c FROM sms_logs" + where, tuple(params))[0]["c"]
    rows = db_query("SELECT * FROM sms_logs" + where + " ORDER BY id DESC LIMIT ? OFFSET ?",
                    tuple(params) + (ps, (page - 1) * ps))
    return jsonify({"total": total, "page": page, "page_size": ps, "items": rows})

@app.route("/api/sms_templates", methods=["GET"])
@login_required
def api_sms_templates():
    return jsonify({"items": db_query("SELECT * FROM sms_templates ORDER BY rowid")})

@app.route("/api/sms_templates/<scene>", methods=["PUT"])
@admin_required
def api_sms_template_save(scene):
    scene = str(scene or "").strip()
    if not scene: return jsonify({"error": "场景不能为空"}), 400
    d = request.get_json(silent=True) or {}
    content = str(d.get("content", "") or "")
    if not content.strip(): return jsonify({"error": "模板内容不能为空"}), 400
    if "【" in content or "】" in content:
        return jsonify({"error": "模板内容不能包含“【】”中括号：运营商会判定为多签名并直接阻断发送（签名由平台侧自动添加）"}), 400
    u = get_cur_user()
    db_exec('UPDATE sms_templates SET "模板内容"=?,"启用"=?,"更新时间"=?,"更新人"=? WHERE "场景"=?',
            (content, str(d.get("enabled", "是") or "是"),
             datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), (u.get("用户名", "") if u else ""), scene))
    write_log("系统配置", "修改短信模板", obj=scene)
    return jsonify({"message": "模板已保存"})

# ─── 回执 / 回复回调（由平台调用，免登录；配置 sms_receipt_token 时校验口令）───
@app.route("/api/sms/callback/receipt", methods=["POST"])
def api_sms_cb_receipt():
    if not _sms_cb_token_ok(): return jsonify({"code": "403", "success": False, "msg": "token invalid"}), 403
    d = request.get_json(silent=True) or {}
    guid = str(d.get("guId", "") or "").strip()
    if guid:
        db_exec('UPDATE sms_logs SET "回执状态码"=?,"回执描述"=? WHERE "外部guId"=?',
                (str(d.get("receiptCode", "") or ""), str(d.get("receiptDesc", "") or ""), guid))
    return jsonify({"code": "200", "success": True, "msg": "OK"})

@app.route("/api/sms/callback/reply", methods=["POST"])
def api_sms_cb_reply():
    if not _sms_cb_token_ok(): return jsonify({"code": "403", "success": False, "msg": "token invalid"}), 403
    d = request.get_json(silent=True) or {}
    guid = str(d.get("guId", "") or "").strip()
    if guid:
        db_exec('UPDATE sms_logs SET "回复时间"=?,"回复内容"=? WHERE "外部guId"=?',
                (str(d.get("replyTime", "") or ""), str(d.get("content", "") or ""), guid))
    return jsonify({"code": "200", "success": True, "msg": "OK"})

# ─── 每周定时任务（后台线程；命中配置的周几/时间点后当天只跑一次）───
def _sms_scheduler_loop():
    while True:
        try:
            cfg = _sms_cfg()
            if cfg["enabled"] and cfg["weekly_enabled"]:
                now = datetime.datetime.now()
                due = (now.isoweekday() == cfg["weekday"] and
                       (now.hour > cfg["hour"] or (now.hour == cfg["hour"] and now.minute >= cfg["minute"])))
                today = now.strftime("%Y-%m-%d")
                if due and str(get_cfg("sms_last_run_at") or "") != today:
                    with _sms_lock:
                        if str(get_cfg("sms_last_run_at") or "") != today:
                            set_cfg("sms_last_run_at", today)   # 先占位，避免重复触发
                            r = sms_weekly_run(operator="系统")
                            _log_exec('INSERT INTO operation_logs ("时间","用户名","姓名","角色","IP","模块","操作类型","对象","详情","结果") '
                                      'VALUES (?,?,?,?,?,?,?,?,?,?)',
                                      (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "系统", "", "", "",
                                       "短信提醒", "定时发送", today, json.dumps(r, ensure_ascii=False),
                                       "成功" if r.get("failed", 0) == 0 else "失败"))
        except Exception as e:
            try: print("[WARN] 短信定时任务异常：%s" % e, file=sys.stderr)
            except Exception: pass
        time.sleep(30)

def _start_sms_scheduler():
    global _sms_sched_started
    if _sms_sched_started: return
    _sms_sched_started = True
    threading.Thread(target=_sms_scheduler_loop, name="sms-scheduler", daemon=True).start()

# ─── Contracts CRUD ───
@app.route("/api/contracts",methods=["GET"])
@login_required
def api_list_contracts():
    cts = _scope_contracts(read_contracts())
    params = {k:request.args.get(k,"").strip() for k in ["unit","dept","person","name","no","category","method","amount_min","amount_max","start_date","end_date","remind_start","remind_end","is_settled","vendor","archived","fund_source"]}
    flt = []
    for c in cts:
        if params["unit"] and params["unit"] not in c.get("\u6240\u5c5e\u90e8\u95e8",""): continue
        if params["dept"] and params["dept"] not in c.get("\u6240\u5c5e\u90e8\u95e8",""): continue
        if params["person"] and params["person"] not in c.get("\u7ecf\u529e\u4eba",""): continue
        if params["name"] and params["name"] not in c.get("\u9879\u76ee\u540d\u79f0",""): continue
        if params["no"] and params["no"] not in c.get("\u7f16\u53f7",""): continue
        if params["category"] and c.get("\u9879\u76ee\u5206\u7c7b","")!=params["category"]: continue
        if params["method"] and c.get("\u91c7\u8d2d\u65b9\u5f0f","")!=params["method"]: continue
        if params["fund_source"] and c.get("\u8d44\u91d1\u6765\u6e90","")!=params["fund_source"]: continue
        amt = _num(c.get("\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09"))
        if params["amount_min"]:
            try:
                if amt is None or amt<float(params["amount_min"]): continue
            except: pass
        if params["amount_max"]:
            try:
                if amt is None or amt>float(params["amount_max"]): continue
            except: pass
        cd = parse_date(c.get("\u91c7\u8d2d\u65f6\u95f4",""))
        if params["start_date"]:
            sd=parse_date(params["start_date"])
            if sd and (cd is None or cd<sd): continue
        if params["end_date"]:
            ed=parse_date(params["end_date"])
            if ed and (cd is None or cd>ed): continue
        nd = parse_date(c.get("\u4e0b\u6b21\u62db\u91c7\u65f6\u95f4",""))
        if params["remind_start"]:
            rs=parse_date(params["remind_start"])
            if rs and (nd is None or nd<rs): continue
        if params["remind_end"]:
            re_=parse_date(params["remind_end"])
            if re_ and (nd is None or nd>re_): continue
        if params["is_settled"] in ("\u662f","\u5426") and c.get("\u662f\u5426\u7ed3\u7b97","")!=params["is_settled"]: continue
        if params["vendor"]:
            if params["vendor"] not in c.get("\u4e2d\u6807\u5355\u4f4d","") and params["vendor"] not in c.get("\u5b9e\u9645\u5b9e\u65bd\u5355\u4f4d",""): continue
        if params["archived"] in ("是","否") and (c.get("是否归档","") or "否") != params["archived"]: continue
        flt.append(c)
    # 列表默认按“编号”降序排列（数字编号按数值降序，非数字编号排后）
    def _no_key(c):
        v = str(c.get("编号", "") or "").strip()
        try:
            return (0, -float(v))
        except Exception:
            return (1, v)
    flt.sort(key=_no_key)
    try: p=int(request.args.get("page",1)); ps=min(int(request.args.get("page_size",20)),200)
    except: p=1; ps=20
    si=(p-1)*ps; ei=si+ps
    att_map = attach_map_by_number()
    items=[]
    for x in flt[si:ei]:
        it = c2r(x)
        it["attachments"] = att_map.get(str(x.get("\u7f16\u53f7","")), [])
        items.append(it)
    return jsonify({"total":len(flt),"page":p,"page_size":ps,"items":items})

@app.route("/api/contracts",methods=["POST"])
@login_required
def api_create_contract():
    u = get_cur_user()
    if u["\u89d2\u8272"]=="\u67e5\u8be2\u7528\u6237": return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
    d = request.get_json(silent=True) or {}
    # 编号：未提供时按《合同编号规则》自动生成（未配置规则的科室/组室不生成、不猜测）
    if not str(d.get("number", "") or "").strip():
        _gen, _err = next_contract_number("project_contracts", "编号", _number_rule_dept(u))
        if _err: return jsonify({"error": "无法自动生成合同编号：" + _err}), 400
        d["number"] = _gen
    else:
        d["number"] = str(d["number"]).strip()
    for f in ["number","category","name","procurement_method","amount","main_content","handler","fund_source"]:
        if not d.get(f,""): return jsonify({"error":f"\u5b57\u6bb5 '{f}' \u4e0d\u80fd\u4e3a\u7a7a"}),400
    for c in read_contracts():
        if str(c.get("\u7f16\u53f7",""))==str(d["number"]): return jsonify({"error":f"\u7f16\u53f7 '{d['number']}' \u5df2\u5b58\u5728"}),400
    rd = {cn:"" for cn in CONTRACT_COLUMNS}
    fm = {"number":"\u7f16\u53f7","category":"\u9879\u76ee\u5206\u7c7b","name":"\u9879\u76ee\u540d\u79f0","decision_basis":"\u51b3\u7b56\u4f9d\u636e","bid_control_price":"\u51b3\u7b56\u7684\u62db\u6807\u63a7\u5236\u4ef7","price_basis":"\u62db\u6807\u63a7\u5236\u4ef7\u51b3\u7b56\u7684\u4f9d\u636e","procurement_method":"\u91c7\u8d2d\u65b9\u5f0f","winner":"\u4e2d\u6807\u5355\u4f4d","actual_implementor":"\u5b9e\u9645\u5b9e\u65bd\u5355\u4f4d","amount":"\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09","fund_source":"\u8d44\u91d1\u6765\u6e90","main_content":"\u5408\u540c\u4e3b\u8981\u5185\u5bb9","procurement_date":"\u91c7\u8d2d\u65f6\u95f4","start_date":"\u5f00\u5de5\u65f6\u95f4","end_date":"\u5b8c\u5de5\u65f6\u95f4","acceptance_date":"\u7ae3\u5de5\u9a8c\u6536\u65f6\u95f4","site_manager":"\u73b0\u573a\u8d1f\u8d23\u4eba","has_change":"\u662f\u5426\u6709\u7b7e\u8bc1\u6216\u53d8\u66f4","change_detail":"\u7b7e\u8bc1\u6216\u53d8\u66f4\u60c5\u51b5","is_settled":"\u662f\u5426\u7ed3\u7b97","settled_amount":"\u7ed3\u7b97\u91d1\u989d","paid_amount":"\u5df2\u652f\u4ed8","unpaid_amount":"\u672a\u652f\u4ed8","term":"\u671f\u9650","handler":"\u7ecf\u529e\u4eba","remark":"\u5907\u6ce8","next_bid_date":"\u4e0b\u6b21\u62db\u91c7\u65f6\u95f4","remind_days":"\u63d0\u9192\u63d0\u524d\u5929\u6570"}
    for rk,ck in fm.items():
        if rk in d and d[rk] is not None: rd[ck]=d[rk]
    rd["\u6240\u5c5e\u90e8\u95e8"]=u["\u6240\u5c5e\u90e8\u95e8"]
    if not rd.get("\u63d0\u9192\u63d0\u524d\u5929\u6570") or str(rd.get("\u63d0\u9192\u63d0\u524d\u5929\u6570","")).strip()=="":
        rd["\u63d0\u9192\u63d0\u524d\u5929\u6570"]=get_default_days()
    app_row(rd)
    write_log("合同台账", "新增", obj=d["number"], detail=rd.get("\u9879\u76ee\u540d\u79f0",""))
    for c in read_contracts():
        if str(c.get("\u7f16\u53f7",""))==str(d["number"]):
            return jsonify({"message":"\u5408\u540c\u521b\u5efa\u6210\u529f","number":d["number"],"id":f"row_{c['_row']}"}),201
    return jsonify({"message":"\u5408\u540c\u521b\u5efa\u6210\u529f","number":d["number"]}),201

# ─── 合同编号（6 位）生成预览（不写库，供新增页面即时展示）───
@app.route("/api/contract_number/next", methods=["GET"])
@login_required
def api_next_contract_number():
    """按《合同编号规则》预览下一个编号。

    参数：kind=project|asset（默认 project）；dept 或 unit/section（默认取当前用户所属部门）。
    返回 {number, error, department, unit, section, prefix}；未配置编号规则时 number 为空并给出 error。
    """
    kind = (request.args.get("kind", "project") or "project").strip().lower()
    dept = request.args.get("dept", "").strip()
    unit = request.args.get("unit", "").strip()
    section = request.args.get("section", "").strip()
    if not dept and not unit and not section:
        dept = _number_rule_dept()
    table, column = ("asset_contracts", "id") if kind == "asset" else ("project_contracts", "编号")
    try:
        p = contract_no.plan(dept=dept, unit=unit or None, section=section or None)
        num = contract_no.next_number(_existing_numbers(table, column), p["prefix"])
    except contract_no.NoRuleError as e:
        return jsonify({"number": "", "error": str(e), "department": dept,
                        "unit": unit, "section": section, "prefix": ""})
    except OverflowError as e:
        return jsonify({"number": "", "error": str(e), "department": dept,
                        "unit": unit, "section": section, "prefix": ""})
    return jsonify({"number": num, "error": "", "department": dept,
                    "unit": p["unit"], "section": p["section"], "prefix": p["prefix"]})

# ─── Export ───
@app.route("/api/contracts/export",methods=["GET"])
@login_required
def api_export():
    cts = _scope_contracts(read_contracts())
    params = {k:request.args.get(k,"").strip() for k in ["unit","dept","person","name","no","category","method","amount_min","amount_max","start_date","end_date","remind_start","remind_end","is_settled","vendor","archived","fund_source"]}
    flt = list(cts)
    if params["unit"]: flt=[c for c in flt if params["unit"] in c.get("\u6240\u5c5e\u90e8\u95e8","")]
    if params["dept"]: flt=[c for c in flt if params["dept"] in c.get("\u6240\u5c5e\u90e8\u95e8","")]
    if params["person"]: flt=[c for c in flt if params["person"] in c.get("\u7ecf\u529e\u4eba","")]
    if params["name"]: flt=[c for c in flt if params["name"] in c.get("\u9879\u76ee\u540d\u79f0","")]
    if params["no"]: flt=[c for c in flt if params["no"] in c.get("\u7f16\u53f7","")]
    if params["category"]: flt=[c for c in flt if c.get("\u9879\u76ee\u5206\u7c7b","")==params["category"]]
    if params["method"]: flt=[c for c in flt if c.get("\u91c7\u8d2d\u65b9\u5f0f","")==params["method"]]
    if params["fund_source"]: flt=[c for c in flt if c.get("\u8d44\u91d1\u6765\u6e90","")==params["fund_source"]]
    if params["amount_min"]:
        try: mn=float(params["amount_min"]); flt=[c for c in flt if _num(c.get("\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09")) is not None and _num(c.get("\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09"))>=mn]
        except: pass
    if params["amount_max"]:
        try: mx=float(params["amount_max"]); flt=[c for c in flt if _num(c.get("\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09")) is not None and _num(c.get("\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09"))<=mx]
        except: pass
    if params["start_date"]:
        sd=parse_date(params["start_date"])
        if sd: flt=[c for c in flt if parse_date(c.get("\u91c7\u8d2d\u65f6\u95f4","")) and parse_date(c.get("\u91c7\u8d2d\u65f6\u95f4",""))>=sd]
    if params["end_date"]:
        ed=parse_date(params["end_date"])
        if ed: flt=[c for c in flt if parse_date(c.get("\u91c7\u8d2d\u65f6\u95f4","")) and parse_date(c.get("\u91c7\u8d2d\u65f6\u95f4",""))<=ed]
    if params["remind_start"]:
        rs=parse_date(params["remind_start"])
        if rs: flt=[c for c in flt if parse_date(c.get("\u4e0b\u6b21\u62db\u91c7\u65f6\u95f4","")) and parse_date(c.get("\u4e0b\u6b21\u62db\u91c7\u65f6\u95f4",""))>=rs]
    if params["remind_end"]:
        re_=parse_date(params["remind_end"])
        if re_: flt=[c for c in flt if parse_date(c.get("\u4e0b\u6b21\u62db\u91c7\u65f6\u95f4","")) and parse_date(c.get("\u4e0b\u6b21\u62db\u91c7\u65f6\u95f4",""))<=re_]
    if params["is_settled"] in ("\u662f","\u5426"): flt=[c for c in flt if c.get("\u662f\u5426\u7ed3\u7b97","")==params["is_settled"]]
    if params["vendor"]: flt=[c for c in flt if params["vendor"] in c.get("\u4e2d\u6807\u5355\u4f4d","") or params["vendor"] in c.get("\u5b9e\u9645\u5b9e\u65bd\u5355\u4f4d","")]
    if params["archived"] in ("是","否"): flt=[c for c in flt if (c.get("是否归档","") or "否")==params["archived"]]

    wb = openpyxl.Workbook(); ws = wb.active; ws.title="\u5408\u540c\u53f0\u8d26"
    for ci,cn in enumerate(CONTRACT_COLUMNS,1): ws.cell(row=1,column=ci,value=cn).font=Font(bold=True)
    for ri,c in enumerate(flt,2):
        for ci,cn in enumerate(CONTRACT_COLUMNS,1):
            v=c.get(cn,"")
            if isinstance(v,float): v=round(v,2)
            ws.cell(row=ri,column=ci,value=v)
    bio=BytesIO(); wb.save(bio); bio.seek(0)
    ts=datetime.datetime.now().strftime("%Y%m%d")
    from urllib.parse import quote
    fn = f"\u5408\u540c\u53f0\u8d26\u5bfc\u51fa_{ts}.xlsx"
    write_log("合同台账", "导出", obj="合同台账", detail="导出 " + str(len(flt)) + " 条")
    return Response(bio.getvalue(),mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition":f"attachment; filename*=UTF-8''{quote(fn)}"})

# ─── Import ───
IMPORT_ENUM_CHECK = {
    "项目分类": ["服务类", "维护维修类", "货物类", "建设项目"],
    "采购方式": ["公开招标", "竞争性磋商", "竞争性谈判", "自行采购"],
    "资金来源": ["年初部门预算", "年中追加预算", "统筹预算", "上级专项转移支付", "上年结转结余", "其他"],
}
IMPORT_YN_FIELDS = ("是否有签证或变更", "是否结算")
IMPORT_DATE_FIELDS = ("采购时间", "开工时间", "完工时间", "竣工验收时间", "下次招采时间")
IMPORT_NUM_FIELDS = ("决策的招标控制价", "合同金额（万元）", "结算金额", "已支付", "未支付")
TEMPLATE_EXAMPLE = {
    "编号": "2026-FW-001", "项目分类": "服务类", "项目名称": "示例合同（导入前请删除此行）",
    "决策的招标控制价": 10, "采购方式": "自行采购",
    "中标单位": "示例供应商", "实际实施单位": "示例供应商",
    "合同金额（万元）": 9.5, "资金来源": "年初部门预算", "合同主要内容": "示例内容",
    "采购时间": "2026-01-01", "开工时间": "2026-02-01", "完工时间": "2026-12-31",
    "现场负责人": "张三", "是否有签证或变更": "否", "是否结算": "否",
    "已支付": 5, "未支付": 4.5, "期限": "一年", "经办人": "张三",
    "备注": "示例行，导入前请删除",
    "所属部门": "湘潭市机关事务管理局本级-办公室",
    "下次招采时间": "2027-01-01", "提醒提前天数": 60,
}

def _norm_header(cn):
    if cn == "端工验收时间": return "竣工验收时间"
    if cn in ("经办人/分管领导",): return "经办人"
    if cn in ("甲方现场负责人",): return "现场负责人"
    return cn

def _build_import_vals(hdr, src):
    vals = {cn: "" for cn in CONTRACT_COLUMNS}
    for ci, cn in enumerate(hdr):
        cn2 = _norm_header(cn)
        if cn2 in vals and ci < len(src):
            vals[cn2] = _norm(src[ci])
    return vals

def _check_import_row(vals, existing, file_seen, cur_row):
    errors, warns = [], []
    number = str(vals.get("编号", "")).strip()
    name = str(vals.get("项目名称", "")).strip()
    if not number: errors.append("编号为空")
    if not name: errors.append("项目名称为空")
    if number:
        if number in existing: errors.append("编号已存在于数据库")
        if number in file_seen: errors.append("编号在文件内重复（首次出现于第%d行）" % file_seen[number])
        else: file_seen[number] = cur_row
    # 必填字段校验（与数据库 NOT NULL 、前端必填星号一致：编号/项目名称已在上方单独校验）
    IMPORT_REQUIRED = ("项目分类", "采购方式", "合同金额（万元）", "资金来源", "合同主要内容", "经办人", "所属部门")
    for cn in IMPORT_REQUIRED:
        if not str(vals.get(cn, "")).strip():
            errors.append("%s为空（必填）" % cn)
    for cn in IMPORT_NUM_FIELDS:
        v = vals.get(cn, "")
        if v not in ("", None) and _num(v) is None:
            errors.append("%s不是有效数字" % cn)
    for cn in IMPORT_DATE_FIELDS:
        v = vals.get(cn, "")
        if v not in ("", None) and parse_date(v) is None:
            errors.append("%s日期格式应为yyyy-mm-dd" % cn)
    rv = vals.get("提醒提前天数", "")
    if rv not in ("", None) and _int(rv) is None:
        errors.append("提醒提前天数应为整数")
    dept = str(vals.get("所属部门", "")).strip()
    if dept and dept not in get_dept_strings():
        warns.append("所属部门不存在于部门列表")
    for cn, allowed in IMPORT_ENUM_CHECK.items():
        v = str(vals.get(cn, "")).strip()
        if v and v not in allowed:
            warns.append("%s取值不在标准选项内" % cn)
    for cn in IMPORT_YN_FIELDS:
        v = str(vals.get(cn, "")).strip()
        if v and v not in ("是", "否"):
            warns.append("%s应为是或否" % cn)
    if errors: return ("error", errors)
    if warns: return ("warn", warns)
    return ("ok", [])

def _load_import_file():
    f = request.files.get("file")
    if not f or not f.filename:
        return None, None, "请选择要导入的文件"
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return None, None, "仅支持 .xlsx 格式文件"
    try:
        wb = openpyxl.load_workbook(BytesIO(f.read()), data_only=True)
    except Exception:
        return None, None, "文件解析失败，请使用有效的 xlsx 文件"
    sheet_name = "合同台账" if "合同台账" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows or len(rows) < 2:
        return None, None, "文件中没有可导入的数据"
    hdr = [_norm_header(str(x).strip() if x is not None else "") for x in rows[0]]
    missing = [cn for cn in ("编号", "项目名称") if cn not in hdr]
    if missing:
        return None, None, "文件表头缺少列：" + "、".join(missing) + "，请使用导出的合同台账模板"
    return rows, hdr, None

def _import_existing_numbers():
    return {str(c.get("编号", "")).strip() for c in read_contracts()}

@app.route("/api/contracts/template", methods=["GET"])
@admin_required
def api_contract_template():
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "合同台账"
    headers = list(CONTRACT_COLUMNS)
    for ci, cn in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=cn)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2E74B5")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for ci, cn in enumerate(headers, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = 16
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"
    # 示例行
    for ci, cn in enumerate(headers, 1):
        v = TEMPLATE_EXAMPLE.get(cn, "")
        if v != "":
            ws.cell(row=2, column=ci, value=v)
    ws.cell(row=2, column=1).font = Font(italic=True, color="999999")
    # 下拉验证
    def add_dv(formula, col_letter):
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add("%s3:%s1000" % (col_letter, col_letter))
    col_of = {cn: openpyxl.utils.get_column_letter(i + 1) for i, cn in enumerate(headers)}
    add_dv('"服务类,维护维修类,货物类,建设项目"', col_of["项目分类"])
    add_dv('"公开招标,竞争性磋商,竞争性谈判,自行采购"', col_of["采购方式"])
    add_dv('"是,否"', col_of["是否有签证或变更"])
    add_dv('"是,否"', col_of["是否结算"])
    add_dv('"年初部门预算,年中追加预算,统筹预算,上级专项转移支付,上年结转结余,其他"', col_of["资金来源"])
    # 填写说明
    ws2 = wb.create_sheet("填写说明")
    tips = [
        ("合同台账导入模板填写说明", True),
        ("", False),
        ("1. 第一行为表头，不可删除；第二行为示例，导入前请删除。", False),
        ("2. 必填字段：编号、项目分类、项目名称、采购方式、合同金额（万元）、资金来源、合同主要内容、经办人、所属部门。", False),
        ("3. 编号全局唯一，与数据库重复的行将被跳过不导入。", False),
        ("4. 日期列（采购时间、开工时间、完工时间、竣工验收时间、下次招采时间）格式为 yyyy-mm-dd。", False),
        ("5. 金额列（合同金额、结算金额、已支付、未支付）填写数字。", False),
        ("6. 经办人多人用顿号、分隔；所属部门格式为“单位名称-科室名称”。", False),
        ("7. 下次招采时间填写后系统会按提醒提前天数进行到期提醒；提醒提前天数为空时使用全局默认值。", False),
        ("8. 导入前可先使用“导入预览”查看逐条校验结果，错误行不会被导入。", False),
    ]
    for i, (txt, bold) in enumerate(tips, 1):
        cell = ws2.cell(row=i, column=1, value=txt)
        if bold:
            cell.font = Font(bold=True, size=13, color="1F4D78")
    ws2.column_dimensions["A"].width = 90
    bio = BytesIO(); wb.save(bio); bio.seek(0)
    from urllib.parse import quote
    fn = "合同台账导入模板.xlsx"
    write_log("合同台账", "下载模板", obj="合同台账", detail="下载导入模板")
    return Response(bio.getvalue(), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": "attachment; filename*=UTF-8''%s" % quote(fn)})

@app.route("/api/contracts/import/preview", methods=["POST"])
@admin_required
def api_import_preview():
    rows, hdr, err = _load_import_file()
    if err: return jsonify({"error": err}), 400
    existing = _import_existing_numbers()
    file_seen = {}
    items = []
    for ri in range(1, len(rows)):
        src = rows[ri]
        if src is None: continue
        if all(v is None or str(v).strip() == "" for v in src): continue
        vals = _build_import_vals(hdr, src)
        number = str(vals.get("编号", "")).strip()
        name = str(vals.get("项目名称", "")).strip()
        status, reasons = _check_import_row(vals, existing, file_seen, ri + 1)
        items.append({"row": ri + 1, "number": number, "name": name, "status": status, "reasons": reasons})
    ok_count = sum(1 for it in items if it["status"] == "ok")
    warn_count = sum(1 for it in items if it["status"] == "warn")
    error_count = sum(1 for it in items if it["status"] == "error")
    return jsonify({
        "filename": request.files.get("file").filename,
        "total": len(items), "ok_count": ok_count, "warn_count": warn_count, "error_count": error_count,
        "importable": ok_count + warn_count, "items": items,
    })

@app.route("/api/contracts/import", methods=["POST"])
@admin_required
def api_import_contracts():
    rows, hdr, err = _load_import_file()
    if err: return jsonify({"error": err}), 400
    existing = _import_existing_numbers()
    file_seen = {}
    to_insert = []
    skipped = []
    for ri in range(1, len(rows)):
        src = rows[ri]
        if src is None: continue
        if all(v is None or str(v).strip() == "" for v in src): continue
        vals = _build_import_vals(hdr, src)
        number = str(vals.get("编号", "")).strip()
        name = str(vals.get("项目名称", "")).strip()
        status, reasons = _check_import_row(vals, existing, file_seen, ri + 1)
        if status == "error":
            skipped.append({"row": ri + 1, "number": number, "name": name, "reasons": reasons})
            continue
        to_insert.append(vals)
        existing.add(number)
    added = 0
    if to_insert:
        with db_lock:
            _backup()
            con = _conn()
            try:
                cols = list(CONTRACT_COLUMNS)
                sql = "INSERT INTO project_contracts (" + ",".join('"'+c+'"' for c in cols) + ") VALUES (" + ",".join("?"*len(cols)) + ")"
                for vals in to_insert:
                    con.execute(sql, [vals[cn] for cn in cols])
                    added += 1
                con.commit()
                _invalidate_caches()
            except sqlite3.IntegrityError as e:
                con.rollback()
                return jsonify({"error": "数据库约束拒绝写入（必填字段为空、编号重复、金额非负、日期格式等）：%s" % e}), 400
            except Exception:
                con.rollback(); raise
            finally:
                con.close()
    msg = "导入完成：新增 %d 条" % added
    if skipped: msg += "，跳过 %d 条" % len(skipped)
    write_log("合同台账", "导入", obj=str(request.files.get("file").filename or ""), detail=msg)
    return jsonify({"message": msg, "added": added, "skipped": skipped}), 200

# ─── Get Single Contract ───
@app.route("/api/contracts/<cid>",methods=["GET"])
@login_required
def api_get_contract(cid):
    u = get_cur_user()
    if not cid.startswith("row_"): return jsonify({"error":"\u65e0\u6548\u7684\u5408\u540cID"}),400
    try: rn=int(cid.split("_")[1])
    except: return jsonify({"error":"\u65e0\u6548\u7684\u5408\u540cID"}),400
    for c in read_contracts():
        if c["_row"]==rn:
            if u["\u89d2\u8272"]!="\u7ba1\u7406\u5458":
                vis = _visible_depts(u)
                if vis is not None and str(c.get("\u6240\u5c5e\u90e8\u95e8","")).strip() not in vis:
                    return jsonify({"error":"\u5408\u540c\u4e0d\u5b58\u5728"}),404
            it = c2r(c)
            it["attachments"] = attach_map_by_number().get(str(c.get("\u7f16\u53f7","")), [])
            return jsonify(it)
    return jsonify({"error":"\u5408\u540c\u4e0d\u5b58\u5728"}),404

# ─── Update ───
@app.route("/api/contracts/<cid>",methods=["PUT"])
@login_required
def api_update_contract(cid):
    u = get_cur_user()
    if u["\u89d2\u8272"]=="\u67e5\u8be2\u7528\u6237": return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
    if not cid.startswith("row_"): return jsonify({"error":"\u65e0\u6548\u7684\u5408\u540cID"}),400
    try: rn=int(cid.split("_")[1])
    except: return jsonify({"error":"\u65e0\u6548\u7684\u5408\u540cID"}),400
    t=None
    for c in read_contracts():
        if c["_row"]==rn: t=c; break
    if not t: return jsonify({"error":"\u5408\u540c\u4e0d\u5b58\u5728"}),404
    if str(t.get("是否归档",""))=="是": return jsonify({"error":"合同已归档，仅可浏览，不可修改"}),403
    if not _can_write(t): return jsonify({"error":"\u53ea\u80fd\u7f16\u8f91\u672c\u90e8\u95e8\u6216\u5206\u7ba1\u90e8\u95e8\u5408\u540c"}),403
    d=request.get_json(silent=True) or {}
    if "fund_source" in d and str(d.get("fund_source","")).strip()=="": return jsonify({"error":"字段 'fund_source' 不能为空"}),400
    ud={}
    fm={"number":"\u7f16\u53f7","category":"\u9879\u76ee\u5206\u7c7b","name":"\u9879\u76ee\u540d\u79f0","decision_basis":"\u51b3\u7b56\u4f9d\u636e","bid_control_price":"\u51b3\u7b56\u7684\u62db\u6807\u63a7\u5236\u4ef7","price_basis":"\u62db\u6807\u63a7\u5236\u4ef7\u51b3\u7b56\u7684\u4f9d\u636e","procurement_method":"\u91c7\u8d2d\u65b9\u5f0f","winner":"\u4e2d\u6807\u5355\u4f4d","actual_implementor":"\u5b9e\u9645\u5b9e\u65bd\u5355\u4f4d","amount":"\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09","fund_source":"\u8d44\u91d1\u6765\u6e90","main_content":"\u5408\u540c\u4e3b\u8981\u5185\u5bb9","procurement_date":"\u91c7\u8d2d\u65f6\u95f4","start_date":"\u5f00\u5de5\u65f6\u95f4","end_date":"\u5b8c\u5de5\u65f6\u95f4","acceptance_date":"\u7ae3\u5de5\u9a8c\u6536\u65f6\u95f4","site_manager":"\u73b0\u573a\u8d1f\u8d23\u4eba","has_change":"\u662f\u5426\u6709\u7b7e\u8bc1\u6216\u53d8\u66f4","change_detail":"\u7b7e\u8bc1\u6216\u53d8\u66f4\u60c5\u51b5","is_settled":"\u662f\u5426\u7ed3\u7b97","settled_amount":"\u7ed3\u7b97\u91d1\u989d","paid_amount":"\u5df2\u652f\u4ed8","unpaid_amount":"\u672a\u652f\u4ed8","term":"\u671f\u9650","handler":"\u7ecf\u529e\u4eba","remark":"\u5907\u6ce8","next_bid_date":"\u4e0b\u6b21\u62db\u91c7\u65f6\u95f4","remind_days":"\u63d0\u9192\u63d0\u524d\u5929\u6570"}
    for rk,ck in fm.items():
        if rk in d and d[rk] is not None: ud[ck]=d[rk]
    old_num = str(t.get("\u7f16\u53f7",""))
    new_num = ud.get("\u7f16\u53f7")
    if new_num is not None and str(new_num)!=old_num:
        for c in read_contracts():
            if c["_row"]!=rn and str(c.get("\u7f16\u53f7",""))==str(new_num):
                return jsonify({"error":f"\u7f16\u53f7 '{new_num}' \u5df2\u5b58\u5728"}),400
        rows = read_attachments(); changed=False
        for r in rows:
            if str(r.get("\u5408\u540c\u7f16\u53f7",""))==old_num: r["\u5408\u540c\u7f16\u53f7"]=str(new_num); changed=True
        if changed: write_attachments(rows)
    upd_row(rn,ud)
    log_field_changes("合同台账", t.get("编号",""), t.get("项目名称",""), t, ud)
    write_log("合同台账", "修改", obj=str(t.get("\u7f16\u53f7","")))
    return jsonify({"message":"\u5408\u540c\u66f4\u65b0\u6210\u529f"})

# ─── Delete ───
@app.route("/api/contracts/<cid>",methods=["DELETE"])
@login_required
def api_delete_contract(cid):
    u=get_cur_user()
    if u["\u89d2\u8272"]=="\u67e5\u8be2\u7528\u6237": return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
    if not cid.startswith("row_"): return jsonify({"error":"\u65e0\u6548\u7684\u5408\u540cID"}),400
    try: rn=int(cid.split("_")[1])
    except: return jsonify({"error":"\u65e0\u6548\u7684\u5408\u540cID"}),400
    t=None
    for c in read_contracts():
        if c["_row"]==rn: t=c; break
    if not t: return jsonify({"error":"\u5408\u540c\u4e0d\u5b58\u5728"}),404
    if str(t.get("是否归档",""))=="是": return jsonify({"error":"合同已归档，仅可浏览，不可修改"}),403
    if not _can_write(t): return jsonify({"error":"\u53ea\u80fd\u5220\u9664\u672c\u90e8\u95e8\u6216\u5206\u7ba1\u90e8\u95e8\u5408\u540c"}),403
    data = request.get_json(silent=True) or {}
    reason = str(data.get("reason","")).strip()
    if not reason: return jsonify({"error":"请填写删除原因"}),400

    remove_contract_attachments(t.get("\u7f16\u53f7",""))
    del_row(rn)
    write_log("合同台账", "删除", obj=str(t.get("\u7f16\u53f7","")), detail="删除原因：" + reason)
    return jsonify({"message":"\u5408\u540c\u5220\u9664\u6210\u529f"})

# ---- Archive ----
@app.route("/api/contracts/<cid>/archive",methods=["PUT"])
@login_required
def api_archive_contract(cid):
    u=get_cur_user()
    if u["角色"]=="查询用户": return jsonify({"error":"权限不足"}),403
    if not cid.startswith("row_"): return jsonify({"error":"无效的合同ID"}),400
    try: rn=int(cid.split("_")[1])
    except: return jsonify({"error":"无效的合同ID"}),400
    t=None
    for c in read_contracts():
        if c["_row"]==rn: t=c; break
    if not t: return jsonify({"error":"合同不存在"}),404
    if str(t.get("是否归档",""))=="是": return jsonify({"message":"该合同已归档"}),200
    if u["角色"]!="管理员":
        pf = str(t.get("经办人",""))
        names = _split_names(pf)
        if u["姓名"] not in names:
            return jsonify({"error":"仅管理员或该合同经办人可执行归档"}),403
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    db_exec('UPDATE project_contracts SET "是否归档"=?, "归档人"=?, "归档时间"=? WHERE id=?',
            ("是", u.get("姓名", u.get("用户名","")), ts, rn))
    log_field_changes("合同台账", t.get("编号",""), t.get("项目名称",""), t,
                      {"是否归档": "是", "归档人": u.get("姓名", u.get("用户名","")), "归档时间": ts})
    write_log("合同台账", "归档", obj=str(t.get("编号","")))
    return jsonify({"message":"合同已归档，进入只读状态"})

# ---- Attachments ----
@app.route("/api/contracts/<cid>/attachments",methods=["POST"])
@login_required
def api_upload_attachment(cid):
    u = get_cur_user()
    if u["\u89d2\u8272"]=="\u67e5\u8be2\u7528\u6237": return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
    if not cid.startswith("row_"): return jsonify({"error":"\u65e0\u6548\u7684\u5408\u540cID"}),400
    try: rn=int(cid.split("_")[1])
    except: return jsonify({"error":"\u65e0\u6548\u7684\u5408\u540cID"}),400
    t=None
    for c in read_contracts():
        if c["_row"]==rn: t=c; break
    if not t: return jsonify({"error":"\u5408\u540c\u4e0d\u5b58\u5728"}),404
    if str(t.get("是否归档",""))=="是": return jsonify({"error":"合同已归档，仅可浏览，不可修改"}),403
    if not _can_write(t): return jsonify({"error":"\u53ea\u80fd\u4e3a\u672c\u90e8\u95e8\u6216\u5206\u7ba1\u90e8\u95e8\u5408\u540c\u4e0a\u4f20\u9644\u4ef6"}),403
    atype = (request.form.get("type","") or "").strip()
    if atype not in ATTACH_TYPES: return jsonify({"error":"\u9644\u4ef6\u7c7b\u578b\u65e0\u6548"}),400
    f = request.files.get("file")
    if not f or not f.filename: return jsonify({"error":"\u8bf7\u9009\u62e9\u8981\u4e0a\u4f20\u7684\u6587\u4ef6"}),400
    display = os.path.basename(f.filename.replace("\\","/")).strip()[:200] or "\u9644\u4ef6"
    ext = os.path.splitext(display)[1].lower()
    if ext not in ATTACH_ALLOWED_EXT:
        return jsonify({"error":"\u4ec5\u652f\u6301\u56fe\u7247\u6216PDF\u683c\u5f0f\uff08jpg/jpeg/png/gif/webp/bmp/pdf\uff09"}),400
    max_mb = get_attach_max_mb()
    max_bytes = max_mb * 1024 * 1024
    store_name = uuid.uuid4().hex + ext
    os.makedirs(ATTACH_DIR,exist_ok=True)
    fp = os.path.join(ATTACH_DIR, store_name)
    f.save(fp)
    size = os.path.getsize(fp)
    if size > max_bytes:
        try: os.remove(fp)
        except Exception: pass
        return jsonify({"error":"附件大小不能超过 %dM（当前 %.2fM）" % (max_mb, size/1048576.0)}),400
    rows = read_attachments()
    rows.append({
        "\u9644\u4ef6ID": uuid.uuid4().hex,
        "\u5408\u540c\u7f16\u53f7": str(t.get("\u7f16\u53f7","")),
        "\u9644\u4ef6\u7c7b\u578b": ATTACH_TYPES[atype],
        "\u663e\u793a\u540d\u79f0": display,
        "\u5b58\u50a8\u6587\u4ef6\u540d": store_name,
        "\u6587\u4ef6\u5927\u5c0f": str(size),
        "\u4e0a\u4f20\u4eba": u.get("\u59d3\u540d", u.get("\u7528\u6237\u540d","")),
        "\u4e0a\u4f20\u65f6\u95f4": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
    write_attachments(rows)
    write_log("合同台账", "上传附件", obj=str(t.get("\u7f16\u53f7","")), detail=display)
    return jsonify({"message":"\u9644\u4ef6\u4e0a\u4f20\u6210\u529f","attachment":attachment_json(rows[-1])}),201

@app.route("/api/attachments/<att_id>/download",methods=["GET"])
@login_required
def api_download_attachment(att_id):
    u = get_cur_user()
    row = next((r for r in read_attachments() if r.get("\u9644\u4ef6ID","")==att_id), None)
    if not row: return jsonify({"error":"\u9644\u4ef6\u4e0d\u5b58\u5728"}),404
    if u["\u89d2\u8272"]!="\u7ba1\u7406\u5458":
        vis = _visible_depts(u)
        cnum = str(row.get("\u5408\u540c\u7f16\u53f7",""))
        if vis is not None and not any(str(c.get("\u7f16\u53f7",""))==cnum and str(c.get("\u6240\u5c5e\u90e8\u95e8","")).strip() in vis for c in read_contracts()):
            return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
    fp = os.path.join(ATTACH_DIR, row.get("\u5b58\u50a8\u6587\u4ef6\u540d",""))
    if not os.path.exists(fp): return jsonify({"error":"\u9644\u4ef6\u6587\u4ef6\u4e0d\u5b58\u5728"}),404
    return send_file(fp, as_attachment=True, download_name=row.get("\u663e\u793a\u540d\u79f0","\u9644\u4ef6"))

@app.route("/api/attachments/<att_id>",methods=["DELETE"])
@login_required
def api_delete_attachment(att_id):
    u=get_cur_user()
    if u["\u89d2\u8272"]=="\u67e5\u8be2\u7528\u6237": return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
    rows = read_attachments()
    row = next((r for r in rows if r.get("\u9644\u4ef6ID","")==att_id), None)
    if not row: return jsonify({"error":"\u9644\u4ef6\u4e0d\u5b58\u5728"}),404
    for c in read_contracts():
        if str(c.get("\u7f16\u53f7",""))==str(row.get("\u5408\u540c\u7f16\u53f7","")):
            if not _can_write(c): return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
            if c.get("是否归档","")=="是": return jsonify({"error":"合同已归档，仅可浏览，不可修改"}),403
            break
    fp = os.path.join(ATTACH_DIR, row.get("\u5b58\u50a8\u6587\u4ef6\u540d",""))
    try:
        if os.path.exists(fp): os.remove(fp)
    except Exception: pass
    write_attachments([r for r in rows if r.get("\u9644\u4ef6ID","")!=att_id])
    write_log("合同台账", "删除附件", obj=str(row.get("\u5408\u540c\u7f16\u53f7","")), detail=str(row.get("\u663e\u793a\u540d\u79f0","")))
    return jsonify({"message":"\u9644\u4ef6\u5df2\u5220\u9664"})

# ─── Clear Remind ───
@app.route("/api/contracts/<cid>/clear_remind",methods=["PUT"])
@login_required
def api_clear_remind(cid):
    u=get_cur_user()
    if u["\u89d2\u8272"]=="\u67e5\u8be2\u7528\u6237": return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
    if not cid.startswith("row_"): return jsonify({"error":"\u65e0\u6548\u7684\u5408\u540cID"}),400
    try: rn=int(cid.split("_")[1])
    except: return jsonify({"error":"\u65e0\u6548\u7684\u5408\u540cID"}),400
    t=None
    for c in read_contracts():
        if c["_row"]==rn: t=c; break
    if not t: return jsonify({"error":"\u5408\u540c\u4e0d\u5b58\u5728"}),404
    if not _can_write(t): return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
    upd_cell(rn,"\u4e0b\u6b21\u62db\u91c7\u65f6\u95f4","")
    log_field_changes("合同台账", t.get("编号",""), t.get("项目名称",""), t, {"下次招采时间": ""})
    write_log("合同台账", "清空提醒", obj=str(t.get("\u7f16\u53f7","")))
    return jsonify({"message":"\u63d0\u9192\u5df2\u6e05\u9664"})

# 资产合同：永久关闭提醒（清空“下次出租时间”）
@app.route("/api/asset_contracts/<aid>/clear_remind", methods=["PUT"])
@login_required
def api_clear_asset_remind(aid):
    u = get_cur_user()
    if u["角色"] == "查询用户": return jsonify({"error":"权限不足"}),403
    rid = _asset_id(aid)
    if rid is None: return jsonify({"error":"无效的资产合同ID"}),400
    t = next((a for a in read_assets() if a["_row"] == rid), None)
    if not t: return jsonify({"error":"资产合同不存在"}),404
    if t.get("archived"): return jsonify({"error":"该资产合同已归档，仅可浏览，不可修改"}),403
    raw = next((r for r in db_query("SELECT * FROM asset_contracts WHERE id=?", (rid,))), None)
    db_exec('UPDATE asset_contracts SET "下次出租时间"=? WHERE id=?', ("", rid))
    if raw:
        log_field_changes("资产合同台账", rid, raw.get("资产名称", ""), raw, {"下次出租时间": ""})
    write_log("资产合同台账", "清空提醒", obj=str(t.get("asset_name", "")))
    return jsonify({"message":"提醒已清除"})

# ═══════════════ 资产合同台账（asset_contracts）═══════════════
def _asset_norm_date(v):
    """资产合同日期规范化：接受 date/datetime 与 2021.11.15 / 2021-11-15 / 2021/11/15（可带尾点）。"""
    if v is None or str(v).strip() == "": return None
    if isinstance(v, datetime.datetime): return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date): return v.strftime("%Y-%m-%d")
    s = str(v).strip().strip(".").strip()
    s = s.replace("年", ".").replace("月", ".").replace("日", "")
    p = s.replace("/", ".").replace("-", ".").replace(" ", "").split(".")
    if len(p) != 3: return None
    try:
        return datetime.date(int(p[0]), int(p[1]), int(p[2])).strftime("%Y-%m-%d")
    except Exception:
        return None

def a2r(r):
    """asset_contracts 数据行 → 接口结构（附件/部门等无关字段不参与）。"""
    it = {"id": "arow_%s" % r.get("id"), "_row": r.get("id")}
    for k, cn in ASSET_FIELDS:
        v = r.get(cn)
        it[k] = "" if v is None else v
    for cn in ASSET_ARCHIVE_COLS:
        v = r.get(cn)
        it[cn] = "" if v is None else v
    it["archived"] = it.get("是否归档", "") == "是"
    it["archive_by"] = it.get("归档人", "")
    it["archive_at"] = it.get("归档时间", "")
    it["短信提醒"] = "" if r.get("短信提醒") is None else r.get("短信提醒")
    it["sms_switch"] = it["短信提醒"] or "开启"
    return it

def read_assets():
    """资产合同列表（含计算列）：paid_total = 子表实缴合计（列表/统计的“已缴租金合计”口径），
    due_total = 子表应缴合计，rent_count = 子表记录数（删除确认用）。"""
    cached = _cache_get("assets")
    if cached is None:
        totals = asset_rent_totals()
        out = []
        for r in db_query("SELECT * FROM asset_contracts ORDER BY id"):
            it = a2r(r)
            t = totals.get(_int(r.get("id")) or 0, {"count": 0, "due": 0.0, "paid": 0.0})
            it["rent_count"] = t["count"]
            it["due_total"] = t["due"]
            it["paid_total"] = t["paid"]
            out.append(it)
        _cache_put("assets", out)
        cached = out
    return [dict(x) for x in cached]

def read_asset_attachments():
    return db_query("SELECT * FROM asset_attachments ORDER BY rowid")

def write_asset_attachments(rows):
    with db_lock:
        _backup()
        con = _conn()
        try:
            con.execute("DELETE FROM asset_attachments")
            for r in rows:
                _insert_dict(con, "asset_attachments", ASSET_ATTACH_HEADERS, r)
            con.commit()
        except Exception:
            con.rollback(); raise
        finally:
            con.close()

def asset_attachment_json(row):
    tlabel = row.get("附件类型", "")
    tkey = next((k for k, v in ASSET_ATTACH_TYPES.items() if v == tlabel), "")
    return {
        "id": row.get("附件ID", ""), "asset_id": row.get("资产合同ID", ""),
        "type": tkey, "type_label": tlabel, "display_name": row.get("显示名称", ""),
        "size": row.get("文件大小", ""), "uploader": row.get("上传人", ""),
        "uploaded_at": row.get("上传时间", ""),
    }

def asset_attach_map():
    m = {}
    for r in read_asset_attachments():
        m.setdefault(_int(r.get("资产合同ID")) or 0, []).append(asset_attachment_json(r))
    return m

def remove_asset_attachments(rid):
    rows = read_asset_attachments()
    keep, changed = [], False
    for r in rows:
        if (_int(r.get("资产合同ID")) or 0) == (_int(rid) or 0):
            fp = os.path.join(ATTACH_DIR, r.get("存储文件名", ""))
            try:
                if os.path.exists(fp): os.remove(fp)
            except Exception:
                pass
            changed = True
        else:
            keep.append(r)
    if changed:
        write_asset_attachments(keep)

def _asset_places(items):
    return sorted({str(a.get("place", "")).strip() for a in items if str(a.get("place", "")).strip()})

# ─── 资产合同租金收取明细（子表 asset_rent_collections：一个合同 N 条收款记录）───
def _rent_vals(d):
    """租金明细：接口字段 → 数据库列值（不含“资产合同ID”，由接口按合同路径注入）。

    校验：收款期次、应缴金额（元）必填；金额为数字、≥0、保留两位；日期 yyyy-mm-dd、可空；
    实缴金额（元）留空按 0 落库（列本身为 NOT NULL DEFAULT 0）。返回 (vals, err)。
    """
    vals = {}
    for k, cn in ASSET_RENT_FIELDS:
        if k == "asset_id":
            continue
        v = d.get(k, "")
        s = "" if v is None else str(v).strip()
        if k in ASSET_RENT_NUM_KEYS:
            if s == "":
                vals[cn] = None if k == "due_amount" else 0.0
            else:
                try:
                    f = float(s.replace(",", ""))
                except ValueError:
                    return None, "%s 不是有效数字" % cn
                if f < 0:
                    return None, "%s 不能为负数" % cn
                vals[cn] = round(f, 2)
        elif k in ASSET_RENT_DATE_KEYS:
            if s == "":
                vals[cn] = None
            else:
                nd = _asset_norm_date(v)
                if not nd:
                    return None, "%s 日期格式应为 yyyy-mm-dd" % cn
                vals[cn] = nd
        elif k == "pay_method":
            if s and s not in ASSET_RENT_PAY_METHODS:
                return None, "%s 只能是：%s" % (cn, "/".join(ASSET_RENT_PAY_METHODS))
            vals[cn] = s
        else:
            vals[cn] = s
    for k in ASSET_RENT_REQUIRED:
        if not str(d.get(k, "") or "").strip():
            return None, "%s 不能为空" % ASSET_RENT_KEY_TO_COL[k]
    if vals.get("应缴金额（元）") is None:
        return None, "应缴金额（元）不能为空"
    if vals.get("实缴金额（元）") is None:
        vals["实缴金额（元）"] = 0.0
    return vals, None

def _rent_row_json(row):
    """asset_rent_collections 数据行 → 接口结构（金额两位小数，空值转空串）。"""
    it = {"id": _int(row.get("id")) or 0}
    for k, cn in ASSET_RENT_FIELDS:
        v = row.get(cn)
        if k == "asset_id":
            it[k] = _int(v) or 0
        elif k in ASSET_RENT_NUM_KEYS:
            it[k] = "" if v is None else round(float(v), 2)
        else:
            it[k] = "" if v is None else v
    return it

def read_asset_rent(contract_id=None):
    """读取租金收取明细；给定 contract_id 时仅读取该合同。"""
    if contract_id is None:
        rows = db_query('SELECT * FROM asset_rent_collections ORDER BY "资产合同ID", id')
    else:
        rows = db_query('SELECT * FROM asset_rent_collections WHERE "资产合同ID"=? ORDER BY id', (int(contract_id),))
    return [_rent_row_json(r) for r in rows]

def asset_rent_map():
    m = {}
    for r in read_asset_rent():
        m.setdefault(_int(r.get("asset_id")) or 0, []).append(r)
    return m

def asset_rent_totals():
    """{合同id: {"count":N,"due":应缴合计,"paid":实缴合计}}——用于列表/统计的“已缴租金合计”计算列。"""
    m = {}
    for r in db_query('SELECT "资产合同ID" AS cid, COUNT(*) AS n, SUM("应缴金额（元）") AS due, '
                      'SUM("实缴金额（元）") AS paid FROM asset_rent_collections GROUP BY "资产合同ID"'):
        m[_int(r.get("cid")) or 0] = {"count": int(r.get("n") or 0),
                                      "due": round(float(r.get("due") or 0), 2),
                                      "paid": round(float(r.get("paid") or 0), 2)}
    return m

def _rent_totals_of(rows):
    """一组租金明细的汇总：应缴合计 / 实缴合计 / 差额（应缴-实缴）。"""
    due = sum(float(r.get("due_amount") or 0) for r in rows)
    paid = sum(float(r.get("paid_amount") or 0) for r in rows)
    return {"count": len(rows), "due": round(due, 2), "paid": round(paid, 2), "diff": round(due - paid, 2)}

def replace_asset_rent(rid, items, allow_clear=False):
    """按“整表替换”语义保存某合同的租金收取明细：新增/修改/删除在一次事务内完成（原子）。

    - items 中带 id 的为已有记录（必须是该合同下的记录），不带 id 的为新增；未出现在 items 中的已有记录被删除。
    - “资产合同ID”一律取接口路径上的合同 id，前端无需（也不允许）指定。
    - 防误删：当 items 为空而该合同下已有记录时，必须显式传入 allow_clear=True（接口层由 confirm_clear 控制），
      否则抛 ValueError——避免“明细未加载成功就保存”把已有收款记录清空的误操作。
    - 校验失败抛 ValueError（接口层转 400），事务内异常回滚。
    返回 (统计 dict, 保存后的明细列表)。
    """
    rid = int(rid)
    if not isinstance(items, list):
        raise ValueError("租金收取明细格式不正确")
    if len(items) > ASSET_RENT_MAX_ROWS:
        raise ValueError("单个资产合同的租金收取记录不能超过 %d 条" % ASSET_RENT_MAX_ROWS)
    prepared = []
    for i, it in enumerate(items, 1):
        if not isinstance(it, dict):
            raise ValueError("第 %d 条租金记录格式不正确" % i)
        vals, err = _rent_vals(it)
        if err:
            raise ValueError("第 %d 条租金记录：%s" % (i, err))
        raw_id = it.get("id")
        try:
            row_id = int(raw_id) if str(raw_id if raw_id is not None else "").strip() not in ("", "None") else None
        except Exception:
            raise ValueError("第 %d 条租金记录：记录ID无效" % i)
        prepared.append((row_id, vals))
    stats = {"inserted": 0, "updated": 0, "deleted": 0}
    with db_lock:
        _backup()
        con = _conn()
        try:
            existing = set(r[0] for r in con.execute(
                'SELECT id FROM asset_rent_collections WHERE "资产合同ID"=?', (rid,)).fetchall())
            if not prepared and existing and not allow_clear:
                raise ValueError("未收到租金收取明细（可能是明细未加载成功）；如确认清空该合同的全部 %d 条租金收取记录，请重新提交并带上确认标记" % len(existing))
            keep = set()
            for row_id, vals in prepared:
                cols = list(vals.keys())
                if row_id is None:
                    allcols = ["资产合同ID"] + cols
                    con.execute("INSERT INTO asset_rent_collections (" + ",".join('"%s"' % c for c in allcols) +
                                ") VALUES (" + ",".join("?" * len(allcols)) + ")",
                                [rid] + [vals[c] for c in cols])
                    stats["inserted"] += 1
                else:
                    if row_id in keep:
                        raise ValueError("租金记录重复提交（id=%d）" % row_id)
                    if row_id not in existing:
                        raise ValueError("租金记录不存在或不属于该资产合同（id=%d）" % row_id)
                    keep.add(row_id)
                    con.execute("UPDATE asset_rent_collections SET " + ",".join('"%s"=?' % c for c in cols) +
                                ' WHERE id=? AND "资产合同ID"=?', [vals[c] for c in cols] + [row_id, rid])
                    stats["updated"] += 1
            for old_id in sorted(existing - keep):
                con.execute('DELETE FROM asset_rent_collections WHERE id=? AND "资产合同ID"=?', (old_id, rid))
                stats["deleted"] += 1
            con.commit()
            _invalidate_caches()      # 租金明细变动会影响“待收/已缴”等汇总
        except Exception:
            con.rollback()
            raise
        finally:
            con.close()
    return stats, read_asset_rent(rid)

# ─── 租金收取明细的字段级变更日志（供“字段级变更日志”页审计）───
def _log_val(v):
    """日志里的值格式化：纯数字去掉多余的 .0 / 尾随零，其余原样（空值记空串）。"""
    if v is None:
        return ""
    s = str(v).strip()
    try:
        f = float(s)
    except Exception:
        return s
    if s.lstrip("-").replace(".", "", 1).isdigit():
        return ("%.2f" % f).rstrip("0").rstrip(".") if f != int(f) else str(int(f))
    return s

def _rent_row_brief(r):
    """一条租金明细的摘要文本（新增/删除行各记一条汇总日志）。"""
    parts = []
    if _log_val(r.get("due_amount")) != "": parts.append("应缴 %s" % _log_val(r.get("due_amount")))
    if _log_val(r.get("paid_amount")) != "": parts.append("实缴 %s" % _log_val(r.get("paid_amount")))
    if str(r.get("pay_date") or "").strip(): parts.append("缴款日期 %s" % r.get("pay_date"))
    if str(r.get("pay_method") or "").strip(): parts.append("缴款方式 %s" % r.get("pay_method"))
    if str(r.get("payee") or "").strip(): parts.append("收款人 %s" % r.get("payee"))
    return "；".join(parts) or "（无金额信息）"

def _log_rent_collection_changes(t, old_rows, new_rows):
    """把本次租金明细的变更写入 field_change_logs（模块=资产合同台账，对象编号=资产合同 id）。

    - 已有记录：逐字段记录变化，字段名 = 租金明细[收款期次]·列名（如 租金明细[2026年第1季度]·实缴金额（元））；
    - 新增/删除记录：各记一条汇总（字段名 = 租金明细[收款期次]，原值/新值记摘要）；
    - 无实际变化时不写日志；写日志异常不影响主流程（log_field_changes 内部已兜底）。
    """
    try:
        old_map, new_map = {}, {}
        new_ids = set()
        for r in (new_rows or []):
            rid = r.get("id")
            new_ids.add(rid)
            label = "租金明细[%s]" % (r.get("period_label") or ("id=%s" % rid))
            prev = (old_rows or {}).get(rid)
            if prev:
                # 注意：read_asset_rent 返回的是“接口字段名”字典，需按 ASSET_RENT_FIELDS 映射到中文列名
                for k, c in ASSET_RENT_FIELDS:
                    if k == "asset_id":
                        continue
                    key = "%s·%s" % (label, c)
                    old_map[key] = _log_val(prev.get(k, ""))
                    new_map[key] = _log_val(r.get(k, ""))
            else:
                old_map[label] = "（空）"
                new_map[label] = "新增：" + _rent_row_brief(r)
        for rid, prev in (old_rows or {}).items():
            if rid in new_ids:
                continue
            label = "租金明细[%s]" % (prev.get("period_label") or ("id=%s" % rid))
            old_map[label] = "删除：" + _rent_row_brief(prev)
            new_map[label] = "（空）"
        if old_map or new_map:
            log_field_changes("资产合同台账", t.get("_row"), t.get("asset_name", ""), old_map, new_map)
    except Exception:
        pass

def _asset_filter(items, q):
    flt = list(items)
    if q.get("place"): flt = [a for a in flt if q["place"] == str(a.get("place", "")).strip()]
    if q.get("tenant"): flt = [a for a in flt if q["tenant"] in str(a.get("tenant", ""))]
    if q.get("asset_name"): flt = [a for a in flt if q["asset_name"] in str(a.get("asset_name", ""))]
    if q.get("dept"): flt = [a for a in flt if q["dept"] in str(a.get("dept", ""))]
    if q.get("person"): flt = [a for a in flt if q["person"] in str(a.get("person", ""))]
    if q.get("unit"): flt = [a for a in flt if q["unit"] in str(a.get("dept", ""))]
    if q.get("archived") in ("是", "否"):
        want = q["archived"] == "是"
        flt = [a for a in flt if bool(a.get("archived")) == want]
    if q.get("seq"):
        sq = _int(q["seq"])
        if sq is not None: flt = [a for a in flt if _int(a.get("seq")) == sq]
    if q.get("lease_start_start"):
        sd = parse_date(q["lease_start_start"])
        if sd: flt = [a for a in flt if parse_date(a.get("lease_start", "")) and parse_date(a.get("lease_start", "")) >= sd]
    if q.get("lease_start_end"):
        ed = parse_date(q["lease_start_end"])
        if ed: flt = [a for a in flt if parse_date(a.get("lease_start", "")) and parse_date(a.get("lease_start", "")) <= ed]
    if q.get("lease_end_start"):
        sd = parse_date(q["lease_end_start"])
        if sd: flt = [a for a in flt if parse_date(a.get("lease_end", "")) and parse_date(a.get("lease_end", "")) >= sd]
    if q.get("lease_end_end"):
        ed = parse_date(q["lease_end_end"])
        if ed: flt = [a for a in flt if parse_date(a.get("lease_end", "")) and parse_date(a.get("lease_end", "")) <= ed]
    return flt

def _asset_query():
    return {k: request.args.get(k, "").strip() for k in ASSET_QUERY_KEYS}

def _asset_id(aid):
    """解析接口路径上的资产合同 id（arow_<编号>）。

    编号统一后为 6 位字符串（写入 asset_contracts.id），历史数据为整数自增 id；
    此处返回字符串原值，调用方按字符串比较（数值比较统一走 _int）。
    """
    if not aid or not str(aid).startswith("arow_"): return None
    sid = str(aid).split("_", 1)[1].strip() if "_" in str(aid) else ""
    return sid or None

def _asset_vals(d, require_all=False):
    """接口字段 → 数据库列值；返回 (vals, 错误信息)。require_all=True 时三个必填字段必须齐全。"""
    vals = {}
    for k, cn in ASSET_FIELDS:
        if k not in d: continue
        v = d[k]
        s = "" if v is None else str(v).strip()
        if k in ASSET_NUM_KEYS:
            if s == "":
                vals[cn] = None
            else:
                try:
                    f = float(s.replace(",", ""))
                except ValueError:
                    return None, "%s 不是有效数字" % cn
                if f < 0: return None, "%s 不能为负数" % cn
                vals[cn] = round(f, 2)
        elif k in ASSET_DATE_KEYS:
            if s == "":
                vals[cn] = None
            else:
                nd = _asset_norm_date(v)
                if not nd: return None, "%s 日期格式应为 yyyy-mm-dd" % cn
                vals[cn] = nd
        elif k == "seq":
            if s == "":
                vals[cn] = None
            else:
                try: vals[cn] = int(float(s))
                except ValueError: return None, "序号 应为整数"
        elif k == "remind_days":
            if s == "":
                vals[cn] = None
            else:
                try: vals[cn] = int(float(s))
                except ValueError: return None, "提醒提前天数 应为整数"
        else:
            vals[cn] = s
    for k in ASSET_REQUIRED:
        cn = ASSET_KEY_TO_COL[k]
        if require_all and not str(d.get(k, "")).strip():
            return None, "%s 不能为空" % cn
        if k in d and not str(vals.get(cn) or "").strip():
            return None, "%s 不能为空" % cn
    return vals, None

def _asset_sort(items):
    """列表/导出的“序号”列取自 asset_contracts.id，故按 id 降序排列。"""
    return sorted(items, key=lambda a: _int(a.get("_row")) or 0, reverse=True)

@app.route("/api/asset_contracts", methods=["GET"])
@login_required
def api_list_assets():
    all_items = _scope_assets(read_assets())
    flt = _asset_sort(_asset_filter(all_items, _asset_query()))
    try:
        p = int(request.args.get("page", 1)); ps = min(int(request.args.get("page_size", 20)), 200)
        if p < 1: p = 1
        if ps < 1: ps = 20
    except Exception:
        p, ps = 1, 20
    si = (p - 1) * ps
    att_map = asset_attach_map()
    for it in flt:
        it["attachments"] = att_map.get(_int(it.get("_row")) or 0, [])
    return jsonify({"total": len(flt), "page": p, "page_size": ps,
                    "items": flt[si:si + ps], "places": _asset_places(all_items)})

@app.route("/api/asset_contracts", methods=["POST"])
@login_required
def api_create_asset():
    u = get_cur_user()
    if u["角色"] == "查询用户": return jsonify({"error":"权限不足"}),403
    d = request.get_json(silent=True) or {}
    vals, err = _asset_vals(d, require_all=True)
    if err: return jsonify({"error": err}),400
    if not vals: return jsonify({"error":"没有需要保存的字段"}),400
    # 资产合同编号写入 id：未提供时按《合同编号规则》自动生成（6 位；未配置规则的科室/组室不生成、不猜测）
    number = str(d.get("number", "") or "").strip()
    dept = str(vals.get("管理部门") or "").strip() or _number_rule_dept(u)
    if not number:
        number, err = next_contract_number("asset_contracts", "id", dept)
        if err: return jsonify({"error": "无法自动生成资产合同编号：" + err}), 400
    elif not (len(number) == 6 and number.isdigit()):
        return jsonify({"error": "资产合同编号应为 6 位数字"}), 400
    if db_query("SELECT id FROM asset_contracts WHERE id=?", (number, )):
        return jsonify({"error": "编号 '%s' 已存在" % number}), 400
    cols = ["id"] + list(vals.keys())
    sql = "INSERT INTO asset_contracts (" + ",".join('"%s"' % c for c in cols) + ") VALUES (" + ",".join("?" * len(cols)) + ")"
    db_exec(sql, [number] + [vals[c] for c in cols[1:]])
    write_log("资产合同台账", "新增", obj=str(vals.get("资产名称", "")),
              detail="编号 %s；%s" % (number, vals.get("承租人", "")))
    return jsonify({"message":"资产合同新增成功", "id":"arow_%s" % number, "number": number}),201

@app.route("/api/asset_contracts/<aid>", methods=["GET"])
@login_required
def api_get_asset(aid):
    rid = _asset_id(aid)
    if rid is None: return jsonify({"error":"无效的资产合同ID"}),400
    for a in _scope_assets(read_assets()):
        if a["_row"] == rid:
            a["attachments"] = asset_attach_map().get(_int(rid) or 0, [])
            a["rent_collections"] = read_asset_rent(rid)
            return jsonify(a)
    return jsonify({"error":"资产合同不存在"}),404

@app.route("/api/asset_contracts/<aid>", methods=["PUT"])
@login_required
def api_update_asset(aid):
    u = get_cur_user()
    if u["角色"] == "查询用户": return jsonify({"error":"权限不足"}),403
    rid = _asset_id(aid)
    if rid is None: return jsonify({"error":"无效的资产合同ID"}),400
    t = next((a for a in read_assets() if a["_row"] == rid), None)
    if not t: return jsonify({"error":"资产合同不存在"}),404
    if t.get("archived"): return jsonify({"error":"该资产合同已归档，仅可浏览，不可修改"}),403
    d = request.get_json(silent=True) or {}
    vals, err = _asset_vals(d)
    if err: return jsonify({"error": err}),400
    if not vals: return jsonify({"error":"没有需要更新的字段"}),400
    raw = next((r for r in db_query("SELECT * FROM asset_contracts WHERE id=?", (rid,))), None)
    sql = "UPDATE asset_contracts SET " + ",".join('"%s"=?' % c for c in vals) + " WHERE id=?"
    db_exec(sql, list(vals.values()) + [rid])
    if raw:
        log_field_changes("资产合同台账", rid, raw.get("资产名称", ""), raw, vals)
    write_log("资产合同台账", "修改", obj=str(t.get("asset_name", "")), detail=str(t.get("tenant", "")))
    return jsonify({"message":"资产合同更新成功"})

@app.route("/api/asset_contracts/<aid>", methods=["DELETE"])
@login_required
def api_delete_asset(aid):
    u = get_cur_user()
    if u["角色"] == "查询用户": return jsonify({"error":"权限不足"}),403
    rid = _asset_id(aid)
    if rid is None: return jsonify({"error":"无效的资产合同ID"}),400
    t = next((a for a in read_assets() if a["_row"] == rid), None)
    if not t: return jsonify({"error":"资产合同不存在"}),404
    if t.get("archived"): return jsonify({"error":"该资产合同已归档，仅可浏览，不可删除"}),403
    data = request.get_json(silent=True) or {}
    reason = str(data.get("reason", "")).strip()
    if not reason: return jsonify({"error":"请填写删除原因"}),400
    # 外键为 ON DELETE RESTRICT：租金收取记录属财务留痕，删除合同前须显式确认“一并删除”
    rent_cnt = _int(t.get("rent_count")) or 0
    if rent_cnt and not data.get("delete_rent"):
        return jsonify({"error": "该资产合同存在 %d 条租金收取记录，删除合同将一并删除这些记录，请确认后重试" % rent_cnt,
                        "rent_count": rent_cnt}), 409
    if rent_cnt:
        db_exec('DELETE FROM asset_rent_collections WHERE "资产合同ID"=?', (rid,))
    remove_asset_attachments(rid)
    db_exec("DELETE FROM asset_contracts WHERE id=?", (rid,))
    write_log("资产合同台账", "删除", obj=str(t.get("asset_name", "")),
              detail="删除原因：" + reason + ("；同时删除租金收取记录 %d 条" % rent_cnt if rent_cnt else ""))
    return jsonify({"message":"资产合同删除成功"})

# ---- 资产合同租金收取明细（子表 asset_rent_collections）----
@app.route("/api/asset_contracts/<aid>/rent_collections", methods=["GET"])
@login_required
def api_list_asset_rent(aid):
    rid = _asset_id(aid)
    if rid is None: return jsonify({"error":"无效的资产合同ID"}),400
    if not next((a for a in _scope_assets(read_assets()) if a["_row"] == rid), None):
        return jsonify({"error":"资产合同不存在"}),404
    rows = read_asset_rent(rid)
    return jsonify({"items": rows, "totals": _rent_totals_of(rows)})

@app.route("/api/asset_contracts/<aid>/rent_collections", methods=["PUT"])
@login_required
def api_save_asset_rent(aid):
    """整表替换保存某合同的租金收取明细（前端弹窗点“保存”时调用，合同 id 由路径提供）。"""
    u = get_cur_user()
    if u["角色"] == "查询用户": return jsonify({"error":"权限不足"}),403
    rid = _asset_id(aid)
    if rid is None: return jsonify({"error":"无效的资产合同ID"}),400
    t = next((a for a in _scope_assets(read_assets()) if a["_row"] == rid), None)
    if not t: return jsonify({"error":"资产合同不存在"}),404
    if t.get("archived"): return jsonify({"error":"该资产合同已归档，仅可浏览，不可修改租金收取明细"}),403
    d = request.get_json(silent=True) or {}
    old_rows = {r["id"]: r for r in read_asset_rent(rid)}      # 变更前的明细（写字段级日志用）
    try:
        stats, rows = replace_asset_rent(rid, d.get("items", []), allow_clear=bool(d.get("confirm_clear")))
    except ValueError as e:
        return jsonify({"error": str(e)}),400
    _log_rent_collection_changes(t, old_rows, rows)
    parts = []
    if stats["inserted"]: parts.append("新增 %d 条" % stats["inserted"])
    if stats["updated"]: parts.append("修改 %d 条" % stats["updated"])
    if stats["deleted"]: parts.append("删除 %d 条" % stats["deleted"])
    if parts:
        write_log("资产合同台账", "修改租金明细", obj=str(t.get("asset_name", "")),
                  detail=str(t.get("tenant", "")) + "：" + "、".join(parts))
    return jsonify({"message": "租金收取明细已保存", "items": rows,
                    "totals": _rent_totals_of(rows), "stats": stats})

# ---- 资产合同归档 ----
@app.route("/api/asset_contracts/<aid>/archive", methods=["PUT"])
@login_required
def api_archive_asset(aid):
    u = get_cur_user()
    if u["角色"] == "查询用户": return jsonify({"error":"权限不足"}),403
    rid = _asset_id(aid)
    if rid is None: return jsonify({"error":"无效的资产合同ID"}),400
    t = next((a for a in read_assets() if a["_row"] == rid), None)
    if not t: return jsonify({"error":"资产合同不存在"}),404
    if t.get("archived"): return jsonify({"message":"该资产合同已归档"}),200
    if u["角色"] != "管理员":
        names = _split_names(str(t.get("person", "")))
        if u["姓名"] not in names:
            return jsonify({"error":"仅管理员或该资产合同责任人可执行归档"}),403
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    raw = next((r for r in db_query("SELECT * FROM asset_contracts WHERE id=?", (rid,))), None)
    db_exec('UPDATE asset_contracts SET "是否归档"=?, "归档人"=?, "归档时间"=? WHERE id=?',
            ("是", u.get("姓名", u.get("用户名", "")), ts, rid))
    if raw:
        log_field_changes("资产合同台账", rid, raw.get("资产名称", ""), raw,
                          {"是否归档": "是", "归档人": u.get("姓名", u.get("用户名","")), "归档时间": ts})
    write_log("资产合同台账", "归档", obj=str(t.get("asset_name", "")), detail=str(t.get("tenant", "")))
    return jsonify({"message":"资产合同已归档，进入只读状态"})

# ---- 资产合同附件：上传 / 下载 / 删除 ----
@app.route("/api/asset_contracts/<aid>/attachments", methods=["POST"])
@login_required
def api_upload_asset_attachment(aid):
    u = get_cur_user()
    if u["角色"] == "查询用户": return jsonify({"error":"权限不足"}),403
    rid = _asset_id(aid)
    if rid is None: return jsonify({"error":"无效的资产合同ID"}),400
    t = next((a for a in read_assets() if a["_row"] == rid), None)
    if not t: return jsonify({"error":"资产合同不存在"}),404
    if t.get("archived"): return jsonify({"error":"该资产合同已归档，仅可浏览，不可修改"}),403
    atype = (request.form.get("type","") or "").strip()
    if atype not in ASSET_ATTACH_TYPES:
        # 附带“收到的类型 + 当前服务端支持的类型”，便于识别“服务端未重启/代码未更新”的情况
        return jsonify({"error": "附件类型无效：收到 type=\"%s\"，当前服务端支持 %s。若刚新增过附件类型，请重启应用后再试"
                                 % (atype, "、".join(ASSET_ATTACH_TYPES.keys()))}), 400
    f = request.files.get("file")
    if not f or not f.filename: return jsonify({"error":"请选择要上传的文件"}),400
    display = os.path.basename(f.filename.replace("\\","/")).strip()[:200] or "附件"
    ext = os.path.splitext(display)[1].lower()
    if ext not in ATTACH_ALLOWED_EXT:
        return jsonify({"error":"仅支持图片或PDF（jpg/jpeg/png/gif/webp/bmp/pdf）"}),400
    max_mb = get_attach_max_mb()
    max_bytes = max_mb * 1024 * 1024
    store_name = uuid.uuid4().hex + ext
    os.makedirs(ATTACH_DIR, exist_ok=True)
    fp = os.path.join(ATTACH_DIR, store_name)
    f.save(fp)
    size = os.path.getsize(fp)
    if size > max_bytes:
        try: os.remove(fp)
        except Exception: pass
        return jsonify({"error":"附件大小不能超过 %dM（当前 %.2fM）" % (max_mb, size/1048576.0)}),400
    rows = read_asset_attachments()
    rows.append({
        "附件ID": uuid.uuid4().hex,
        "资产合同ID": rid,
        "附件类型": ASSET_ATTACH_TYPES[atype],
        "显示名称": display,
        "存储文件名": store_name,
        "文件大小": str(size),
        "上传人": u.get("姓名", u.get("用户名","")),
        "上传时间": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    })
    write_asset_attachments(rows)
    write_log("资产合同台账", "上传附件", obj=str(t.get("asset_name","")), detail=display)
    return jsonify({"message":"附件上传成功","attachment":asset_attachment_json(rows[-1])}),201

@app.route("/api/asset_attachments/<att_id>/download", methods=["GET"])
@login_required
def api_download_asset_attachment(att_id):
    row = next((r for r in read_asset_attachments() if r.get("附件ID","") == att_id), None)
    if not row: return jsonify({"error":"附件不存在"}),404
    # 数据范围：附件所属资产合同须在当前用户可见部门内（管理员不受限）
    visible = {_int(a["_row"]) or 0 for a in _scope_assets(read_assets())}
    if (_int(row.get("资产合同ID")) or 0) not in visible: return jsonify({"error":"附件不存在"}),404
    fp = os.path.join(ATTACH_DIR, row.get("存储文件名",""))
    if not os.path.exists(fp): return jsonify({"error":"附件文件缺失"}),404
    return send_file(fp, as_attachment=True, download_name=row.get("显示名称","附件"))

@app.route("/api/asset_attachments/<att_id>", methods=["DELETE"])
@login_required
def api_delete_asset_attachment(att_id):
    u = get_cur_user()
    if u["角色"] == "查询用户": return jsonify({"error":"权限不足"}),403
    rows = read_asset_attachments()
    row = next((r for r in rows if r.get("附件ID","") == att_id), None)
    if not row: return jsonify({"error":"附件不存在"}),404
    t = next((a for a in read_assets() if (_int(a["_row"]) or 0) == (_int(row.get("资产合同ID")) or 0)), None)
    if t and t.get("archived"): return jsonify({"error":"该资产合同已归档，仅可浏览，不可修改"}),403
    fp = os.path.join(ATTACH_DIR, row.get("存储文件名",""))
    try:
        if os.path.exists(fp): os.remove(fp)
    except Exception: pass
    write_asset_attachments([r for r in rows if r.get("附件ID","") != att_id])
    write_log("资产合同台账", "删除附件", obj=str(t.get("asset_name","") if t else ""), detail=str(row.get("显示名称","")))
    return jsonify({"message":"附件已删除"})

@app.route("/api/asset_contracts/export", methods=["GET"])
@login_required
def api_export_assets():
    flt = _asset_sort(_asset_filter(_scope_assets(read_assets()), _asset_query()))
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "资产合同台账"
    headers = [cn for _, cn in ASSET_FIELDS]
    for ci, cn in enumerate(headers, 1): ws.cell(row=1, column=ci, value=cn).font = Font(bold=True)
    for ri, a in enumerate(flt, 2):
        for ci, (k, cn) in enumerate(ASSET_FIELDS, 1):
            v = a.get(k, "")
            if isinstance(v, float): v = round(v, 2)
            ws.cell(row=ri, column=ci, value=v)
    bio = BytesIO(); wb.save(bio); bio.seek(0)
    ts = datetime.datetime.now().strftime("%Y%m%d")
    from urllib.parse import quote
    fn = "资产合同台账导出_%s.xlsx" % ts
    write_log("资产合同台账", "导出", obj="资产合同台账", detail="导出 %d 条" % len(flt))
    return Response(bio.getvalue(), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": "attachment; filename*=UTF-8''%s" % quote(fn)})

# ─── 资产合同台账：模板 / 导入（管理员）───
ASSET_TEMPLATE_EXAMPLE = {
    "位置": "市民之家", "序号": 1, "资产名称": "示例资产（导入前请删除此行）", "承租人": "示例承租人",
    "联系电话": "13800000000", "管理部门": "湘潭市机关事务管理局本级-办公室", "责任人": "示例责任人",
    "面积（㎡）": 30, "租期开始日": "2026-01-01", "租期结束日": "2026-12-31",
    "押金（元）": 0, "年租金（元）": 10000, "季租（元）": 2500,
    "备注": "示例行，导入前请删除",
}
ASSET_HEADER_ALIASES = {"面积": "面积（㎡）", "押金(元)": "押金（元）", "季租(元)": "季租（元）",
                        "年租金(元)": "年租金（元）", "年租（元）": "年租金（元）", "年租(元)": "年租金（元）"}

def _asset_header_name(cn):
    import re as _re
    s = _re.sub(r"\s+", "", str(cn if cn is not None else ""))
    return ASSET_HEADER_ALIASES.get(s, s)

def _load_asset_import_file():
    f = request.files.get("file")
    if not f or not f.filename: return None, None, "请选择要导入的文件"
    if not f.filename.lower().endswith((".xlsx", ".xlsm")): return None, None, "仅支持 .xlsx 格式文件"
    try:
        wb = openpyxl.load_workbook(BytesIO(f.read()), data_only=True)
    except Exception:
        return None, None, "文件解析失败，请使用有效的 xlsx 文件"
    sheet = "资产合同台账" if "资产合同台账" in wb.sheetnames else wb.sheetnames[0]
    rows = list(wb[sheet].iter_rows(values_only=True))
    if not rows or len(rows) < 2: return None, None, "文件中没有可导入的数据"
    hdr = [_asset_header_name(x) for x in rows[0]]
    if "资产名称" not in hdr: return None, None, "文件表头缺少列：资产名称，请使用下载的资产合同台账模板"
    return rows, hdr, None

def _asset_row_check(vals):
    errors = []
    for cn in ("位置", "资产名称", "承租人"):
        if not str(vals.get(cn, "")).strip(): errors.append("%s为空（必填）" % cn)
    for k, cn in ASSET_FIELDS:
        v = vals.get(cn, "")
        if v in ("", None): continue
        if k in ASSET_NUM_KEYS and _num(v) is None: errors.append("%s不是有效数字" % cn)
        elif k == "seq" and _int(v) is None: errors.append("序号应为整数")
        elif k == "remind_days" and _int(v) is None: errors.append("%s应为整数" % cn)
        elif k in ASSET_DATE_KEYS and not _asset_norm_date(v): errors.append("%s日期格式应为yyyy-mm-dd" % cn)
    return errors

def _asset_import_items(rows, hdr):
    """解析导入行 → [(excel 行号, 列值字典, 状态, 说明)]"""
    out = []
    for ri in range(1, len(rows)):
        src = rows[ri]
        vals = {cn: "" for _, cn in ASSET_FIELDS}
        for ci, cn in enumerate(hdr):
            if cn in vals and ci < len(src): vals[cn] = _norm(src[ci])
        if not any(str(v).strip() for v in vals.values()): continue
        errors = _asset_row_check(vals)
        out.append((ri + 1, vals, ("error" if errors else "ok"), errors))
    return out

@app.route("/api/asset_contracts/template", methods=["GET"])
@admin_required
def api_asset_template():
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "资产合同台账"
    headers = [cn for _, cn in ASSET_FIELDS]
    for ci, cn in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=cn)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2E74B5")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = 16
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = "A2"
    for ci, cn in enumerate(headers, 1):
        v = ASSET_TEMPLATE_EXAMPLE.get(cn, "")
        if v != "": ws.cell(row=2, column=ci, value=v)
    ws.cell(row=2, column=1).font = Font(italic=True, color="999999")
    ws2 = wb.create_sheet("填写说明")
    tips = [
        ("资产合同台账导入模板填写说明", True),
        ("", False),
        ("1. 第一行为表头，不可删除；第二行为示例行，导入前请删除。", False),
        ("2. 必填字段：位置、资产名称、承租人。", False),
        ("3. 日期列（租期开始日、租期结束日、下次出租时间）格式为 yyyy-mm-dd。", False),
        ("4. 金额列（面积、押金、年租金、季租、总合同额）填写数字。", False),
        ("6. 导入前可先“导入预览”查看逐条校验结果，错误行不会被导入。", False),
        ("7. 管理部门（填“单位名称-科室名称”，与合同台账口径一致）、责任人为选填项，可留空。", False),
        ("8. 租金收取明细（收款期次/应缴金额/实缴金额/缴款日期/缴款方式/收款人）在“编辑资产合同”弹窗内单独维护，不通过本模板导入导出。", False),
    ]
    for i, (txt, bold) in enumerate(tips, 1):
        cell = ws2.cell(row=i, column=1, value=txt)
        if bold: cell.font = Font(bold=True, size=13, color="1F4D78")
    ws2.column_dimensions["A"].width = 90
    bio = BytesIO(); wb.save(bio); bio.seek(0)
    from urllib.parse import quote
    fn = "资产合同台账导入模板.xlsx"
    write_log("资产合同台账", "下载模板", obj="资产合同台账", detail="下载导入模板")
    return Response(bio.getvalue(), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": "attachment; filename*=UTF-8''%s" % quote(fn)})

@app.route("/api/asset_contracts/import/preview", methods=["POST"])
@admin_required
def api_asset_import_preview():
    rows, hdr, err = _load_asset_import_file()
    if err: return jsonify({"error": err}),400
    items = _asset_import_items(rows, hdr)
    ok = sum(1 for it in items if it[2] != "error")
    return jsonify({"items": [{"row": it[0], "place": str(it[1].get("位置", "")),
                               "asset_name": str(it[1].get("资产名称", "")), "tenant": str(it[1].get("承租人", "")),
                               "status": it[2], "messages": it[3]} for it in items],
                    "total": len(items), "ok": ok, "error": len(items) - ok})

@app.route("/api/asset_contracts/import", methods=["POST"])
@admin_required
def api_asset_import():
    rows, hdr, err = _load_asset_import_file()
    if err: return jsonify({"error": err}),400
    items = _asset_import_items(rows, hdr)
    added, skipped = 0, []
    for rno, vals, status, msgs in items:
        if status == "error":
            skipped.append({"row": rno, "asset_name": str(vals.get("资产名称", "")), "errors": msgs}); continue
        d = {k: vals.get(cn, "") for k, cn in ASSET_FIELDS}
        clean, verr = _asset_vals(d, require_all=True)
        if verr:
            skipped.append({"row": rno, "asset_name": str(vals.get("资产名称", "")), "errors": [verr]}); continue
        cols = list(clean.keys())
        sql = "INSERT INTO asset_contracts (" + ",".join('"%s"' % c for c in cols) + ") VALUES (" + ",".join("?" * len(cols)) + ")"
        db_exec(sql, [clean[c] for c in cols])
        added += 1
    msg = "导入完成：新增 %d 条" % added
    if skipped: msg += "，跳过 %d 条" % len(skipped)
    write_log("资产合同台账", "导入", obj=str(os.path.basename((request.files.get("file").filename or ""))), detail=msg)
    return jsonify({"message": msg, "added": added, "skipped": skipped})

# ─── Statistics ───
# ─── 资产合同统计（资产合同查询统计选项卡）───
def _asset_stat_filter(items, unit, dept, person, sd, ed):
    """资产统计筛选：单位/科室按“管理部门”包含匹配，经办人按“责任人”模糊匹配，出租时间取“租期开始日”。"""
    out = []
    for a in items:
        dep = str(a.get("dept", ""))
        if unit and unit not in dep: continue
        if dept and dept not in dep: continue
        if person and person not in str(a.get("person", "")): continue
        d = parse_date(a.get("lease_start", ""))
        if sd:
            s_ = parse_date(sd)
            if s_ and (d is None or d < s_): continue
        if ed:
            e_ = parse_date(ed)
            if e_ and (d is None or d > e_): continue
        out.append(a)
    return out

def _asset_stat_newg():
    return {"count": 0, "area": 0.0, "rent": 0.0, "paid": 0.0}

def _asset_stat_acc(g, a):
    g["count"] += 1
    for key, src in (("area", "area"), ("rent", "annual_rent"), ("paid", "paid_total")):
        v = _num(a.get(src))
        if v: g[key] += v

def _asset_stat_row(dim_label, key, g):
    return {dim_label: key, "资产数": g["count"], "面积合计(㎡)": round(g["area"], 2),
            "年租金合计(元)": round(g["rent"], 2), "已缴租金合计(元)": round(g["paid"], 2)}

@app.route("/api/asset_statistics")
@login_required
def api_asset_statistics():
    st = request.args.get("type", "by_dept")
    items = _asset_stat_filter(_scope_assets(read_assets()),
                               request.args.get("unit", "").strip(),
                               request.args.get("dept", "").strip(),
                               request.args.get("person", "").strip(),
                               request.args.get("start_date", "").strip(),
                               request.args.get("end_date", "").strip())
    if st == "by_dept":
        grp = {}
        for a in items:
            g = grp.setdefault(str(a.get("dept", "") or "未指定"), _asset_stat_newg()); _asset_stat_acc(g, a)
        return jsonify({"type": st, "data": [_asset_stat_row("管理部门", k, v)
                                             for k, v in sorted(grp.items(), key=lambda x: -x[1]["count"])]})
    if st == "by_person":
        grp = {}
        for a in items:
            for p_ in (_split_names(str(a.get("person", ""))) or ["未指定"]):
                g = grp.setdefault(p_, _asset_stat_newg()); _asset_stat_acc(g, a)
        return jsonify({"type": st, "data": [_asset_stat_row("责任人", k, v)
                                             for k, v in sorted(grp.items(), key=lambda x: -x[1]["count"])]})
    if st == "by_time":
        grp = {}
        for a in items:
            d = parse_date(a.get("lease_start", ""))
            if not d: continue
            q = (d.month - 1) // 3 + 1
            g = grp.setdefault("%d年第%d季度" % (d.year, q), _asset_stat_newg()); _asset_stat_acc(g, a)
        return jsonify({"type": st, "data": [_asset_stat_row("时间", k, v) for k, v in sorted(grp.items())]})
    return jsonify({"error": "不支持的统计类型"}), 400

def _stat_filter(cts, unit, dept, person, year, sd, ed):
    out = []
    for c in cts:
        if unit and unit not in c.get("\u6240\u5c5e\u90e8\u95e8",""): continue
        if dept and dept not in c.get("\u6240\u5c5e\u90e8\u95e8",""): continue
        if person and person not in c.get("\u7ecf\u529e\u4eba",""): continue
        cd = parse_date(c.get("\u91c7\u8d2d\u65f6\u95f4",""))
        if year and (cd is None or str(cd.year)!=year): continue
        if sd:
            s_ = parse_date(sd)
            if s_ and (cd is None or cd < s_): continue
        if ed:
            e_ = parse_date(ed)
            if e_ and (cd is None or cd > e_): continue
        out.append(c)
    return out


@app.route("/api/statistics")
@login_required
def api_statistics():
    st = request.args.get("type","by_dept")
    dim = request.args.get("dim","dept")
    # 维度拆分：按部门/项目支付情况、按部门/项目预算执行情况
    if st in ("pay_dept","pay_project","budget_dept","budget_project"):
        dim = "project" if st.endswith("project") else "dept"
        st = "unpaid" if st.startswith("pay_") else "budget"
    cts = _stat_filter(_scope_contracts(read_contracts()),
                       request.args.get("unit","").strip(),
                       request.args.get("dept","").strip(),
                       request.args.get("person","").strip(),
                       request.args.get("year","").strip(),
                       request.args.get("start_date","").strip(),
                       request.args.get("end_date","").strip())
    dim_label = "\u9879\u76ee" if dim=="project" else "\u90e8\u95e8"
    def _gkey(c, dim):
        if dim=="project":
            return c.get("\u9879\u76ee\u540d\u79f0","") or "\u672a\u77e5"
        return c.get("\u6240\u5c5e\u90e8\u95e8","") or "\u672a\u77e5"
    if st in ("unpaid","budget"):
        grp={}
        for c in cts:
            k=_gkey(c, dim)
            if k not in grp: grp[k]={"count":0,"amt":0.0,"paid":0.0,"unpaid":0.0,"bid":0.0}
            g=grp[k]; g["count"]+=1
            a=_num(c.get("\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09")) or 0.0
            g["amt"]+=a
            p_=_num(c.get("\u5df2\u652f\u4ed8")) or 0.0
            g["paid"]+=p_
            up=_num(c.get("\u672a\u652f\u4ed8"))
            if up is None:
                up = (a - p_) if (a or p_) else 0.0
            g["unpaid"]+=up
            b=_num(c.get("\u51b3\u7b56\u7684\u62db\u6807\u63a7\u5236\u4ef7")) or 0.0
            g["bid"]+=b
        rows=[]
        for k,v in sorted(grp.items(), key=lambda x:-x[1]["count"]):
            if st=="unpaid":
                rows.append({dim_label:k,"\u5408\u540c\u6570":v["count"],"\u5408\u540c\u91d1\u989d":round(v["amt"],2),"\u5df2\u652f\u4ed8":round(v["paid"],2),"\u672a\u652f\u4ed8":round(v["unpaid"],2)})
            else:
                diff=v["amt"]-v["bid"]
                rate=round(diff/v["bid"]*100,2) if v["bid"] else None
                rows.append({dim_label:k,"\u5408\u540c\u6570":v["count"],"\u62db\u6807\u63a7\u5236\u4ef7":round(v["bid"],2),"\u5408\u540c\u91d1\u989d":round(v["amt"],2),"\u5dee\u989d":round(diff,2),"\u504f\u5dee\u7387":rate})
        return jsonify({"type":st,"dim":dim,"data":rows})
    if st=="by_dept":
        grp={}
        for c in cts:
            d=c.get("\u6240\u5c5e\u90e8\u95e8","\u672a\u77e5")
            if d not in grp: grp[d]={"count":0,"amt":0.0}
            grp[d]["count"]+=1
            a=_num(c.get("\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09"))
            if a: grp[d]["amt"]+=a
        return jsonify({"type":"by_dept","data":[{"\u90e8\u95e8":k,"\u5408\u540c\u6570":v["count"],"\u5408\u540c\u91d1\u989d":round(v["amt"],2)} for k,v in sorted(grp.items(),key=lambda x:-x[1]["count"])]})
    elif st=="by_category":
        grp={}
        for c in cts:
            cat=c.get("\u9879\u76ee\u5206\u7c7b","\u672a\u77e5")
            if cat not in grp: grp[cat]={"count":0,"amt":0.0}
            grp[cat]["count"]+=1
            a=_num(c.get("\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09"))
            if a: grp[cat]["amt"]+=a
        return jsonify({"type":"by_category","data":[{"\u5206\u7c7b":k,"\u5408\u540c\u6570":v["count"],"\u5408\u540c\u91d1\u989d":round(v["amt"],2)} for k,v in sorted(grp.items(),key=lambda x:-x[1]["count"])]})
    elif st=="by_fund_source":
        grp={}
        for c in cts:
            fs=c.get("资金来源","") or "未指定"
            if fs not in grp: grp[fs]={"count":0,"amt":0.0}
            grp[fs]["count"]+=1
            a=_num(c.get("合同金额（万元）"))
            if a: grp[fs]["amt"]+=a
        return jsonify({"type":"by_fund_source","data":[{"资金来源":k,"合同数":v["count"],"合同金额":round(v["amt"],2)} for k,v in sorted(grp.items(),key=lambda x:-x[1]["count"])]})
    elif st=="by_person":
        grp={}
        for c in cts:
            for p_ in _split_names(c.get("\u7ecf\u529e\u4eba","")):
                if p_ not in grp: grp[p_]={"count":0,"amt":0.0}
                grp[p_]["count"]+=1
                a=_num(c.get("\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09"))
                if a: grp[p_]["amt"]+=a
        return jsonify({"type":"by_person","data":[{"\u7ecf\u529e\u4eba":k,"\u5408\u540c\u6570":v["count"],"\u5408\u540c\u91d1\u989d":round(v["amt"],2)} for k,v in sorted(grp.items(),key=lambda x:-x[1]["count"])]})
    elif st=="by_time":
        grp={}
        for c in cts:
            cd=parse_date(c.get("\u91c7\u8d2d\u65f6\u95f4",""))
            if not cd: continue
            q=(cd.month-1)//3+1; k=f"{cd.year}\u5e74\u7b2c{q}\u5b63\u5ea6"
            if k not in grp: grp[k]={"count":0,"amt":0.0}
            grp[k]["count"]+=1
            a=_num(c.get("\u5408\u540c\u91d1\u989d\uff08\u4e07\u5143\uff09"))
            if a: grp[k]["amt"]+=a
        return jsonify({"type":"by_time","data":[{"\u65f6\u95f4":k,"\u5408\u540c\u6570":v["count"],"\u5408\u540c\u91d1\u989d":round(v["amt"],2)} for k,v in sorted(grp.items())]})
    return jsonify({"error":"\u4e0d\u652f\u6301\u7684\u7edf\u8ba1\u7c7b\u578b"}),400

# ─── User Management ───
@app.route("/api/users",methods=["GET"])
@login_required
def api_list_users():
    u = get_cur_user()
    if u["\u89d2\u8272"] != "\u7ba1\u7406\u5458":
        return jsonify({"items":[{"username":u["\u7528\u6237\u540d"],"name":u["\u59d3\u540d"],"department":u["\u6240\u5c5e\u90e8\u95e8"],"role":u["\u89d2\u8272"],"managed_departments":u.get("\u5206\u7ba1\u90e8\u95e8",""),"phone":u.get("\u7535\u8bdd","")}]})
    return jsonify({"items":[{"username":x["\u7528\u6237\u540d"],"name":x["\u59d3\u540d"],"department":x["\u6240\u5c5e\u90e8\u95e8"],"role":x["\u89d2\u8272"],"managed_departments":x.get("\u5206\u7ba1\u90e8\u95e8",""),"phone":x.get("\u7535\u8bdd","")} for x in get_all_users()]})

@app.route("/api/users",methods=["POST"])
@admin_required
def api_create_user():
    d=request.get_json(silent=True) or {}
    un=d.get("username","").strip(); nm=d.get("name","").strip(); dp=d.get("department","").strip(); rl=d.get("role","普通用户").strip()
    if not un or not nm or not dp: return jsonify({"error":"用户名、姓名、所属部门不能为空"}),400
    if rl not in ("管理员","普通用户","查询用户"): return jsonify({"error":"角色无效"}),400
    if get_user(un): return jsonify({"error":"用户名已存在"}),400
    if dp not in get_dept_strings(): return jsonify({"error":"部门不存在，请先添加部门"}),400
    managed = ""
    try:
        phone = norm_phone(d.get("phone",""))
    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    if is_leadership_dept(dp) and d.get("managed_departments"):
        try:
            managed = norm_managed(d["managed_departments"])
        except ValueError as e:
            return jsonify({"error": str(e)}), 400
    db_exec('INSERT INTO users ("用户名","密码","姓名","所属部门","角色","分管部门","电话") VALUES (?,?,?,?,?,?,?)',
            (un, hash_pw("123456"), nm, dp, rl, managed, phone))
    write_log("用户管理", "新增用户", obj=un, detail=nm + " / " + dp)
    return jsonify({"message":"用户创建成功，默认密码 123456"}),201

@app.route("/api/users/<username>",methods=["PUT"])
@admin_required
def api_update_user(username):
    d=request.get_json(silent=True) or {}
    cur_user = get_user(username)
    if not cur_user: return jsonify({"error":"用户不存在"}),404
    final_dept = d["department"].strip() if (d.get("department") and d["department"].strip()) else cur_user.get("所属部门","")
    sets=[]; params=[]
    if "name" in d and d["name"]:
        sets.append('"姓名"=?'); params.append(d["name"].strip())
    if "department" in d and d["department"]:
        if d["department"].strip() not in get_dept_strings(): return jsonify({"error":"部门不存在"}),400
        sets.append('"所属部门"=?'); params.append(d["department"].strip())
    if "role" in d and d["role"]:
        if d["role"] not in ("管理员","普通用户","查询用户"): return jsonify({"error":"角色无效"}),400
        sets.append('"角色"=?'); params.append(d["role"])
    if "phone" in d:
        try:
            phone = norm_phone(d["phone"])
        except ValueError as e:
            return jsonify({"error": str(e)}), 400
        sets.append('"电话"=?'); params.append(phone)
    if "reset_password" in d and d["reset_password"]:
        sets.append('"密码"=?'); params.append(hash_pw("123456"))
    if "managed_departments" in d:
        if is_leadership_dept(final_dept):
            try:
                managed = norm_managed(d["managed_departments"])
            except ValueError as e:
                return jsonify({"error": str(e)}), 400
            sets.append('"分管部门"=?'); params.append(managed)
        else:
            sets.append('"分管部门"=?'); params.append("")
    if sets:
        params.append(username)
        db_exec("UPDATE users SET " + ",".join(sets) + ' WHERE "用户名"=?', params)
    if "reset_password" in d and d["reset_password"]:
        write_log("用户管理", "重置密码", obj=username)
    write_log("用户管理", "修改用户", obj=username)
    return jsonify({"message":"用户更新成功"})

@app.route("/api/users/<username>/self",methods=["PUT"])
@login_required
def api_update_self(username):
    u = get_cur_user()
    if not u or u["\u7528\u6237\u540d"] != username:
        return jsonify({"error":"\u6743\u9650\u4e0d\u8db3"}),403
    cur = get_user(username)
    if not cur: return jsonify({"error":"\u7528\u6237\u4e0d\u5b58\u5728"}),404
    d = request.get_json(silent=True) or {}
    new_un = d.get("new_username","").strip()
    old_pw = str(d.get("old_password","") or "")
    new_pw = str(d.get("new_password","") or "")
    if new_pw and not old_pw:
        return jsonify({"error":"\u4fee\u6539\u5bc6\u7801\u9700\u586b\u5199\u539f\u5bc6\u7801"}),400
    if old_pw:
        stored = cur.get("\u5bc6\u7801\u54c8\u5e0c", cur.get("\u5bc6\u7801",""))
        if stored != hash_pw(old_pw):
            return jsonify({"error":"\u539f\u5bc6\u7801\u4e0d\u6b63\u786e"}),400
    if not new_un and not new_pw:
        return jsonify({"error":"\u8bf7\u586b\u5199\u8981\u4fee\u6539\u7684\u7528\u6237\u540d\u6216\u5bc6\u7801"}),400
    if new_un and new_un != username:
        if get_user(new_un):
            return jsonify({"error":"\u7528\u6237\u540d\u5df2\u5b58\u5728"}),400
        db_exec('UPDATE users SET "用户名"=? WHERE "用户名"=?', (new_un, username))
        username = new_un
    if new_pw:
        db_exec('UPDATE users SET "密码"=? WHERE "用户名"=?', (hash_pw(new_pw), username))
    session["username"] = username
    write_log("用户管理", "修改个人信息", obj=username, detail="用户名或密码变更")
    return jsonify({"message":"\u4fee\u6539\u6210\u529f","username":username})

@app.route("/api/users/<username>",methods=["DELETE"])
@admin_required
def api_delete_user(username):
    if username=="admin": return jsonify({"error":"不能删除默认管理员"}),400
    if not get_user(username): return jsonify({"error":"用户不存在"}),404
    db_exec('DELETE FROM users WHERE "用户名"=?', (username,))
    write_log("用户管理", "删除用户", obj=username)
    return jsonify({"message":"用户删除成功"})

# ─── Department Management ───
@app.route("/api/departments",methods=["GET"])
@login_required
def api_list_depts():
    """部门列表接口：返回 unit / dept / display 三项（供前端下拉与展示使用）。

    兼容性：display 由 _dept_display() 生成，不在 f-string 表达式内书写 Unicode 转义（形如 \\uXXXX），
            因此 Python 3.8（服务器）与 3.12（开发机）均可运行；
            返回结构与字段名（unit/dept/display）与改造前完全一致，前端无需改动。
    """
    # 原实现（仅 Python 3.12+ 可用：display 的 f-string 表达式内含 \uXXXX 反斜杠，3.11 及以前报 SyntaxError）
    # 为便于回溯保留在此，禁止启用：
    # return jsonify({"items":[{"unit":d["\u5355\u4f4d\u540d\u79f0"],"dept":d["\u79d1\u5ba4\u540d\u79f0"],"display":f"{d['\u5355\u4f4d\u540d\u79f0']}-{d['\u79d1\u5ba4\u540d\u79f0']}"} for d in get_all_depts()]})
    items = []
    for d in get_all_depts():
        items.append({
            "unit": d["\u5355\u4f4d\u540d\u79f0"],   # 单位名称
            "dept": d["\u79d1\u5ba4\u540d\u79f0"],   # 科室名称
            "display": _dept_display(d),              # 「单位名称-科室名称」（显示用）
        })
    return jsonify({"items": items})

@app.route("/api/departments",methods=["POST"])
@admin_required
def api_create_dept():
    d=request.get_json(silent=True) or {}; un=d.get("unit","").strip(); dn=d.get("dept","").strip()
    if not un or not dn: return jsonify({"error":"单位名称和科室名称不能为空"}),400
    for x in get_all_depts():
        if x["单位名称"]==un and x["科室名称"]==dn: return jsonify({"error":"该部门组合已存在"}),400
    db_exec('INSERT INTO departments ("单位名称","科室名称") VALUES (?,?)', (un,dn))
    write_log("部门管理", "新增部门", obj=un + "-" + dn)
    return jsonify({"message":"部门创建成功"}),201

@app.route("/api/departments",methods=["DELETE"])
@admin_required
def api_delete_dept():
    un=request.args.get("unit","").strip(); dn=request.args.get("dept","").strip()
    if not un or not dn: return jsonify({"error":"单位名称和科室名称不能为空"}),400
    ds=f"{un}-{dn}"
    for u in get_all_users():
        if u.get("所属部门","")==ds: return jsonify({"error":f"部门 '{ds}' 下仍有用户，无法删除"}),400
    depts=get_all_depts()
    if not any(x["单位名称"]==un and x["科室名称"]==dn for x in depts): return jsonify({"error":"部门不存在"}),404
    db_exec('DELETE FROM departments WHERE "单位名称"=? AND "科室名称"=?', (un,dn))
    write_log("部门管理", "删除部门", obj=un + "-" + dn)
    return jsonify({"message":"部门删除成功"})

@app.route("/api/departments",methods=["PUT"])
@admin_required
def api_update_dept():
    d=request.get_json(silent=True) or {}
    un=d.get("unit","").strip(); dn=d.get("dept","").strip()
    n_un=d.get("new_unit","").strip(); n_dn=d.get("new_dept","").strip()
    if not un or not dn or not n_un or not n_dn:
        return jsonify({"error":"单位名称和部门名称不能为空"}),400
    old=f"{un}-{dn}"; new=f"{n_un}-{n_dn}"
    depts=get_all_depts()
    if not any(x["单位名称"]==un and x["科室名称"]==dn for x in depts):
        return jsonify({"error":"部门不存在"}),404
    if old!=new and any(x["单位名称"]==n_un and x["科室名称"]==n_dn for x in depts):
        return jsonify({"error":"该部门组合已存在"}),400
    db_exec('UPDATE departments SET "单位名称"=?, "科室名称"=? WHERE "单位名称"=? AND "科室名称"=?',(n_un,n_dn,un,dn))
    if old!=new:
        db_exec('UPDATE users SET "所属部门"=? WHERE "所属部门"=?',(new,old))
        for u in get_all_users():
            mg=str(u.get("分管部门") or "").strip()
            if not mg: continue
            parts=[p.strip() for p in mg.replace("、",",").replace("；",";").split(",") if p.strip()]
            if old in parts:
                parts=[new if p==old else p for p in parts]
                db_exec('UPDATE users SET "分管部门"=? WHERE "用户名"=?',("、".join(parts),u["用户名"]))
        db_exec('UPDATE project_contracts SET "所属部门"=? WHERE "所属部门"=?',(new,old))
    write_log("部门管理","编辑部门",obj=old+" → "+new)
    return jsonify({"message":"部门修改成功"})

# ─── Config ───
@app.route("/api/config",methods=["GET"])
@login_required
def api_get_config():
    cfg = _sms_cfg()
    return jsonify({"default_remind_days":get_default_days(),"log_keep_days":get_log_keep_days(),
                    "default_attach_max_mb":get_attach_max_mb(),
                    "sms":{"enabled":cfg["enabled"],"base_url":cfg["base_url"],"access_key":cfg["ak"],
                           "weekly_enabled":cfg["weekly_enabled"],"weekday":cfg["weekday"],"hour":cfg["hour"],
                           "minute":cfg["minute"],"max_per_run":cfg["max_per_run"],"retry_max":cfg["retry_max"],
                           "receipt_token":cfg["receipt_token"],"timeout":cfg["timeout"],
                           "include_sign":cfg["include_sign"],"segment_len":cfg["segment_len"],
                           "sign_suffix":cfg["sign_suffix"],
                           "max_segments":cfg["max_segments"],"block_check":cfg["block_check"],
                           "block_words":",".join(cfg["block_words"]),
                           "daily_per_phone":cfg["daily_per_phone"],"min_interval_sec":cfg["min_interval_sec"],
                           "last_run_at":get_cfg("sms_last_run_at") or ""}})

@app.route("/api/config",methods=["PUT"])
@admin_required
def api_update_config():
    d=request.get_json(silent=True) or {}
    updated=[]
    if "default_remind_days" in d:
        try:
            v=int(d["default_remind_days"])
            if v<1 or v>365: return jsonify({"error":"\u63d0\u9192\u5929\u6570\u5e94\u57281-365\u4e4b\u95f4"}),400
            set_cfg("default_remind_days",v)
            write_log("系统配置", "修改配置", obj="default_remind_days", detail=str(v))
            updated.append("default_remind_days")
        except: return jsonify({"error":"\u65e0\u6548\u7684\u6570\u503c"}),400
    if "log_keep_days" in d:
        try:
            v=int(d["log_keep_days"])
            if v<1 or v>3650: return jsonify({"error":"\u65e5\u5fd7\u4fdd\u7559\u5929\u6570\u5e94\u57281-3650\u4e4b\u95f4"}),400
            set_cfg("log_keep_days",v)
            write_log("系统配置", "修改配置", obj="log_keep_days", detail=str(v))
            updated.append("log_keep_days")
        except: return jsonify({"error":"\u65e0\u6548\u7684\u6570\u503c"}),400
    if "default_attach_max_mb" in d:
        try:
            v=int(d["default_attach_max_mb"])
            if v<1 or v>200: return jsonify({"error":"附件大小限制应在1-200(M)之间"}),400
            set_cfg("default_attach_max_mb",v)
            write_log("系统配置", "修改配置", obj="default_attach_max_mb", detail=str(v))
            updated.append("default_attach_max_mb")
        except: return jsonify({"error":"无效的数值"}),400
    # ── 短信平台配置（v20 起）──
    if "sms_sign_suffix" in d:
        _sfx = str(d.get("sms_sign_suffix") or "")
        if "【" in _sfx or "】" in _sfx:
            return jsonify({"error": "署名文本不能包含“【】”中括号（平台会判定多签名并阻断发送），请使用全角括号（）"}), 400
    for k in ("sms_base_url","sms_access_key","sms_receipt_token","sms_block_words","sms_sign_suffix"):
        if k in d:
            set_cfg(k, str(d[k] if d[k] is not None else "").strip()); updated.append(k)
    if str(d.get("sms_secret_key","")).strip() not in ("", "******"):
        set_cfg("sms_secret_key", str(d["sms_secret_key"]).strip()); updated.append("sms_secret_key")
    for k in ("sms_enabled","sms_weekly_enabled","sms_include_sign","sms_block_check"):
        if k in d:
            set_cfg(k, "1" if str(d[k]).strip().lower() in ("1","true","yes","on","是","开启") else "0")
            updated.append(k)
    for k, lo, hi in (("sms_weekly_weekday",1,7),("sms_weekly_hour",0,23),("sms_weekly_minute",0,59),
                      ("sms_max_per_run",1,1000),("sms_retry_max",0,5),("sms_timeout_sec",1,120),
                      ("sms_segment_len",20,500),("sms_max_segments",1,10),
                      ("sms_daily_per_phone",0,100),("sms_min_interval_sec",0,3600)):
        if k in d:
            try:
                v = int(float(d[k]))
            except Exception:
                return jsonify({"error": "%s 应为数字" % k}), 400
            if v < lo or v > hi:
                return jsonify({"error": "%s 应在 %d-%d 之间" % (k, lo, hi)}), 400
            set_cfg(k, v); updated.append(k)
    if any(str(x).startswith("sms_") for x in updated):
        write_log("系统配置", "修改短信配置", obj="sms",
                  detail=",".join(str(x) for x in updated if str(x).startswith("sms_")))
    if not updated: return jsonify({"error":"\u7f3a\u5c11\u914d\u7f6e\u9879"}),400
    return jsonify({"message":"\u914d\u7f6e\u66f4\u65b0\u6210\u529f","default_remind_days":get_default_days(),"log_keep_days":get_log_keep_days(),"default_attach_max_mb":get_attach_max_mb()})

# ─── 操作日志（审计） ───
def _log_filter():
    # 分级查看：管理员可看全量；普通/查询用户仅看本人
    u = get_cur_user()
    a = request.args
    start_date = a.get("start_date","").strip()
    end_date = a.get("end_date","").strip()
    username = a.get("username","").strip()
    module = a.get("module","").strip()
    action = a.get("action","").strip()
    conds=[]; params=[]
    if u and u["\u89d2\u8272"] != "\u7ba1\u7406\u5458":
        conds.append('"用户名" = ?'); params.append(u["\u7528\u6237\u540d"])
    elif username:
        conds.append('"用户名" LIKE ?'); params.append("%" + username + "%")
    if start_date: conds.append('"时间" >= ?'); params.append(start_date + " 00:00:00")
    if end_date: conds.append('"时间" <= ?'); params.append(end_date + " 23:59:59")
    if module: conds.append('"模块" = ?'); params.append(module)
    if action: conds.append('"操作类型" = ?'); params.append(action)
    where = (" WHERE " + " AND ".join(conds)) if conds else ""
    return where, params

@app.route("/api/logs")
@login_required
def api_list_logs():
    where, params = _log_filter()
    a = request.args
    try:
        page = max(1, int(a.get("page", 1))); ps = min(200, max(1, int(a.get("page_size", 20))))
    except Exception:
        page = 1; ps = 20
    total = db_query('SELECT COUNT(*) c FROM operation_logs' + where, params)[0]["c"]
    rows = db_query('SELECT * FROM operation_logs' + where + ' ORDER BY id DESC LIMIT ? OFFSET ?',
                    params + [ps, (page-1)*ps])
    return jsonify({"total": total, "page": page, "page_size": ps, "items": rows})

def _field_log_filter():
    """字段级变更日志筛选：管理员看全量；普通/查询用户仅看本人操作；支持 编号/名称、模块、字段、时间范围。"""
    u = get_cur_user()
    a = request.args
    conds, params = [], []
    if u and u["角色"] != "管理员":
        conds.append('"用户名" = ?'); params.append(u["用户名"])
    else:
        username = a.get("username", "").strip()
        if username:
            conds.append('("用户名" LIKE ? OR "姓名" LIKE ?)')
            params += ["%" + username + "%", "%" + username + "%"]
    obj = a.get("obj", "").strip()
    if obj:
        conds.append('("对象编号" LIKE ? OR "对象名称" LIKE ?)')
        params += ["%" + obj + "%", "%" + obj + "%"]
    module = a.get("module", "").strip()
    if module:
        conds.append('"模块" = ?'); params.append(module)
    field = a.get("field", "").strip()
    if field:
        conds.append('"字段" LIKE ?'); params.append("%" + field + "%")
    start_date = a.get("start_date", "").strip()
    end_date = a.get("end_date", "").strip()
    if start_date:
        conds.append('"时间" >= ?'); params.append(start_date + " 00:00:00")
    if end_date:
        conds.append('"时间" <= ?'); params.append(end_date + " 23:59:59")
    where = (" WHERE " + " AND ".join(conds)) if conds else ""
    return where, params

@app.route("/api/field_logs")
@login_required
def api_field_logs():
    where, params = _field_log_filter()
    a = request.args
    try:
        page = max(1, int(a.get("page", 1))); ps = min(500, max(1, int(a.get("page_size", 20))))
    except Exception:
        page, ps = 1, 20
    total = db_query('SELECT COUNT(*) c FROM field_change_logs' + where, params)[0]["c"]
    rows = db_query('SELECT * FROM field_change_logs' + where + ' ORDER BY id DESC LIMIT ? OFFSET ?',
                    params + [ps, (page - 1) * ps])
    fields = [r["字段"] for r in db_query('SELECT DISTINCT "字段" FROM field_change_logs ORDER BY "字段"')]
    return jsonify({"total": total, "page": page, "page_size": ps, "items": rows, "fields": fields})

@app.route("/api/logs/export")
@admin_required
def api_export_logs():
    where, params = _log_filter()
    rows = db_query('SELECT * FROM operation_logs' + where + ' ORDER BY id DESC', params)
    headers = ["时间","用户名","姓名","角色","IP地址","模块","操作类型","对象","详情","结果"]
    header_keys = {"IP地址": "IP"}
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "操作日志"
    for ci, h in enumerate(headers, 1):
        ws.cell(row=1, column=ci, value=h).font = Font(bold=True)
    for ri, r in enumerate(rows, 2):
        for ci, h in enumerate(headers, 1):
            ws.cell(row=ri, column=ci, value=r.get(header_keys.get(h, h), ""))
    bio = BytesIO(); wb.save(bio); bio.seek(0)
    ts = datetime.datetime.now().strftime("%Y%m%d")
    from urllib.parse import quote
    fn = f"操作日志_export_{ts}.xlsx"
    write_log("日志审计", "导出", obj="操作日志", detail="导出 " + str(len(rows)) + " 条")
    return Response(bio.getvalue(), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(fn)}"})

# âââ Root Route âââ
@app.route("/")
def index():
    idx_path = os.path.join(BASE_DIR, "templates", "index.html")
    if os.path.exists(idx_path):
        return send_from_directory(os.path.join(BASE_DIR, "templates"), "index.html")
    html = '<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8"><title>潭合同 台账管理系统</title><style>body{font-family:sans-serif;display:flex;justify-content:center;align-items:center;min-height:100vh;margin:0;background:linear-gradient(135deg,#0A2E5C,#1E90FF);color:#fff}.card{text-align:center;padding:40px;background:rgba(255,255,255,0.1);border-radius:16px;backdrop-filter:blur(10px)}h1{font-size:28px;margin:0 0 8px}p{font-size:14px;opacity:.8}</style></head><body><div class="card"><h1>潭合同</h1><p>台账管理系统</p><p>湘潭市机关事务管理局</p><hr style="width:40px;margin:16px auto;border-color:rgba(255,255,255,.3)"><p style="font-size:12px;opacity:.6">请将前端 index.html 放置于 templates/ 目录</p></div></body></html>'
    return html

# âââ Error Handlers âââ
@app.errorhandler(404)
def not_found(e):
    if request.path.startswith("/api/"):
        return jsonify({"error":"\u63a5\u53e3\u4e0d\u5b58\u5728"}),404
    return index()

@app.errorhandler(500)
def server_error(e): return jsonify({"error":"\u670d\u52a1\u5668\u5185\u90e8\u9519\u8bef"}),500


# ═══════════════════════════════════════════════════════════════════════
init_db()
_maybe_cleanup_logs()
_start_sms_scheduler()

if __name__=="__main__":
    lp=os.path.join(STATIC_DIR,"logo.jpg")
    if not os.path.exists(lp): print("[WARN] static/logo.jpg 不存在，请将单位LOGO文件放置于此路径。")
    print("[INFO] 潭合同 台账管理系统 启动中...")
    print(f"[INFO] 数据库: {DB_PATH}")
    print("[INFO] 访问地址: http://localhost:5000")
    print("[INFO] 默认管理员: admin / 123456")
    import webbrowser; webbrowser.open("http://localhost:5000")
    app.run(host="0.0.0.0",port=5000,debug=False,threaded=True)
