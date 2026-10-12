# 资产合同台账表（asset_contracts）说明

> 目标：在现有“潭合同”台账管理系统中新增**资产合同台账**数据表 `asset_contracts`。
> 建表依据：《资产合同台账.xlsx》（湘潭市国有资产租金缴纳情况明细表，2026 年度，`Sheet3`）；
> 设计参照：现有合同主表 `project_contracts`（原 `contracts`，2026-09-20 改名）的命名与约束约定（中文列名、`id` 自增主键、金额 `REAL CHECK(>=0)`、日期 `TEXT CHECK(yyyy-mm-dd)`）。
> 范围：仅新增 1 张表 + 1 个版本化迁移 + 1 个一次性导入脚本，**其余表结构、接口与页面均未改动**。
> **更新（V1.1，2026-09-23）**：本文第 1~6 节记录建表时的 v4 变更与导入过程；第 7 节补充后续 v6~v10 追加的管理部门/责任人、归档列、下次出租时间、总合同额、提醒提前天数，以及新增的 `asset_attachments` 附件表。
> **更新（V1.2，2026-09-29，租金收缴改造）**：`asset_contracts` **当前为 21 列（id + 20 个业务列）**：v13 删除原固定季度列（已缴租金1~4季度（元）/1~4季度交款时间/已缴租金开票对应租期/已缴租金合计（元）共 10 列，采用官方重建表流程，其余列数据不丢）、v14 将 `年租（元）` 改名为 `年租金（元）`、v15 删除合同级 `缴款方式`（改由子表按收款记录维护）；同一批次新增 **`asset_rent_collections` 租金收缴子表（v12）** 与 4 个统计/提醒索引（v18）。当前索引共 **7 个**（原 6 个 + `idx_asset_contracts_next_lease_date`）。最新字段字典见《数据字典表》第 10、14 节，页面说明见《资产合同台账页面说明.md》。

---

## 1. 变更内容

| 项 | 内容 |
|---|---|
| 数据库表 | `data/合同台账.db` 新增 `asset_contracts`（24 列，含 `id`） |
| 代码 | `app.py` 新增 `ASSET_CONTRACTS_DDL` / `ASSET_CONTRACTS_INDEXES` / `_ensure_asset_contracts()`；`_SCHEMA` 追加该表 DDL（全新库直接建表） |
| 迁移 | `MIGRATIONS` 追加 v4：`_mig_v4_asset_contracts`（幂等建表 + 建索引），启动时自动执行并记入 `schema_version` |
| 索引 | `idx_asset_contracts_location`（位置）、`idx_asset_contracts_tenant`（承租人）、`idx_asset_contracts_end_date`（租期结束日） |
| 数据 | 由 `_import_asset_contracts.py` 导入源表 5 条记录（2026 年度） |

## 2. 字段对应关系（源表列 → asset_contracts 列）

源表为**两行合并表头**（“已缴租金”跨列，下挂 1~4 季度金额与交款时间），下表按合并关系展开。

| 源表列 | asset_contracts 列 | 类型 / 约束 |
|---|---|---|
| — | `id` | INTEGER PRIMARY KEY AUTOINCREMENT |
| A 位置 | 位置 | TEXT NOT NULL |
| B 序号 | 序号 | INTEGER |
| C 资产名称 | 资产名称 | TEXT NOT NULL |
| D 承租人 | 承租人 | TEXT NOT NULL |
| E 联系电话 | 联系电话 | TEXT |
| F 面积 | 面积（㎡） | REAL CHECK ≥ 0 |
| G 租期合同起止日（起） | 租期开始日 | TEXT CHECK `yyyy-mm-dd` |
| H 租期合同起止日（止） | 租期结束日 | TEXT CHECK `yyyy-mm-dd` |
| I 押金(元) | 押金（元） | REAL CHECK ≥ 0 |
| J 年租(元) | 年租（元） | REAL CHECK ≥ 0 |
| K 季租(元) | 季租（元） | REAL CHECK ≥ 0 |
| （原表“已缴租金”区域填写“年交”等频次） | 缴款方式 | TEXT |
| L 已缴租金 / 1季度（元） | 已缴租金1季度（元） | REAL CHECK ≥ 0 |
| M 已缴租金 / 交款时间 | 已缴租金1季度交款时间 | TEXT CHECK `yyyy-mm-dd` |
| O 已缴租金 / 2季度（元） | 已缴租金2季度（元） | REAL CHECK ≥ 0 |
| P 已缴租金 / 交款时间 | 已缴租金2季度交款时间 | TEXT CHECK `yyyy-mm-dd` |
| Q 已缴租金 / 3季度（元） | 已缴租金3季度（元） | REAL CHECK ≥ 0 |
| R 已缴租金 / 交款时间 | 已缴租金3季度交款时间 | TEXT CHECK `yyyy-mm-dd` |
| S 已缴租金 / 4季度（元） | 已缴租金4季度（元） | REAL CHECK ≥ 0 |
| T 已缴租金 / 交款时间 | 已缴租金4季度交款时间 | TEXT CHECK `yyyy-mm-dd` |
| N 已缴租金 / 开票对应租期 | 已缴租金开票对应租期 | TEXT |
| U 已缴租金合计 | 已缴租金合计（元） | REAL CHECK ≥ 0 |
| V 备注 | 备注 | TEXT |

## 3. 数据导入与规范化

命令（在项目目录下执行）：

```bash
python _import_asset_contracts.py            # 首次导入；源文件自动查找 资产合同台账.xlsx / 资产合同表.xlsx
python _import_asset_contracts.py --force    # 清空后重导
python _import_asset_contracts.py D:\路径\资产合同台账.xlsx
```

规范化规则（只统一格式，不做业务推断）：
- 面积：`29.4㎡` → `29.4`（单位体现在列名）；
- 日期：`2021.11.15` / `2022.1.31` / `2026.11.14.` → `2021-11-15` / `2022-01-31` / `2026-11-14`（与 project_contracts 日期列一致）；
- 金额：`*`、空值 → NULL；
- 缴款频次：原表“已缴租金”区域填写的“年交”归入 `缴款方式`，季度金额列只存数字；
- 文本：去首尾空白、内部连续空白合并为一个空格。

已导入 5 条（位置均为“市民之家”，承租人：农业银行 / 长沙银行 / 建设银行湘潭市分行 / 湖南银行 / 华龙证券），
并在 `operation_logs` 写入“资产合同台账 / 导入”审计记录（每次执行导入写 1 条）。

## 4. 与合同主表 project_contracts 的差异（有意为之）

- 无 `编号` 及唯一索引：源表未提供合同编号，资产合同以 `id` 唯一标识（页面“序号”列即取 `id`）；避免臆造业务值。
- 部门归属列名为 **`管理部门`**（而非 `所属部门`），且**没有** `所属单位/所属科室` 派生列与外键/触发器（v6 追加的 `管理部门` 为普通 TEXT 列，数据权限过滤直接按该列做字符串匹配）。

## 5. 验收结果（已实测）

1. 表结构：`asset_contracts` 24 列，列名/类型/约束与上表一致；3 个索引建成。
2. 启动迁移：首次启动输出 `[MIGRATE] schema_version 4: 新增资产合同台账表 asset_contracts…`，`schema_version` 记录 v1~v4；二次启动无重复迁移。
3. 约束生效：写入非法日期（`2021.11.15`）、负金额、`资产名称` 为空均被数据库拒绝。
4. 存量数据零改动：`project_contracts`(639)/`users`(27)/`departments`(13)/`attachments`(6) 与改动前备份逐行哈希一致。
5. 导入幂等：重复执行导入脚本输出 `[SKIP] asset_contracts 已有 5 条数据`。

## 6. 回滚方式

```sql
DROP TABLE IF EXISTS asset_contracts;          -- 删除新表
DELETE FROM schema_version WHERE version = 4;  -- 如需重跑 v4 迁移
```

改动前整库备份：`backups/合同台账_新增asset_contracts前_20260920_130005.db`

---

## 7. 后续版本追加的字段与表（v6~v10，2026-09-21 ~ 2026-09-22）

### 7.1 追加字段（均为 `ALTER TABLE` 幂等迁移）

| 版本 | 追加列 | 类型/约束 | 业务用途 |
|---|---|---|---|
| v6 | **管理部门** | TEXT | 资产合同的部门归属；页面筛选下拉、列表/统计/导出/提醒与首页卡片的**数据范围过滤**均以此列为准 |
| v6 | **责任人** | TEXT | 资产日常管理责任人（多人用顿号分隔）；列表“责任人”列、资产统计“按经办人”维度、归档权限判定 |
| v7 | **是否归档 / 归档人 / 归档时间** | TEXT | 归档后记录只读；页面显示“已归档，仅浏览”徽章，且不进入导入/模板/导出列 |
| v8 | **下次出租时间** | DATE（允许为空） | 首页“合同到期提醒”的资产提醒触发日；编辑表单默认＝租期结束日 + 1 天，可手工修改 |
| v8 | **总合同额** | DECIMAL(12,2) DEFAULT 0 | 资产合同金额；编辑表单按“年租（元）× 租期年数”自动计算（可手改）；首页“总合同额”卡片按该列求和后 ÷10000 折算万元 |
| v10 | **提醒提前天数** | INTEGER（可空，默认 60） | 与项目合同同一口径：为空时取系统全局默认提醒天数 |

> 说明：v8 的 `下次出租时间` 使用 DATE 类型、`总合同额` 使用 DECIMAL(12,2) DEFAULT 0，是本项目首次使用非 TEXT 的日期/定点金额列；写入由接口层统一格式化（日期 `yyyy-mm-dd`、金额保留两位小数）。

### 7.2 索引（共 6 个）

| 索引 | 列 | 版本 |
|---|---|---|
| idx_asset_contracts_location | 位置 | v4 |
| idx_asset_contracts_tenant | 承租人 | v4 |
| idx_asset_contracts_end_date | 租期结束日 | v4 |
| idx_asset_contracts_dept | 管理部门 | v6 |
| idx_asset_contracts_person | 责任人 | v6 |
| idx_asset_contracts_archived | 是否归档 | v7 |

### 7.3 新增附件表 asset_attachments（v9）

```sql
CREATE TABLE IF NOT EXISTS asset_attachments (
  "附件ID" TEXT PRIMARY KEY,
  "资产合同ID" INTEGER NOT NULL,
  "附件类型" TEXT, "显示名称" TEXT, "存储文件名" TEXT, "文件大小" TEXT, "上传人" TEXT, "上传时间" TEXT,
  FOREIGN KEY ("资产合同ID") REFERENCES asset_contracts(id) ON DELETE RESTRICT ON UPDATE RESTRICT
);
CREATE INDEX IF NOT EXISTS idx_asset_attachments_contract ON asset_attachments("资产合同ID");
```

- 附件类型：`决策依据附件`（键 `decision`）、`合同附件`（键 `contract`）；**2026-09-29 起新增 `资产现状图片`（键 `condition`，仅允许图片格式），共三种**；
- 选填；仅支持图片（jpg/jpeg/png/gif/webp/bmp）与 PDF，单文件不超过“全局默认上传最大附件”（默认 20M）；
- 物理文件与项目合同附件同目录（`data/attachments/`），删除资产合同时级联清理附件记录与文件；
- 字段与表结构另见《数据字典表》第 11 节；**v12~v18 的删列/改名/子表变更见文首“更新（V1.2，2026-09-29）”**。

### 7.4 当前字段字典

完整字段字典（接口字段名、类型、必填、取值、业务说明）见《数据字典表》第 10 节（asset_contracts 字段字典）；代码中的单一事实来源为 `app.py` 的 `ASSET_CONTRACTS_DDL`、`ASSET_FIELDS`、`ASSET_QUERY_KEYS`、`ASSET_ARCHIVE_COLS`。

### 7.5 验收结果（v6~v10，已实测）

1. 迁移：线上库 `schema_version` 依次记录 v6~v10，启动日志输出对应 `[MIGRATE]` 行；二次启动无重复执行，存量 5 条资产合同数据完整保留。
2. 约束与默认值：`总合同额` 未填时落库为 `0`（DECIMAL(12,2)）；`提醒提前天数` 于 id1~4 为 60、id5 为 55；`下次出租时间` 于 id1~3 为 `2026-11-15`、id4/id5 为空。
3. 接口回归（数据库副本）：资产列表（5 条，按 id 降序）、筛选（位置/资产名称/承租人/管理部门/责任人/序号/缴款方式/租期结束日区间/归档状态）、详情、新增/修改/删除、归档（归档后禁改禁删）、附件上传/下载/删除、导出 xlsx、模板下载、导入预览与导入、`/api/asset_statistics` 三种维度；查询用户对写接口返回 403；部门范围过滤（他部门用户列表 0 条、详情/附件下载 404）均符合预期。
4. 页面（真实浏览器 + 库副本）：菜单「资产合同」→「资产合同台账」列表 12 列、新增/编辑弹窗三段式表单与附件区、归档徽章、导入预览流程均正常，控制台无报错。

### 7.6 回滚方式（v6~v10）

```sql
DELETE FROM schema_version WHERE version IN (6,7,8,9,10);  -- 如需重跑对应迁移（迁移幂等）
DROP INDEX IF EXISTS idx_asset_contracts_dept;            -- 结构回退需手工执行（SQLite 不支持 DROP COLUMN 前的 3.35 以下版本）
DROP TABLE IF EXISTS asset_attachments;                   -- 回退 v9
```

> 结构级回退建议直接使用 `backups/` 中对应时点的整库备份（如 `合同台账_新增提醒提前天数前_20260922_231043.db`），避免手工改表。
