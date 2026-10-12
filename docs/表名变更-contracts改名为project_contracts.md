# 表名变更：contracts → project_contracts

> 目标：把合同台账主表 `contracts` 改名为 `project_contracts`，数据库与代码中与该表名相关的对象一并更新，保证系统功能不变。
> 范围：1 张表改名 + 索引/触发器改名 + 外键引用更新 + 代码 SQL 全面替换；前端接口路径、页面、字段名与业务逻辑均未改动。
> 变更日期：2026-09-20

---

## 1. 数据库变更（data/合同台账.db）

| 对象 | 变更前 | 变更后 |
|---|---|---|
| 表 | `contracts` | `project_contracts` |
| 唯一索引 | `idx_contracts_no` | `idx_project_contracts_no` |
| 普通索引 | `idx_contracts_department` / `_procurement_date` / `_next_bid_date` / `_archived` / `_handler` | `idx_project_contracts_*`（同上 5 个） |
| 触发器 | `trg_contracts_dept_sync_i` / `trg_contracts_dept_sync_u` | `trg_project_contracts_dept_sync_i` / `trg_project_contracts_dept_sync_u` |
| 外键 | `attachments.合同编号 → contracts("编号")` | `attachments.合同编号 → project_contracts("编号")` |
| 自增序列 | `sqlite_sequence.contracts` | `sqlite_sequence.project_contracts` |

外键、索引与触发器中对表的引用由 SQLite 在 `ALTER TABLE ... RENAME TO` 时自动改写（索引/触发器不支持直接改名，按其原定义重建后删除旧名）。

## 2. 代码变更（app.py）

- `_SCHEMA`：主表建表语句、attachments 外键改为 `project_contracts`；索引/触发器名同步。
- 版本化迁移新增 **v5 `_mig_v5_rename_contracts_table`**（幂等）：
  1. 若库中只有旧表 `contracts`：执行改名，并同步索引/触发器名；
  2. 若库中同时存在 `contracts` 与 `project_contracts`，且新表为空、无自定义索引（旧库首次启动时 `_SCHEMA` 会先按新名建出空表）：删除该空表后按第 1 步改名，保证旧库数据落到新表名下；
  3. 若两表都有数据则打印告警并跳过，交人工确认。
- 历史迁移 v1/v2 通过新增的 `_contracts_table(con)` 读取当前表名，兼容仍为旧名 `contracts` 的历史库。
- 其余所有 SQL（A-01/A-03/A-05/A-06 约束块、CRUD、导入、导出、统计、归档、部门同步等）统一改为 `project_contracts`。
- `_migrate_a08.py`（A-08 一次性清洗脚本）中的 SQL 同步改名，脚本仍可复用。

## 3. 未变更（有意保留）

- 前端与接口路径：`templates/index.html`、`/api/contracts*` 路径、前端变量/类名、JSON 字段名一律不变（用户与浏览器端无感知）。
- Python 内部函数名（`read_contracts`、`_scope_contracts`、`api_list_contracts` 等）与导出 Excel 的列名不变。
- `asset_contracts`（资产合同台账）与 `users`/`departments`/`attachments`/`operation_logs`/`settings` 表结构不变。

## 4. 验收结果（已实测）

**旧库升级（用改动前备份模拟）**
1. 首次启动自动执行 v5：表改名成功，639 条合同、27 用户、13 部门、6 附件、操作日志全部保留；旧名 `contracts` 不再存在。
2. 9 个索引、5 个触发器全部为新名；attachments 外键指向 `project_contracts`；`sqlite_sequence` 同步。
3. 约束照旧生效：编号唯一索引拒绝重复编号；`所属部门` 触发器自动回填 `所属单位/所属科室`；attachments 悬空合同编号被外键拒绝。
4. `schema_version` 记录 v1~v5，二次启动无重复迁移。

**接口回归（真实 Flask 应用 + 数据库副本，全部通过）**
登录 / 当前用户 / 首页 / 仪表盘 / 合同列表（639 条）/ 按编号搜索 / 合同详情 / 提醒 / 待支付 / 统计 / 导出 xlsx / 下载模板 / 新增合同（含字段校验与重复编号拦截）/ 修改 / 附件上传下载删除 / 归档（归档后禁改禁删）/ 删除 / 导入预览 + 导入 / 部门改名级联同步合同 `所属部门`（29 条）/ 用户 / 部门 / 配置 / 操作日志 / 日志导出 / 退出登录。

## 5. 回滚方式

```sql
ALTER TABLE project_contracts RENAME TO contracts;
-- 索引/触发器名回退（如确需）：
DROP INDEX idx_project_contracts_no;  -- 其余 idx_project_contracts_* 同理
CREATE UNIQUE INDEX idx_contracts_no ON contracts("编号");
DELETE FROM schema_version WHERE version = 5;   -- 如需重跑 v5
```

同时把代码中 `project_contracts` 改回 `contracts`（含 v5 迁移函数与 `_contracts_table`）。

改动前整库备份：`backups/合同台账_改表名project_contracts前_20260920_131020.db`

## 6. 遗留说明

- 2026-09-20 后续界面文案调整：前端菜单「合同台账」→「项目合同」、合同操作区标题「合同台账」→「项目合同台账」（仅 `templates/index.html` 文案，接口路径与交互逻辑未动）。
- `docs/` 下既有文档（如《数据库规范化评估报告》《完整性规范化建议》等 A-01~A-08 治理期文档）成文于改名之前，正文中仍称该表为 `contracts`，属历史记录未作改写；**面向交付的《数据字典表》《系统开发文档》《用户操作手册》已于 2026-09-23 同步更新为 `project_contracts`**（并重新生成对应 .docx）。
- 数据库内部表名已变更，如外部系统/脚本直接用 SQL 访问该库，需同步改用 `project_contracts`。
