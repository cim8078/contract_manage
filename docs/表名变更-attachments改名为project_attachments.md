# 表名变更：attachments → project_attachments

> 目标：把项目合同附件表 `attachments` 改名为 `project_attachments`，数据库与代码中与该表名相关的对象一并更新，保证系统功能不变。
> 范围：1 张表改名 + 外键与自动索引随表名更新 + 代码 SQL 全面替换；表结构、字段、前端接口路径与业务逻辑均未改动。
> 变更日期：2026-10-12（迁移版本 **v28**）

---

## 1. 数据库变更（data/合同台账.db）

| 对象 | 变更前 | 变更后 |
|---|---|---|
| 表 | `attachments` | `project_attachments` |
| 主键自动索引 | `sqlite_autoindex_attachments_1` | `sqlite_autoindex_project_attachments_1` |
| 外键 | `attachments.合同编号 → project_contracts("编号")` | `project_attachments.合同编号 → project_contracts("编号")` |
| 自定义索引 / 触发器 | 无 | 无（该表原本没有自定义索引与触发器，故无需重建对象） |

改名由 `ALTER TABLE attachments RENAME TO project_attachments` 完成：SQLite 会自动改写表定义中的外键子句，主键自动索引也随之改名。
表结构未变：`附件ID`(TEXT PK) / `合同编号` / `附件类型` / `显示名称` / `存储文件名` / `文件大小`(INTEGER) / `上传人` / `上传时间`，共 8 列。

## 2. 代码变更

### app.py

| 位置 | 变更 |
|---|---|
| `_SCHEMA` | 建表语句改为 `CREATE TABLE IF NOT EXISTS project_attachments (...)`（外键仍 → `project_contracts`） |
| 版本化迁移 | 新增 **v28 `_mig_v28_rename_attachments_table`**（幂等）：旧表存在且新表为空时删空表后改名；两表都有数据时打印告警并跳过，交人工确认 |
| A-03 外键兜底块 | 重建附件表、悬空 `合同编号` 校验、`PRAGMA table_info` 全部改用 `project_attachments` |
| A-06 类型规范化块 | `文件大小 TEXT→INTEGER` 重建逻辑改用 `project_attachments` |
| 附件读写 | `read_attachments()` / `write_attachments()` / Excel 导入的附件 Sheet 写入全部改用 `project_attachments` |

### _import_httz.py（一次性增量导入脚本）

源库 `httz.db` 仍为旧表名 `attachments`，目标库改为 `project_attachments`，脚本内 `common("attachments","project_attachments")` 兼容两库表名差异，其余 SQL 与打印文案同步更新。

### 未变更（有意保留）

- **数据库文件目录 `data/attachments/`**：附件物理文件的存放目录名未变，改目录会导致存量文件全部失联，故保留。
- **接口路径**：`/api/contracts/<cid>/attachments`、`/api/attachments/<id>/download`、`/api/attachments/<id>`、`/api/asset_contracts/<aid>/attachments` 一律不变（与 v5 改名时"前端与接口路径不变"的口径一致）。
- **前端字段**：`templates/index.html` 中的 `attachments` 为 JSON 字段名与前端变量，不属于表名，未改动。
- `asset_attachments`（资产合同附件表）与其余表结构不变。

## 3. 验收结果（已实测）

**数据库**

1. `sqlite_master` 中已无 `attachments` 表；`project_attachments` 存在，且仅有主键自动索引 `sqlite_autoindex_project_attachments_1`。
2. 建表语句中外键为 `REFERENCES "project_contracts"("编号")`；`PRAGMA foreign_key_check` 全库无违规。
3. 数据零改动：7 条附件记录与改名前的整库备份逐行一致（逐字段比对通过）。
4. `schema_version` 记录 v28，二次启动不重复执行。

**接口回归（真实 Flask 应用 + 生产库）**

登录 / 项目合同列表（含附件列表渲染）/ 附件下载（HTTP 200）/ 附件上传（201，记录数 7→8，列表中可见）/ 附件删除（200，记录数 8→7，记录与磁盘文件同步清理）。

## 4. 回滚方式

```sql
ALTER TABLE project_attachments RENAME TO attachments;
DELETE FROM schema_version WHERE version = 28;
```

同时把代码中的 `project_attachments` 改回 `attachments`（含 `_SCHEMA`、v28 迁移函数、A-03/A-06 兜底块与附件读写函数、`_import_httz.py`）。

改动前整库备份：`backups/合同台账_附件表改名project_attachments前_20261012_164422.db`

## 5. 遗留说明

1. `docs/` 下 A-01/A-05/A-06 等治理期文档，以及 v5 改名文档《表名变更-contracts改名为project_contracts.md》成文于本次改名之前，正文中仍称该表为 `attachments`，属历史记录未作改写。
2. 本次同步更新的交付文档：`docs/数据字典表.md`、`docs/系统开发文档.md`（对应的 .docx 未重新生成）。
3. 数据库内部表名已变更，如外部系统/脚本直接用 SQL 访问该库，需同步改用 `project_attachments`。
4. 附带发现（与本次改名无关，改名前后一致）：项目附件记录 `067e244909904106967ef5ac16fc101c`（合同 246610「市民之家16台报价单.pdf」）在库中有记录，但 `data/attachments/` 下已无对应文件，下载会返回"附件文件不存在"，建议按需重新上传。
