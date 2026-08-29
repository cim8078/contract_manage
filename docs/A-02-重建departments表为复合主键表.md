# A-02 重建 departments 表为复合主键表

> 目标：为 `departments` 表添加复合主键（单位名称, 科室名称），杜绝重复部门组合，使部门组合唯一性在数据库层兜底。
> 背景：`departments` 表（13 行）原无主键、无唯一约束，组合唯一性仅靠应用层保证。即使绕过 API 直接写入，数据库层也能拒绝重复组合。

**1. 修改范围**
- 仅修改 `app.py` 的 `init_db()` 迁移逻辑（A-02 块）；部门管理的增删改接口行为不变。
- SQLite 无法通过 `ALTER TABLE` 直接加主键，采用“建新表 → 拷贝数据 → DROP 旧表 → RENAME”的标准重建流程，`BEGIN…COMMIT` 事务包裹，重建前先备份。

**2. 数据库结构修改（1 项）**
- 重建 `departments` 表（当前实况）：

```sql
CREATE TABLE "departments" (
  "单位名称" TEXT NOT NULL,
  "科室名称" TEXT NOT NULL,
  PRIMARY KEY ("单位名称","科室名称")
);
```

- 变化点：
  - 两列均加 `NOT NULL`；
  - 新增复合主键（单位名称, 科室名称），SQLite 自动生成主键唯一索引 `sqlite_autoindex_departments_1`（实测 `unique=1`、origin=`pk`）；
- 迁移前备份：原表 DDL 与 13 行 INSERT 语句写入 `data/departments_ddl_backup.sql`（SQL 脚本，供回滚）。

**3. 数据库记录修改（0 条）**
- 迁移前先扫描存量重复组合：

```sql
SELECT "单位名称","科室名称",COUNT(*) AS c FROM departments GROUP BY "单位名称","科室名称" HAVING c > 1;
```

- 实测 13 行无重复；若存在重复，脚本会报错终止并输出清单，由人工处理后重跑，不做静默删除或合并。
- 本次迁移仅把 13 行原样拷入新表，未改任何记录内容。

**4. 解决的问题**
- 改动前：部门组合（单位名称, 科室名称）唯一性仅靠应用层校验，绕过 API（直接 SQL 写入、并发请求）可插入重复组合或空值。
- 改动后：数据库层兜底——任何写入路径插入重复组合或 NULL 科室名都会触发 `UNIQUE/NOT NULL constraint failed`（`sqlite3.IntegrityError`）。
- 行为不变：部门列表/新增/删除接口照常工作，应用层友好提示不变。
- 附加收益：复合主键为后续 A-03 外键（`contracts`/`users` 的 所属单位/所属科室 → `departments`）提供“唯一、非空”的引用目标；A-03 启用外键后，被合同/用户引用的部门禁止删除（RESTRICT），避免悬空引用。

**5. 验收结果（已实测）**
1. `PRAGMA table_info(departments)` 显示两列均为复合主键（pk=1、pk=2）且 `notnull=1`。
2. `PRAGMA index_list(departments)` 显示主键自动索引 `sqlite_autoindex_departments_1`（unique=1）。
3. 直接插入已存在的（单位名称, 科室名称）组合被拒绝（`UNIQUE constraint failed`）。
4. 部门列表 API 与新增/删除部门回归正常。

**6. 回滚方式**
- 用迁移前备份 `data/departments_ddl_backup.sql` 恢复原表结构与 13 行数据（先 `DROP TABLE departments` 再按脚本重建），数据无丢失。
- 或恢复迁移前的整库备份（`backups/` 下 A-02 重建前快照）。
