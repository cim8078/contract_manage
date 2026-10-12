# 数据库版本化迁移幂等性测试报告

- **测试时间**：2026-09-30 13:30:31
- **被测对象**：`app.py` 的 `MIGRATIONS`（共 19 版，v1~v19）+ `_run_versioned_migrations()`
- **源库**：`data/合同台账.db`（schema_version 已应用 19 版）
- **工作库**：`test_data/合同台账.db`（复制自源库，固定目录；`test_data/` 已加入 `.gitignore`）
- **测试方式**：复制源库到工作库后，依次删除 v13、v15、v17 版本行并重启迁移，每次迁移后执行 `PRAGMA foreign_key_check` + 行数断言

## 断言目标

| 表 | 期望行数 |
|---|---|
| project_contracts | 639 |
| asset_contracts | 5 |
| asset_rent_collections | 16 |
| users | 29 |
| departments | 15 |

- `PRAGMA foreign_key_check` 违规数 = 0；
- 迁移后 `schema_version` 仍为 19 版（被删版本行重新写入）。

## 测试结果

### 场景 1：删除 v13（删除资产合同台账旧租金收缴列（已缴租金1~4季度/交款时间/开票对应租期/已缴租金合计，重建表，旧数据不迁移））→ PASS

- 删除版本行数：1
- 迁移触发输出：
  ```
  [INFO] v13：asset_contracts 已无旧租金列，跳过重建
  [MIGRATE] schema_version 13: 删除资产合同台账旧租金收缴列（已缴租金1~4季度/交款时间/开票对应租期/已缴租金合计，重建表，旧数据不迁移）
  ```
- 迁移前 `schema_version`：v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18, v19
- 迁移后 `schema_version`：v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18, v19
- `foreign_key_check` 违规数：0（迁移前 0）
- 行数（迁移后）：`project_contracts=639(✓)  asset_contracts=5(✓)  asset_rent_collections=16(✓)  users=29(✓)  departments=15(✓)`
- 结论：**PASS**

### 场景 2：删除 v15（删除资产合同台账“缴款方式”列（含数据；缴款方式改由租金收取明细按枚举维护））→ PASS

- 删除版本行数：1
- 迁移触发输出：
  ```
  [INFO] v15：asset_contracts 已无“缴款方式”列，跳过重建
  [MIGRATE] schema_version 15: 删除资产合同台账“缴款方式”列（含数据；缴款方式改由租金收取明细按枚举维护）
  ```
- 迁移前 `schema_version`：v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18, v19
- 迁移后 `schema_version`：v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18, v19
- `foreign_key_check` 违规数：0（迁移前 0）
- 行数（迁移后）：`project_contracts=639(✓)  asset_contracts=5(✓)  asset_rent_collections=16(✓)  users=29(✓)  departments=15(✓)`
- 结论：**PASS**

### 场景 3：删除 v17（租金收取明细“缴款方式”默认值设为“直缴非税”）→ PASS

- 删除版本行数：1
- 迁移触发输出：
  ```
  [INFO] v17：asset_rent_collections.“缴款方式”已有默认值，跳过重建
  [MIGRATE] schema_version 17: 租金收取明细“缴款方式”默认值设为“直缴非税”
  ```
- 迁移前 `schema_version`：v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18, v19
- 迁移后 `schema_version`：v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16, v17, v18, v19
- `foreign_key_check` 违规数：0（迁移前 0）
- 行数（迁移后）：`project_contracts=639(✓)  asset_contracts=5(✓)  asset_rent_collections=16(✓)  users=29(✓)  departments=15(✓)`
- 结论：**PASS**

## 总结论

**全部通过 ✓**：v13 / v15 / v17 三个重建表迁移均通过幂等性验证（删除版本行后重启迁移，表结构与数据行数保持不变，无外键违规，schema_version 恢复为 19 版）。

