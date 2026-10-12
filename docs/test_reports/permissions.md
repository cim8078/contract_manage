# 权限边界测试报告

- **测试时间**：2026-09-30 13:49:13
- **服务**：http://127.0.0.1:5000（本地 Flask，`python app.py`）
- **账号**：章伟/谢玮/曾明峰/熊玉婷（默认密码 123456）

## 用例结果

| # | 用例 | 预期 | 结果 | 详情 |
|---|------|------|------|------|
| 1 | 登录 章伟（角色=普通用户） | 200 且 role 匹配 | **PASS** | role=普通用户 / name=章伟 / dept=湘潭市机关事务管理局后勤中心-主任班子 |
| 2 | 登录 谢玮（角色=查询用户） | 200 且 role 匹配 | **PASS** | role=查询用户 / name=谢玮 / dept=湘潭市机关事务管理局后勤中心-主任班子 |
| 3 | 登录 曾明峰（角色=普通用户） | 200 且 role 匹配 | **PASS** | role=普通用户 / name=曾明峰 / dept=湘潭市机关事务管理局资产中心-国资组 |
| 4 | 登录 熊玉婷（角色=普通用户） | 200 且 role 匹配 | **PASS** | role=普通用户 / name=熊玉婷 / dept=湘潭市机关事务管理局资产中心-办公用房组 |
| 5 | GET /api/contracts 可见范围（章伟） | 200 / total=357 / 无越权部门 | **PASS** | total=357 越权部门=无 |
| 6 | GET /api/contracts 可见范围（谢玮） | 200 / total=166 / 无越权部门 | **PASS** | total=166 越权部门=无 |
| 7 | GET /api/contracts 可见范围（曾明峰） | 200 / total=0 / 无越权部门 | **PASS** | total=0 越权部门=无 |
| 8 | GET /api/contracts 可见范围（熊玉婷） | 200 / total=0 / 无越权部门 | **PASS** | total=0 越权部门=无 |
| 9 | GET /api/asset_contracts 可见范围（章伟） | 200 / total=0 / 无越权部门 | **PASS** | total=0 越权部门=无 |
| 10 | GET /api/asset_contracts 可见范围（谢玮） | 200 / total=0 / 无越权部门 | **PASS** | total=0 越权部门=无 |
| 11 | GET /api/asset_contracts 可见范围（曾明峰） | 200 / total=5 / 无越权部门 | **PASS** | total=5 越权部门=无 |
| 12 | GET /api/asset_contracts 可见范围（熊玉婷） | 200 / total=0 / 无越权部门 | **PASS** | total=0 越权部门=无 |
| 13 | PUT /api/contracts/row_4 越权（章伟，编号=216201，部门=湘潭市机关事务管理局后勤中心-公车管理组） | 403 | **PASS** | HTTP 403 只能编辑本部门或分管部门合同 |
| 14 | PUT /api/asset_contracts/arow_1/archive 责任人限制（熊玉婷，非责任人） | 403 | **PASS** | HTTP 403 仅管理员或该资产合同责任人可执行归档 |
| 15 | DELETE /api/asset_contracts/arow_1 无 delete_rent（租金记录 3 条） | 409 | **PASS** | HTTP 409 / rent_count=3 |

## 失败详情

无。

## 总结

通过 15 / 共 15。全部通过 ✓

