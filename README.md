潭合同-台账管理系统
实现合同信息全过程、信息化管理，提供合同全要素信息、附件添加、修改、删除，提供按角色权限查询、统计分析，合同到期提醒，合同框选支付提醒等功能。



启动命令

cd C:\\Users\\jghq\\Desktop\\contract\_manager

pip install -r requirements.txt

python app.py


发版前检查（完整清单见 docs/系统部署手册.md 第 5.4 节「发版（升级）检查单」）

1) 先备份生产库：把 data/合同台账.db 复制到 backups/合同台账_发版前_<时间戳>.db
2) 语法自检：python -m py_compile app.py contract_no.py
3) 启动服务后跑前端逐视图冒烟（退出码必须为 0）：
   node tools/browser_smoke.mjs --base http://127.0.0.1:5000
   脚本用真实浏览器登录后逐个切换左侧 9 个菜单视图，捕捉控制台报错，可拦住“整页空白/单个视图空白”类前端渲染故障。

