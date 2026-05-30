# 采购询价智能助手云端部署包

这是采购询价智能助手的网页原型部署包。

## 云端启动

Docker 环境会自动运行：

```bash
python web_app.py --cloud
```

服务会读取云平台提供的 `PORT` 环境变量。

## 本地启动

```bash
pip install -r requirements.txt
python web_app.py
```

## 当前限制

- 当前只支持 `.xlsx` 文件，旧版 `.xls` 请先另存为 `.xlsx`。
- 这是原型演示版，云端免费服务的文件存储不适合作为正式长期数据仓库。
- 正式上线前建议增加账号权限、项目隔离、数据库/对象存储和操作日志。
