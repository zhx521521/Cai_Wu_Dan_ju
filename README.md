# 财务单据智能风险审核系统

这是一个可运行的 Python 后端和浏览器前端演示系统，对应 `Demand/spec.md` 与 `.scratch/financial-document-risk-review/issues/01-09`。

## 运行

```bash
python -m venv .venv
# Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
uvicorn app.main:app --reload --port 8000
```

使用 `uv` 时也可以执行：

```powershell
uv venv .venv
uv pip install --python .venv\\Scripts\\python.exe --index-url https://pypi.tuna.tsinghua.edu.cn/simple -r requirements.txt
.venv\\Scripts\\python.exe -m uvicorn app.main:app --reload --port 8000
```

运行自动验收：

```powershell
& ".\\.venv\\Scripts\\python.exe" -m pytest -q
```

浏览器访问 <http://127.0.0.1:8000>。顶部可以切换演示 SSO 用户；真实部署时用反向代理把企业 SSO 身份映射为 `X-User`，服务端仍会执行角色与资源归属校验。

## 主要入口

- `GET /api/health`：服务健康检查。
- `POST /api/auth/login`、`GET /api/auth/me`：SSO 映射与当前角色。
- `POST/GET/PATCH /api/documents`：五类单据创建、查询与草稿编辑。
- `POST /api/documents/{id}/attachments`、`POST /api/documents/{id}/analyze`：PDF/PNG/JPG 上传、OCR 适配器解析和证据坐标。
- `POST /api/documents/{id}/run-checks`、`POST /api/documents/{id}/external-checks`：规则引擎与征信/市场价适配器。
- `POST /api/documents/{id}/actions`：提交、撤回、退回、驳回、通过、作废等状态机动作。
- `POST /api/import/xlsx`：固定列模板批量导入，返回逐行错误。
- `/api/reports/{id}/export.pdf`、`/api/reports/{id}/export.xlsx`：含风险明细和审计轨迹的报告导出。
- `GET /api/audit`：追加写入、哈希指纹的不可变审计记录。

演示适配器使用可重复的本地结果，不冒充真实外部服务；文件名包含 `timeout`、`nodata` 或供应商包含 `unavailable` 时可验收超时、无数据和可重试失败分支。

## 固定 Excel 列

`type,title,department,amount,detail_total,invoice_total,payee,supplier,reason,contract_ratio`。`type` 取 `payment`、`prepayment`、`batch_payment`、`expense`、`travel`。
