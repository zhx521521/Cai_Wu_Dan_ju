# 九个工单验收记录

| 工单 | 需求 -> 实现 | 技术验收入口 | 当前状态 |
| --- | --- | --- | --- |
| 01 平台/SSO/审计 | `X-User` 角色映射、服务端资源归属、追加审计哈希 | `POST /api/auth/login`、`GET /api/audit` | 已实现，待部署企业 SSO 配置 |
| 02 单据生命周期 | 五类单据、草稿编辑、提交/撤回/退回/重提/驳回/作废和版本号 | `POST /api/documents`、`POST /api/documents/{id}/actions` | 已实现 |
| 03 凭证/OCR/校正 | PDF/PNG/JPG 20MB 限制、适配器状态、字段坐标/置信度、人工校正 | `/attachments`、`/analyze`、`/evidence/{field}/correct` | 已实现，适配器为可替换演示实现 |
| 04 Excel 导入 | 固定列模板、逐行校验、重复识别、错误反馈、审计 | `POST /api/import/xlsx` | 已实现 |
| 05 核心规则 | 金额、明细、发票、合同比例、费用标准、重复发票/批量收款人 | `POST /api/documents/{id}/run-checks` | 已实现，规则版本 `rules-v1.0` |
| 06 外部风险联动 | 征信/市场价四种结果，失败不等于无风险，查询摘要和版本 | `POST /api/documents/{id}/external-checks` | 已实现，适配器为可替换演示实现 |
| 07 三栏工作台 | 左结构化单据、中证据预览、右风险与审批动作，幂等审批 | 浏览器“智能审核工作台” | 已实现 |
| 08 报告/审计管理 | 总览、台账、在线查看、PDF/Excel/CSV、全量审计轨迹 | `/api/reports`、`/export.pdf`、`/export.xlsx` | 已实现 |
| 09 端到端场景 | 大额技术服务费对公付款演示数据与全链路入口 | `GET /api/e2e/demo` | 已实现，需部署环境提供真实外部凭证后做生产验收 |

## 可重复验收顺序

1. 以 `applicant@example.com` 创建付款单，上传一个 PDF/PNG/JPG；`timeout`、`nodata` 文件名可验证外部失败分支。
2. 对低置信度字段执行人工校正后提交；规则会生成合同付款比例等风险项。
3. 切换 `reviewer@example.com`，在工作台执行外部查询、确认/忽略风险、退回或通过。
4. 切换 `approver@example.com` 完成审批，再到报告中心导出 PDF/Excel，并在审计页核对追加事件和快照指纹。

本轮已使用 uv 提供的 CPython 3.12 完成 `py_compile`、`node --check frontend/app.js` 和真实 Uvicorn HTTP 检查；契约与端到端测试结果为 `4 passed`。真实 OCR、征信、市场价服务仍需部署环境提供正式适配器配置。
