# 物流服务商对账服务 (WF-C 引擎)

发货进度管理台子系统。n8n 定时 HTTP 触发 `POST /recon/run`，扫物流对账单任务台
`触发对账=true` 的记录 → 下载服务商对账单 Excel 附件 → openpyxl 按配置解析按运单号聚合
→ join 物流对账明细（服务商运单号）→ 我方侧按『服务商对账映射配置.我方汇总公式』逐票算
→ 与服务商侧逐票比 → 回写汇总/差异/结果/快照 + 关联命中明细 → 飞书卡片通知物流。

所有凭据走 Zeabur 环境变量，仓库不含密钥。`DRY_RUN=1` 时通知只发 Frankie。

## 环境变量

`FEISHU_APP1_ID/SECRET`(聪哥1号 数据) `FEISHU_APP3_ID/SECRET`(聪哥3号 发卡)
`BITABLE_APP_TOKEN` `TBL_RECON` `TBL_DETAIL` `TBL_TASK` `TBL_CFG`
`FRANKIE_OID` `LOGI_DEPT_ID` `LOGI_JOB` `DRY_RUN`(默认1) `BEARER`(端点鉴权)

## 端点

- `GET /health` 健康检查
- `POST /recon/run` （Header `Authorization: Bearer <BEARER>`）执行对账
