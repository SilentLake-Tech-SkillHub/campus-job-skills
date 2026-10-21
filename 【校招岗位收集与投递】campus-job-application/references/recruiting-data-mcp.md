# 招聘数据 MCP 与固定脚本

`search_mcp.py`（注册名 `campus-search`）在搜索队列工具之外，提供以下数据接口。它们都是 `scripts/json_excel_sync.py` 的薄封装，与命令行完全一致；接口不登录、不填写、不提交招聘申请。部署时 `CAMPUS_RECRUITING_WORKSPACE` 指向用户授权的私有工作区根；标签与字段规则在工作区 `产品管理/投递状态标签.json`、`产品管理/表格字段规则.json`，不写入公开 Skill。

| MCP 工具 | 对应命令 | 作用 |
|---|---|---|
| `recruiting_application_requirements_begin` | `application_requirements.py begin` | 初始化并读取当前回合需求；同回合默认不重复问，新回合重新初始化 |
| `recruiting_application_requirements_set` | `application_requirements.py set` | 覆盖本轮需求与用户原话；空对象选择默认路线 |
| `recruiting_application_forward` | `company_store.py forward` | 从临时表组批；无公司参数默认最新10家；成功后清空需求表 |
| `recruiting_data_query` | `query` | 只读：按公司或工作表分页读取主数据 |
| `recruiting_company_report` | `report` | 只读：公司汇报——校招主链接、计划数与名称、管培/产培计划、岗位总数及按投递状态计数与列表、用户已选（计入名额）、投递名额与剩余 |
| `recruiting_status_check` | `check` | 只读：标签、字段、名额表、未确认状态变更、JSON 与 Excel 差异 |
| `recruiting_status_propose` | `propose` | 只读：生成“当前状态是X，更新状态是Y”（含名额收口行），必须原样给用户看 |
| `recruiting_status_confirm` | `confirm` | 用户明确确认后执行，记用户原话到确认台账并改主数据 |
| `recruiting_rows_upsert` | `upsert` | 新增岗位行或改非状态字段；整行过字段门禁（岗位名称不得含“志愿”等）；不能更换已有岗位的链接 |
| `recruiting_link_exception_preview` | `link-exception --preview` | 只读：官网无独立详情页时，登记或撤销链接例外前生成备注“当前 → 更新”，必须原样给用户看 |
| `recruiting_link_exception_confirm` | `link-exception --quote` | 用户确认后写固定格式的链接例外备注并记入台账；forward 据此只免“列表页”检查 |
| `recruiting_link_change_preview` | `relink --preview` | 只读：更换已有岗位链接前生成“岗位链接 X → Y”，必须原样给用户看 |
| `recruiting_link_change_confirm` | `relink --quote` | 用户确认后换链接：状态不变、无需重新确认，确认记录、志愿顺序等投递数据随行迁移；岗位在已交接批次中时拒绝 |
| `recruiting_json_write_excel` | `write` | 唯一写 Excel 的入口，写后逐格读回 |
| `recruiting_sync_pending` | `pending` / `diff` | 是否有未写入 Excel 的改动及逐格差异 |
| `recruiting_write_defer` | `defer` | 用户选择“下一轮再写入”时登记原话 |

只有命令行提供的维护命令：

- `set-quota-rules`：按公司整体覆盖写“投递名额”表。
- `reset-rows`：删除无效岗位行，公司退回“待搜索”；`--mark-research` 标明需重新搜索；有真实投递记录的行不能删。
- `set-sheet-rows`：整页覆盖说明类工作表（如“分类说明”），岗位表和名额表不能用。
- `add-column` / `drop-column`、`clear-orphans`。

数据记录规则：每类信息只有一个存放位置，只保留最新一版，更新即覆盖；不另存草稿、副本、历史或备份，回滚靠 git。存放位置见 [数据建设与固定流程](storage-workflow.md)。
