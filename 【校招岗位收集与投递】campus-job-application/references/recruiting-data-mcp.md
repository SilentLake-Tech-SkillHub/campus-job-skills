# 招聘数据 MCP：初始化、JSON、Excel 与用户核对

同一 `search_mcp.py` 服务保留搜索调度工具，并提供以下固定数据接口。`recruiting_data.py` 是薄适配层，复用 `company_handoff.py`、`workbook_store.py` 的同一字段契约、备份、增量写入、读回、完整业务 JSON、固定 HTML 模板和确认门禁，不维护另一套业务格式。

部署时设置 `CAMPUS_RECRUITING_WORKSPACE` 为用户授权的私有工作区根，`CAMPUS_SEARCH_PRIVATE_DIR` 为其内部运行目录。两者不写入公开 Skill。`workbook` 使用相对工作区的现有 xlsx 路径。接口不登录、填写或提交招聘申请。

| 调用 | 输入与结果 | 流程作用 |
|---|---|---|
| `recruiting_data_open` | workbook、batch；返回当前 Hash、表头、行数、公司数、完整 JSON 同步问题及字段契约 | 初始化快速展开，纯读 Excel，不创建明细表，不把初始化算成研究 |
| `recruiting_data_query` | workbook、company 或 sheet、offset、limit；分页返回原列及业务行 | 按公司/表快速展开已同步 JSON，避免反复读取整本内容；旧 JSON 拒绝读取 |
| `recruiting_draft_put` | workbook、batch、data；返回私有 draft_ref、draft_sha256、workbook_sha256 | 将收集整理后的公司、计划、目录、JD、规则、内推及两类志愿保存为结构化 UTF-8 JSON 交换输入；此时未写 Excel，业务校验尚待 commit |
| `recruiting_draft_get` | workbook、batch、draft_ref、draft_sha256 | 读回精确草稿，检查工作簿/批次归属和篡改 |
| `recruiting_data_initialize` | workbook、batch、sheet、scope、expected_sha256，可选 company | 到达已授权写入阶段后，调用固定初始化脚本，保留历史行并建立研究待查记录，刷新派生文件；不能提前绕过批次统一写入时机 |
| `recruiting_data_commit` | workbook、batch、draft_ref、draft_sha256、expected_sha256 | 校验输入、锁定串行写入、备份和增量写 Excel，读回并重导整本 JSON、所有既有批次 JSON/HTML；返回文件 Hash 与研究缺口，不返回整本内容 |
| `recruiting_review_refresh` | workbook、batch、expected_sha256 | 手动修改 Excel 后重新导出完整业务 JSON、批次 JSON、固定格式 HTML；不从旧 JSON 覆盖 Excel |
| `recruiting_review_confirm` | workbook、batch、company_key、evidence_ref、html_sha256、expected_sha256 | 仅在用户明确确认当前 HTML 后调用，绑定已有私有确认依据；执行完整研究、用户志愿依据与版本门禁，不授予申请权 |
| `recruiting_handoff_check` | workbook、batch，可选岗位 url | 检查 Excel/全量 JSON/批次 JSON/HTML、研究和用户确认；未通过返回具体缺项 |

草稿 `data` 仍是固定写入协议：`companies` 或 `company_patches` 二选一，可同时含 `workbook_updates`。字段与类型见[数据建设](storage-workflow.md)。草稿不可原地修改；用户调整后另存新草稿、赋新公司版本，带当前工作簿 Hash 提交。Excel 保存后成为主记录，JSON 草稿不是独立长期台账。写入自动撤销受影响公司的旧确认，重新生成 HTML 后再请用户核对。

HTML 继续使用同一固定模板：计划要求、完整名称目录、三档候选及完整 JD、推荐、规则矩阵、内推、官网志愿与用户公司/BU排序、既有记录和待核事项。MCP 返回文件位置及 Hash，宿主展示文件；文件生成或打开都不等于用户确认。接口不能通过自己生成确认文本、填入 `confirmed` 或使用推荐次序制造用户批准；`evidence_ref` 必须引用真实用户对该版本的明确回复，宿主负责核对其内容与授权范围。只有测试可以使用明确标注的合成证据。

初始顺序：open → 必要时 refresh → query → 研究并 draft_put/get → 到达批次写入门禁后 initialize/commit → 读回 query → 展示固定 HTML → 用户调整则新草稿 commit、重新展示 → 明确回复后 confirm → handoff_check。首轮未查齐时保存草稿和缺口，不能为测试接口而写真实主表。

现有 `campus-search` 注册名保持兼容；它已是研究和数据核对服务，不仅是网页搜索。新增接口以实际 MCP 工具发现及协议调用为准；修改服务文件不证明当前聊天工具菜单已热加载。已有固定 CLI 保留，数据和门禁结果一致。
