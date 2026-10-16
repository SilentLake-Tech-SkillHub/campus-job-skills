# 搜索 MCP 与双浏览器并发

用户明确选择 Chrome 与内置浏览器并发时使用此入口；这是同一执行者的两个浏览器通道，不自动授权创建多个 AI 或改变申请权限。批次未作该选择时沿用已确认方式。搜索 MCP 通过 `scripts/search_mcp.py` 提供 stdio 工具，依赖见 `scripts/search-mcp-requirements.txt`。私有目录通过 `CAMPUS_SEARCH_PRIVATE_DIR` 设置，禁止写入公开包。

MCP 管理持久化任务、两个通道的独占租约、公司／共享来源互斥、证据收回和分轮检查。实际网页由宿主现有浏览器 MCP 执行器操作，MCP 不伪装成本地程序直接拥有内置浏览器。`search_dispatch` 返回的两个任务应并发交给 Chrome/IAB 执行，不是连续调用或同一公司双重抓取。派发成功不等于访问成功。工具接口与宿主桥接、当前服务注册、真实双通道执行必须分别验证。

此服务也承接[完整数据处理与用户核对](recruiting-data-mcp.md)：初始化 Excel/JSON 快速展开、结构化草稿保存、固定脚本写 Excel、全量/批次 JSON 同步、固定 HTML 生成、用户调整及明确确认。数据接口与下方搜索队列接口并列；搜索轮次推进不自动写表或确认。

## 调用闭环

1. `search_batch_create(manifest)` 冻结 batch_id、scope、authorization_ref、workbook_sha256、companies。companies 每项保存 company_key、name、source_key（相同官方来源范围使用相同键），可带已观察入口和证据索引。授权文件须存在于私有目录。不得把历史初次访问误归作新首轮；本批可执行的明确补查范围写入 scope。
2. `search_dispatch(batch_id, channels=["chrome","iab"])` 原子领取最多两个独立任务，保存 task_id、lease_token、company、channel、stage、期限。两个通道各自执行，其等待互不阻塞；同一公司／来源不能同时被领取。租约过期转受阻，不自动重试原页。
3. 将 DOM/截图或结构化来源原文保存到私有目录。证据 JSON 含 task_id、company_key、channel、stage、observed_at、source_urls、observations、artifacts（path/sha256）、checks、gaps。gaps 是阻止本轮验收的未完成工作；已研究但未披露的规则另明确记 unknown 与来源，不冒充未开展研究。第一轮通过还须 inventory.plans/roles 的完整名称目录、各记录 id、岗位 plan_ids 和 expected_plans/expected_roles，数量须相等。网页真实性与覆盖充分性由执行者逐项核验，不能用 true 代替未执行的检查。
4. `search_result_record` 只接受活跃租约对应公司和通道的实际证据，验证文件 Hash、必要检查和计数。第三／四轮还须 company_payload_ref，指向按公司交接协议填写并校验的公司对象，列入 artifacts。不得直接并发写 Excel。
5. `search_round_check` 输出计数、前若干个异常和完整检查文件路径。四个研究轮次依次为 collection、verification、structuring、recommendation。`audited` 表示本轮每个任务已有真实证据及检查结论，`passed` 表示全部核验通过，两者独立。`search_round_advance` 仅在 audited 后推进，保留上一轮报告及具体阻塞数，不把受阻公司改成已通过；下一轮继续各项独立可做工作。第三／四轮允许将明确未知与缺口结构化，不能因缺口阻断所有数据整理；公司是否可交接仍由 company_handoff 的完整研究与用户确认门禁判断。任务小批按公司原序每十家分组；小批数量不改变真实业务覆盖要求。
6. 所有研究轮完成后，按批次已批准的统一写入时机调用 company_handoff.py，读回 Excel 并刷新整本 JSON、批次 JSON 和 HTML。最终用户核对仍是独立阶段，MCP 不代填确认。用户公司／BU志愿及官网机制分别核对。`search_task_retry` 仅在检查原页／证据、保存恢复依据后显式重排受阻任务。

可通过 `scripts/search_mcp_client.py --input <UTF-8请求.json>` 用官方 SDK 对服务作实际 stdio 调用。输入为 `{ "tool": "search_round_check", "arguments": { "batch_id": "..." } }`；工具发现用 `{ "method": "tools/list" }`。此客户端与宿主加载的 MCP 工具菜单是两种接入方式，不应声称修改配置即代表当前聊天已热加载。

## 用户志愿单独分组

`company.priority_scope=company/bu/unknown`；已知分组规则保存 priority_scope_source/priority_scope_checked_at。用户志愿记录在第九张共享表“用户志愿”及 `user_priorities`：role_id、scope、group_id、bu、rank、status、confirmation_ref、confirmation_sha256、confirmed_at。不分 BU 时 group_id 为公司键且 bu 空；按 BU 时为公司键::BU，每个 BU 从一开始连续排序。同组序号唯一，不同 BU 可都为第一志愿。未知分组保持 unknown，不能从未发现 BU 限制推断公司不分 BU。

未确认 status=pending/unknown、rank=null；用户确认后记录正整数序号和私有依据及 Hash。官网是平行志愿或没有志愿机制时，用户自身顺序仍须独立保存。roles.order 仍是处理次序，官网 preferences.rank 仍按机制决定是否可为空。已选岗位缺用户确认顺序，或依据文件改变时，不可通过交接。用户一次明确确认可以同时覆盖公司 HTML 和分组排序，不重复索要相同批准。
