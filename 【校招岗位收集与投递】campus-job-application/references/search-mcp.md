# 搜索 MCP 与双浏览器并发

按 `产品管理/求职配置.md`「浏览器」双开（默认）时使用。这是同一执行者的两个浏览器通道，不自动创建多个 AI，也不授予任何投递权限。搜索 MCP 由 `scripts/search_mcp.py` 提供（依赖见 `scripts/search-mcp-requirements.txt`），任务队列放在 `CAMPUS_SEARCH_PRIVATE_DIR` 指定的私有运行目录，不进公开包。

MCP 管理任务、每个通道内各公司的独占租约、证据收回和分轮检查；真实网页操作由宿主的浏览器工具执行。派发成功不等于访问成功。

## 调用流程

1. `search_batch_create(manifest)`：冻结 batch_id、scope、authorization_ref 和公司列表。每家公司有 company_key、name、source_key（同一官方来源用同一个键）。公司列表取自主数据的“待搜索”队列或用户点名的公司。
2. `search_dispatch(batch_id, channels=["chrome","iab"], companies_per_channel=10)`：每通道默认领取至多十家公司，可调低到一家。同一公司的所有计划、岗位和详情页共用一个公司租约；同一公司或来源不会同时被两个通道领取。租约过期转为受阻，不自动重试。
3. 页面截图或结构化原文存进私有目录。证据 JSON 含 task_id、company_key、channel、stage、observed_at、source_urls、observations、artifacts（path/sha256）、checks、gaps。gaps 是本轮未完成的工作；已查但官方未披露的事项写明“未披露”和来源，不当作未做。第一轮通过还需要计划与岗位的完整名称目录及数量。
4. `search_result_record`：只接受活跃租约对应公司和通道的证据，校验文件摘要、必需检查项和数量。
5. `search_round_check` / `search_round_advance`：两轮依次为 collection（计划与岗位目录）和 verification（JD、资格、规则、内推、硬范围）。`audited` 表示每个任务都有证据和检查结论，`passed` 表示全部通过，两者独立；受阻公司带着缺口进入下一轮，不改成通过。
6. 两轮结束后，由执行者（多 AI 时为协调者）把核实结果用 `json_excel_sync.py upsert` 写入主数据，状态走 `propose` → 用户确认 → `confirm`。MCP 不写主数据，也不代替用户确认。
7. `search_task_retry`：检查原页面并保存恢复依据后，显式重排一个受阻任务。

`scripts/search_mcp_client.py --input <请求.json>` 可通过官方 SDK 实际调用该服务，例如 `{"tool": "search_round_check", "arguments": {"batch_id": "..."}}`；工具列表用 `{"method": "tools/list"}`。修改配置不代表当前对话已重新加载服务。

同一服务也提供主数据读写接口，见 [招聘数据 MCP 与固定脚本](recruiting-data-mcp.md)。
