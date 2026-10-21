---
name: campus-job-application
description: Use for every campus-job company search in this project — work through the master JSON's 待搜索 queue (or companies the user names), verify official plans and in-scope roles, write results only through the fixed script, get user confirmation for every status change, and hand confirmed roles to the application component through forward review. It never fills or submits applications.
metadata:
  version: "2.0.0"
---

# 校招岗位搜索与主表维护

## 开工前必读

- `产品管理/求职配置.md`：运行配置（主数据位置、地区表、目标届别、公司池基线、搜索顺序、浏览器）。缺失时先问用户并按 [模板](references/profile-template.md) 建立。
- `产品管理/求职目标.json`：用户求职目标与选岗规则的**唯一来源**（岗位顺位、排除红线、城市、用工类型、特殊通道、决策红线、在用简历）。每次会话开始按文件里的“每次会话开始”向用户复述并确认方向和岗位顺位；用户纠正时直接改写该文件对应条目，不另记日志。
- [数据建设与固定流程](references/storage-workflow.md)：每类信息的唯一存放位置和写入命令。
- [主表字段约定](references/workbook-schema.md)。

个人取值只放在工作区私有文件里，不写进本 Skill。

## 数据规则（Hook 强制）

- 主数据是 `岗位交接/工作簿全量.json`。AI 只读 JSON，不读不改任何 `.xlsx`；Excel 只能由 `scripts/json_excel_sync.py write` 写出并逐格读回。
- 记录唯一、覆盖更新：每类信息只有一个存放位置、只保留最新一版。禁止追加旧结论、另存副本/历史/备份、按批次另起同类文件；搜索过程文件放工作区以外的临时目录，结论写回后即删除。回滚只靠 git。
- 投递状态只能用标签组文件里的标签（一组一个文件）：岗位行用 `产品管理/投递状态标签.json`，公司行用 `产品管理/公司状态标签.json`；每行一个、不得为空，写入前由脚本统一检查。每家公司在每张地区表固定一行公司行，有没有岗位都保留。任何状态变化（含首次赋值）先 `propose`，把“当前状态是X，更新状态是Y”原样给用户看，用户明确确认后 `confirm`（附用户原话）。`upsert` 只改非状态字段。
- 名额收口：用户的选择占满某公司“投递名额”表里“公司总量”的上限后，同公司剩余“岗位待筛选”在同一次确认中改为“岗位不符合用户要求”。用户选满之前，任何 Agent 不得因名额剔除岗位。
- 对话结束前若 JSON 有未写入 Excel 的改动，展示逐格差异并问用户“现在写入 Excel，还是下一轮再写入”；现在写入 → `write`，下一轮 → `defer`（附原话）。

## 搜索队列

- **第一轮**：按主表行序处理状态为“待搜索”的公司，或用户本批点名的公司；按 `求职目标.json` 的“集团公司规则”，集团已投递过的旗下公司不再单独搜索，“不再搜索的公司”直接跳过。已经搜过的公司（有岗位行、标为“公司无合适岗位”、或“校招计划”表有本届核查记录）不重开，即使有缺口。
- **补查**：第一轮之后，只处理用户点名的公司和具体缺口。一句笼统的“继续”只推进第一轮，不授权补查。
- 每个新批次开始前运行 [campus-job-preference-intake](../【校招批次范围确认】campus-job-preference-intake/SKILL.md)，确认本批公司范围、本批意图（信息收集 / 准备投递）、执行方式（一个 AI / 多个 AI）。求职目标里已有的长期规则不重复问。
- 用户确认多个 AI 共同搜索时，加载 [campus-job-parallel-search](skills/campus-job-parallel-search/SKILL.md)；用户确认双浏览器并发时读 [搜索 MCP](references/search-mcp.md)。

## 每家公司的固定步骤

1. **对照已有数据**：用 `json_excel_sync.py query --company <公司>` 读该公司的岗位行、校招主链接、校招计划和投递名额，只查缺的范围，不重复已查页面。
2. **计划盘点**：运行 [campus-job-plan-inventory](../【校招计划盘点】campus-job-plan-inventory/SKILL.md)，列出本届适用计划及计划间关系。
3. **筛选**：按求职目标的硬范围在官网筛选（用工类型、城市、岗位方向）；排序偏好只用于排列，不缩小搜索范围。标题和分类只用来发现候选，是否纳入按完整 JD 的实际职责判断：标题写“产品经理”但职责是运营/支持的不纳入，标题不含“产品”但承担产品职责的要纳入详核。零结果要覆盖同义标题和相关分类。
4. **详情核验**：每个范围内岗位从官网结果卡片打开详情页，核对岗位名、完整 JD、届别或应届要求、工作地点。404、空白、过期或跳到首页的不算已核岗位；不按列表 ID 拼链接，不用第三方招聘网站代替官网。重试一次仍失败的列为未核。
5. **任职资格**：按 [任职资格核对](references/eligibility-check.md) 分别判断职责是否在范围内、个人是否满足条件。
6. **写入主数据**：
   - 每个已核岗位用 `upsert` 写一行：岗位名称只写官网原名，岗位链接是该岗位的 https 官网详情页，填工作地点、完整 JD、所属BU（无则“未披露”）、校招计划、投递截止（日期或“公司未披露”）。新行状态为空，随后 `propose` 为“岗位待筛选”，用户确认后 `confirm`。
   - 已有岗位换链接（官网换了详情页，或原链接不是详情页）：`relink --preview` 把“岗位链接 X → Y”原样给用户看，确认后 `relink --quote`；投递状态不变、无需重新确认，志愿顺序等投递数据随行迁移。不能用 `upsert` 直接改已有行的链接。岗位在已交接的投递批次里时须先完成投递与反向评审。
   - 公司行：搜完有岗位的 `propose`“已有岗位”；没有符合要求岗位的 `propose`“公司无合适岗位”；用户确认后 `confirm`。新增岗位不占用公司行；新公司由 `upsert` 自动补公司行（状态为空，同样须确认）。
   - 官方校招入口写“校招主链接”，计划写“校招计划”，公司投递规则用 `set-quota-rules` 整家覆盖写“投递名额”。
   - 中国大陆公司按 [内推码子 Skill](skills/campus-job-referral-code/SKILL.md) 查内推码。
7. **读回后关窗**：`query` 读回该公司，确认岗位链接和 JD 已保存后，才按 [浏览器规则](references/browser-and-submission-gates.md) 关闭该公司由 Agent 打开的标签（信息收集模式）。无法核实的公司保持待查并写明缺口。
8. **汇报**：逐岗列出公司、岗位、直链、BU（或未披露）、城市、用工类型、方向、JD 依据和纳入/排除理由；按计划列出结果数、打开的详情数、无效页和未查页。只有所有相关计划都查完，才能下“无合适岗位”的结论。

## 浏览器

开页前读 [浏览器与页面核验](references/browser-and-submission-gates.md) 和 [标签分组](references/chrome-window-grouping.md)。浏览器按 `求职配置.md` 的设定使用；切换浏览器先说明失败情况并取得用户确认。每个标签记录所属公司，不占用用户的页面和窗口。

## 进度汇报

查询、续跑和汇报前按 [研究进度核对](references/research-progress.md) 从主数据统计覆盖情况，已查但有缺口的公司不能说成“未开始”或“已完成”。向用户汇报公司情况时用 `json_excel_sync.py report --company <公司>`（MCP `recruiting_company_report`）的结果，数字以它为准，不手工统计。

## 交给投递组件

选岗由用户决定。用户确认“岗位待投递”后，按 [正向评审与反向评审](references/company-review-handoff.md) 生成投递批次和正向评审页，用户确认后才交给投递组件；投递结果经反向评审、用户确认后才写回主表。本 Skill 不填表、不上传简历、不提交申请。

## 数据接口

[招聘数据 MCP 与固定脚本](references/recruiting-data-mcp.md) 列出 MCP 工具与命令行的对应关系。接口调用成功不等于事实已查齐或用户已确认。
