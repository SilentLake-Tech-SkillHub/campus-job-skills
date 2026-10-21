---
name: campus-job-plan-inventory
description: Identify and inventory a company's applicable campus-recruitment plans once during the first-pass search, recording plans, plan relationships and application rules in the master JSON; after the first pass, investigate plan gaps only when the user selects them.
metadata:
  version: "2.0.0"
---

# 校招计划盘点

在 `campus-job-application` 打开或筛选某公司岗位之前使用。先读 [主表字段约定](../【校招岗位收集与投递】campus-job-application/references/workbook-schema.md)。所有写入经 `json_excel_sync.py`：计划写“校招计划”表（`upsert`），名额与投递规则写“投递名额”表（`set-quota-rules`，按公司整体覆盖），官方入口写“校招主链接”。

## 第一轮：每家公司盘点一次

1. 用 `json_excel_sync.py query --company <公司>` 读该公司已有的校招主链接、校招计划、投递名额和岗位行。已有本届核查记录（即使有缺口）就不重开，缺口留给补查。
2. 只对“未访问”的公司（见 [研究进度核对](../【校招岗位收集与投递】campus-job-application/references/research-progress.md)）打开当前官方校招入口。一家公司可能有多个计划：正式校招、提前批、BU 专项、管培/产培、实习等。只记录当前官方来源支持的计划；往届计划只作线索，不当作本届。
3. 每个已核实的计划写一行“校招计划”，填计划链接、依据链接、核验日期和盘点状态。没有核实到本届计划时保留空计划的覆盖行，并写明已查来源。
4. 官方直接披露的计划间关系写进 `投递关系`；未披露的写 `待核实` 并列入补查清单，不为逐对核实计划关系而拖延第一轮队列。
5. 主动查清投递规则：公司总量、各计划或 BU 的上限、能否同时进行多个流程、能否分 BU 平行投递、终止后能否续投、冷静期与刷新、投递后能否修改。写入“投递名额”表：规则原文逐字摘抄官方原文，同时把保存的来源原文文件作为 `原文出处文件` 交给脚本核对；来源写成 `来源：<公司>-<岗位或页面> 核验日期：YYYY-MM-DD`。查不到的取值写“公司未披露”。不能用一个公司总上限代替各范围的规则。
6. 特殊计划（如管培/产培）先向用户展示培养路径、要求和名额影响，再问是否考虑，并对照 `求职目标.json` 的特殊通道规则；不能只抽象地问“要不要管培”。

## 补查

第一轮结束后，汇报四类覆盖数量和具体缺口，等用户点名公司和缺口。只刷新缺失、过期、变化或被点名的证据；复用已查页面、筛选、结果和岗位，不重复已完成的搜索。只有该公司所有相关计划和范围内结果都查完，才能下“没有匹配岗位”的结论，并列出仍未打开或未查的页面与计划。

## 证据与安全

- 以官方页面、官方接口和雇主公告为准。高校转发的雇主公告可作佐证，记录准确链接和原文。第三方汇总只是线索，不能单独证明名额或计划冲突。
- 不因有不同的申请按钮或计划名就推断计划独立，也不因共用登录或账号就推断冲突。
- 名额和刷新等时效性规则，在投递组件填写或提交前由其重新核对登录账号显示的剩余机会。

## 浏览器

使用父 Skill 的浏览器规则和 [标签分组](../【校招岗位收集与投递】campus-job-application/references/chrome-window-grouping.md)。每个标签标明所属公司，不占用其他公司或用户的标签与窗口。完成后按父 Skill 的读回后关窗规则，只处理当前公司的标签。盘点不授权打开申请表、上传、填写或提交。

用户选择双浏览器并发时，按父 Skill 的 [搜索 MCP](../【校招岗位收集与投递】campus-job-application/references/search-mcp.md) 执行，两个浏览器处理不同公司。
