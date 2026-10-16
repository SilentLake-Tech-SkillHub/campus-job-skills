# 数据建设与固定脚本（协议 v2）

本文件是公司交接契约的一部分。`scripts/`、测试和引用文档都属于 Skill，发布时与入口一起同步。实际公司资料和用户证据只保存在私有工作区。

## 每家公司必须走完的建设流程

1. 从已确认范围和原表建立公司研究记录。已有信息全部保留为来源线索；初始化不等于核验，不新增用户选择或批准。
2. 查官方入口、全部适用招聘计划与要求，查计划的完整岗位名称目录和分页；记录计划数、目录数及来源。招聘轮次、BU、计划属于不同维度，不能混为同一个名额上限。
3. 对候选读取完整 JD，分别记录职责档位、资格依据、推荐与用户选择。特殊计划展示要求并取得是否考虑的决定。
4. 查完公司/计划/轮次/BU/共享名额/同时流程/冷冻期/撤回换岗/志愿规则；按每个计划与 BU 查内推码。没有找到码也记录已查来源、日期和范围，不能空着跳过。
5. 核实是否设置志愿，记录志愿组、岗位、序号、名称、对应 BU/地点。无此机制时 `rank: null`、`status: not_applicable`；未查清为 `unknown`。`roles.order` 是处理次序，绝不能充当志愿序号。同一有序志愿组内序号唯一，不同组可以重复；平行志愿不填序号。未知志愿阻止已选岗位交接。
6. 用固定写入脚本更新原表和八张共享表，读回后同步全量 JSON、全部现有批次 JSON 与 HTML。运行 `audit` 列出准确缺口；完成检查后才向用户交付正式核对页，过程中可以展示清楚标注的草稿。
7. 用户核对当前 HTML 全部内容后，保存确认依据，运行 `confirm`；进入投递前运行 `check` 和投递阶段门禁。投递端继承确认的志愿；提交目标确认中的“意向”必须等于 `preference_target` 输出（无机制为“无”，有机制为“组/序号/名称/BU/地点”，多项用中文分号分隔）。页面选项变化先回写和重新确认，不现场自行换志愿。

`研究检查` 必须含 plans、catalog、jd、eligibility、rules、referrals、preferences 七项。每项状态为 todo / checked / blocked，并保存来源、核验时间或具体缺口。checked 表示已经执行这项研究，不代表所有事实都已得到肯定答案。计划和目录检查还需记录应有数量、实际数量，与明细行数相等；分页或官网总量不明则保留待查，不能宣称完整。规则“已查但未披露”可记录 unknown 和已查来源，依交接契约向用户披露；不能把未开展的研究直接标 checked。脚本检查字段、关系、数量与同步；官网来源是否充分由执行者逐项核对，脚本不能替代取证。

## 完整 JSON 的范围

`岗位交接/工作簿全量.json` 是**整本当前工作簿的完整业务数据**，包含所有工作表、所有有值单元格及坐标、原列名、历史与当前行、公式及可读取缓存值、单元格备注、真实超链接、合并区域及表范围。公司、招聘计划、全部岗位目录、JD、规则、内推、志愿、研究状态、推荐、选择、确认以及原表其他业务列均保留；不能仅导出当前候选、部分列或一个批次。

此文件不承诺重建 Excel 的样式、图片、图表或宏。它不是第二份独立维护的台账。整本业务 JSON 与 `岗位交接/<批次>/公司方案.json` 同时存在；后者是完整数据中的公司审核视图及相关原表行，方便交接，不能冒充全量。两者均绑定工作簿 SHA-256、数据摘要。HTML 从同版公司视图生成。

## 固定命令

在 Skill 脚本目录下调用，路径和批次由私有配置提供；输入 JSON 必须 UTF-8。

```bash
python workbook_store.py init --workbook <账本.xlsx> --batch <批次> --sheet <地区表> --scope <已确认范围>
python company_handoff.py write --workbook <账本.xlsx> --batch <批次> --input <增量.json>
python company_handoff.py audit --workbook <账本.xlsx> --batch <批次>
python workbook_store.py query --workbook <账本.xlsx> --company <公司> --sheet <指定表>
python workbook_store.py refresh --workbook <账本.xlsx>
python workbook_store.py check --workbook <账本.xlsx>
```

init 可加 `--company`，只建立指定公司；旧协议升级保留旧明细与回执，补研究检查和未知志愿，撤销旧协议确认，等待新版核对。新建公司不会把历史岗位重新认定为当前已核验目录；旧事实仍在原表、全量 JSON 及公司 existing_data 中，需按来源确认后增量写入新明细。

write 接收 `companies` 完整公司记录，或者 `company_patches` 增量（两者互斥）；可在同一输入内加 `workbook_updates` 写地区岗位、原校招计划、校招主链接及其他既有表。增量只提供变化字段，新条目仍需补足其必填字段；每次业务修改提供新 version，受影响公司旧确认清除。明细按 id 更新；需要删除条目时提供审阅后的 companies 完整公司记录。未知字段、重复键、过长单元格、覆盖已有公式或重置已提交会报错，禁止静默丢弃。

```json
{
  "company_patches": [{
    "company_key": "已有公司键",
    "version": "v2",
    "roles": [{"id": "已有岗位ID", "recommended": true}]
  }],
  "workbook_updates": [{
    "sheet": "地区岗位",
    "key": {"公司": "示例公司", "岗位链接": "https://example.test/job"},
    "values": {"内推码": "已核实的码及来源记录"}
  }]
}
```

原表键必须包含足以唯一定位的字段：岗位用公司和直链；计划用地区、公司、届别、计划（没有届别值时用可靠的其他非空身份字段）；主链接用地区和公司。不要依赖旧行号。首岗填公司占位行，追加岗位保持公司块；含公式表的中间插行不自动移动引用，需专门处理。新设已提交须提供私有 `receipt_ref`，实际回执仍按投递端证据流程核实。不得用脚本替代事实核验。

统一写入先验证输入、备份，保存后生成派生文件。Excel 与多个文件无法一起原子替换；中断时门禁拦截不一致文件，重跑 refresh 恢复，禁止回写旧 JSON。手动修改 Excel 后同样 refresh；没有后台监听。每次刷新会更新全部已经建立的批次，保留历史快照。`query` 只输出指定公司/表的已同步业务记录，避免每次把整本数据送入上下文。

## 全部业务维度的结构化要求

Excel 和 JSON 使用同一份字段契约。不得以备注、Markdown 或 `text` 原文代替下表字段；未知维度写明确状态和缺口，不写成无限制、不适用或零。完整 JD、官方原文和解释仍保留在文本字段，用作证据；执行条件另行拆分。

| 维度 | Excel 所在表 / JSON 字段 |
|---|---|
| 用户需求范围 | 公司核对 / scope_status、regions、cohort、employment_types、role_families、hard_cities、preferred_cities |
| 招聘计划与要求 | 计划明细 / category、window_status、window_start/end、requirements_status、degree_levels、majors、graduation_start/end；其他条件拆入条件明细，requirements 仅保留完整原文依据 |
| 岗位与三档 | 岗位目录 / id、plan_ids、round_id、title、bu、cities 原文、city_options、employment、tier、functional_fit、eligibility、jd、basis |
| 推荐、选择、处理次序 | 岗位目录 / recommended、selected、order，三者独立 |
| 投递限制适用对象 | 投递规则 / dimension、role_ids、plan_ids、bu_ids、round_id；多种范围按交集匹配，空列表表示该维度不再细分，scope 只是展示说明 |
| 公司/计划/BU/轮次/共享限额 | 投递规则 / limit_type、limit、used、usage_status、shared_group；limit_type 为 count / unlimited / unknown / not_applicable，数量必须是数值 |
| 轮次窗口与重置 | 投递规则 / round_id、round_start/end、reset_policy、reset_at；日期用 ISO 日期，reset_policy 为 none / round / annual / date / unknown / not_applicable |
| 跨计划和 BU 关系 | 投递规则 / relation 为 independent / shared / exclusive / conditional / unknown / not_applicable；plan_ids、bu_ids 记录实际对象 |
| 冷冻期 | 投递规则 / cooldown_triggers、cooldown_duration 数值、cooldown_unit、cooldown_until、cooldown_release；触发动作逐项记录，不能只留原文 |
| 撤回、换岗、重投 | 投递规则 / change_actions、quota_effect；影响为 retained / released / consumed / conditional / unknown / not_applicable |
| 内推码与适用性 | 内推信息 / code、cohort、plan_ids、bu_ids、role_ids、valid_from、expires、status、source、checked_at |
| 志愿 | 志愿明细 / role_id、group_id、rank、label、bu、cities、status；rank 可为 null，但必须区别已知无机制和未知 |
| 研究进度、来源、确认 | 研究检查七项、各明细 source/checked_at、公司确认字段和版本，独立于已提交状态 |

字段与类型由 `scripts/company_handoff.py` 及 `scripts/structured_dimensions.py` 共同定义。列表在 Excel 专用列以 JSON 数组保存，读回后恢复为真正的数组；数量/布尔/空值恢复对应 JSON 类型，不能变成一段自然语言说明。全量 JSON 的 `company_batches` 提供所有已建批次的结构化公司对象，`sheets` 保留所有原表业务单元格。旧记录仅被原样导出，不等于已拆分成新业务字段；初始化将其列为待补并清除旧协议批准。完成结构化拆分和来源核验后，才可将研究检查置为 checked、进入用户确认。

计划类别 category 为 regular / technical / product_trainee / management_trainee / other。`window_status` 为 known / not_disclosed / unknown；`requirements_status` 为 structured / not_disclosed / unknown。未披露必须有原始来源支持，不能用来代替未查。公司范围 `scope_status=verified` 后，空城市硬范围表示无城市硬限制；unknown 状态的空值不作该解释。内推的适用届别必须与公司本批届别一致。


不能归入现有专用列的要求也必须结构化写入 `条件明细`，例如培养期、轮岗次数、技术要求、语言、实习经历或特殊项目资格。每行只放一个条件：关联公司/计划/岗位/规则、具体维度 field、比较关系 operator、真实类型 condition_value、单位、条件组及 all/any 关系、必需/优先/说明、核验状态、来源和原文。例如培养两年写为 training_duration / eq / 2 / years；不能将“其他要求”整段塞进一个字段后标完成。比较关系支持 eq、neq、in、not_in、gte、lte、between、exists。Excel 条件值列保存 JSON 类型编码，JSON 中恢复数值、布尔、字符串或数组；数值比较不得存“若干年”之类文本。原文留在 raw，只作核对依据。

志愿机制必须独立记录 `mode` 和规则的 `preference_mode`：ordered（有先后次序）、parallel（平行志愿）、not_applicable（无此机制）、unknown（未核实）。平行志愿的 rank 保持 null，不能为了有数字而编造第一/第二优先级；目标确认使用“组/parallel/名称/BU/地点”。机制已查清但用户未选择时用 pending_selection；这与未知机制不同，也不能让已选岗位直接通过交接。志愿明细必须与适用规则的机制一致。

原表的日期字段可用 `{ "type": "date", "value": "2001-01-02" }` 传入，固定脚本写为真正 Excel 日期；数字和布尔传原生 JSON 类型，不能传字符串冒充。复杂条件依然逐条拆分为条件明细。
