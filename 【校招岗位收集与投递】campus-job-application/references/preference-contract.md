# 范围、独立偏好与逐岗证据契约

Query、计划盘点、收集、Prepare、Review、续跑与 closeout 均读取本契约。仅提供机制；实际取值、用户原话和确认依据留在私有工作区。空模板复制后才能填写，不把个人偏好、生日、政治民族、联系方式或简历带入公共包。

## Query：硬范围与三个轴分别确认

- `hard_scope` 只记录明确排除的公司/届别/用工/城市/岗位职责等边界。`cities: []` 表示不限城市；不限城市不等于城市没有排序。低偏好项仍在范围内，不能当作硬排除，也不能把站点偏好过滤后的零结果当作公司没有岗位。
- `axes.employment`、`axes.city`、`axes.direction` 各自保存用户原话和独立关系，不计算合成权重。每条关系为 `{higher, lower, relation}`，`relation` 可为 `strict`（>）、`weak`（>=，只知不低于）、`tied`（=）、`unknown`（未决）。弱优先不擅改严格优先，不把不同轴的关系拼成一个总顺序。
- 按就业路径 → 城市 → 方向分组展示属于信息组织。平级保持平级，未知标待确认，多城市岗位只展示一个身份，可附多个地点。跨轴冲突列两边 JD/城市/用工事实，请用户取舍；不得自动加权、用能力分破局、以默认第一项代选志愿或替换已明确选岗。
- `匹配度` 只评价完整 JD 与已确认能力材料的适配，不掺城市/路径偏好；城市只在确实改变 JD 能力要求且有依据时影响能力适配。Top5 是既有高亮展示容量，分数相同用原行序仅安排展示；不意味着自动保留/排除、优先投递或关闭其他岗位。
- 用户已给本批取值则复用，仅问剩余歧义。更新记录 `version`、生效时间、确认引用、替代版本及复查范围。保留已投状态、回执和明确选岗；不撤回、不重投、不自动重启首轮扫站。先复用已保存完整 JD，仅在用户指定公司/缺口内复查；旧版本增量拒收并等待人工协调，不能改版本号冒充新结果。

## 默认次序与公司动作前确认

每批开始时先向用户要本批默认岗位类别次序；同批已有明确回答可引用，新批不得沿用上一批。续跑的首回合若缺本批确认，也必须先问，确认后才能写行、重排或生成汇总。`default_role_order` 保存用户原话，`default_order_confirmation_ref` 保存确认依据，`default_order_batch_id` 绑定本批。

每家公司开始申请操作前，用同一份逐岗报告先展示公司、BU 或未披露、办公地、岗位及直链，再问“相对本批默认次序有没有变动？”确认后继续。该报告就是 `pre_form_report_ref`，公司次序确认保存为 `company_order_confirmation_ref`，`company_order_report_ref` 必须引用同一报告；`company_order_decision` 为 unchanged/changed，改变时保存 `company_role_order_raw`。不能拿仅登录或旧批确认代替。校园收集流程在交给投递助手时传递同一报告和确认，不重复建立独立确认清单。最终提交仍审核当前具体目标与表单版本；前面的次序确认不能授权提交，目标变化须重新确认。

## 就业路径：可配置的转正实习次序

`internship_policy` 由用户在私有项目选择，没有公共默认值：`excluded` 不收实习；`separate` 表示用户明确选择独立实习批次；`conversion_last` 只纳入官方明确转正机会、资格和条件且方向符合的实习，与其他岗位在同一份汇报里排最后。该配置改变展示优先级，不把岗位排除出主清单；有合适正式岗也保留转正实习，不以正式覆盖充分或“无合适正式岗”为展示前提。普通实习在 conversion_last 下不纳入，转正含糊/详情不可达标待核，不保证转正。

正式覆盖记录仍说明实际已查范围与缺口；覆盖不足不能宣称公司无正式岗，但不妨碍已核转正实习进入同一汇报。roles 的候选均为 primary；按同一报告顺序存放，conversion_last 的实习候选排在其他候选之后且 `display_position: last`。公开空模板不预填实际个人次序或政策。旧 conditional/fallback 快照先保留原件、按新确认版本适配，不静默改版本或重置历史。

## 标题＋完整 JD：职责决定方向

标题和分类用来发现岗位，不能用标题黑名单。特别是“数据产品”：仅数仓/ETL/报表职责可依据已确认方向偏好排除（保留 JD 原文依据）；实际负责 AI 产品设计落地、Agent 需求/功能/迭代的可以候选。不能仅凭标题纳入或排除；完整 JD 不可达标待核，不把摘要当完整 JD。脚本只检查字段与声明之间的明显矛盾，不能验证自然语言 JD 真伪、职责理解或“覆盖充分”的现实真实性。

## 职责符合范围与个人任职资格

搜索与投递必须执行[资格核对](eligibility-check.md)：匹配分只表达能力适配，任职资格另作符合/不符合/待核判断，逐项保留JD条件性质和个人材料证据。职责符合范围的研究岗位可以保留资格缺口，但不直接成为可投候选。投递前继承并复核；前置岗位包和Review都展示决定性条件、资格结论、未决事实及用户具体尝试指令（如有）。不把必需和优先混为一谈，不把材料没写当成不满足，不改变现有Excel枚举。

## 身份、逐岗前置报告与阶段证据

身份键优先 `company_key + ':id:' + official_job_id`（company_key 含官方 ATS/招聘主体）；无官方 ID 则 `company_key + ':url:' + canonical_url`。规范直链只去已知跟踪参数 utm_* / ref / source，保留岗位、BU、计划等业务参数；不同 ID 的同名不同 BU/城市不可合并，一个官方多城市 ID 不膨胀成多个岗位。身份待核时保留待核证据，不造 ID 或猜直链。

每个已核岗位在表单前的报告和最终报告必须包括：公司、岗位、直链、BU（缺失写 `未披露`）、官方城市、用工性质、方向、完整 JD 与捕获引用、决定性 JD 依据、纳入/排除/待核理由、偏好版本、覆盖缺口与转正依据、实际填写/保存/提交状态及证据。未执行明确写 `not_started`，不能只报已登录。被排除或待核条目也保留依据和缺口，不误计为已核/已准备。

`pre_form_report_ref` 指表单动作前已经展示的逐岗包，`report_precedes_form: true` 是执行者声明，必须现场核验；发现、信息收集或排除记录也要有逐岗报告，未开表单时该布尔字段为 true 表示仍在动作前。Review 还包含材料、字段读回、表单版本、账户名额与未决项，目标和表单版本的既有审核仍生效。`execution` 三个独立项 `fill/save/submit` 各为 `{status, evidence_ref}`；允许 `not_started/in_progress/complete/blocked/uncertain`，除未开始外必须有证据引用。登录不等于填写，显示已填不等于保存，点击不等于回执。`submitted` 仍须父校验器的回执要求；`submit.complete` 也必须为 submitted 且有回执。

## 私有 JSON 字段（不新增 Excel 列或枚举）

所有新快照包含 `preferences`；批次台账放在 `batch.preferences`，搜索进度/并行 manifest 放顶层。它含非空 `version/confirmation_ref/effective_at`、`supersedes_version`（null 或旧版本）、`recheck_scope`（用户指定缺口列表；空列表不授权复查）、对象 `hard_scope`（`cities` 为字符串列表）、三个轴各含 `raw/relations`、本批 `default_role_order/default_order_confirmation_ref/default_order_batch_id`、`internship_policy` 及 `conflict_policy: user_decides`。各岗位和增量携带同一 `preference_version`。

`roles` 保存 `role_key/company_key/official_job_id`（无 ID 用 canonical 直链）、`company/title/direct_url/bu/cities/employment_type/direction/jd_full_text/jd_capture_ref/jd_basis/decision_reason/decision_basis`（`jd` 或 `scope`）、`hard_scope_pass`（布尔）、`disposition`（primary/pending/excluded）、`preference_version/pre_form_report_ref/report_precedes_form/execution`。conversion_last 的实习候选另有 `conversion_ref` 与 `display_position: last`；公司有无合适正式岗结论通过 `no_suitable_formal_found: true` 明确声明。只有这类无正式岗结论才需要 `formal_coverage: sufficient`、非空 `formal_coverage_refs`、`suitable_formal_found:false` 和 `no_suitable_formal_ref`；它们不是转正实习展示门槛。

跨轴冲突在岗位上记录 `cross_axis_conflict: true`；选择/替代它需要 `selected: true` 与 `selection_ref`，否则仅展示且不执行表单。`prior_roles` 是私有上一快照的身份/`stage/receipt_ref/selected`；当前快照必须保留已投及明确选择，任何获明确新指令的更改留 `change_authorization_ref`。偏好更新本身不是这项授权。缺少旧版本字段或逐岗报告的历史增量先适配和人工复核，不能作为新完整交付接收。

运行父流程列出的校验脚本检查新快照。它们是技能内结构校验工具，不是平台强制 hook、锁或真实网站验收；人工仍核验来源、覆盖、前置报告时间和旧记录完整性。

## 搜索中的只读历史保留

本批新发现岗位保持 discovered/verified（或未填写 stage）及 execution 三项 not_started。保留既有已投岗位时使用 historical_readonly:true；原 stage/receipt_ref/execution/材料与审核确认身份不变，prior_roles 保存其完整执行快照以逐字段比对，另用 current_batch_execution 的 fill/save/submit 全 {status:not_started} 明确本批未操作。顶层和 delta 都校验此例外。没有完整历史快照的旧记录先保留原件并适配/人工核验，不能重置原提交证据、冒称历史或以历史标记授权新表单动作。
