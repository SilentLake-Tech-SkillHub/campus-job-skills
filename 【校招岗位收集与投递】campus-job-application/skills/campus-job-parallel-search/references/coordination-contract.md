# Private search coordination record

Use workspace-owned records, not this public package, for actual batch values. The allocation validator accepts this manifest contract:

| Field | Meaning |
|---|---|
| workflow | `search` |
| batch_id / query_version | Stable private batch and confirmed scope version |
| mode / mode_approval_ref | `single` or `parallel`, with the user's explicit choice evidence |
| execution_channel | `platform` or `manual_handoff`; channel choice is not proof of launch |
| participants / coordinator | Distinct participant ID strings; coordinator belongs to the list |
| output_roots | Participant-to-relative-directory map; roots must not overlap |
| source_hash | SHA256 of the current source snapshot, verified separately before merge |
| assignment_version | Positive integer; each delta references this exact version |
| assignments | Allocation list described below |

Each assignment contains `assignment_id`, `company_key` (includes region/cohort when needed), `source_scope` (canonical ATS/plan/filter scope), `owner`, and `output_dir` inside that owner's root. A company key and canonical source scope each have one assignment; aliases belong in `aliases`, not duplicate allocations. Truly independent scopes need distinct composite company keys supported by evidence. Optional `handoff` contains `from_owner`, `to_owner`, `stopped_ref`, `accepted_ref`; both owners are declared participants, the new owner matches `owner` and neither acknowledgement can be inferred from file age.

Each delta carries the manifest's `batch_id`, `query_version`, `assignment_version`, `assignment_id`, `owner`, source Hash and checked time. Reject stale or foreign-version deltas until reconciled; do not silently update their version numbers. Preserve official company entry, applicable plans/relationships, current-cycle preflight, page/filter coverage, outcome, unresolved scope and `application_actions: none`. Each role carries stable identity, direct official URL, title, location and full JD (`jd_full_text` plus source capture reference). An optional concise summary cannot replace full JD. A list entry marked full text still requires semantic comparison with the source.

Merge evidence records delta validation, current workbook Hash, backup reference, affected IDs, saved Hash, company/role link and JD readback. Delta receipt, accepted evidence and workbook readback are different stages; retain failures and unmerged results. Use only the workspace's existing workbook status vocabulary. Allocation validation grants no browsing authority, checks no external source and creates no cross-process lock. Operational exclusivity requires all participants to stop when their assignment version is superseded.

## Allocation decision evidence

Keep a private readable proposal alongside the manifest: the user's explicit requirements or ideas, Query/progress references, chosen grouping and count rationale, estimated workload and uncertainty, participant/worker/coordinator counts, actual concurrent capacity, wave/handoff arrangement where needed, configurable numeric or nonnumeric targets and stopping conditions, and user confirmation reference. These decision records precede dispatch. The snapshot validator does not judge allocation quality, confirm a proposal, optimize participant counts or enforce platform capacity. Reuse an unchanged confirmed proposal; version material changes and preserve their effective time and approval.

Delta acceptance and user reports also require the [progress contract](../../../references/research-progress.md): result status, current-cycle attempt evidence, coverage bucket, exact gaps and merge/readback state are independent required fields. A legacy delta lacking them is received evidence pending adaptation and review, not an accepted completed delivery.

## 独立偏好与执行证据

manifest 顶层 preferences、prior_roles；assignment/role/delta 必填 preference_version 与当前确认版本一致。roles 按 [偏好契约](../../../references/preference-contract.md) 携带所有逐岗字段及 pre_form_report_ref、execution；仅搜索的 fill/save/submit 为 not_started。候补公司覆盖字段放 assignment；旧版本拒收，已投/选岗保护，用户指定复查范围外不重查。

## 搜索中的只读历史保留

本批新发现岗位保持 discovered/verified（或未填写 stage）及 execution 三项 not_started。保留既有已投岗位时使用 historical_readonly:true；原 stage/receipt_ref/execution/材料与审核确认身份不变，prior_roles 保存其完整执行快照以逐字段比对，另用 current_batch_execution 的 fill/save/submit 全 {status:not_started} 明确本批未操作。顶层和 delta 都校验此例外。没有完整历史快照的旧记录先保留原件并适配/人工核验，不能重置原提交证据、冒称历史或以历史标记授权新表单动作。
