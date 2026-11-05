# Search coordination record

Keep actual batch values out of this public package: the manifest and deltas are process files in a temporary directory outside the workspace, deleted once results are written to the master JSON. The allocation validator accepts this manifest contract:

| Field | Meaning |
|---|---|
| workflow | `search` |
| batch_id / query_version | Stable batch ID and confirmed scope version |
| mode / mode_approval_ref | `single` or `parallel`, with the user's explicit choice evidence |
| execution_channel | `platform` or `manual_handoff`; channel choice is not proof of launch |
| participants / coordinator | Distinct participant ID strings; coordinator belongs to the list |
| output_roots | Participant-to-relative-directory map; roots must not overlap |
| source_hash | SHA256 of the current source snapshot, verified separately before merge |
| assignment_version | Positive integer; each delta references this exact version |
| assignments | Allocation list described below |

Each assignment contains `assignment_id`, `company_key` (includes region/cohort when needed), `source_scope` (canonical ATS/plan/filter scope), `owner`, and `output_dir` inside that owner's root. A company key and canonical source scope each have one assignment; aliases belong in `aliases`, not duplicate allocations. Truly independent scopes need distinct composite company keys supported by evidence. Optional `handoff` contains `from_owner`, `to_owner`, `stopped_ref`, `accepted_ref`; both owners are declared participants, the new owner matches `owner` and neither acknowledgement can be inferred from file age.

Each delta carries the manifest's `batch_id`, `query_version`, `assignment_version`, `assignment_id`, `owner`, source Hash and checked time. Reject stale or foreign-version deltas until reconciled; do not silently update their version numbers. Preserve official company entry, applicable plans/relationships, current-cycle preflight, page/filter coverage, outcome, unresolved scope and `application_actions: none`. Each role carries stable identity, direct official URL, title, location and full JD (`jd_full_text` plus source capture reference). An optional concise summary cannot replace full JD. A list entry marked full text still requires semantic comparison with the source.

Merge evidence records delta validation, affected IDs and the company/role link and JD readback from the master JSON. Delta receipt, accepted evidence and master-JSON readback are different stages; retain failures and unmerged results until resolved. Statuses use only the workspace label configuration and the propose → confirm flow. Allocation validation grants no browsing authority, checks no external source and creates no cross-process lock. Operational exclusivity requires all participants to stop when their assignment version is superseded.

## Allocation decision evidence

Keep a readable proposal alongside the manifest: the user's explicit requirements or ideas, Query/progress references, chosen grouping and count rationale, estimated workload and uncertainty, participant/worker/coordinator counts, actual concurrent capacity, wave/handoff arrangement where needed, configurable numeric or nonnumeric targets and stopping conditions, and user confirmation reference. These decision records precede dispatch. The snapshot validator does not judge allocation quality, confirm a proposal, optimize participant counts or enforce platform capacity. Reuse an unchanged confirmed proposal; version material changes and preserve their effective time and approval.

Delta acceptance and user reports also follow the [progress contract](../../../references/research-progress.md): current-cycle attempt evidence, coverage class, exact gaps and merge/readback state are independent facts.

## 搜索端只读

搜索增量里的岗位只处于发现/核实阶段（`stage` 为空、`discovered` 或 `verified`），`execution` 的 fill/save/submit 全部为 `{"status": "not_started"}`。已投递的岗位不在搜索中改动。
