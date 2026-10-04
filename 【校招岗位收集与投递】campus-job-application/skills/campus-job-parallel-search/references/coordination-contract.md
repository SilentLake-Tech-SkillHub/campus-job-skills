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
