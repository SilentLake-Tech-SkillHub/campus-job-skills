---
name: campus-job-parallel-search
description: Coordinate user-approved multi-Agent campus-job research with exclusive source assignments, independent evidence outputs and serial tracker reconciliation. Use after this search batch selects multiple AI participants; it does not prepare or submit applications.
metadata:
  version: "1.0.0"
---

# 多 Agent 共同搜索岗位

Load the [search parent](../../SKILL.md) and its current-action references first. This subskill changes coordination, not the confirmed job scope, research stopping rules, browser policy or authority. Keep personal values and actual company allocations in the workspace's private batch records.

## Query and dispatch

Before a new search batch opens employer pages, ask: `本轮岗位搜索由一个 AI 执行，还是由多个 AI 共同执行？` Combine this with other missing batch choices. Reuse an explicit choice for this same batch; if the user already requested multiple AI, ask only remaining decisions. An unanswered question does not enable parallel execution. Continue independent offline work while awaiting an answer.

Single AI returns to the parent workflow. Multiple AI requires the participant identities or requested count, coordinator, resource limit and execution channel: available platform subagents or task packages the user will hand to other AI. Use the actual available capabilities; do not invent sessions, insist on a fixed count or launch user-owned chats without permission. A package produced is not a participant started. Explicit multi-AI selection authorizes the agreed research delegation only.

Before dispatch, read [the private record contract](references/coordination-contract.md), reconcile existing research and generate versioned, exclusive assignments. Run `python3 scripts/validate_assignments.py <private-manifest.json>` relative to this subskill. It checks the allocation snapshot, not external research truth or a live lock. The coordinator records ownership changes and all participants acknowledge the latest version before operating.

## Execute and reconcile

- Allocate company/region/cohort scope to one owner. Canonical ATS (recruitment platform), plan and source keys expose aliases sharing the same search; map known identical sources rather than searching twice. Separate source scopes only with recorded official evidence. Existing visited, blocked or ambiguous-cycle entries follow the parent's stopping rules; they are not automatically new work.
- Each participant uses its own output directory and company browser windows. Preserve user and other participants' pages, especially review, login and unfinished forms. Browser retention/closure follows the parent's batch intent and actual tool policy.
- For each company, save official plans, company campus entry, every verified role's direct URL and **complete visible JD** immediately in the participant's delta. Preserve source sections, qualifications and material conditions; a summary alone is insufficient. Include filters/pages inspected, zero-result evidence, source mapping and exact remaining gaps. This delta is pending merge, not proof of workbook completion.
- Only the coordinator writes the shared workbook and public ledgers. Before each merge, reread its current Hash, back up the current copy, reconcile company/plan/job identity and accept only validated deltas. If the file changes during merge, discard the stale write candidate and reconcile against the newest copy. Use the workbook tooling and schema required by the parent.
- Read back saved company and role links, full JD, plan and status before reporting merged completion or releasing an eligible company window. If readback is delayed, protect the evidence/tab, record the gap and continue an independent allocated company when allowed. Zero-role outcomes still need plan/result coverage evidence.

## Handoff and report

On interruption, persist output and exact unfinished scope. The coordinator records a new assignment version, old owner stopping acknowledgement, new owner and recovery step; the new owner accepts that version before reopening sources. Uncertain ownership stops only the conflicting scope. Resume from saved evidence instead of restarting searches.

Report allocated companies, documented visits, verified coverage, mapped sources, received deltas, accepted deltas, workbook-readback successes, blocked/unprocessed scopes and outstanding gaps separately. A file count or blank row proves none of these. Search cooperation never authorizes form filling, résumé upload or submission; route a separately authorized application batch to the independent application workflow, with its own execution-mode Query.
