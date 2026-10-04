---
name: campus-job-parallel-search
description: Coordinate user-approved multi-Agent campus-job research with exclusive source assignments, independent evidence outputs and serial tracker reconciliation. Use after this search batch selects multiple AI participants; it does not prepare or submit applications.
metadata:
  version: "1.0.2"
---

# 多 Agent 共同搜索岗位

Load the [search parent](../../SKILL.md) and its current-action references first. This subskill changes coordination, not the confirmed job scope, research stopping rules, browser policy or authority. Keep personal values and actual company allocations in the workspace's private batch records.

## Query and dispatch

Before a new search batch opens employer pages, ask: `本轮岗位搜索由一个 AI 执行，还是由多个 AI 共同执行？` Combine this with other missing batch choices. Reuse an explicit choice for this same batch; if the user already requested multiple AI, ask only remaining decisions. An unanswered question does not enable parallel execution. Continue independent offline work while awaiting an answer.

Single AI returns to the parent workflow. Multiple AI requires the participant identities or requested count, coordinator, resource limit and execution channel: available platform subagents or task packages the user will hand to other AI. Use the actual available capabilities; do not invent sessions, insist on a fixed count or launch user-owned chats without permission. A package produced is not a participant started. Explicit multi-AI selection authorizes the agreed research delegation only.

Before dispatch, read [the private record contract](references/coordination-contract.md), reconcile existing research and generate versioned, exclusive assignments. Run `python3 scripts/validate_assignments.py <private-manifest.json>` relative to this subskill. It checks the allocation snapshot, not external research truth or a live lock. The coordinator records ownership changes and all participants acknowledge the latest version before operating.

## User requirements and allocation proposal

Before proposing a split, proactively ask `对 AI 人数、分工方式、优先级、时间或任务数量，你有明确要求吗？` Reuse explicit answers already recorded for this batch. When there is no definite requirement, invite the user's own ideas: `你希望怎样推进？可以说说最看重的方向、覆盖面、速度，或你愿意投入的时间；暂时没有具体想法也可以告诉我，我会给你建议。` Do not make the user supply a finished allocation or a numeric answer. A user explicitly asking for recommendations may receive a reasoned proposal without another redundant question; silence does not confirm it. Continue independent evidence reconciliation while awaiting missing answers.

Derive the proposal from the latest confirmed Query, the user's ideas, existing progress, uncovered independent scopes, priority/deadline, estimated effort and uncertainty, browser/account dependencies, available participants and actual tool capacity. Choose grouping dimensions that fit this batch; employer size, industry and historical task packages are possible evidence, never universal buckets. Search groups canonical source/plan coverage; application groups recruiting entities and shared account/history/official-quota dependencies. Keep dependent work together, balance estimated effort rather than company counts alone, and disclose estimates that have not been measured.

If the user supplies a participant count or grouping requirement, honor it within actual execution constraints and explain any conflict. Otherwise recommend a justified count and allocation; do not demand that the user choose a number first or reuse a previous batch's count. Distinguish total participants, active workers, coordinator work and simultaneous platform capacity. More participants than current capacity can work in successive waves or through user-managed handoff, if supported and confirmed. A tool's measured capacity is an execution constraint, not a universal Skill limit; never claim unavailable agents exist or bypass tool controls.

Present a readable proposal with the Query/evidence references, proposed participants and coordinator, each owner's scope and outputs, grouping/count rationale, estimated workload, time/resource constraints, quantity/stop conditions and known gaps. Ask the user to confirm or adjust it before dispatch. An execution-mode choice alone does not approve an inferred allocation. Record the confirmed proposal and evidence in private batch records, then produce the versioned manifest. Reuse the unchanged confirmed allocation on continuation; material scope, ownership or count changes require a revised proposal and the handoff controls below.

Business targets are batch choices. Ask whether the user wants a numerical target, coverage of a confirmed list, a time window, or another explicit stopping rule. Do not impose a fixed number of companies, roles or applications, force equal-size groups, or silently fill missing numbers from earlier batches. A numeric target is a goal unless the user explicitly makes it a ceiling or stopping condition. Changing a confirmed goal needs an effective-time record and user agreement; absence of a numeric target does not expand the confirmed scope. Official application/account limits and exact-role submission approval remain binding.

## Execute and reconcile

- Allocate company/region/cohort scope to one owner. Canonical ATS (recruitment platform), plan and source keys expose aliases sharing the same search; map known identical sources rather than searching twice. Separate source scopes only with recorded official evidence. Existing visited, blocked or ambiguous-cycle entries follow the parent's stopping rules; they are not automatically new work.
- Each participant uses its own output directory and company browser windows. Preserve user and other participants' pages, especially review, login and unfinished forms. Browser retention/closure follows the parent's batch intent and actual tool policy.
- For each company, save official plans, company campus entry, every verified role's direct URL and **complete visible JD** immediately in the participant's delta. Preserve source sections, qualifications and material conditions; a summary alone is insufficient. Include filters/pages inspected, zero-result evidence, source mapping and exact remaining gaps. This delta is pending merge, not proof of workbook completion.
- Only the coordinator writes the shared workbook and public ledgers. Before each merge, reread its current Hash, back up the current copy, reconcile company/plan/job identity and accept only validated deltas. If the file changes during merge, discard the stale write candidate and reconcile against the newest copy. Use the workbook tooling and schema required by the parent.
- Read back saved company and role links, full JD, plan and status before reporting merged completion or releasing an eligible company window. If readback is delayed, protect the evidence/tab, record the gap and continue an independent allocated company when allowed. Zero-role outcomes still need plan/result coverage evidence.

## Handoff and report

On interruption, persist output and exact unfinished scope. The coordinator records a new assignment version, old owner stopping acknowledgement, new owner and recovery step; the new owner accepts that version before reopening sources. Uncertain ownership stops only the conflicting scope. Resume from saved evidence instead of restarting searches.

Report allocated companies, documented visits, verified coverage, mapped sources, received deltas, accepted deltas, workbook-readback successes, blocked/unprocessed scopes and outstanding gaps separately. A file count or blank row proves none of these. Search cooperation never authorizes form filling, résumé upload or submission; route a separately authorized application batch to the independent application workflow, with its own execution-mode Query.

Before accepting a delta or reporting progress, load [the mandatory progress contract](../../references/research-progress.md) and run its validator. Read current tracker notes and dated evidence; a raw incomplete label cannot erase a documented first-pass visit. Legacy outputs require an evidence-preserving audit adapter before acceptance.
