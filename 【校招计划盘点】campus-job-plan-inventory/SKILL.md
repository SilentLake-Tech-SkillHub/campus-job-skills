---
name: campus-job-plan-inventory
description: Inventory current campus-recruitment plans for companies in the career workbook, verify plan-to-plan application relationships, and maintain the 校招计划 sheet before role filtering.
---

# 校招计划初始化与刷新

Use only in a job-search workspace with a job-search profile (`产品管理/求职配置.md`), before `campus-job-application` opens or filters a company's roles. Read the workbook schema in `../【校招岗位收集与投递】campus-job-application/references/workbook-schema.md` before editing.

## One-time project initialization

1. Read the company list from both regional sheets. Create one coverage entry for every distinct region–company identity in `校招计划` before claiming initialization coverage. A blank `计划` with `盘点状态=未盘点` is a queue item, not a discovered plan.
2. Check each company's current official campus entry for plans applicable to the user's target graduating class. A company may have multiple plans, business-unit programs, graduate tracks, and dual-eligible internship programs. Record only what the current source supports; do not treat a historical program as current.
3. Replace a company's unexamined coverage entry with one row per verified plan. Retain a company-level coverage row if no plan can be verified, with `计划` blank and an explicit `盘点状态` such as `访问受阻` or `未发现当前计划`. Never turn a failed search into proof that the company has no plan.
4. For every pair of current plans, research whether applications are independent, mutually exclusive, share a quota, or have another condition. Put the counterpart plan and relationship in `投递关系`. If current evidence is missing, write `与<计划>：待核实`. Keep within-plan application limits in `计划内投递限制`, separate from cross-plan relationships.
5. Preserve source URL and date, report counts of unexamined, verified, unavailable, and relationship-unknown companies. A queue of hundreds of unexamined companies is not a completed research pass. Resume in batches without overwriting verified work.

## Company lookup and refresh

For each company in the current Query-confirmed region and company/industry scope, refresh that company's plans before role search. First reuse `校招主链接`; verify it is current, or search for and verify the official campus URL when absent. Then enumerate every relevant plan on the official page or current employer announcement. Match the user's cohort, role and city preferences to plans; do not assume a plan lacking a matching job is the company's only plan. Within each relevant plan, filter roles separately, retain that plan's filtered master page during verification, and hand every in-scope result (including all available pagination) to `campus-job-application` for separate detail-tab opening, verification and **immediate regional role-row recording** in both information-only and application-preparation modes. `校招计划` rows alone never constitute completed role coverage. Only conclude “no matching role found” after all relevant, accessible plans and their in-scope results have been checked; name any plan or result page that remained inaccessible or unsearched.

### Internship-to-full-time conversion candidates

- When the active Query targets formal/full-time campus roles but the official campus entry or a relevant target-family result surfaces internships, flag only the relevant target-family internships for detail-level conversion checks; do not expand to unrelated internship families.
- Inspect the official JD for explicit conversion/retention wording (`实习转正`, `留用`, `转正机会`), target cohort, minimum duration/attendance and performance or other conditions. Hand off a candidate only when the target-family JD is in-scope and the official page explicitly states a conversion opportunity. The role remains an internship, is separately labeled `实习转正候选` in the role-table `备注`, and is never counted as a formal/full-time role or guaranteed conversion.
- Ordinary internships without explicit conversion evidence remain out of scope. Ambiguous or inaccessible wording is recorded as unresolved/pending, not inferred. Track internship plan identity, cross-plan application relationship and limits separately; unknown remains unknown.

Use one dedicated Chrome window per company. Keep that company's master, plan-filtered results and role pages inside it; do not repurpose another company's window. Keep useful pages open while the company investigation remains active. At its end, hand control to `campus-job-application`'s **Close a completed company window** gate: the Query-confirmed `信息收集` mode authorizes closing an agent-created, unprotected company window after recording its outcome; `准备投递` retains it. Protected or user-owned pages remain open. Capturing plan metadata does not authorize opening application forms, uploading a résumé, filling, submitting, or closing user pages.

## Evidence and safety

- Prefer official pages, official APIs and employer announcements. A university-hosted employer announcement may corroborate a relationship, but record its exact URL and wording. Third-party summaries are leads, not sole proof of a quota or plan conflict.
- Do not infer independence from separate apply buttons or separate plan names. Do not infer conflict from a shared login or one account.
- Recheck time-sensitive rules and the logged-in account's remaining opportunities immediately before filling or submitting under the main Skill.
- Use the Spreadsheets Skill for every workbook edit, with a recoverable backup before structural changes and saved-file/visual validation afterward.
