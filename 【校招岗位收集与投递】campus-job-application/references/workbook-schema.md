# Workbook schema

Read this reference before creating or updating role rows in the workbook named in the job-search profile (`产品管理/求职配置.md`).

## Sheets

- Regional role sheets: one per region listed in the job-search profile (for example `中国大陆` and `香港`), each holding the source companies of that region.
- `分类说明`: preserve the source classification guidance and add only concise notes needed to explain role-row behavior.
- `校招主链接`: reusable company-level index of verified official campus-recruitment landing pages. It is separate from the regional role tables and must not contain role-level JD, status, or application data.
- `校招计划`: one row per verified company–plan, plus one blank-plan coverage row for each company not yet verified. It stores cross-plan application relationships separately from within-plan limits. This sheet does not replace the company master-link index or the role table.

Do not merge the regional sheets or create per-company sheets.

## Campus recruitment master-link index

`校招主链接` uses these columns:

1. `公司`
2. `地区`
3. `校招主链接`
4. `适用届别`
5. `投递窗口`

- Store one current verified official campus-recruitment entry page per company. Use it as the first navigation source for future Open/filter work.
- A `校招主链接` is a company-level landing, directory, or current-cycle campus page; it is never a direct role page and does not replace `岗位链接` in a role row.
- Keep the visible cohort or cycle only when the official URL/page makes it clear. If the page is current but has no cohort label, leave `适用届别` blank rather than guessing.
- Write a fully disclosed window as `YYYY年MM月DD日-YYYY年MM月DD日可投递`. If an official source provides a refresh date but no closing date, record the date and `截止日未披露`; do not derive a date from a previous cohort, a quota rule, or a third-party listing.
- Refresh a stale master link in place after reporting the change. Do not create a duplicate company entry or populate any role fields in this index.

## Column order

1. `地区`
2. `公司`
3. `主要行业`
4. `企业性质`
5. `岗位名称`
6. `工作地点`
7. `岗位JD`
8. `所属BU`
9. `岗位链接`
10. `投递状态`
11. `匹配度`
12. `匹配原因`
13. `校招冷静期`
14. `校招计划`
15. `备注`

`校招计划` in a real role row names the verified source plan. Leave it blank when the source plan cannot be identified; do not infer it from the company. Keep a placeholder company's role-plan cell blank.

`备注` records material, evidence-based role facts that do not fit the existing fields, especially an explicitly stated internship-to-full-time conversion opportunity, cohort eligibility, minimum duration/attendance and conversion conditions. Preserve `实习` in the title and plan; label an eligible case `实习转正候选` and quote the official evidence concisely. State that conversion is an opportunity, not a guarantee. Ordinary internships without explicit official conversion wording remain out of a formal/full-time batch. Leave the note blank when no material note is supported. This field does not imply user selection or application authorization.

## Campus plan inventory

The `校招计划` sheet begins with the user-specified columns:

1. `公司`
2. `计划`
3. `投递关系`
4. `地区`
5. `适用届别`
6. `计划链接`
7. `计划内投递限制`
8. `依据链接`
9. `核验日期`
10. `盘点状态`

- Identify a row by region, company, cohort, and plan name. Several current plans for one company occupy separate rows.
- A blank `计划` with `盘点状态=未盘点` is an initialization queue item, not a plan. Keep an explicit status such as `访问受阻` or `未发现当前计划` after an unsuccessful attempt; never claim no plan exists from an inaccessible page.
- Write `投递关系` against named counterpart plans, e.g. `与JDS-新星计划：可同时投递；与新锐之星：待核实`. If a company has only one verified plan, write `暂无其他已核实计划` rather than asserting no others exist. The relationship can be independent, mutually exclusive, shared-quota, conditional, or pending verification only when supported by current evidence.
- Record plan-internal quota, one-time opportunity, cooling period or refresh in `计划内投递限制`, not as cross-plan conflict. Link the exact current source in `依据链接` and record `核验日期` as an Excel date. If one source does not support every assertion, list the specific supporting links in that cell.
- The one-time initialization covers every company in the approved region/scope with either verified plan rows or a blank-plan coverage row. For the first-pass sweep, count prior research only when it covers the target graduating cohort/current recruiting cycle. Distinguish companies/source scopes with no such coverage from already visited scopes whose outcome is partial, blocked, or relationship-unknown. Reuse coverage across aliases/business-unit rows when the official ATS tenant, plan, filters, result pages and job IDs match, and document the mapping. Only uncovered scopes enter the first-pass queue. Report first-pass-unvisited and first-pass-visited counts separately, then list unresolved items for possible user-selected follow-up. Prior-cycle research remains useful source context but does not count as current-cycle coverage; such a company still needs its first pass for the current cycle. A `部分盘点`/`访问受阻` status with an unclear cycle or evidence is deferred as already visited to prevent accidental repetition. A partial status, an old plan row, or a missing role row does not automatically trigger follow-up. At the end of Stage 1, report gaps and wait for user selection; a generic `继续` does not authorize Stage 2. In a user-directed follow-up, inspect only the selected company and exact gap.

`岗位名称` is user-approved and required to distinguish several roles at one company.

## Row semantics

- Each actual role occupies exactly one row.
- A company with no **verified in-scope role** keeps one placeholder row containing only the four company-level fields; all role-specific fields remain blank. In information-collection mode, user selection is **not** required to create a row for each verified in-scope role.
- For the first role at a company, populate the existing placeholder row.
- For additional roles, insert an adjacent row and copy the four company-level fields.
- Use the direct official role URL as the primary duplicate key. If the site changes the URL, compare company, role title, BU, location text inside the JD, and visible job ID before deciding whether it is new.
- Preserve submitted rows. A refresh may update a changed JD or URL only after reporting the observed difference; never replace submission evidence with a newer listing silently.

## Values

- `岗位JD`: copy the complete visible job description faithfully. Preserve headings and paragraph separation in one wrapped cell. Do not summarize unless the site blocks full access, in which case record the visible content and limitation.
- `工作地点`: copy the explicit official work-location text exactly as shown on the role page. If absent, leave blank or use `未披露`; do not infer a work location from the BU, company headquarters, office network, or incidental geographic text in the JD.
- `所属BU`: use the site's explicit business unit, department, team, product line, or employing entity. If absent, use `未披露` rather than guessing.
- `岗位链接`: for a verified role row, use its actual direct official role-detail URL, opened from an official result card or equivalent site navigation. Do not substitute a result-list, guessed, expired or generic URL and call that role complete. If no direct detail URL can be verified, retain the official result URL as separate research evidence, mark that role/company scope pending and report the limitation; never fabricate a direct link.
- `投递状态`: blank for company placeholders; `待投递` for every real role until a verified receipt exists; `已提交` only after verified success. `待投递` means recorded but unsubmitted, not selected by the user or approved for form filling.
- `匹配度`: integer from 0 to 100, based on the complete JD, current user-confirmed résumé, verified project portfolio, and explicit user self-assessment. Leave blank when the required candidate material is unavailable. It is a relative prioritization score, not a hiring probability.
- `匹配原因`: concise evidence-based explanation using `匹配：...；缺口：...`. Separate résumé-visible evidence from portfolio-only evidence when that difference could affect screening.
- `校招冷静期`: verified official constraints on applications for the relevant campus-recruitment cycle, including a quota, cooling period, annual refresh date, or an immutable post-submission rule. Leave blank for company placeholders and where no official rule has been verified; blank never means unrestricted. Before filling or submitting, recheck the logged-in account's displayed remaining count. Reuse verified rules recorded in the job-search profile.

Do not infer mastery from project labels or technology keywords. Follow the user's self-assessment recorded in the job-search profile when it conflicts with keyword evidence.

## Visual priority marking

- In a reviewed batch with scored real roles, apply pale-yellow fill across the full row of only the five highest `匹配度` rows; when fewer than five are scored, mark all scored rows. Use current worksheet row order only to resolve a tie.
- Do not color company placeholders or change any non-top-five row. This is a visual cue only: it must not alter `匹配度`, `投递状态`, direct links, or existing status validation and conditional formatting.

Do not add extra status values without a user-approved schema change.

## Workbook safety

- Before structural edits, record the source SHA-256 and make one recoverable backup.
- Preserve all source companies and classification text. Check company counts against the expected baseline recorded in the job-search profile.
- Use filters, freeze the header row, wrap JD cells, keep hyperlinks readable, and use a list validation for nonblank role statuses where supported.
- On every edit, reload the latest workbook, update the smallest affected region, recalculate once, inspect key values and formula errors, render the changed sheet, export to the original workbook path, and verify the saved file.
- Before a researched company is marked complete or its research window is closed, reopen the saved workbook and read back both the company's verified `校招主链接` row and every verified in-scope regional role row's direct `岗位链接` and JD. If no role was found, preserve the official plan/result-page and filter evidence for the zero-result conclusion. A missing direct link, unsaved row, inaccessible detail or unsearched plan keeps the affected scope pending.
