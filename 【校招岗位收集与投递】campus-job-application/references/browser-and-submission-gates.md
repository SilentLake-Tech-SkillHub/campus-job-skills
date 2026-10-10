# Browser and submission gates

Read this reference before opening/filtering roles, filling, or submitting any application.

## Browser selection and retained tabs

- Browser priority is **in-app browser → Chrome + computer use → Playwright framework**. Start in the in-app browser. Before each downgrade, show the actual failure, proposed next tier and effect on login/unsaved work, then obtain the user's explicit confirmation for that transition. Confirmation of one transition does not authorize the next. Preserve company tabs and state while waiting; do not automatically switch or bypass a site safety refusal. The in-app browser tool's own Playwright-named control API remains the first tier; it is not use of the separate Playwright framework.
- Read [browser tab grouping](chrome-window-grouping.md); the ten-page Chrome window rule applies only to an explicitly selected Chrome branch. Each tab keeps its company identity; never repurpose another company's tab or a user window.
- Confirm the batch intent before opening company windows. In `准备投递` mode, keep filtered result pages and role details open for user selection and review. In `信息收集` mode, keep them open during that company's investigation; save and read back the official company campus URL plus **every verified in-scope direct role URL and JD** in the workbook, or record the exact pending blocker, before closing only that company's eligible agent-created tabs and advancing. Window closure does not turn a blocked or partially recorded company into completed coverage. This choice never authorizes résumé upload, form filling or application submission.
- Do not close or repurpose a tab the user opened. Open a new tab when ownership is unclear.
- Exception for reviewed non-priority role tabs: after the user has confirmed the exact retained highlighted roles and then separately confirms the closing action, close only tabs that match recorded direct role links outside that retained set. Keep the master-link page, filtered result pages, login/account pages, retained tabs, and unidentified user tabs open.
- The user's confirmed `信息收集` batch choice authorizes closure of the investigated company's unprotected **agent-created** tabs after the dual-link saved-workbook readback (or explicit pending-blocker record), company identity check and tab count. In `准备投递` mode, require explicit company-specific closure authorization; a request to move to the next company is not that authorization.
- Before closing a company window, treat user-selected/retained role pages, a form awaiting review, submission receipts, and login/account pages required for a next approved action as protected. Name them and obtain a separate explicit confirmation for that exact window; never close user-owned or ownership-unclear windows. Existing windows created in an earlier batch cannot be bulk-closed on an unverified ownership assumption: inventory their company, ownership, and protected tabs first, and request confirmation for those that cannot be safely classified. Preserve other companies in shared windows. Close a whole window only when all its tabs are Agent-owned, eligible and unprotected; verify the exact target tabs are gone and report their company and count.
- Recovery and downgrade: make one evidenced recovery attempt within the current tier while protecting unsaved work. If still blocked, ask for approval of the next tier in the stated priority. After Chrome + computer use fails, ask separately before the Playwright framework. A real failure is evidence for a request, not permission to downgrade. If approval or a supported next tier is unavailable, keep the company/URL/stage handoff and report the blocker.
- Visible retention: retain the entire confirmed company group in preparation/submission mode, including unfinished companies, accounts, forms and receipts. Before ending the turn, use the tool's documented visible-retention mechanism and read back company, URL and stage from the tab list. Where a visible deliverable and an internal handoff are distinct, retained pages must be visible deliverables; a handoff call alone is not proof of visible retention. Missing pages must be reported and recovered from original URLs, with saved content checked; reopening does not restore unsaved edits. Finishing one role or yielding for user action never authorizes cleanup of the remaining group.

## Page and source checks

- Confirm the domain or page branding belongs to the company or its named recruiting platform.
- Confirm the visible cohort, recruitment type, location, role title, and role category whenever the site exposes them.
- For every filtered plan view, including a broad target-family view with no user preference, retain the filtered result page as the master page. Enumerate all visible in-scope results across available pagination and open each through its visible official result card in a separate detail tab (or the site's equivalent non-destructive navigation). Verify a loaded detail page with a role title and substantive job description before calling that role opened. Retain its actual direct URL and the master result page; a search-result card or a guessed URL is not detail-page evidence.
- If the detail page is 404, blank, expired, unavailable, or redirects to a generic page, do not count it as a verified role or fabricate its direct link or JD. Retry once from the visible official result card; if still unresolved, report the failed page and leave that role and the batch's affected scope pending. Do not silently treat the failure as zero matching roles.
- If the in-scope result count is too large to open and inspect without causing unreasonable tab volume, do not sample or omit roles. Preserve the master page, report the count, and ask the user to narrow the role or city criteria before completing the batch.
- Take reviewable screenshots with enough browser context to identify the company, page state, and decision point.

## Filling boundary

- Filling is authorized only after the user has selected the specific role and asked Codex to continue.
- Use facts from the authorized résumé and other user-approved materials. Do not infer personal data from old or protected files.
- Leave credentials, OTPs, CAPTCHA, electronic signatures, sensitive identity declarations, and person-only attestations to the user.
- If the site auto-saves a draft or creates an application shell during filling, report that external state.
- Before stopping for review, check all visible required fields and upload controls, but do not activate the final submission control.

## Review gate

Provide a concise review package for each role:

- company and role title;
- direct role URL;
- résumé file used;
- material answers or uploads completed;
- unresolved or user-only fields;
- screenshots of the completed form and final submit area.

Keep the form open. `待投递` remains the workbook status.

## Submission gate

- Require explicit user approval after the review package. Approval applies only to the named roles in that package.
- Re-check the role title, company, résumé, and unresolved warnings immediately before each submission.
- Submit one role, wait for its result, and verify its receipt before moving to the next.
- On a timeout, blank page, ambiguous confirmation, or recoverable error, stop retries that could create duplicate applications. Inspect the application center or visible site state first.
- Mark `已提交` only when a visible receipt or application status verifies success.

## Screenshots and sensitive data

- Capture the minimum area that proves the state while retaining enough context to identify the site and step.
- Avoid exposing phone numbers, email addresses, addresses, identification numbers, birth dates, signatures, or uploaded document contents when the same evidence can be captured without them.
- Do not save passwords, cookies, OTPs, tokens, or identity-document contents in project files.

## 独立偏好与执行证据

Open/filter 依 [偏好契约](preference-contract.md) 仅用硬范围筛除，排序不限制覆盖。填写前逐岗展示公司、岗位、直链、BU 或未披露、城市、用工、方向、完整 JD 依据及筛选理由，持久化 pre_form_report_ref；退出/交付报告 fill/save/submit 的实际状态与证据，不以登录代替准备。Review 保留表单版本，提交继续精确目标与版本审核。

每批先向用户要默认岗位次序，保存同批确认，不沿用上一批。每家公司开始申请操作前，先用同一逐岗包展示公司/岗位/BU或未披露/办公地，确认相对默认次序有无变化；company_order_confirmation_ref 和 pre_form_report_ref 对应同一包。交接复用这份记录，最终具体目标和当前表单版本继续审核，不另建三套脱节清单。
