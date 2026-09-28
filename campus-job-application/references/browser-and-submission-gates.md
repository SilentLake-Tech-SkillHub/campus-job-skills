# Browser and submission gates

Read this reference before opening/filtering roles, filling, or submitting any application.

## Chrome state

- Use the connected Chrome browser requested by the user so existing login state and user-opened tabs remain available.
- Open each company in a new, dedicated Chrome window. Keep that company's result, role-detail, and application tabs in its own window. Do not change an existing company's window or tabs into another company's page, even when they appear idle; cross-company reuse can cause Chrome to freeze. If a new window cannot be created, report the blocker instead of using an existing company window.
- Confirm the batch intent before opening company windows. In `准备投递` mode, keep filtered result pages and role details open for user selection and review. In `信息收集` mode, keep them open during that company's investigation; save and read back the official company campus URL plus **every verified in-scope direct role URL and JD** in the workbook, or record the exact pending blocker, before closing an eligible agent-created window and advancing. Window closure does not turn a blocked or partially recorded company into completed coverage. This choice never authorizes résumé upload, form filling or application submission.
- Do not close or repurpose a tab the user opened. Open a new tab when ownership is unclear.
- Exception for reviewed non-priority role tabs: after the user has confirmed the exact retained highlighted roles and then separately confirms the closing action, close only tabs that match recorded direct role links outside that retained set. Keep the master-link page, filtered result pages, login/account pages, retained tabs, and unidentified user tabs open.
- The user's confirmed `信息收集` batch choice authorizes closure of each investigated, unprotected **agent-created, company-dedicated** Chrome window after the dual-link saved-workbook readback (or explicit pending-blocker record), company identity check and tab count. In `准备投递` mode, require explicit company-specific closure authorization; a request to move to the next company is not that authorization.
- Before closing a company window, treat user-selected/retained role pages, a form awaiting review, submission receipts, and login/account pages required for a next approved action as protected. Name them and obtain a separate explicit confirmation for that exact window; never close user-owned or ownership-unclear windows. Existing windows created in an earlier batch cannot be bulk-closed on an unverified ownership assumption: inventory their company, ownership, and protected tabs first, and request confirmation for those that cannot be safely classified. Verify each closed window is gone and report its company and tab count.
- Do not switch to the in-app browser, another browser, or headless automation unless the user changes the requirement.

## Page and source checks

- Confirm the domain or page branding belongs to the company or its named recruiting platform.
- Confirm the visible cohort, recruitment type, location, role title, and product category whenever the site exposes them.
- For every filtered plan view, including a broad product view with no user preference, retain the filtered result page as the master page. Enumerate all visible in-scope results across available pagination and open each through its visible official result card in a separate detail tab (or the site's equivalent non-destructive navigation). Verify a loaded detail page with a role title and substantive job description before calling that role opened. Retain its actual direct URL and the master result page; a search-result card or a guessed URL is not detail-page evidence.
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
