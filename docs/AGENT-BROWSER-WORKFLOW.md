# Browser checks for agents

Prefer semantic checks. Review tenant screenshots and traces before sharing.

The npm package includes this guide. Repository probes (`node scripts/...`) require a source checkout; npm consumers can run `@spfx-fast/playwright` recipes.

## Signed-in SharePoint

- Use a page set aside for tests. Do not edit a page another person is testing.
- Run one job per signed-in browser profile. The SPFx Fast runner locks the profile. Close manually opened Chrome first.
- Use synthetic data. Get approval before writing. Read saved state back after a write.
- A successful REST check proves access to that web only. It does not prove Graph access, web-part load, or write success.
- For app-catalog upload, deployment, and site installation, follow the [ALM deployment guide](AGENT-SHAREPOINT-DEPLOYMENT.md). Do not use guessed catalog URLs or confuse catalog row IDs, package ProductIds, and web-part component IDs.

```sh
pnpm add -D @spfx-fast/playwright @playwright/test playwright
pnpm exec spfx-fast-playwright init
pnpm exec spfx-fast-playwright validate tests/e2e/sharepoint-smoke.recipe.json
pnpm exec spfx-fast-playwright run tests/e2e/sharepoint-smoke.recipe.json \
  --base-url "$SPFX_BASE_URL" \
  --profile-dir "$SPFX_TEST_PROFILE" \
  --artifacts test-results/sharepoint-smoke
```

Runs are headless. Use `--headed` to watch. Use `--fast` during fix loops; it skips success screenshots and traces. Run once without it after fixes.

Set `--base-url` to the target web. Recipes use explicit paths, e.g. `/sites/your-test-site/SitePages/E2E-Test.aspx`.
With npm, use `npm install -D @spfx-fast/playwright @playwright/test playwright` and `npx spfx-fast-playwright` for the same commands.

The [dashboard CRUD recipe](../packages/playwright/examples/tenant-crud-validation.recipe.json) creates, verifies, updates, verifies, and deletes one item. Use a disposable list, test page, and account. It needs `Sites.ReadWrite.All`; inspect the result and clean up manually if needed.

## Failed checks

Before changing selectors, inspect the screenshot, trace, console, failed requests, and page. Prefer roles and labels. After two similar failures, inspect the DOM/accessibility tree and network response.

Confirm the expected bundle on the published page; editor drafts and Workbench alone can miss SharePoint wire-format differences.

Find the page in **Site Pages**; do not assume `Page.aspx`. `scripts/sharepoint-validation-page.mjs` creates or verifies the dashboard page from inspected list IDs. Confirm the component and current bundle before testing.

Run `pnpm test:spfx-fixture` after changing fixture components. Root TypeScript and browser checks do not catch SPFx Heft lint rules, older compiler targets, or fixture `rootDir` imports. Keep web-part components inside the fixture's `src` tree and use APIs available in its configured target.

Graph drive delta is current state, not event history. Label modified-date lists as inventory.

`/versions/{id}/content` downloads prior versions. Use `/items/{id}/content` for current bytes. Timestamps can tie; sort by version ID as a tie-breaker.

The tenant fixture checks 10, 50, 100, and 128 MiB + 1 byte Graph upload round-trips. The dashboard's **Test upload recovery** simulates a lost response. For a real browser-offline interruption, run:

```sh
node scripts/sharepoint-upload-offline-check.mjs \\
  --site "$SPFX_TEST_SITE" \\
  --page "$SPFX_VALIDATION_PAGE" \\
  --profile-dir "$SPFX_TEST_PROFILE"

node scripts/sharepoint-upload-offline-check.mjs \\
  --site "$SPFX_TEST_SITE" \\
  --page "$SPFX_VALIDATION_PAGE" \\
  --profile-dir "$SPFX_TEST_PROFILE" \\
  --post-window
```

Both modes disconnect after SharePoint accepts the first 10 MiB chunk, then verify upload, fresh download, and cleanup. Default reconnects after one failed status read. `--post-window` waits for all six failures, resumes the same session, verifies name and SHA-256, and checks the UI. Fixture 1.0.0.53 passed with 128 MiB + 1 byte. The profile must be signed in; run one job per profile and confirm before upload.

Run `node scripts/sharepoint-upload-final-response-check.mjs` with the same `--site`, `--page`, and `--profile-dir` to test a lost final response. It verifies bytes and cleanup. Add `--replace` for overwrite recovery or `--rename` for Graph's duplicate-name behavior. Fixture 1.0.0.60 passed with 20 MiB + 1 byte; recovery accepts only a bounded name match with identical bytes.

If retries expire, `uploadDriveItem` throws `DriveUploadInterruptedError`. Keep it in page state and call `resume()` after reconnecting. It holds the Blob and private session URL in memory; it cannot survive reload. Never serialize it or expose the URL.

Delta workloads use synthetic data: 250 rows with memory cache; 1,000–3,000 mixed-field rows and 3,000 folders with IndexedDB. They test volume, not business-list performance. Separate seed, sync, and query timings. Run `node scripts/sharepoint-delta-workload-check.mjs --list-only` or `--drive-only` to repeat one workload.

Calendar checks need `Calendars.ReadBasic` and an active mailbox enabled for REST. `MailboxNotEnabledForRESTAPI` indicates mailbox setup, not missing consent.

## Images and OCR

If `agent-tools` is installed, make a smaller overview. Keep the source image:

```sh
agent-tools image-prepare screenshot.png /tmp/overview.png 800
agent-tools image-text screenshot.png --json --max-chars 3000
agent-tools image-crop screenshot.png /tmp/toolbar.png 0 0 900 180
```

Try PDF text before OCR. OCR only the needed pages:

```sh
agent-tools pdf-text guide.pdf 1 20
agent-tools pdf-ocr guide.pdf 3 5
```

OCR can miss text and layout. Check the source when accuracy matters. Smaller images may use fewer tokens, but savings depend on the model.

## SharePoint gotchas

- Do not add `loadSPFX=true`; it can stop the page from loading the web part.
- Handle the debug-manifest prompt after each navigation.
- Close the toolbox before searching; two open toolboxes make locators ambiguous.
- A page edit may stay in draft. Readers see the published page.
- Search the toolbox for the exact manifest title and verify the component ID. A generic name such as `Dashboard` may be a built-in SharePoint web part.
- `200` from app-catalog `Deploy` or `Install` is only request acceptance. Read `AvailableApps` from the target site and confirm `InstalledVersion === AppCatalogVersion`; `GetAppInstanceById` gave a false negative for an app present in Site Contents. Then verify the expected component and loaded bundle on a published page.
- If an install call says the app instance already exists while site read-back says it does not, stop and inspect the app state. Repeating Install can leave the same conflict.

These checks complement tenant-specific permission and write validation. `agent-tools` is optional and not a runtime dependency.
