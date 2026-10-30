# SPFx Playwright instructions

For any SPFx project. This dev tool does not require `@spfx-fast/core`.

## Start with the existing project

- Reuse the project's Playwright version, config, test account, and site.
- Set `SPFX_BASE_URL` to the SharePoint **web URL**. Set `SPFX_STORAGE_STATE` only to a local ignored file when that project already uses saved state.
- Never automate sign-in, inspect cookies, copy a daily profile, or commit auth state.
- Run one job per persistent profile. `run` and `record` lock it; close manually opened browsers first.

## Generate and run

- Record with `spfx-fast-playwright record`; review selectors and add user-visible assertions.
- For repeatable checks, validate a JSON recipe, then `generate` or `run` it.
- Prefer role/label locators and explicit waits. Do not use generic `networkidle` on SharePoint.
- Use `run --fast` during fixes, then run once without it.
- `assertAuthenticated` checks REST without reading account details.
- A click may set `confirm: true` to accept exactly one expected JavaScript confirmation; it fails if the dialog is missing or has another type.
- Keep navigation on the configured origin. Sign-in redirects fail the precondition.
- Test the real QA page and required data; missing web parts or zero-to-zero comparisons prove nothing.
- Before tenant deployment, read [the SPFx ALM deployment guide](../../../docs/AGENT-SHAREPOINT-DEPLOYMENT.md). Verify the catalog package state and the manifest component ID; don't choose a web part by a generic name.

## Prove and report

- Prefer read-only checks. Writes need dedicated records and cleanup; never change production data.
- The [tenant CRUD recipe](../examples/tenant-crud-validation.recipe.json) is only for a disposable list on a dedicated test page and account.
- Review generated specs. Inspect `summary.json`, screenshots, and trace after each run.
- Screenshots show appearance only. Assert API or saved/published state separately.
- After three failed attempts to reach a control, stop and report the blocker with evidence.
- Check artifacts for tenant/personal data before sharing.
- For image overviews, OCR, PDF text, and screenshot diagnostics, see [the browser, image, and OCR workflow](../../../docs/AGENT-BROWSER-WORKFLOW.md). OCR does not prove visual layout or complete text capture.
