# @spfx-fast/playwright

Playwright helpers for any SPFx project. This dev tool stays out of the web-part bundle and uses the approved signed-in test browser. It does not depend on `@spfx-fast/core` or a UI framework.

It adds agent-friendly recipes, same-origin navigation, and authenticated runs with reviewable evidence. Playwright provides the test runner, locators, recorder, traces, and screenshots.

## Install in an SPFx project

```sh
npm install --save-dev @spfx-fast/playwright @playwright/test playwright
```

SPFx is not a dependency. Production web parts do not import this package by default.

Scaffold a serial Playwright setup with a starter recipe, generated test spec, agent instructions, screenshots, traces, and an ignored auth-state folder:

```sh
npx spfx-fast-playwright init
```

The config uses storage state only when `SPFX_STORAGE_STATE` is set. Keep it in an ignored folder; never commit it. The scaffold adds `SPFX-PLAYWRIGHT.md` and links it from `AGENTS.md` without replacing existing instructions.

## Record a test with the Playwright recorder

Use a dedicated signed-in test profile. The command calls Playwright Codegen and writes an ordinary Playwright spec:

```sh
npx spfx-fast-playwright record \
  --url "$SPFX_BASE_URL" \
  --profile-dir "$SPFX_TEST_PROFILE" \
  --out tests/e2e/home-page.spec.ts
```

Review selectors and add business assertions; Codegen records actions only.

## Generate a test mechanically from a small recipe

Edit the scaffolded `tests/e2e/sharepoint-smoke.recipe.json` for the target site, then:

```sh
npx spfx-fast-playwright validate tests/e2e/sharepoint-smoke.recipe.json
npx spfx-fast-playwright generate tests/e2e/sharepoint-smoke.recipe.json --out tests/e2e/sharepoint-smoke.spec.mjs --force
```

`init` already creates a spec at that output path; `--force` regenerates it from the edited recipe.

The output is a normal Playwright Test spec. Run it with the project's `SPFX_BASE_URL` and `storageState`:

```sh
SPFX_BASE_URL="https://tenant.sharepoint.com/sites/your-site" npx playwright test
```

Keep state files out of source control. They contain impersonation-capable session data. Use a dedicated least-privilege account and test site.

## Run a recipe with a persistent authenticated browser profile

This path is useful when a team already has a signed-in Playwright profile and does not want to export its storage state:

```sh
npx spfx-fast-playwright run tests/e2e/sharepoint-smoke.recipe.json \
  --base-url "$SPFX_BASE_URL" \
  --profile-dir "$SPFX_TEST_PROFILE" \
  --artifacts test-results/sharepoint-smoke
```

Runs headless, enforces the origin, and saves requested screenshots, a trace, and `summary.json`. Use `--headed` to debug. The command locks the profile and releases it after browser exit. It does not sign in, inspect cookies, or copy/remove profile files.

Use `--fast` during fix loops: assertions remain, successful screenshots/traces are skipped, and failures capture a screenshot. Run once without it after fixes. Use one job per profile; close any manually opened Chrome first.

## Recipe actions

Recipes are plain JSON so agents can generate them without writing browser plumbing:

`navigate` accepts a `query` map. `debugManifestsFile` is limited to local HTTPS hosts for SPFx debug builds.

| Action | Purpose |
| --- | --- |
| `navigate` | Open a path on the configured site origin |
| `assertPageReady` | Wait for the document body and SharePoint origin |
| `assertAuthenticated` | Check the current web's SharePoint REST endpoint without reading user data |
| `click` | Click a semantic locator; `confirm: true` accepts an expected browser confirmation |
| `fill` | Fill an accessible input |
| `select` | Select an option in a native select |
| `setChecked` | Set a checkbox or radio state |
| `assertVisible` | Wait for a semantic target to appear |
| `assertText` | Check exact or contained visible text |
| `assertValue` | Check an input value |
| `screenshot` | Save a named PNG to the local artifact directory |

Prefer role/label locators. Recipes do not allow raw JavaScript or arbitrary URLs.

For a known browser confirmation, a click can set `"confirm": true`. The runner accepts only a JavaScript `confirm` dialog and fails if no confirmation appears or a different dialog type appears. Use this only for an explicitly approved test action; it does not bypass the confirmation shown to the user.

### Validate one Graph list CRUD operation

The [tenant CRUD recipe](examples/tenant-crud-validation.recipe.json) clicks the dashboard's **Test list CRUD** button, accepts its explicit warning, and requires the create/read/update/read/delete success message. Copy it from the installed package into the project, then replace the site and page path and configure a test list in the web part:

```sh
cp node_modules/@spfx-fast/playwright/examples/tenant-crud-validation.recipe.json tests/e2e/tenant-crud-validation.recipe.json
npx spfx-fast-playwright validate tests/e2e/tenant-crud-validation.recipe.json
npx spfx-fast-playwright run tests/e2e/tenant-crud-validation.recipe.json \
  --base-url "$SPFX_BASE_URL" \
  --profile-dir "$SPFX_TEST_PROFILE" \
  --artifacts test-results/tenant-crud
```

Use a disposable test list and an account approved for Graph `Sites.ReadWrite.All`. The test creates one temporary item and deletes it; inspect the result and remove any item manually if cleanup reports failure. Do not point this recipe at production data.

For a known browser confirmation, a click can set `"confirm": true`. The runner accepts only a JavaScript `confirm` dialog and fails if no confirmation appears or a different dialog type appears. Use this only for an explicitly approved test action; it does not bypass the confirmation shown to the user.

## What this does not do

- No sign-in automation, token extraction, or permission bypass.
- No content writes unless explicitly added and approved.
- Screenshots do not prove API success or persistence; assert those separately.
- Codegen records actions; add meaningful assertions yourself.

Artifacts may contain tenant data. Keep results, auth files, and profiles local; review before sharing.

For a full workflow with profile, SharePoint page, image, and OCR guidance, read [the agent browser workflow](../../docs/AGENT-BROWSER-WORKFLOW.md).

## Package status

This is the early `0.1.0` companion package. It is framework-agnostic across SPFx projects and ships independently from `@spfx-fast/core`.
