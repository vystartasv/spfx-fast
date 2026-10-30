# Deploy an SPFx package

Deploy to a tenant or site app catalog with the authenticated ALM helper.

This guide ships in npm. `node scripts/...` commands require the source checkout.

## Fast path

```sh
pnpm test:spfx-fixture
node scripts/sharepoint-app-lifecycle.mjs deploy \
  --catalog "$SPFX_APP_CATALOG_URL" \
  --package examples/fixtures/spfx-fast-dashboard/sharepoint/solution/spfx-fast-dashboard.sppkg \
  --config examples/fixtures/spfx-fast-dashboard/config/package-solution.json \
  --profile-dir "$SPFX_TEST_PROFILE"
node scripts/sharepoint-app-lifecycle.mjs status \
  --web "$SPFX_TEST_SITE" \
  --config examples/fixtures/spfx-fast-dashboard/config/package-solution.json \
  --profile-dir "$SPFX_TEST_PROFILE"
node scripts/sharepoint-app-lifecycle.mjs install \
  --site "$SPFX_TEST_SITE" \
  --config examples/fixtures/spfx-fast-dashboard/config/package-solution.json \
  --profile-dir "$SPFX_TEST_PROFILE"
```

The profile must be signed in. The script uses the in-page SharePoint session, locks the profile, and never copies auth data. Run one job per profile.

Use the supplied site path for REST, e.g. `/sites/team/_api`; `/_api` targets the tenant root.

## Keep the deployment scoped

- Find the tenant app catalog through SharePoint admin center or `/_api/SP_TenantSettings_Current`; do not guess `/sites/appcatalog`.
- Check `sitecollectionappcatalog` too. If IDs overlap across catalogs, verify the page's asset URL.
- `deploy` uploads and publishes to the selected catalog. `install` targets one site. It stops if SharePoint says the package is invalid, disabled, undeployed, or already has an unverified instance.
- For web part tests, choose **Only enable this app** and install on the test site. Copilot UX instead requires **Enable this app and add it to all sites**, then **Add to Teams** to sync its agent. This does not install a Teams tab.
- Approve only the exact Graph scopes requested by this package. Do not approve other pending packages.
- Catalog row `ID`, package `ProductId`, and component ID differ. Use the catalog row for ALM and component ID for the page.
- A `200` is not proof. Read target `AvailableApps` and verify `InstalledVersion === AppCatalogVersion`. `GetAppInstanceById` falsely reported a known Site Contents app as uninstalled.
- After upgrades, verify the published page loads the current `ClientSideAssets` bundle; matching versions can still serve stale assets. Increment `solution.version` before rebuilding, or SharePoint may keep serving the old bundle.
- If install reports an existing instance but read-back cannot find it, stop and inspect Site Contents before retrying.

## Add a test page and verify

1. Create a new page under **Site Pages**. SharePoint may name it `Page(1).aspx`; give it a clear title, save, and publish it.
2. Search the toolbox for the exact web-part title from its manifest. “Dashboard” can select SharePoint’s native Dashboard web part. Confirm the component ID against `GetClientSideWebParts` before adding it.
3. Configure it with IDs read from the site, not guessed values. Use a disposable list and synthetic records.
4. Open the published page, confirm the expected web part and signed-in UI, then run one CRUD recipe. Verify create/read/update/read/delete and cleanup.
5. Keep screenshots, traces, and tenant data local until reviewed.

The [dashboard CRUD recipe](../packages/playwright/examples/tenant-crud-validation.recipe.json) uses a test page/list. `scripts/sharepoint-validation-page.mjs` creates or verifies a page from inspected list IDs. See the [browser workflow](AGENT-BROWSER-WORKFLOW.md).

Batch-check legacy APIs with `node scripts/sharepoint-legacy-api-batch.mjs --site "$SPFX_TEST_SITE" --profile-dir "$SPFX_TEST_PROFILE"`. It tests navigation, CAML, stream rendering, change tokens, `GetChanges`, views, and audience targeting on a disposable list, then deletes it.

## ALM API shape

```text
POST /_api/web/{tenantappcatalog|sitecollectionappcatalog}/Add(overwrite=true,url='name.sppkg')
POST /_api/web/{scope}appcatalog/AvailableApps/GetById('{catalog-app-id}')/Deploy
POST /_api/web/{scope}appcatalog/AvailableApps/GetById('{catalog-app-id}')/Install
GET  /_api/web/{scope}appcatalog/AvailableApps
```

POST calls need `X-RequestDigest` and the catalog URL. See Microsoft's [ALM API reference](https://learn.microsoft.com/en-us/sharepoint/dev/apis/alm-api-for-spfx-add-ins).

`GetAppInstanceById` reported a known Site Contents app as “not installed”; use the `AvailableApps` version check instead.
