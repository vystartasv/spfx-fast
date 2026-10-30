# Sample web parts

Four small SPFx components for a fictional North Campus workplace service:

- `RequestDashboard.tsx`: cache-first request list and create form.
- `TeamTaskQueue.tsx`: filtered, paged task list.
- `KnowledgePanel.tsx`: cached Markdown rendered safely with GFM tables.
- `ApprovalInbox.tsx`: pending approvals with a SharePoint update and Teams dark theme.

Pass the signed-in `client`, site ID, and list IDs from the web part. `SampleWebpartsGallery.tsx` uses synthetic data and an in-memory client for local tests. It makes no network calls. Samples use Fluent UI v9 web light; approvals use Teams dark. Priority, status, and map areas use distinct semantic colors.

The sample client is for demos only. Replace it with `createSPFxServices` in a real SPFx web part. Build with `pnpm run build` or `npm run build`; capture screenshots with `SPFX_FAST_BROWSER_SCREENSHOT_DIR=test-results/screenshots pnpm exec node scripts/browser-a11y.mjs` (npm: `npm run test:browser`).

Pass `sharepoint={services.sharepoint}` to `RequestDashboard` to show attachments after create when SharePoint returns a numeric item ID; pass `attachmentItemId` for an existing item. Power Automate examples are generic patterns for customers to adapt.

See the [sample screenshots](../../docs/SCREENSHOTS.md).
