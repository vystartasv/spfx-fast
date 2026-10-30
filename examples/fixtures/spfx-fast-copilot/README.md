# SPFx Fast Copilot UX list browser

SPFx 1.24.0 Copilot UX fixture using `@spfx-fast/core`, React 18, and Fluent UI v9. It reads up to 50 items from an exact list and can send the first ten titles to Copilot after user action.

## Requirements

- Node.js 22.14+, 24.x, or 26.x; SPFx 1.24.0 and SharePoint App Catalog.
- Microsoft 365 Copilot add-on for each user, including Workbench testers.
- Admin approval for Graph `Sites.Read.All`.

## Build and test in Copilot Workbench

From the repository root:

```sh
pnpm test:copilot-fixture
```

Builds the library and fixture, runs Heft checks, validates the agent schema, and packages the `.sppkg`.

Use `npm run test:copilot-fixture` if npm runs your repository scripts.

For the local Workbench loop, from this fixture folder:

```sh
pnpm install --frozen-lockfile
SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com pnpm start --nobrowser
```

Using npm: run `npm ci` and `SPFX_SERVE_TENANT_DOMAIN=contoso.sharepoint.com npm start -- --nobrowser` in this fixture. Replace `contoso.sharepoint.com` with your tenant host.

Open `https://contoso.sharepoint.com/_layouts/15/copilotworkbench.aspx` with a licensed account and allow the debug manifest. In the component's Raw properties tab, enter an exact site URL and list title, then select **Fire turn**:

```json
{"siteUrl":"https://contoso.sharepoint.com/sites/team","listTitle":"Exact list title"}
```

Use the exact site URL and list title supplied by the developer; omit `siteUrl` to use the Workbench's current site. Invalid URLs and page/list paths are rejected. Confirm the component loads the list and displays its bounded results. Use **Share first 10 titles with Copilot** only with data approved for that conversation. See Microsoft's [Workbench tutorial](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/copilot/get-started/build-your-first-copilot-app) and [SPFx 1.24 release notes](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/release-1.24.0) for current licensing and client support.

## Deploy

Deploy `sharepoint/solution/spfx-fast-copilot.sppkg`, approve Graph permission, and sync the declarative agent. Redeploy tool/schema changes; refresh may take minutes.

## Security and limits

- `MSGraphClientV3` supplies auth; tokens are not stored. Reads use the signed-in user's access and cap at 50 items. The component is read-only.
- Only the first ten titles are shared with Copilot after explicit user action; context stays below 10 KB.
- Cache scope does not change authorization.
- At GA, Copilot UX supports browser/desktop Microsoft 365 Copilot and Copilot in Teams, not mobile. Workbench also needs the Copilot add-on.
- Separate fixture: Microsoft documents a GA issue syncing Teams when Copilot UX and SPFx Teams features share a solution.

## Key files

- `src/copilotComponents/workItems/WorkItemsCopilotComponent.tsx`: host lifecycle, authenticated services, URL validation, and exact list resolution.
- `src/copilotComponents/workItems/WorkItemsCopilotComponentProperties.ts`: Zod schema consumed by the agent tool.
- `src/copilotComponents/workItems/components/WorkItems.tsx`: Fluent UI v9 list UX, bounded paging, and explicit Copilot context action.
- `copilot/instruction.txt`: agent routing and no-guessing instructions.
- `config/copilot-agent.json`: component-to-agent registration.
