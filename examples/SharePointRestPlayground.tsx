import * as React from "react";
import { Button, Card, Field, FluentProvider, Input, MessageBar, MessageBarBody, Text, webLightTheme, type Theme } from "@fluentui/react-components";
import { createLegacySharePointApi } from "@spfx-fast/core/sharepoint";
import type { SharePointRestTransport } from "@spfx-fast/core/spfx";
import { ItemAttachmentsPanel } from "./ItemAttachmentsPanel.js";

/** Exercise classic SharePoint REST helpers inside a signed-in SPFx web part. */
export function SharePointRestPlayground({ rest, listId, itemId, theme = webLightTheme }: {
  rest: SharePointRestTransport;
  listId: string;
  itemId: number;
  theme?: Theme;
}): React.ReactElement {
  const api = React.useMemo(() => createLegacySharePointApi(rest), [rest]);
  const [query, setQuery] = React.useState("Project Plan");
  const [result, setResult] = React.useState<unknown>();
  const [error, setError] = React.useState<string>();
  const [busy, setBusy] = React.useState(false);

  async function run(operation: () => Promise<unknown>) {
    setBusy(true);
    setError(undefined);
    try { setResult(await operation()); }
    catch (reason) { setError(reason instanceof Error ? reason.message : String(reason)); }
    finally { setBusy(false); }
  }

  return (
    <FluentProvider theme={theme}>
      <main aria-busy={busy}>
        <h1>SharePoint REST playground</h1>
        <Text>Runs through the SPFx host's authenticated SPHttpClient. Use a disposable test list for the audience-targeting operation.</Text>
        <Card>
          <h2>Read operations</h2>
          <Field label="Search terms">
            <Input value={query} onChange={(_, data) => setQuery(data.value)} />
          </Field>
          <Button disabled={busy} onClick={() => void run(() => api.search(query, { rowLimit: 10, selectProperties: ["Title", "Path"] }))}>Search SharePoint</Button>
          <Button disabled={busy} onClick={() => void run(() => api.getItemsByCaml(listId, "<View><Query><OrderBy><FieldRef Name='ID' Ascending='FALSE' /></OrderBy></Query><RowLimit>20</RowLimit></View>"))}>Run CAML query</Button>
        </Card>
        <ItemAttachmentsPanel api={api} listId={listId} itemId={itemId} />
        <Card>
          <h2>List setting</h2>
          <Text>Audience targeting changes a list property. Use only a disposable list where you have list-management rights.</Text>
          <Button disabled={busy} onClick={() => void run(() => api.setAudienceTargeting(listId, true))}>Enable audience targeting</Button>
        </Card>
        {error ? <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar> : null}
        {result !== undefined ? <pre aria-live="polite" style={{ overflow: "auto", whiteSpace: "pre-wrap" }}>{JSON.stringify(result, null, 2)}</pre> : null}
      </main>
    </FluentProvider>
  );
}
