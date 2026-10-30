import * as React from "react";
import { Badge, Button, Card, MessageBar, MessageBarBody, Spinner, Text } from "@fluentui/react-components";
import { type GraphDriveItem, type createClient } from "@spfx-fast/core";
import { useDriveItems } from "@spfx-fast/core/react";

type Client = ReturnType<typeof createClient>;
const CATEGORIES = [
  { name: "Documents", color: "var(--colorPaletteBlueBackground2, #cfe4fa)", extensions: ["doc", "docx", "md", "pdf", "txt"] },
  { name: "Spreadsheets", color: "var(--colorPaletteGreenBackground2, #dff6dd)", extensions: ["csv", "xls", "xlsx"] },
  { name: "Presentations", color: "var(--colorPalettePurpleBackground2, #e8daff)", extensions: ["ppt", "pptx"] },
  { name: "Images", color: "var(--colorPaletteMarigoldBackground2, #fff4ce)", extensions: ["gif", "jpeg", "jpg", "png", "svg", "webp"] },
  { name: "Other files", color: "var(--colorNeutralBackground3, #f0f0f0)", extensions: [] },
];

/** Cache-first drive inventory and resumable delta status for a signed-in SPFx host. */
export function DriveDeltaPanel({ client, driveId, maxPages = 100 }: { client: Client; driveId: string; maxPages?: number }): React.ReactElement {
  const drive = useDriveItems(client, driveId, { maxPages });
  const [actionError, setActionError] = React.useState<string>();
  const files = React.useMemo(() => drive.items.filter((item) => item.file && !item.deleted), [drive.items]);
  const folders = React.useMemo(() => drive.items.filter((item) => item.folder && !item.deleted), [drive.items]);
  const categories = React.useMemo(() => {
    const counts = CATEGORIES.map(() => 0);
    for (const file of files) {
      const extension = file.name?.split(".").pop()?.toLocaleLowerCase() ?? "";
      const index = CATEGORIES.findIndex((category) => category.extensions.indexOf(extension) >= 0);
      counts[index < 0 ? CATEGORIES.length - 1 : index]++;
    }
    return counts;
  }, [files]);
  const modified = React.useMemo(() => [...files].sort((left, right) =>
    (Date.parse(right.lastModifiedDateTime ?? "") || 0) - (Date.parse(left.lastModifiedDateTime ?? "") || 0)), [files]);
  const total = categories.reduce((sum, count) => sum + count, 0);
  let end = 0;
  const stops = categories.map((count, index) => {
    const start = end;
    end += total ? count / total * 100 : 0;
    return count ? `${CATEGORIES[index].color} ${start}% ${end}%` : undefined;
  }).filter(Boolean).join(", ");

  async function download(item: GraphDriveItem): Promise<void> {
    try {
      const content = await client.downloadDriveItem(driveId, item.id);
      const objectUrl = URL.createObjectURL(content);
      const anchor = document.createElement("a");
      anchor.href = objectUrl;
      anchor.download = item.name || "download";
      anchor.hidden = true;
      document.body.append(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(objectUrl), 60_000);
    } catch {
      // `useDriveItems` owns sync errors; download errors are surfaced below via a separate state.
      setActionError("Could not download this file. Check access and try again.");
    }
  }

  return <Card>
    <header style={{ display: "flex", flexWrap: "wrap", alignItems: "center", justifyContent: "space-between", gap: 12 }}>
      <div>
        <h2>Drive sync</h2>
        <Text>{drive.source === "cache" ? "Showing cached inventory" : drive.source === "network" ? "Inventory refreshed from SharePoint" : "Signed-in drive inventory"}</Text>
      </div>
      <Button appearance="primary" disabled={drive.loading || !driveId} onClick={async () => { await drive.refresh(); }}>
        {drive.loading ? "Syncing…" : drive.complete === false ? "Continue sync" : "Sync now"}
      </Button>
    </header>
    <div role="status" aria-live="polite" style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 8, margin: "12px 0" }}>
      {drive.loading ? <Spinner size="tiny" label="Synchronizing drive changes" /> : null}
      <Badge appearance="tint" color={drive.error ? "danger" : drive.complete === false ? "warning" : "success"}>
        {drive.error ? "Sync needs attention" : drive.loading ? "Syncing changes" : drive.complete === false ? "More changes to sync" : "Drive is up to date"}
      </Badge>
      <Text>{drive.items.length} items · {files.length} {files.length === 1 ? "file" : "files"} · {folders.length} {folders.length === 1 ? "folder" : "folders"} · {drive.source ?? "loading"}</Text>
    </div>
    {drive.error ? <MessageBar intent="error"><MessageBarBody>Drive sync failed. Check your access and retry. {drive.error instanceof Error ? drive.error.message : ""}</MessageBarBody></MessageBar> : null}
    {actionError ? <MessageBar intent="error"><MessageBarBody>{actionError}</MessageBarBody></MessageBar> : null}
    {total > 0 ? <section aria-label="Drive file types" style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 20, margin: "12px 0" }}>
      <div aria-hidden="true" style={{ display: "grid", placeItems: "center", width: 112, height: 112, borderRadius: "50%", background: `conic-gradient(${stops})` }}>
        <div style={{ display: "grid", placeContent: "center", width: 68, height: 68, borderRadius: "50%", textAlign: "center", background: "var(--colorNeutralBackground1, #fff)" }}><strong>{total}</strong><span style={{ fontSize: 12 }}>files</span></div>
      </div>
      <ul style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: "4px 20px", flex: "1 1 320px", listStyle: "none", padding: 0, margin: 0 }}>
        {CATEGORIES.map((category, index) => <li key={category.name} style={{ display: "flex", alignItems: "center", gap: 8 }}><span aria-hidden="true" style={{ width: 10, height: 10, borderRadius: 2, background: category.color }} /><span>{category.name}</span><strong>{categories[index]}</strong></li>)}
      </ul>
    </section> : null}
    {drive.loading && drive.items.length === 0 ? <Spinner label="Loading drive inventory" /> : null}
    {!drive.loading && drive.items.length === 0 && !drive.error ? <Text>No drive items found.</Text> : null}
    <h3>Latest modified files</h3>
    {modified.length ? <ul aria-label="Latest modified drive files" style={{ paddingLeft: 20 }}>
      {modified.slice(0, 6).map((item) => <li key={item.id} style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 8, padding: "6px 0" }}>
        <span style={{ flex: "1 1 200px" }}>{item.name ?? "Unnamed file"}{item.lastModifiedDateTime ? <Text size={200}> · {new Date(item.lastModifiedDateTime).toLocaleDateString()}</Text> : null}</span>
        {item.webUrl ? <Button as="a" href={item.webUrl} target="_blank" rel="noreferrer" appearance="subtle">Open</Button> : null}
        <Button appearance="subtle" onClick={async () => { await download(item); }}>Download</Button>
      </li>)}
    </ul> : null}
    <Text size={200}>Delta sync stores the current drive state. “Latest modified” sorts that inventory by file modified date; it is not an edit history.</Text>
  </Card>;
}
