import * as React from "react";
import { Button, Card, MessageBar, MessageBarBody, Spinner, Text } from "@fluentui/react-components";
import type { createLegacySharePointApi } from "@spfx-fast/core/sharepoint";

type SharePointApi = ReturnType<typeof createLegacySharePointApi>;
type Attachment = { FileName: string; ServerRelativeUrl?: string; Length?: string; ETag?: string };

function attachmentItems(value: unknown): Attachment[] {
  if (!value || typeof value !== "object") return [];
  const record = value as { value?: unknown; d?: { results?: unknown } };
  const items = Array.isArray(value) ? value : Array.isArray(record.value) ? record.value : record.d?.results;
  if (!Array.isArray(items)) return [];
  return items.filter((item): item is Attachment => Boolean(item && typeof item === "object" && typeof (item as Attachment).FileName === "string"));
}

/** Authenticated list-item attachment workflow for use inside an SPFx web part. */
export function ItemAttachmentsPanel({ api, listId, itemId }: { api: SharePointApi; listId: string; itemId: number }): React.ReactElement {
  const [attachments, setAttachments] = React.useState<Attachment[]>([]);
  const [file, setFile] = React.useState<File>();
  const [loading, setLoading] = React.useState(true);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string>();
  const [status, setStatus] = React.useState("");
  const [deleteConfirmation, setDeleteConfirmation] = React.useState<string>();
  const [inputKey, setInputKey] = React.useState(0);

  const refresh = React.useCallback(async () => {
    setLoading(true);
    setError(undefined);
    try {
      setAttachments(attachmentItems(await api.listItemAttachments(listId, itemId)));
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "Could not load attachments.");
    } finally {
      setLoading(false);
    }
  }, [api, itemId, listId]);

  React.useEffect(() => { void refresh(); }, [refresh]);

  async function mutate(operation: () => Promise<unknown>, message: string): Promise<void> {
    setBusy(true);
    setError(undefined);
    setStatus("");
    try {
      await operation();
      setFile(undefined);
      setInputKey((key) => key + 1);
      setDeleteConfirmation(undefined);
      await refresh();
      setStatus(message);
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "The attachment action failed.");
    } finally {
      setBusy(false);
    }
  }

  async function download(attachment: Attachment): Promise<void> {
    setBusy(true);
    setError(undefined);
    try {
      const blob = await api.readItemAttachment(listId, itemId, attachment.FileName);
      const url = URL.createObjectURL(blob);
      const anchor = document.createElement("a");
      anchor.href = url;
      anchor.download = attachment.FileName;
      anchor.hidden = true;
      document.body.append(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
      setStatus(`${attachment.FileName} downloaded.`);
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : "Could not download this attachment.");
    } finally {
      setBusy(false);
    }
  }

  const existingFile = file && attachments.some((attachment) => attachment.FileName === file.name);

  return <Card>
    <header style={{ display: "flex", flexWrap: "wrap", justifyContent: "space-between", alignItems: "center", gap: 8 }}>
      <div><h2>Item attachments</h2><Text>List item {itemId} · {attachments.length} {attachments.length === 1 ? "attachment" : "attachments"}</Text></div>
      <Button appearance="secondary" disabled={loading || busy} style={{ color: "var(--colorNeutralForeground1, #242424)" }} onClick={() => void refresh()}>Refresh attachments</Button>
    </header>
    <div style={{ display: "grid", gap: 4, margin: "12px 0" }}>
      <label htmlFor={`item-attachment-${itemId}`}>Choose a file to attach</label>
      <input id={`item-attachment-${itemId}`} key={inputKey} type="file" disabled={busy} onChange={(event) => setFile(event.currentTarget.files?.[0])} />
    </div>
    <Button appearance="primary" disabled={!file || busy} onClick={() => {
      if (!file) return;
      const selected = file;
      const operation = existingFile
        ? () => api.updateItemAttachment(listId, itemId, selected.name, selected)
        : () => api.addItemAttachment(listId, itemId, selected.name, selected);
      void mutate(operation, `${selected.name} ${existingFile ? "replaced" : "uploaded"}.`);
    }}>{existingFile ? "Replace attachment" : "Add attachment"}</Button>
    {loading ? <Spinner size="tiny" label="Loading attachments" /> : null}
    {error ? <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar> : null}
    {status ? <div role="status">{status}</div> : null}
    {!loading && !error && attachments.length === 0 ? <Text>No attachments on this item.</Text> : null}
    <ul aria-label="Item attachments">
      {attachments.map((attachment) => <li key={attachment.FileName} style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 8, padding: "6px 0" }}>
        <span>{attachment.FileName}{attachment.Length ? ` · ${formatBytes(Number(attachment.Length))}` : ""}</span>
        <Button appearance="subtle" disabled={busy} onClick={() => void download(attachment)}>Download {attachment.FileName}</Button>
        <Button appearance="subtle" disabled={busy} onClick={() => setDeleteConfirmation((name) => name === attachment.FileName ? undefined : attachment.FileName)}>
          {deleteConfirmation === attachment.FileName ? `Confirm delete ${attachment.FileName}` : `Delete ${attachment.FileName}`}
        </Button>
        {deleteConfirmation === attachment.FileName ? <Button appearance="secondary" disabled={busy} onClick={() => void mutate(() => api.deleteItemAttachment(listId, itemId, attachment.FileName, attachment.ETag), `${attachment.FileName} deleted.`)}>Delete now</Button> : null}
      </li>)}
    </ul>
  </Card>;
}

function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return "size unavailable";
  if (bytes < 1024) return `${bytes} B`;
  return `${(bytes / 1024).toFixed(1)} KB`;
}
