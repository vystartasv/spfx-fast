import * as React from "react";
import { Button, Card, CardHeader, MessageBar, MessageBarBody, ProgressBar, Spinner, Text } from "@fluentui/react-components";
import { isDriveUploadInterruptedError, type DriveUploadInterruptedError, type GraphDriveItem, type createClient } from "@spfx-fast/core";
import { useMarkdownFile } from "@spfx-fast/core/react";
import { MarkdownContent } from "@spfx-fast/core/markdown";

type Client = ReturnType<typeof createClient>;
const MAX_FILES = 500;
const FILE_CATEGORIES = [
  { name: "Documents", color: "var(--colorPaletteBlueBackground2, #cfe4fa)", extensions: ["doc", "docx", "md", "pdf", "txt"] },
  { name: "Spreadsheets", color: "var(--colorPaletteGreenBackground2, #dff6dd)", extensions: ["csv", "xls", "xlsx"] },
  { name: "Presentations", color: "var(--colorPalettePurpleBackground2, #e8daff)", extensions: ["ppt", "pptx"] },
  { name: "Images", color: "var(--colorPaletteMarigoldBackground2, #fff4ce)", extensions: ["gif", "jpeg", "jpg", "png", "svg", "webp"] },
  { name: "Other files", color: "var(--colorNeutralBackground3, #f0f0f0)", extensions: [] },
];

/** Lists one configured document-library folder with safe Markdown previews and file actions. */
export function DocumentLibraryPanel({ client, driveId, folderItemId }: { client: Client; driveId: string; folderItemId: string }): React.ReactElement {
  const [files, setFiles] = React.useState<GraphDriveItem[]>([]);
  const [cursor, setCursor] = React.useState<string>();
  const [source, setSource] = React.useState<string>();
  const [loading, setLoading] = React.useState(true);
  const [loadingMore, setLoadingMore] = React.useState(false);
  const [error, setError] = React.useState<string>();
  const [uploadProgress, setUploadProgress] = React.useState<number>();
  const [uploading, setUploading] = React.useState(false);
  const [uploadError, setUploadError] = React.useState<string>();
  const [uploadRecovery, setUploadRecovery] = React.useState<DriveUploadInterruptedError>();
  const [uploadInputKey, setUploadInputKey] = React.useState(0);
  const [markdownFile, setMarkdownFile] = React.useState<GraphDriveItem>();
  const requestId = React.useRef(0);
  const markdown = useMarkdownFile(client, driveId, markdownFile?.id ?? "", { enabled: Boolean(markdownFile) });
  const categoryCounts = React.useMemo(() => {
    const counts = FILE_CATEGORIES.map(() => 0);
    for (const file of files) {
      if (file.folder) continue;
      const extension = file.name?.split(".").pop()?.toLocaleLowerCase() ?? "";
      const index = FILE_CATEGORIES.findIndex((category) => category.extensions.indexOf(extension) >= 0);
      counts[index < 0 ? FILE_CATEGORIES.length - 1 : index]++;
    }
    return counts;
  }, [files]);
  const fileCount = categoryCounts.reduce((sum, count) => sum + count, 0);
  let pieEnd = 0;
  const pieStops = categoryCounts.map((count, index) => {
    const start = pieEnd;
    pieEnd += fileCount ? count / fileCount * 100 : 0;
    return count ? `${FILE_CATEGORIES[index].color} ${start}% ${pieEnd}%` : undefined;
  }).filter(Boolean).join(", ");

  const refresh = React.useCallback(async () => {
    const id = ++requestId.current;
    setLoading(true);
    setLoadingMore(false);
    setError(undefined);
    setFiles([]);
    setCursor(undefined);
    try {
      const page = await client.getDriveItemPage(driveId, folderItemId, { pageSize: 50 });
      if (id !== requestId.current) return;
      setFiles(page.items.slice(0, MAX_FILES));
      setCursor(page.items.length >= MAX_FILES ? undefined : page.nextCursor);
      setSource(page.source);
      page.refresh.then((fresh) => {
        if (id === requestId.current) {
          setFiles(fresh.value.slice(0, MAX_FILES));
          setCursor(fresh.value.length >= MAX_FILES ? undefined : fresh["@odata.nextLink"]);
          setSource("network");
        }
      }).catch((reason: unknown) => {
        if (id === requestId.current) setError(reason instanceof Error ? reason.message : String(reason));
      });
    } catch (reason) {
      if (id === requestId.current) setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      if (id === requestId.current) setLoading(false);
    }
  }, [client, driveId, folderItemId]);

  React.useEffect(() => {
    refresh().then(() => undefined, () => undefined);
    return () => { requestId.current++; };
  }, [refresh]);

  async function loadMore(): Promise<void> {
    if (!cursor || loadingMore || files.length >= MAX_FILES) return;
    const id = requestId.current;
    setLoadingMore(true);
    try {
      const page = await client.getDriveItemPage(driveId, folderItemId, { cursor, pageSize: 50 });
      if (id !== requestId.current) return;
      setFiles((current) => [...current, ...page.items].slice(0, MAX_FILES));
      setCursor(page.items.length + files.length >= MAX_FILES ? undefined : page.nextCursor);
      setSource(page.source);
    } catch (reason) {
      if (id === requestId.current) setError(reason instanceof Error ? reason.message : String(reason));
    } finally {
      if (id === requestId.current) setLoadingMore(false);
    }
  }

  async function uploadFile(file?: File): Promise<void> {
    if (!file || uploading) return;
    setUploading(true);
    setUploadError(undefined);
    setUploadRecovery(undefined);
    setUploadProgress(0);
    try {
      await client.uploadDriveItem(driveId, folderItemId, file.name, file, {
        onProgress: (sent, total) => setUploadProgress(total ? sent / total : 1),
      });
      setUploadInputKey((current) => current + 1);
      await refresh();
    } catch (reason) {
      setUploadError(reason instanceof Error ? reason.message : String(reason));
      setUploadRecovery(isDriveUploadInterruptedError(reason) ? reason : undefined);
    } finally {
      setUploading(false);
      setUploadProgress(undefined);
    }
  }

  async function resumeUpload(): Promise<void> {
    if (!uploadRecovery || uploading) return;
    setUploading(true);
    setUploadError(undefined);
    setUploadProgress(0);
    try {
      await uploadRecovery.resume();
      setUploadRecovery(undefined);
      setUploadInputKey((current) => current + 1);
      await refresh();
    } catch (reason) {
      setUploadError(reason instanceof Error ? reason.message : String(reason));
      setUploadRecovery(isDriveUploadInterruptedError(reason) ? reason : undefined);
    } finally {
      setUploading(false);
      setUploadProgress(undefined);
    }
  }

  async function downloadFile(file: GraphDriveItem): Promise<void> {
    setError(undefined);
    try {
      const content = await client.downloadDriveItem(driveId, file.id);
      const objectUrl = URL.createObjectURL(content);
      const anchor = document.createElement("a");
      anchor.href = objectUrl;
      anchor.download = file.name || "download";
      anchor.hidden = true;
      document.body.append(anchor);
      anchor.click();
      anchor.remove();
      window.setTimeout(() => URL.revokeObjectURL(objectUrl), 60_000);
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason));
    }
  }

  return (
    <Card>
      <CardHeader header={<Text>Document library</Text>} description={<Text>{source ? `Folder contents · ${source}` : "Loading folder contents"}</Text>} />
      <Button appearance="secondary" disabled={loading} style={{ color: "var(--colorNeutralForeground1, #242424)" }} onClick={refresh}>Refresh files</Button>
      <label>
        Upload a file
        <input key={uploadInputKey} type="file" disabled={uploading} onChange={async (event) => { await uploadFile(event.currentTarget.files?.[0]); }} />
      </label>
      {uploadRecovery ? <Button appearance="primary" disabled={uploading} onClick={resumeUpload}>Resume upload</Button> : null}
      {uploading && uploadProgress !== undefined ? <ProgressBar value={uploadProgress} aria-label="File upload progress" /> : null}
      {uploadError ? <MessageBar intent="error"><MessageBarBody>{uploadError}</MessageBarBody></MessageBar> : null}
      {loading && files.length === 0 ? <Spinner label="Loading files" /> : null}
      {error ? <MessageBar intent="error"><MessageBarBody>{error}</MessageBarBody></MessageBar> : null}
      {fileCount > 0 ? <section aria-label="Library file types" style={{ display: "flex", flexWrap: "wrap", alignItems: "center", gap: 20, margin: "12px 0" }}>
        <div aria-hidden="true" style={{ display: "grid", placeItems: "center", width: 116, height: 116, borderRadius: "50%", background: `conic-gradient(${pieStops})` }}>
          <div style={{ display: "grid", placeContent: "center", width: 70, height: 70, borderRadius: "50%", textAlign: "center", background: "var(--colorNeutralBackground1, #fff)" }}>
            <strong>{fileCount}</strong><span style={{ fontSize: 12 }}>files</span>
          </div>
        </div>
        <ul style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: "4px 20px", flex: "1 1 320px", listStyle: "none", padding: 0, margin: 0 }}>
          {FILE_CATEGORIES.map((category, index) => <li key={category.name} style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <span aria-hidden="true" style={{ width: 10, height: 10, borderRadius: 2, background: category.color }} />
            <span>{category.name}</span><strong>{categoryCounts[index]}</strong>
          </li>)}
        </ul>
      </section> : null}
      {!loading && files.length === 0 && !error ? <Text>No files were returned for this folder.</Text> : null}
      <ul aria-label="Files in document library folder">
        {files.map((file) => <li key={file.id}>
          {file.name?.toLocaleLowerCase().endsWith(".md")
            ? <Button appearance="subtle" aria-pressed={markdownFile?.id === file.id} style={{ fontWeight: markdownFile?.id === file.id ? 600 : undefined }} onClick={() => setMarkdownFile(file)}>{file.name}</Button>
            : file.webUrl ? <a href={file.webUrl} target="_blank" rel="noreferrer">{file.name ?? "Open file"}</a> : <Text>{file.name}</Text>}
          {!file.folder ? <Button appearance="subtle" disabled={uploading} onClick={async () => { await downloadFile(file); }}>Download {file.name ?? "file"}</Button> : null}
          {file.folder ? <Text> · Folder ({file.folder.childCount ?? 0} items)</Text> : null}
        </li>)}
      </ul>
      {cursor && files.length < MAX_FILES ? <Button disabled={loadingMore} onClick={loadMore}>{loadingMore ? "Loading…" : "Load more files"}</Button> : null}
      {markdownFile ? <section aria-label="Selected Markdown file">
        <header style={{ display: "flex", flexWrap: "wrap", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
          <h3>{markdownFile.name}</h3>
          <span style={{ display: "flex", gap: 8 }}>
            {markdownFile.webUrl ? <Button as="a" href={markdownFile.webUrl} target="_blank" rel="noreferrer" appearance="subtle">Open in SharePoint</Button> : null}
            <Button appearance="subtle" onClick={() => setMarkdownFile(undefined)}>Close preview</Button>
          </span>
        </header>
        {markdown.loading && !markdown.markdown ? <Spinner label="Loading Markdown file" /> : null}
        {markdown.error ? <MessageBar intent="warning"><MessageBarBody>The file could not be refreshed. Cached Markdown may still be displayed.</MessageBarBody></MessageBar> : null}
        {markdown.markdown !== undefined ? <MarkdownContent>{markdown.markdown}</MarkdownContent> : null}
        <Text size={200}>Safe Markdown · {markdown.source ?? "loading"}</Text>
      </section> : null}
    </Card>
  );
}
