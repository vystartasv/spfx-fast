import * as React from "react";
import { Button, Text } from "@fluentui/react-components";
import { createClient, createIndexedDbStore, createMemoryStore, getCacheStorageStatus, isDriveUploadInterruptedError, type DriveUploadInterruptedError } from "@spfx-fast/core";
import { createSPFxGraphTransport, type createSPFxServices } from "@spfx-fast/core/spfx";

type Services = Awaited<ReturnType<typeof createSPFxServices>>;
type BatchClient = Pick<Services["client"], "batch">;
type BatchRequest = Parameters<BatchClient["batch"]>[0][number];
type BatchResponse = Awaited<ReturnType<BatchClient["batch"]>>[number];
type ValidationCheck = { name: string; status: "pass" | "info" | "fail"; durationMs: number; detail: string };
type PendingUploadRecovery = { error: DriveUploadInterruptedError; driveId: string; fileName: string; sizeMiB: number; expectedHash: string; startedAt: number };

async function batchWithThrottledSubrequestRetry(client: BatchClient, requests: BatchRequest[], onThrottleRetry?: (count: number) => void): Promise<BatchResponse[]> {
  const responses = new Map<string, BatchResponse>();
  let pending = requests;
  for (let attempt = 0; pending.length && attempt <= 5; attempt++) {
    const result = await client.batch(pending);
    const throttled: BatchRequest[] = [];
    let retryAfterMs: number | undefined;
    for (const response of result) {
      responses.set(response.id, response);
      if (response.status !== 429) continue;
      const request = pending.find((candidate) => candidate.id === response.id);
      if (request) throttled.push(request);
      const raw = response.headers?.["Retry-After"] ?? response.headers?.["retry-after"];
      const seconds = Number(raw);
      const dateDelay = raw ? Date.parse(raw) - Date.now() : Number.NaN;
      const delay = Number.isFinite(seconds) ? Math.max(0, seconds * 1000) : Number.isFinite(dateDelay) ? Math.max(0, dateDelay) : Math.min(500 * 2 ** attempt, 8000);
      retryAfterMs = Math.max(retryAfterMs ?? 0, delay);
    }
    if (!throttled.length || attempt === 5) break;
    onThrottleRetry?.(throttled.length);
    await new Promise((resolve) => setTimeout(resolve, retryAfterMs ?? Math.min(500 * 2 ** attempt, 8000)));
    pending = throttled;
  }
  return requests.map((request) => responses.get(request.id) ?? { id: request.id, status: 0 });
}

async function check(name: string, run: () => Promise<string>): Promise<ValidationCheck> {
  const start = performance.now();
  try {
    const detail = await run();
    return { name, status: "pass", durationMs: performance.now() - start, detail };
  } catch (error) {
    const status = error && typeof error === "object" && ("status" in error || "statusCode" in error)
      ? Number((error as { status?: unknown; statusCode?: unknown }).status ?? (error as { statusCode?: unknown }).statusCode)
      : undefined;
    const errorName = error instanceof Error ? error.name : "RequestError";
    return {
      name,
      status: "fail",
      durationMs: performance.now() - start,
      detail: error && typeof error === "object" && "validationDetail" in error && typeof (error as { validationDetail?: unknown }).validationDetail === "string"
        ? (error as { validationDetail: string }).validationDetail
        : Number.isInteger(status) ? `Request failed (HTTP ${status}).` : `Request failed (${errorName}).`,
    };
  }
}

/** Read-only checks run only after an SPFx page supplies its signed-in services. */
export function TenantValidationPanel({ services, siteId, listId, driveId, folderItemId = "root" }: {
  services: Services;
  siteId: string;
  listId: string;
  driveId?: string;
  folderItemId?: string;
}): React.ReactElement {
  const [running, setRunning] = React.useState(false);
  const [checks, setChecks] = React.useState<ValidationCheck[]>([]);
  const [roundTripRunning, setRoundTripRunning] = React.useState(false);
  const [roundTrip, setRoundTrip] = React.useState<ValidationCheck>();
  const [uploadRecovery, setUploadRecovery] = React.useState<PendingUploadRecovery>();
  const [uploadResumeRunning, setUploadResumeRunning] = React.useState(false);
  const [crudRunning, setCrudRunning] = React.useState(false);
  const [crudResult, setCrudResult] = React.useState<ValidationCheck>();
  const [deltaRunning, setDeltaRunning] = React.useState(false);
  const [deltaResult, setDeltaResult] = React.useState<ValidationCheck>();
  const [listDeltaRunning, setListDeltaRunning] = React.useState(false);
  const [listDeltaResult, setListDeltaResult] = React.useState<ValidationCheck>();
  const [referenceRunning, setReferenceRunning] = React.useState(false);
  const [referenceResult, setReferenceResult] = React.useState<ValidationCheck>();
  const [versionRunning, setVersionRunning] = React.useState(false);
  const [versionResult, setVersionResult] = React.useState<ValidationCheck>();
  const [recoveryRunning, setRecoveryRunning] = React.useState(false);
  const [recoveryResult, setRecoveryResult] = React.useState<ValidationCheck>();

  const run = async (): Promise<void> => {
    if (running || referenceRunning || listDeltaRunning || versionRunning || recoveryRunning) return;
    setRunning(true);
    setChecks([]);
    const results = await Promise.all([
      check("SharePoint REST", async () => {
        const web = await services.rest.get("/_api/web?$select=Id,Title") as { Id?: string };
        if (!web?.Id) throw new Error("The hosting web response did not include its ID.");
        return "Authenticated hosting-web request succeeded; title and identifiers are not logged.";
      }),
      check("Current user", async () => {
        const user = await services.sharepoint.getCurrentUser();
        if (!Number.isSafeInteger(user.Id)) throw new Error("SharePoint did not return a current-user ID.");
        return "Authenticated current-user request succeeded; identity fields are not logged.";
      }),
      check("Graph list discovery", async () => {
        const lists = await services.client.getLists(siteId);
        return `Read ${lists.length} list metadata entr${lists.length === 1 ? "y" : "ies"}.`;
      }),
      check("Graph $batch (5 reads)", async () => {
        const site = `/sites/${encodeURIComponent(siteId)}`;
        const list = `${site}/lists/${encodeURIComponent(listId)}`;
        const responses = await services.client.batch([
          { id: "site", method: "GET", url: `${site}?$select=id` },
          { id: "list", method: "GET", url: `${list}?$select=id,displayName` },
          { id: "columns", method: "GET", url: `${list}/columns?$top=1` },
          { id: "contentTypes", method: "GET", url: `${list}/contentTypes?$top=1` },
          { id: "items", method: "GET", url: `${list}/items?$top=1&$select=id` },
        ]);
        const expectedIds = ["site", "list", "columns", "contentTypes", "items"];
        const receivedIds = responses.map((response) => response.id).sort();
        const failed = responses.filter((response) => !Number.isInteger(response.status) || response.status < 200 || response.status >= 300);
        if (responses.length !== expectedIds.length || expectedIds.some((id) => receivedIds.indexOf(id) < 0) || failed.length) {
          throw new Error(`Graph batch returned ${failed.length} failed or invalid subrequest(s).`);
        }
        return `All ${responses.length} site, list, column, content type, and item reads succeeded.`;
      }),
      check("Outlook calendar view", async () => {
        const now = Date.now();
        const start = new Date(now - 7 * 24 * 60 * 60 * 1000).toISOString();
        const end = new Date(now + 7 * 24 * 60 * 60 * 1000).toISOString();
        const page = await services.client.getCalendarView(start, end, { maxItems: 50, maxPages: 1 });
        return `Authenticated Graph read returned ${page.items.length} event(s) in one bounded 14-day window${page.complete ? "" : "; more results remain"}. Event details were not logged or cached.`;
      }),
      check("Graph list schema", async () => {
        const result = await services.client.inspectList(siteId, listId);
        if (!Array.isArray(result.columns)) throw new Error("Graph did not return list columns.");
        return `Read list metadata and ${result.columns.length} column(s).`;
      }),
      check("Graph content types", async () => {
        const contentTypes = await services.client.getListContentTypes(siteId, listId);
        return `Read ${contentTypes.length} list content type(s).`;
      }),
      check("SharePoint list views", async () => {
        const page = await services.sharepoint.getListViewsPage(listId, { pageSize: 1 });
        return `Read ${page.items.length} view(s) in one bounded page.`;
      }),
      check("Browser storage", async () => {
        const storage = await getCacheStorageStatus();
        if (storage.quota === undefined && storage.usage === undefined) return "Storage estimate API is unavailable in this browser.";
        const usage = storage.usage === undefined ? "unknown" : `${Math.round(storage.usage / 1024 / 1024)} MiB`;
        const quota = storage.quota === undefined ? "unknown" : `${Math.round(storage.quota / 1024 / 1024)} MiB`;
        return `Origin estimate: ${usage} used of ${quota}; persistent: ${storage.persistent ?? "unknown"}.`;
      }),
      ...(driveId ? [check("Document library", async () => {
        const page = await services.client.getDriveItemPage(driveId, folderItemId, { pageSize: 1 });
        return `Authenticated folder read succeeded (${page.items.length} item(s), ${page.source}).`;
      })] : []),
    ]);
    const firstPage = await check("Bounded list page", async () => {
      const page = await services.client.getItemPage(siteId, listId, { pageSize: 1, fields: ["Title"] });
      await page.refresh;
      return `Read ${page.items.length} item(s); initial source: ${page.source}.`;
    });
    results.push(firstPage);
    results.push(await check("Cache repeat", async () => {
      const first = await services.client.getItemPage(siteId, listId, { pageSize: 1, fields: ["Title"] });
      await first.refresh.catch(() => undefined);
      let source = first.source;
      for (let attempt = 0; source !== "cache" && attempt < 3; attempt++) {
        await new Promise((resolve) => window.setTimeout(resolve, 25));
        const repeat = await services.client.getItemPage(siteId, listId, { pageSize: 1, fields: ["Title"] });
        source = repeat.source;
        await repeat.refresh.catch(() => undefined);
      }
      return source === "cache"
        ? "Repeated list read was served from cache."
        : "Repeated list read succeeded; persistent or memory storage did not return a warm page.";
    }));
    setChecks(results);
    setRunning(false);
  };

  const runFileRoundTrip = async (sizeMiB: 10 | 20 | 50 | 100 | 128 = 10, simulateInterruption = false, conflictBehavior?: "replace"): Promise<void> => {
    if (!driveId || roundTripRunning || uploadRecovery || running || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning) return;
    const recoveryNote = simulateInterruption
      ? " It simulates a lost response after SharePoint accepts the first chunk, then checks that the helper resumes from the server-reported byte range."
      : conflictBehavior === "replace" ? " It first creates a small temporary file with the same name, then replaces it." : "";
    const accepted = window.confirm(`This check uploads a temporary ${sizeMiB} MiB + 1 byte file to the selected document library folder, downloads and verifies it, then deletes it.${recoveryNote} Continue?`);
    if (!accepted) return;
    setRoundTripRunning(true);
    setRoundTrip(undefined);
    const start = performance.now();
    let uploadedItemId: string | undefined;
    let fileName = "";
    let payload: Uint8Array | undefined;
    let pausedForRecovery = false;
    let detail = "";
    let failure: unknown;
    const originalFetch = window.fetch;
    let uploadSessionUrl: string | undefined;
    let interruptionInjected = false;
    let recoveryRangeRead = false;
    let replacedItemId: string | undefined;
    let finalResponseDropped = false;
    try {
      const name = `SPFxFastRoundTrip-${crypto.randomUUID()}.txt`;
      fileName = name;
      if (conflictBehavior === "replace") {
        const existing = await services.client.uploadDriveItem(driveId, folderItemId, name, new Blob([`Synthetic original ${crypto.randomUUID()}`], { type: "text/plain" }));
        replacedItemId = existing.id;
        uploadedItemId = existing.id;
      }
      const expected = `SPFx Fast ${sizeMiB} MiB upload check ${crypto.randomUUID()}`;
      const filePayload = new Uint8Array(sizeMiB * 1024 * 1024 + 1);
      payload = filePayload;
      filePayload.set(new TextEncoder().encode(expected));
      filePayload[filePayload.length - 1] = 0x5a;
      if (simulateInterruption) {
        window.fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
          const requestUrl = input instanceof Request ? input.url : input.toString();
          const method = init?.method ?? (input instanceof Request ? input.method : "GET");
          const headers = new Headers(init?.headers ?? (input instanceof Request ? input.headers : undefined));
          if (method === "PUT" && headers.get("Content-Range")?.startsWith("bytes 0-")) {
            uploadSessionUrl = requestUrl;
            const response = await originalFetch.call(window, input, init);
            if (response.status === 202) {
              interruptionInjected = true;
              await response.body?.cancel().catch(() => undefined);
              throw new TypeError("Simulated connection loss after SharePoint accepted the first upload chunk.");
            }
            return response;
          }
          const response = await originalFetch.call(window, input, init);
          if (interruptionInjected && method === "GET" && requestUrl === uploadSessionUrl && response.ok) recoveryRangeRead = true;
          return response;
        };
      } else if (conflictBehavior === "replace") {
        window.fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
          const requestUrl = input instanceof Request ? input.url : input.toString();
          const method = init?.method ?? (input instanceof Request ? input.method : "GET");
          const headers = new Headers(init?.headers ?? (input instanceof Request ? input.headers : undefined));
          const response = await originalFetch.call(window, input, init);
          if (method === "PUT" && headers.get("Content-Range")?.startsWith("bytes 0-") && response.status === 202) uploadSessionUrl = requestUrl;
          if (method === "PUT" && requestUrl === uploadSessionUrl && (response.status === 200 || response.status === 201)) {
            finalResponseDropped = true;
            await response.body?.cancel().catch(() => undefined);
            throw new TypeError("Simulated lost final response after SharePoint committed the replacement.");
          }
          return response;
        };
      }
      const item = await services.client.uploadDriveItem(driveId, folderItemId, name, new Blob([filePayload], { type: "text/plain" }), { chunkSize: 10 * 1024 * 1024, conflictBehavior });
      window.fetch = originalFetch;
      if (conflictBehavior === "replace" && !finalResponseDropped) throw new Error("The final successful replacement response was not dropped.");
      if (replacedItemId && item.id !== replacedItemId) throw new Error("Replacing the temporary file created a different drive item.");
      uploadedItemId = item.id;
      if (simulateInterruption && (!interruptionInjected || !recoveryRangeRead)) throw new Error("The upload did not demonstrate a server-confirmed resume after the simulated interruption.");
      const downloaded = await services.client.downloadDriveItem(driveId, item.id);
      const bytes = new Uint8Array(await downloaded.arrayBuffer());
      if (bytes.length !== filePayload.length || bytes.some((value, index) => value !== filePayload[index])) throw new Error("Downloaded file content did not match the uploaded check payload.");
      detail = conflictBehavior === "replace"
        ? `Graph replaced the existing temporary file after its final response was lost; ${sizeMiB} MiB + 1 byte matched after download.`
        : simulateInterruption
        ? `Graph upload recovered from a simulated lost response after the first accepted chunk; server range was read, ${sizeMiB} MiB + 1 byte matched after download.`
        : `Graph resumable upload and fresh binary download succeeded for ${sizeMiB} MiB + 1 byte; content matched.`;
    } catch (error) {
      const recoveryPayload = payload;
      if (isDriveUploadInterruptedError(error) && fileName && recoveryPayload) {
        const hash = new Uint8Array(await crypto.subtle.digest("SHA-256", recoveryPayload));
        const expectedHash = Array.from(hash, (byte) => ("0" + byte.toString(16)).slice(-2)).join("");
        setUploadRecovery({ error, driveId, fileName, sizeMiB, expectedHash, startedAt: start });
        setRoundTrip({ name: "File upload/download/cleanup", status: "info", durationMs: performance.now() - start, detail: "Upload paused after status retries. Reconnect, then resume this upload to verify its bytes and cleanup." });
        pausedForRecovery = true;
      } else {
        failure = error;
      }
    } finally {
      window.fetch = originalFetch;
      if (uploadedItemId) {
        try {
          await services.client.requestGraph(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(uploadedItemId)}`, { method: "DELETE" });
          detail += " Temporary file deleted.";
        } catch (cleanupError) {
          failure = new Error(`Temporary file cleanup failed (${cleanupError instanceof Error ? cleanupError.name : "RequestError"}); remove the SPFxFastRoundTrip-* file manually.`);
        }
      }
    }
    if (pausedForRecovery) return;
    if (failure) {
      const errorName = failure instanceof Error ? failure.name : "RequestError";
      const httpStatus = failure && typeof failure === "object" && "status" in failure ? Number((failure as { status?: unknown }).status) : undefined;
      const message = failure instanceof Error && failure.message.startsWith("Temporary file cleanup failed")
        ? failure.message
        : `Upload round-trip failed (${Number.isInteger(httpStatus) ? `HTTP ${httpStatus}` : errorName}).${uploadedItemId ? " Temporary file cleanup was attempted." : " No completed temporary file was returned."}`;
      setRoundTrip({ name: "File upload/download/cleanup", status: "fail", durationMs: performance.now() - start, detail: message });
    } else {
      setRoundTrip({ name: "File upload/download/cleanup", status: "pass", durationMs: performance.now() - start, detail });
    }
    setRoundTripRunning(false);
  };

  const resumePendingUpload = async (): Promise<void> => {
    const pending = uploadRecovery;
    if (!pending || uploadResumeRunning) return;
    setUploadResumeRunning(true);
    setRoundTrip({ name: "File upload/download/cleanup", status: "info", durationMs: performance.now() - pending.startedAt, detail: "Resuming the existing upload session and verifying the downloaded file…" });
    let uploadedItemId: string | undefined;
    let pausedAgain = false;
    try {
      const item = await pending.error.resume();
      uploadedItemId = item.id;
      if (item.name !== pending.fileName) throw new Error("Resumed file name did not match the requested upload name.");
      const downloaded = await services.client.downloadDriveItem(pending.driveId, item.id);
      if (downloaded.size !== pending.sizeMiB * 1024 * 1024 + 1) throw new Error("Downloaded file size did not match the upload check.");
      const hash = new Uint8Array(await crypto.subtle.digest("SHA-256", await downloaded.arrayBuffer()));
      const actualHash = Array.from(hash, (byte) => ("0" + byte.toString(16)).slice(-2)).join("");
      if (actualHash !== pending.expectedHash) throw new Error("Downloaded file bytes did not match the upload check.");
      await services.client.requestGraph(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(pending.driveId)}/items/${encodeURIComponent(item.id)}`, { method: "DELETE" });
      uploadedItemId = undefined;
      setUploadRecovery(undefined);
      setRoundTrip({ name: "File upload/download/cleanup", status: "pass", durationMs: performance.now() - pending.startedAt, detail: `Graph upload resumed from its live session, the ${pending.sizeMiB} MiB + 1 byte download matched by SHA-256, and the temporary file was deleted.` });
    } catch (error) {
      if (uploadedItemId) {
        await services.client.requestGraph(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(pending.driveId)}/items/${encodeURIComponent(uploadedItemId)}`, { method: "DELETE" }).catch(() => undefined);
        uploadedItemId = undefined;
      }
      if (isDriveUploadInterruptedError(error)) {
        pausedAgain = true;
        setUploadRecovery({ ...pending, error });
        setRoundTrip({ name: "File upload/download/cleanup", status: "info", durationMs: performance.now() - pending.startedAt, detail: "Upload is still paused. Reconnect and retry Resume upload." });
      } else {
        setUploadRecovery(undefined);
        setRoundTrip({ name: "File upload/download/cleanup", status: "fail", durationMs: performance.now() - pending.startedAt, detail: `Upload resume check failed (${error instanceof Error ? error.name : "RequestError"}). Temporary file cleanup was attempted.` });
      }
    } finally {
      setUploadResumeRunning(false);
      if (!pausedAgain) setRoundTripRunning(false);
    }
  };

  const runRenameConflictCheck = async (): Promise<void> => {
    if (!driveId || roundTripRunning || uploadRecovery || running || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning) return;
    if (!window.confirm("This check creates two temporary files with the same requested name using Graph conflictBehavior=rename, verifies the returned name and downloaded bytes, then deletes both. Continue?")) return;
    setRoundTripRunning(true);
    setRoundTrip(undefined);
    const started = performance.now();
    let originalId: string | undefined;
    let renamedId: string | undefined;
    const originalFetch = window.fetch;
    let uploadSessionUrl: string | undefined;
    let finalResponseDropped = false;
    let tempName = "";
    let payload: Uint8Array | undefined;
    try {
      const name = `SPFxFastRename-${crypto.randomUUID()}.txt`;
      tempName = name;
      const originalText = `Original ${crypto.randomUUID()}`;
      const replacementText = `Renamed ${crypto.randomUUID()}`;
      const original = await services.client.uploadDriveItem(driveId, folderItemId, name, new Blob([originalText], { type: "text/plain" }));
      originalId = original.id;
      const filePayload = new Uint8Array(20 * 1024 * 1024 + 1);
      payload = filePayload;
      filePayload.set(new TextEncoder().encode(replacementText));
      filePayload[filePayload.length - 1] = 0x5a;
      window.fetch = async (input: RequestInfo | URL, init?: RequestInit): Promise<Response> => {
        const requestUrl = input instanceof Request ? input.url : input.toString();
        const method = init?.method ?? (input instanceof Request ? input.method : "GET");
        const headers = new Headers(init?.headers ?? (input instanceof Request ? input.headers : undefined));
        const response = await originalFetch.call(window, input, init);
        if (method === "PUT" && headers.get("Content-Range")?.startsWith("bytes 0-") && response.status === 202) uploadSessionUrl = requestUrl;
        if (method === "PUT" && requestUrl === uploadSessionUrl && (response.status === 200 || response.status === 201)) {
          finalResponseDropped = true;
          await response.body?.cancel().catch(() => undefined);
          throw new TypeError("Simulated lost final response after SharePoint committed the renamed file.");
        }
        return response;
      };
      const renamed = await services.client.uploadDriveItem(driveId, folderItemId, name, new Blob([filePayload], { type: "text/plain" }), { chunkSize: 10 * 1024 * 1024, conflictBehavior: "rename" });
      window.fetch = originalFetch;
      renamedId = renamed.id;
      if (!finalResponseDropped) throw new Error("The final successful rename response was not dropped.");
      if (renamed.id === original.id || renamed.name === name || !renamed.name?.startsWith(name.replace(/\.txt$/i, ""))) {
        throw new Error("Graph did not return a distinct, renamed file for the existing-name conflict.");
      }
      const downloaded = await services.client.downloadDriveItem(driveId, renamed.id);
      const bytes = new Uint8Array(await downloaded.arrayBuffer());
      if (bytes.length !== filePayload.length || bytes.some((value, index) => value !== filePayload[index])) throw new Error("The renamed file content did not match the upload.");
      setRoundTrip({ name: "Graph rename conflict", status: "pass", durationMs: performance.now() - started, detail: `After the final response was dropped, Graph name lookup recovered “${renamed.name}”; its distinct item ID and 20 MiB + 1 byte download matched. Both temporary files were deleted.` });
    } catch (error) {
      setRoundTrip({ name: "Graph rename conflict", status: "fail", durationMs: performance.now() - started, detail: `Rename conflict check failed (${error instanceof Error ? error.name : "RequestError"}).${finalResponseDropped ? " Inspect the folder for an SPFxFastRename-* file if recovery failed." : ""}` });
    } finally {
      window.fetch = originalFetch;
      if (!renamedId && originalId && finalResponseDropped && tempName && payload) {
        try {
          const extensionIndex = tempName.lastIndexOf(".");
          const stem = extensionIndex > 0 ? tempName.slice(0, extensionIndex) : tempName;
          const extension = extensionIndex > 0 ? tempName.slice(extensionIndex) : "";
          const parentId = folderItemId === "root"
            ? (await services.client.requestGraph(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/root`) as { id?: string }).id
            : folderItemId;
          if (parentId) {
            const candidates: string[] = [];
            for (let index = 1; index <= 5; index++) candidates.push(`${stem} ${index}${extension}`, `${stem} (${index})${extension}`);
            for (const candidateName of candidates) {
              let item: { id?: string; name?: string; size?: number; file?: unknown };
              try {
                item = await services.client.requestGraph(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}:/${encodeURIComponent(candidateName)}:?$select=id,name,size,file`) as typeof item;
              } catch { continue; }
              if (!item.id || item.name !== candidateName || !item.file || item.size !== payload.length) continue;
              const content = await services.client.downloadDriveItem(driveId, item.id);
              const bytes = new Uint8Array(await content.arrayBuffer());
              if (bytes.length === payload.length && bytes.every((value, index) => value === payload![index])) { renamedId = item.id; break; }
            }
          }
        } catch { /* Keep the validation failure visible; cleanup below still removes known items. */ }
      }
      for (const itemId of [renamedId, originalId]) if (itemId) {
        try {
          await services.client.requestGraph(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`, { method: "DELETE" });
        } catch {
          setRoundTrip({ name: "Graph rename conflict", status: "fail", durationMs: performance.now() - started, detail: "Temporary file cleanup failed; remove the SPFxFastRename-* files manually." });
        }
      }
      setRoundTripRunning(false);
    }
  };

  const runDriveDelta = async (): Promise<void> => {
    if (!driveId || deltaRunning || listDeltaRunning || running || roundTripRunning || crudRunning || referenceRunning || versionRunning || recoveryRunning) return;
    setDeltaRunning(true);
    setDeltaResult(await check("Drive delta (one page)", async () => {
      const result = await services.client.syncDrive(driveId, { maxPages: 1 });
      return result.complete
        ? `Delta completed in ${result.pages} page(s); ${result.changed} item(s) synchronized.`
        : `Read one page (${result.changed} item(s)); checkpoint saved. Run again to continue.`;
    }));
    setDeltaRunning(false);
  };

  const runDriveDeltaPaging = async (): Promise<void> => {
    if (!driveId || deltaRunning || listDeltaRunning || running || roundTripRunning || crudRunning || referenceRunning || versionRunning || recoveryRunning) return;
    if (!window.confirm("This check creates one temporary folder with 3,000 synthetic child folders, resumes Graph drive delta one page at a time, then deletes the temporary folder. Continue?")) return;
    setDeltaRunning(true);
    setDeltaResult(undefined);
    setDeltaResult(await check("Graph drive delta paging", async () => {
      const suffix = crypto.randomUUID().slice(0, 8);
      const prefix = `SPFx Fast drive delta ${suffix} `;
      const parentName = `SPFxFastDriveDelta_${suffix}`;
      const itemCount = 3000;
      let parentId: string | undefined;
      let failure: unknown;
      let cleanupFailed = false;
      let detail = "";
      let stage = "create parent folder";
      let createMs = 0;
      let syncMs = 0;
      let cleanupMs = 0;
      let throttleRetries = 0;
      try {
        const parent = await services.client.createDriveFolder(driveId, folderItemId, parentName);
        parentId = parent.id;
        const childrenPath = `/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}/children`;
        const folderBatches = Array.from({ length: Math.ceil(itemCount / 20) }, (_, batchIndex) => {
          const start = batchIndex * 20;
          const end = Math.min(start + 20, itemCount);
          return Array.from({ length: end - start }, (_, offset) => {
            const index = start + offset;
            return {
              id: `folder-${index}`,
              method: "POST" as const,
              url: childrenPath,
              body: {
                name: `${prefix}${index}`,
                folder: {},
                "@microsoft.graph.conflictBehavior": "fail",
              },
            };
          });
        });
        const createStarted = performance.now();
        for (let offset = 0; offset < folderBatches.length; offset += 3) {
          const results = await Promise.all(folderBatches.slice(offset, offset + 3).map((batch) => batchWithThrottledSubrequestRetry(services.client, batch, (count) => { throttleRetries += count; })));
          let failed = 0;
          let firstFailure: { status: number } | undefined;
          for (const batchResults of results) for (const result of batchResults) if (result.status < 200 || result.status >= 300) {
            failed++;
            if (!firstFailure) firstFailure = result;
          }
          if (failed) {
            throw Object.assign(new Error(`Graph batch failed to create ${failed} folder(s); attempted ${Math.min(itemCount, (offset + 3) * 20)} of ${itemCount}.`), { status: firstFailure?.status });
          }
        }
        createMs = performance.now() - createStarted;
        stage = "sync delta";
        const store = createMemoryStore();
        const probeClient = createClient({ scope: services.scope, store, ...createSPFxGraphTransport(services.graph) });
        const syncStarted = performance.now();
        let sync = await probeClient.syncDrive(driveId, { maxPages: 1 });
        let pages = sync.pages;
        for (let attempt = 1; !sync.complete && attempt < 40; attempt++) {
          sync = await probeClient.syncDrive(driveId, { maxPages: 1 });
          pages += sync.pages;
        }
        syncMs = performance.now() - syncStarted;
        stage = "verify cached items and paging";
        const cachedItems = await probeClient.getDriveItems(driveId);
        const cachedProbeItems = cachedItems.filter((item) => item.name?.startsWith(prefix));
        if (!sync.complete || pages < 2 || cachedProbeItems.length !== itemCount) {
          throw new Error(`Graph drive delta completed=${sync.complete}, pages=${pages}, cached=${cachedProbeItems.length}/${itemCount}.`);
        }
        detail = `Created ${itemCount} synthetic folders in ${createMs.toFixed(0)} ms (batch throttle retries: ${throttleRetries}); bounded delta cached all items across ${pages} pages in ${syncMs.toFixed(0)} ms.`;
      } catch (error) {
        failure = error;
      } finally {
        const cleanupStarted = performance.now();
        if (parentId) {
          try {
            await services.client.requestGraph(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentId)}`, { method: "DELETE" });
            detail += " Temporary folder and children deleted.";
          } catch {
            cleanupFailed = true;
          }
        }
        cleanupMs = performance.now() - cleanupStarted;
        if (detail) detail += ` Cleanup: ${cleanupMs.toFixed(0)} ms.`;
      }
      if (cleanupFailed) throw new Error(`Temporary drive cleanup failed. Remove ${parentName} if it remains.`);
      if (failure) {
        const status = failure && typeof failure === "object" && "status" in failure
          ? Number((failure as { status?: unknown }).status)
          : undefined;
        throw Object.assign(new Error("Graph drive delta probe did not pass."), {
          validationDetail: `Probe failed during ${stage}${Number.isInteger(status) ? ` (HTTP ${status})` : ""}.`,
        });
      }
      return detail;
    }));
    setDeltaRunning(false);
  };

  const runListDeltaPaging = async (): Promise<void> => {
    if (listDeltaRunning || deltaRunning || running || roundTripRunning || crudRunning || referenceRunning || versionRunning || recoveryRunning) return;
    if (!window.confirm("This check creates a disposable list, adds 3,000 synthetic work tickets in Graph batches, tests multi-page delta and IndexedDB queries, then deletes the list and cache. Continue?")) return;
    setListDeltaRunning(true);
    setListDeltaResult(undefined);
    setListDeltaResult(await check("3,000-item work-queue benchmark", async () => {
      const suffix = crypto.randomUUID().slice(0, 8);
      const title = `SPFxFastWorkload_${suffix}`;
      const itemCount = 3000;
      const statuses = ["New", "Assigned", "In Progress", "Resolved"];
      const requestTypes = ["Access", "Hardware", "Software", "Facilities"];
      const departments = ["Finance", "HR", "IT", "Operations", "Sales"];
      let temporaryListId: string | undefined;
      let failure: unknown;
      let cleanupFailed = false;
      let detail = "";
      let stage = "create disposable list";
      const scope = `${services.scope}:workload:${suffix}`;
      const store = createIndexedDbStore();
      try {
        const response = await services.rest.request("/_api/web/lists", { method: "POST", headers: {
          Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose", "OData-Version": "3.0",
        }, body: { __metadata: { type: "SP.List" }, BaseTemplate: 100, Title: title } }) as { Id?: string; d?: { Id?: string } };
        const created = response.d ?? response;
        if (!created?.Id) throw new Error("SharePoint did not return a temporary list ID.");
        temporaryListId = created.Id.replace(/[{}]/g, "");

        const addField = async (schema: string): Promise<void> => {
          await services.rest.request(`/_api/web/lists(guid'${temporaryListId}')/fields/CreateFieldAsXml`, { method: "POST", headers: {
            Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose", "OData-Version": "3.0",
          }, body: { parameters: { __metadata: { type: "SP.XmlSchemaFieldCreationInformation" }, SchemaXml: schema, Options: 0 } } });
        };
        const choiceField = (name: string, choices: string[]): string => `<Field ID="{${crypto.randomUUID()}}" Type="Choice" Name="${name}" StaticName="${name}" DisplayName="${name}" Format="Dropdown"><CHOICES>${choices.map((choice) => `<CHOICE>${choice}</CHOICE>`).join("")}</CHOICES></Field>`;
        stage = "create workload fields";
        await addField(choiceField("Status", statuses));
        await addField(choiceField("RequestType", requestTypes));
        await addField(`<Field ID="{${crypto.randomUUID()}}" Type="Number" Name="Priority" StaticName="Priority" DisplayName="Priority" Min="1" Max="5" Decimals="0" />`);
        await addField(`<Field ID="{${crypto.randomUUID()}}" Type="Text" Name="Department" StaticName="Department" DisplayName="Department" MaxLength="80" />`);
        await addField(`<Field ID="{${crypto.randomUUID()}}" Type="DateTime" Name="DueDate" StaticName="DueDate" DisplayName="DueDate" Format="DateOnly" />`);

        const listPath = `/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(temporaryListId)}/items`;
        const createStarted = performance.now();
        const batches = Array.from({ length: Math.ceil(itemCount / 20) }, (_, batchIndex) => {
          const start = batchIndex * 20;
          const end = Math.min(start + 20, itemCount);
          return Array.from({ length: end - start }, (_, offset) => {
            const index = start + offset;
            return {
              id: `ticket-${index}`,
              method: "POST" as const,
              url: listPath,
              body: { fields: {
                Title: `Ticket ${suffix}-${index + 1}`,
                Status: statuses[index % statuses.length],
                RequestType: requestTypes[Math.floor(index / statuses.length) % requestTypes.length],
                Priority: (index % 5) + 1,
                Department: departments[Math.floor(index / 5) % departments.length],
                DueDate: new Date(Date.now() + (index % 60) * 86_400_000).toISOString(),
              } },
            };
          });
        });
        stage = `seed ${itemCount} work tickets`;
        for (let offset = 0; offset < batches.length; offset += 3) {
          const results = await Promise.all(batches.slice(offset, offset + 3).map((batch) => batchWithThrottledSubrequestRetry(services.client, batch)));
          const failed: { status: number }[] = [];
          for (const batchResults of results) for (const result of batchResults) if (result.status < 200 || result.status >= 300) failed.push(result);
          if (failed.length) throw Object.assign(new Error(`Graph batch failed to create ${failed.length} of ${Math.min(itemCount, (offset + 3) * 20)} attempted work tickets.`), { status: failed[0].status });
        }
        const createMs = performance.now() - createStarted;
        const probeClient = createClient({ scope, store, ...createSPFxGraphTransport(services.graph) });
        stage = "sync list delta into IndexedDB";
        const syncStarted = performance.now();
        const sync = await probeClient.syncList(siteId, temporaryListId, { fields: ["Title", "Status", "RequestType", "Priority", "Department", "DueDate"], maxPages: 20 });
        const syncMs = performance.now() - syncStarted;
        const cachedItems = await store.getItems(scope, siteId, temporaryListId);
        if (sync.pages < 2 || cachedItems.length !== itemCount) {
          throw new Error(`Graph delta returned ${sync.pages} page(s) and cached ${cachedItems.length} of ${itemCount} items.`);
        }
        if (new Set(cachedItems.map((item) => item.fields?.Status)).size !== statuses.length || new Set(cachedItems.map((item) => item.fields?.RequestType)).size !== requestTypes.length) {
          throw new Error("Graph delta did not return all expected Status and RequestType values.");
        }
        stage = "query IndexedDB";
        const query = { equals: { Status: "Assigned" }, limit: 100 };
        const queryStarted = performance.now();
        const filtered = await probeClient.queryItems(siteId, temporaryListId, query);
        const queryMs = performance.now() - queryStarted;
        const warmStarted = performance.now();
        const warm = await probeClient.queryItems(siteId, temporaryListId, query);
        const warmMs = performance.now() - warmStarted;
        if (filtered.length !== 100 || warm.length !== 100) throw new Error(`Local query returned ${filtered.length} rows and warm query returned ${warm.length}; expected 100 each.`);
        detail = `Created ${itemCount} synthetic tickets in Graph batches of 20 (${createMs.toFixed(0)} ms). IndexedDB delta read ${sync.pages} pages in ${syncMs.toFixed(0)} ms. Local Status=Assigned query returned 100 rows in ${queryMs.toFixed(1)} ms; warm repeat took ${warmMs.toFixed(1)} ms. Rows included choice, number, text, and date fields.`;
      } catch (error) {
        failure = error;
      } finally {
        if (temporaryListId) {
          try {
            await services.rest.request(`/_api/web/lists(guid'${temporaryListId}')`, { method: "DELETE", headers: { "If-Match": "*" } });
            detail += " Temporary list deleted.";
          } catch {
            cleanupFailed = true;
          }
        }
        try { await store.clearScope(scope); }
        catch { cleanupFailed = true; }
      }
      if (cleanupFailed) throw new Error(`Temporary list cleanup failed. Remove ${title} if it remains.`);
      if (failure) {
        const status = failure && typeof failure === "object" && "status" in failure
          ? Number((failure as { status?: unknown }).status)
          : undefined;
        const cause = failure instanceof Error ? `: ${failure.message.replace(/https?:\/\/\S+/g, "[url]").replace(/[\r\n]+/g, " ").slice(0, 120)}` : "";
        throw Object.assign(new Error("Graph list workload did not pass."), {
          validationDetail: `Probe failed during ${stage}${Number.isInteger(status) ? ` (HTTP ${status})` : failure instanceof Error ? ` (${failure.name})` : ""}${cause}.`,
        });
      }
      return detail;
    }));
    setListDeltaRunning(false);
  };

  const runListCrud = async (): Promise<void> => {
    if (crudRunning || running || roundTripRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning) return;
    const accepted = window.confirm("This check will create one uniquely named list item, read it back, update it, and delete it. Continue?");
    if (!accepted) return;
    setCrudRunning(true);
    setCrudResult(undefined);
    const start = performance.now();
    const title = `SPFx Fast CRUD ${crypto.randomUUID()}`;
    let itemId: string | undefined;
    let deleted = false;
    let stage = "create";
    let cleanupNote = "";
    let failure: unknown;
    try {
      const created = await services.client.createItem(siteId, listId, { Title: title });
      itemId = created.id;
      const itemUrl = new URL(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}`);
      itemUrl.searchParams.set("$expand", "fields($select=Title)");
      stage = "read after create";
      const createdRead = await services.client.requestGraph(itemUrl.toString()) as { fields?: Record<string, unknown> };
      if (createdRead.fields?.Title !== title) throw new Error("Created item did not read back with the expected title.");
      stage = "update";
      await services.client.updateItem(siteId, listId, itemId, { Title: `${title} updated` }, { eTag: created.eTag });
      stage = "read after update";
      const updatedRead = await services.client.requestGraph(itemUrl.toString()) as { fields?: Record<string, unknown> };
      if (updatedRead.fields?.Title !== `${title} updated`) throw new Error("Updated item did not read back with the expected title.");
      stage = "delete";
      await services.client.deleteItem(siteId, listId, itemId);
      deleted = true;
    } catch (error) {
      failure = error;
    } finally {
      if (itemId && !deleted) {
        try {
          await services.client.deleteItem(siteId, listId, itemId);
          cleanupNote = " Temporary item cleanup succeeded.";
        } catch {
          cleanupNote = ` Temporary item cleanup failed; remove “${title}” from the selected list if it was created.`;
        }
      }
    }
    const httpStatus = failure && typeof failure === "object" && "status" in failure ? Number((failure as { status?: unknown }).status) : undefined;
    setCrudResult({
      name: "Graph list item CRUD",
      status: failure ? "fail" : "pass",
      durationMs: performance.now() - start,
      detail: failure
        ? `Failed during ${stage}${Number.isInteger(httpStatus) ? ` (HTTP ${httpStatus})` : ` (${failure instanceof Error ? failure.name : "RequestError"})`}.${cleanupNote}`
        : "Create, read, update, read, and delete succeeded for one temporary list item.",
    });
    setCrudRunning(false);
  };

  const runPersonLookupProbe = async (): Promise<void> => {
    if (referenceRunning || running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || versionRunning || recoveryRunning) return;
    if (!window.confirm("This check creates two temporary lists, adds one Person field and one Lookup field, tests a synthetic Graph item, then deletes the item and both lists. Continue?")) return;
    setReferenceRunning(true);
    setReferenceResult(await check("Person/Lookup Graph write", async () => {
      const suffix = crypto.randomUUID().slice(0, 8);
      const sourceTitle = `SPFxFastRefProbeSource_${suffix}`;
      const targetTitle = `SPFxFastRefProbeTarget_${suffix}`;
      let sourceListId: string | undefined;
      let targetListId: string | undefined;
      let itemId: string | undefined;
      let itemDeleted = false;
      let failure: unknown;
      let detail = "";
      const cleanupErrors: unknown[] = [];
      try {
        const createList = async (title: string): Promise<string> => {
          const response = await services.rest.request("/_api/web/lists", { method: "POST", headers: {
            Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose", "OData-Version": "3.0",
          }, body: {
            __metadata: { type: "SP.List" }, BaseTemplate: 100, Title: title,
          } }) as { Id?: string; d?: { Id?: string } };
          const result = response.d ?? response;
          if (!result?.Id) throw new Error("SharePoint did not return a temporary list ID.");
          return result.Id.replace(/[{}]/g, "");
        };
        sourceListId = await createList(sourceTitle);
        targetListId = await createList(targetTitle);
        const lookupField = `<Field ID="{${crypto.randomUUID()}}" Type="Lookup" List="{${sourceListId}}" ShowField="Title" Name="Related" StaticName="Related" DisplayName="Related" Required="FALSE" />`;
        const personField = `<Field ID="{${crypto.randomUUID()}}" Type="User" List="UserInfo" ShowField="ImnName" Name="Owner" StaticName="Owner" DisplayName="Owner" UserSelectionMode="PeopleOnly" Required="FALSE" />`;
        const addField = async (id: string, schema: string): Promise<void> => {
          await services.rest.request(`/_api/web/lists(guid'${id}')/fields/CreateFieldAsXml`, { method: "POST", headers: {
            Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose", "OData-Version": "3.0",
          }, body: { parameters: {
            __metadata: { type: "SP.XmlSchemaFieldCreationInformation" }, SchemaXml: schema, Options: 0,
          } } });
        };
        await addField(targetListId, lookupField);
        await addField(targetListId, personField);
        const sourceType = await services.rest.get(`/_api/web/lists(guid'${sourceListId}')?$select=ListItemEntityTypeFullName`) as { ListItemEntityTypeFullName?: string };
        const sourceResponse = await services.rest.request(`/_api/web/lists(guid'${sourceListId}')/items`, { method: "POST", headers: {
          Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose", "OData-Version": "3.0",
        }, body: {
          __metadata: { type: sourceType.ListItemEntityTypeFullName }, Title: "Synthetic lookup target",
        } }) as { Id?: number; d?: { Id?: number } };
        const sourceItem = sourceResponse.d ?? sourceResponse;
        if (!Number.isSafeInteger(sourceItem?.Id)) throw new Error("SharePoint did not return the synthetic lookup item's numeric ID.");
        const user = await services.sharepoint.getCurrentUser();
        if (!Number.isSafeInteger(user.Id)) throw new Error("SharePoint did not return the signed-in user's numeric ID.");
        const targetItem = await services.client.createItem(siteId, targetListId, {
          Title: `SPFx Fast reference probe ${suffix}`,
          OwnerLookupId: user.Id,
          RelatedLookupId: sourceItem.Id,
        });
        itemId = targetItem.id;
        const itemUrl = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(targetListId)}/items/${encodeURIComponent(itemId)}?$expand=fields($select=OwnerLookupId,RelatedLookupId)`;
        const createdFields = (await services.client.requestGraph(itemUrl) as { fields?: Record<string, unknown> }).fields;
        if (String(createdFields?.OwnerLookupId) !== String(user.Id) || String(createdFields?.RelatedLookupId) !== String(sourceItem.Id)) throw new Error("Graph did not read back the requested Person and Lookup IDs.");
        await services.client.updateItem(siteId, targetListId, itemId, { Title: `SPFx Fast reference probe ${suffix} updated` }, { eTag: targetItem.eTag });
        const updatedFields = (await services.client.requestGraph(itemUrl) as { fields?: Record<string, unknown> }).fields;
        if (String(updatedFields?.OwnerLookupId) !== String(user.Id) || String(updatedFields?.RelatedLookupId) !== String(sourceItem.Id)) throw new Error("Graph did not preserve the Person and Lookup IDs after update.");
        await services.client.deleteItem(siteId, targetListId, itemId);
        itemDeleted = true;
        detail = "Graph created, read, updated, and deleted a synthetic item with matching Person and Lookup IDs; both temporary lists were deleted.";
      } catch (error) {
        failure = error;
      } finally {
        if (itemId && !itemDeleted) {
          try { await services.client.deleteItem(siteId, targetListId!, itemId); }
          catch (error) { cleanupErrors.push(error); }
        }
        for (const id of [targetListId, sourceListId]) if (id) {
          try { await services.rest.request(`/_api/web/lists(guid'${id}')`, { method: "DELETE", headers: { "If-Match": "*" } }); }
          catch (error) { cleanupErrors.push(error); }
        }
      }
      if (cleanupErrors.length) throw new Error(`Temporary list cleanup failed. Remove ${sourceTitle} and ${targetTitle} if they remain.`);
      if (failure) throw failure;
      return detail;
    }));
    setReferenceRunning(false);
  };

  const runDeltaRecoveryProbe = async (): Promise<void> => {
    if (recoveryRunning || running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning) return;
    if (!window.confirm("This check creates a temporary list, seeds a malformed cursor for that same list, verifies Graph rejects it without clearing cached data, then deletes the list. Continue?")) return;
    setRecoveryRunning(true);
    setRecoveryResult(await check("Malformed same-list delta cursor handling", async () => {
      const suffix = crypto.randomUUID().slice(0, 8);
      const title = `SPFxFastDeltaRecovery_${suffix}`;
      let temporaryListId: string | undefined;
      let failure: unknown;
      let cleanupFailed = false;
      try {
        const response = await services.rest.request("/_api/web/lists", { method: "POST", headers: {
          Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose", "OData-Version": "3.0",
        }, body: { __metadata: { type: "SP.List" }, BaseTemplate: 100, Title: title } }) as { Id?: string; d?: { Id?: string } };
        const created = response.d ?? response;
        if (!created?.Id) throw new Error("SharePoint did not return the temporary list ID.");
        temporaryListId = created.Id.replace(/[{}]/g, "");

        const store = createMemoryStore();
        const scope = `${services.scope}:delta-recovery:${suffix}`;
        const listPath = `/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(temporaryListId)}/items/delta`;
        const invalidCursor = new URL(listPath, "https://graph.microsoft.com");
        invalidCursor.searchParams.set("token", `invalid-${suffix}`);
        await store.setCursor(scope, siteId, temporaryListId, invalidCursor.toString());
        await store.putItems(scope, siteId, temporaryListId, [{ id: "cached-sentinel", fields: { Title: "Synthetic cached row" } }]);
        const transport = createSPFxGraphTransport(services.graph);
        const recoveryClient = createClient({ scope, store, ...transport });
        let status: number | undefined;
        try { await recoveryClient.syncList(siteId, temporaryListId, { maxPages: 2 }); }
        catch (error) {
          status = error && typeof error === "object" && "status" in error ? Number((error as { status?: unknown }).status)
            : error && typeof error === "object" && "statusCode" in error ? Number((error as { statusCode?: unknown }).statusCode) : undefined;
        }
        if (status !== 400) throw new Error(`Expected Graph to reject the malformed cursor with HTTP 400; received ${Number.isInteger(status) ? `HTTP ${status}` : "no error"}.`);
        const cachedItems = await store.getItems(scope, siteId, temporaryListId);
        const savedCursor = await store.getCursor(scope, siteId, temporaryListId);
        if (cachedItems.length !== 1 || cachedItems[0]?.id !== "cached-sentinel" || savedCursor !== invalidCursor.toString()) {
          throw new Error("The rejected cursor changed cached rows or its checkpoint.");
        }
      } catch (error) {
        failure = error;
      } finally {
        if (temporaryListId) {
          try { await services.rest.request(`/_api/web/lists(guid'${temporaryListId}')`, { method: "DELETE", headers: { "If-Match": "*" } }); }
          catch { cleanupFailed = true; }
        }
      }
      if (cleanupFailed) throw new Error(`Temporary-list cleanup failed. Remove ${title} if it remains.`);
      if (failure) throw failure;
      return "Graph rejected the malformed same-list cursor with HTTP 400 and the client preserved its cached row and checkpoint. This is not an expired-token reset test.";
    }));
    setRecoveryRunning(false);
  };

  const runExpiredDriveCursorProbe = async (): Promise<void> => {
    if (!driveId || recoveryRunning || running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning) return;
    setRecoveryRunning(true);
    setRecoveryResult(await check("Simulated drive delta 410 recovery", async () => {
      const suffix = crypto.randomUUID().slice(0, 8);
      const scope = `${services.scope}:drive-410-recovery:${suffix}`;
      const store = createMemoryStore();
      const driveStoreKey = "__spfx_fast_drive_items__";
      const sentinelId = `expired-cursor-sentinel-${suffix}`;
      const staleCursor = new URL(`/v1.0/drives/${encodeURIComponent(driveId)}/root/delta`, "https://graph.microsoft.com");
      staleCursor.searchParams.set("token", `expired-${suffix}`);
      await store.putItems(scope, driveStoreKey, driveId, [{ id: sentinelId, fields: { name: "Synthetic stale cache row" } }]);
      await store.setCursor(scope, driveStoreKey, driveId, staleCursor.toString());
      const transport = createSPFxGraphTransport(services.graph);
      let injected = false;
      const probeClient = createClient({
        scope,
        store,
        ...transport,
        get: async (url, signal) => {
          if (!injected && url === staleCursor.toString()) {
            injected = true;
            throw Object.assign(new Error("Synthetic expired delta cursor."), { status: 410 });
          }
          return transport.get(url, signal);
        },
      });
      const result = await probeClient.syncDrive(driveId, { maxPages: 1 });
      const cached = await probeClient.getDriveItems(driveId);
      if (!injected || result.pages !== 1 || cached.some((item) => item.id === sentinelId)) {
        throw new Error("The simulated 410 did not clear stale rows and restart a bounded delta read.");
      }
      return `Injected one synthetic HTTP 410, cleared the stale cache, then read ${result.changed} item(s) from the signed-in drive in one page (complete: ${result.complete}). This validates client recovery, not Graph's tenant-generated 410 response.`;
    }));
    setRecoveryRunning(false);
  };

  const runVersionProbe = async (): Promise<void> => {
    if (!driveId || versionRunning || running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || recoveryRunning) return;
    if (!window.confirm("This check creates one temporary versioned list and one temporary versioned document library, verifies list and file history, downloads the first file version, then deletes both. Continue?")) return;
    setVersionRunning(true);
    setVersionResult(undefined);
    setVersionResult(await check("List/file version history", async () => {
      const suffix = crypto.randomUUID().slice(0, 8);
      const listTitle = `SPFxFastVersionsList_${suffix}`;
      const libraryTitle = `SPFxFastVersionsLibrary_${suffix}`;
      let listId: string | undefined;
      let libraryId: string | undefined;
      let failure: unknown;
      let cleanupFailed = false;
      let stage = "create versioned list";
      let detail = "";
      const createList = async (title: string, baseTemplate: number): Promise<string> => {
        const response = await services.rest.request("/_api/web/lists", { method: "POST", headers: {
          Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose", "OData-Version": "3.0",
        }, body: { __metadata: { type: "SP.List" }, BaseTemplate: baseTemplate, Title: title } }) as { Id?: string; d?: { Id?: string } };
        const created = response.d ?? response;
        if (!created?.Id) throw new Error("SharePoint did not return a temporary list ID.");
        return created.Id.replace(/[{}]/g, "");
      };
      const enableVersioning = async (id: string): Promise<void> => {
        await services.rest.request(`/_api/web/lists(guid'${id}')`, { method: "POST", headers: {
          Accept: "application/json;odata=verbose", "Content-Type": "application/json;odata=verbose", "OData-Version": "3.0", "X-HTTP-Method": "MERGE", "If-Match": "*",
        }, body: { __metadata: { type: "SP.List" }, EnableVersioning: true } });
        const settings = await services.rest.get(`/_api/web/lists(guid'${id}')?$select=EnableVersioning`) as { EnableVersioning?: boolean };
        if (settings.EnableVersioning !== true) throw new Error("SharePoint did not enable version history on the temporary list.");
      };
      try {
        listId = await createList(listTitle, 100);
        stage = "enable list versioning";
        await enableVersioning(listId);
        stage = "create list item versions";
        const created = await services.client.createItem(siteId, listId, { Title: `Version ${suffix} one` });
        const updated = await services.client.updateItem(siteId, listId, created.id, { Title: `Version ${suffix} two` }, { eTag: created.eTag });
        await services.client.updateItem(siteId, listId, created.id, { Title: `Version ${suffix} three` }, { eTag: updated.eTag });
        stage = "read list item versions";
        const listVersions = await services.client.getListItemVersions(siteId, listId, created.id, { maxItems: 10 });
        if (listVersions.length < 3) throw new Error(`Graph returned ${listVersions.length} list version(s); expected at least 3.`);

        stage = "create versioned document library";
        libraryId = await createList(libraryTitle, 101);
        stage = "enable library versioning";
        await enableVersioning(libraryId);
        stage = "resolve library drive";
        const drive = await services.client.requestGraph(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(libraryId)}/drive`) as { id?: string };
        if (!drive.id) throw new Error("Graph did not return a drive for the temporary document library.");
        const fileName = `SPFxFastVersion_${suffix}.txt`;
        const firstText = `Synthetic file version one ${suffix}`;
        const currentText = `Synthetic file version two ${suffix}`;
        stage = "upload file versions";
        const firstFile = await services.client.uploadDriveItem(drive.id, "root", fileName, new Blob([firstText], { type: "text/plain" }));
        const currentFile = await services.client.uploadDriveItem(drive.id, "root", fileName, new Blob([currentText], { type: "text/plain" }), { conflictBehavior: "replace" });
        if (currentFile.id !== firstFile.id) throw new Error("Replacing the synthetic file created a different drive item.");
        stage = "list file versions";
        const fileVersions = await services.client.getDriveItemVersions(drive.id, currentFile.id, { maxItems: 10 });
        if (fileVersions.length < 2) throw new Error(`Graph returned ${fileVersions.length} file version(s); expected at least 2.`);
        const compareVersionIds = (left: string, right: string): number => {
          const a = left.split(".").map(Number);
          const b = right.split(".").map(Number);
          return a.every(Number.isFinite) && b.every(Number.isFinite)
            ? (a[0] ?? 0) - (b[0] ?? 0) || (a[1] ?? 0) - (b[1] ?? 0)
            : 0;
        };
        const priorVersion = fileVersions
          .filter((version) => Number.isFinite(Date.parse(version.lastModifiedDateTime ?? "")))
          .sort((a, b) => Date.parse(a.lastModifiedDateTime!) - Date.parse(b.lastModifiedDateTime!) || compareVersionIds(a.id, b.id))[0]
          ?? fileVersions[fileVersions.length - 1];
        stage = "download historical file version";
        const priorBytes = await services.client.downloadDriveItemVersion(drive.id, currentFile.id, priorVersion.id);
        stage = "download current file";
        const currentBytes = await services.client.downloadDriveItem(drive.id, currentFile.id);
        stage = "verify file version contents";
        if (await priorBytes.text() !== firstText || await currentBytes.text() !== currentText) throw new Error("Graph did not return the expected current and prior file contents.");
        detail = `Graph returned ${listVersions.length} list versions and ${fileVersions.length} file versions; the prior file bytes and current file bytes matched. Both temporary objects were deleted.`;
      } catch (error) {
        failure = error;
      } finally {
        for (const id of [libraryId, listId]) if (id) {
          try { await services.rest.request(`/_api/web/lists(guid'${id}')`, { method: "DELETE", headers: { "If-Match": "*" } }); }
          catch { cleanupFailed = true; }
        }
      }
      if (cleanupFailed) throw new Error(`Temporary version-history cleanup failed. Remove ${listTitle} and ${libraryTitle} if they remain.`);
      if (failure) {
        const status = failure && typeof failure === "object" && "status" in failure ? Number((failure as { status?: unknown }).status) : undefined;
        throw Object.assign(new Error("Version-history check failed."), {
          validationDetail: `Failed during ${stage}${Number.isInteger(status) ? ` (HTTP ${status})` : ` (${failure instanceof Error ? failure.name : "RequestError"})`}. Temporary objects were cleaned up.`,
        });
      }
      return detail;
    }));
    setVersionRunning(false);
  };

  return <section aria-label="Signed-in tenant validation">
    <h2>Tenant validation</h2>
    <Text>The signed-in checks below use the current user&apos;s SPFx authentication. Drive delta caches up to 200 items per run. List CRUD and file upload ask before creating temporary data.</Text>
    <p><Button appearance="secondary" disabled={running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning || !siteId || !listId} onClick={run}>{running ? "Running checks…" : "Run signed-in checks"}</Button></p>
    {checks.length ? <ul aria-live="polite" aria-label="Validation results">{checks.map((result) => <li key={result.name}>
      <strong>{result.status.toUpperCase()}</strong> {result.name} ({result.durationMs.toFixed(0)} ms): {result.detail}
    </li>)}</ul> : null}
    <p><Button data-testid="tenant-validation-run-list-crud" appearance="secondary" disabled={running || crudRunning || roundTripRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning || !siteId || !listId} onClick={runListCrud}>{crudRunning ? "Testing list CRUD…" : "Test list CRUD"}</Button></p>
    {crudResult ? <p data-testid="tenant-validation-list-crud-result" aria-live="polite"><strong>{crudResult.status.toUpperCase()}</strong> {crudResult.name} ({crudResult.durationMs.toFixed(0)} ms): {crudResult.detail}</p> : null}
    <p><Button appearance="secondary" disabled={running || crudRunning || roundTripRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning || !siteId} onClick={runPersonLookupProbe}>{referenceRunning ? "Testing Person/Lookup…" : "Test Person/Lookup writes"}</Button></p>
    {referenceResult ? <p aria-live="polite"><strong>{referenceResult.status.toUpperCase()}</strong> {referenceResult.name} ({referenceResult.durationMs.toFixed(0)} ms): {referenceResult.detail}</p> : null}
    <p><Button appearance="secondary" disabled={running || crudRunning || roundTripRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning || !siteId} onClick={runDeltaRecoveryProbe}>{recoveryRunning ? "Testing cursor handling…" : "Test malformed delta cursor"}</Button></p>
    <Text size={200}>Seeds a malformed cursor in a disposable list and verifies Graph rejection preserves cached data. This does not simulate token expiry.</Text>
    <p><Button appearance="secondary" disabled={running || crudRunning || roundTripRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning || !driveId} onClick={runExpiredDriveCursorProbe}>{recoveryRunning ? "Testing 410 recovery…" : "Test simulated drive 410 recovery"}</Button></p>
    <Text size={200}>Injects one synthetic 410 for a local stale cursor, then confirms a bounded read resumes through the signed-in Graph client. It does not simulate Graph expiring a real tenant token.</Text>
    {recoveryResult ? <p aria-live="polite"><strong>{recoveryResult.status.toUpperCase()}</strong> {recoveryResult.name} ({recoveryResult.durationMs.toFixed(0)} ms): {recoveryResult.detail}</p> : null}
    <p><Button appearance="secondary" disabled={running || crudRunning || roundTripRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning || !siteId || !driveId} onClick={runVersionProbe}>{versionRunning ? "Testing version history…" : "Test list/file versions"}</Button></p>
    <Text size={200}>Creates disposable versioned SharePoint objects, verifies list/file history and prior file bytes, then deletes them.</Text>
    {versionResult ? <p aria-live="polite"><strong>{versionResult.status.toUpperCase()}</strong> {versionResult.name} ({versionResult.durationMs.toFixed(0)} ms): {versionResult.detail}</p> : null}
    <p><Button appearance="secondary" disabled={running || crudRunning || roundTripRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning || !siteId} onClick={runListDeltaPaging}>{listDeltaRunning ? "Testing list delta paging…" : "Test list delta paging"}</Button></p>
    <Text size={200}>Creates 3,000 synthetic work tickets across choice, number, text, and date fields; benchmarks paged Graph delta and IndexedDB queries, then deletes the list and cached rows.</Text>
    {listDeltaResult ? <p aria-live="polite"><strong>{listDeltaResult.status.toUpperCase()}</strong> {listDeltaResult.name} ({listDeltaResult.durationMs.toFixed(0)} ms): {listDeltaResult.detail}</p> : null}
    {driveId ? <div>
      <p><Button appearance="secondary" disabled={running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning} onClick={runDriveDelta}>{deltaRunning ? "Syncing one delta page…" : "Test drive delta (one page)"}</Button></p>
      {deltaResult ? <p aria-live="polite"><strong>{deltaResult.status.toUpperCase()}</strong> {deltaResult.name} ({deltaResult.durationMs.toFixed(0)} ms): {deltaResult.detail}</p> : null}
      <p><Button appearance="secondary" disabled={running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning} onClick={runDriveDeltaPaging}>{deltaRunning ? "Testing drive delta paging…" : "Test drive delta paging"}</Button></p>
      <Text size={200}>Creates a disposable folder with 3,000 synthetic child folders, resumes bounded delta pages, then removes it.</Text>
      <p><Button appearance="secondary" disabled={running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning} onClick={() => runFileRoundTrip(10)}>{roundTripRunning ? "Testing file upload…" : "Test 10 MiB upload and download"}</Button></p>
      <p><Button appearance="secondary" disabled={running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning} onClick={() => runFileRoundTrip(50)}>{roundTripRunning ? "Testing file upload…" : "Test 50 MiB upload and download"}</Button></p>
      <p><Button appearance="secondary" disabled={running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning} onClick={() => runFileRoundTrip(100)}>{roundTripRunning ? "Testing file upload…" : "Test 100 MiB upload and download"}</Button></p>
      <p><Button appearance="secondary" disabled={running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning} onClick={() => runFileRoundTrip(128)}>{roundTripRunning ? "Testing file upload…" : "Test 128 MiB upload and download"}</Button></p>
      <p><Button appearance="secondary" disabled={running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning} onClick={() => runFileRoundTrip(20, true)}>{roundTripRunning ? "Testing file upload…" : "Test upload recovery"}</Button></p>
      <p><Button appearance="secondary" disabled={running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning} onClick={() => runFileRoundTrip(20, false, "replace")}>{roundTripRunning ? "Testing file upload…" : "Test replace final response recovery"}</Button></p>
      <p><Button appearance="secondary" disabled={running || roundTripRunning || crudRunning || deltaRunning || listDeltaRunning || referenceRunning || versionRunning || recoveryRunning} onClick={runRenameConflictCheck}>{roundTripRunning ? "Testing file upload…" : "Test Graph rename conflict"}</Button></p>
      <Text size={200}>Requires write access to the selected folder. A confirmation appears before the temporary file is created.</Text>
      {uploadRecovery ? <p><Button appearance="primary" disabled={uploadResumeRunning} onClick={resumePendingUpload}>{uploadResumeRunning ? "Resuming upload…" : "Resume upload and verify"}</Button></p> : null}
      {roundTrip ? <p aria-live="polite"><strong>{roundTrip.status.toUpperCase()}</strong> {roundTrip.name} ({roundTrip.durationMs.toFixed(0)} ms): {roundTrip.detail}</p> : null}
    </div> : null}
  </section>;
}
