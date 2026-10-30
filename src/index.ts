export interface GraphCollection<T> {
  value: T[];
  "@odata.nextLink"?: string;
  "@odata.deltaLink"?: string;
}

export interface GraphListItem {
  id: string;
  eTag?: string;
  fields?: Record<string, unknown>;
  deleted?: { state: string };
  [key: string]: unknown;
}

export interface GraphCalendarEvent {
  id: string;
  subject?: string;
  start?: { dateTime: string; timeZone: string };
  end?: { dateTime: string; timeZone: string };
  isAllDay?: boolean;
  location?: { displayName?: string; [key: string]: unknown };
  webLink?: string;
  [key: string]: unknown;
}

export interface CalendarViewOptions {
  calendarId?: string;
  maxItems?: number;
  maxPages?: number;
  cursor?: string;
  signal?: AbortSignal;
}

export interface CalendarViewResult {
  items: GraphCalendarEvent[];
  nextCursor?: string;
  complete: boolean;
}

export interface GraphContentType {
  id: string;
  name?: string;
  description?: string;
  group?: string;
  hidden?: boolean;
  isBuiltIn?: boolean;
  parentId?: string;
  readOnly?: boolean;
  sealed?: boolean;
  [key: string]: unknown;
}

export interface GraphTermStoreGroup {
  id: string;
  displayName?: string;
  description?: string;
  scope?: "global" | "site" | string;
  [key: string]: unknown;
}

export interface GraphTermStoreSet {
  id: string;
  description?: string;
  localizedNames?: Array<{ name: string; languageTag: string }>;
  [key: string]: unknown;
}

export interface GraphTermStoreTerm {
  id: string;
  labels?: Array<{ name: string; languageTag: string; isDefault?: boolean }>;
  descriptions?: Array<{ description: string; languageTag: string }>;
  isDeprecated?: boolean;
  lastModifiedDateTime?: string;
  [key: string]: unknown;
}

export interface GraphDriveItem {
  id: string;
  name?: string;
  size?: number;
  eTag?: string;
  webUrl?: string;
  file?: Record<string, unknown>;
  folder?: { childCount?: number };
  parentReference?: { driveId?: string; id?: string; name?: string; path?: string; [key: string]: unknown };
  deleted?: { state?: string; [key: string]: unknown };
  lastModifiedDateTime?: string;
  [key: string]: unknown;
}

/** A large upload paused after Graph's session state could not be confirmed. Keep this error and call resume() when the service is available. */
export class DriveUploadInterruptedError extends Error {
  readonly code = "SPFX_FAST_DRIVE_UPLOAD_INTERRUPTED";
  readonly cause?: unknown;

  constructor(private readonly resumeUpload: () => Promise<GraphDriveItem>, cause?: unknown) {
    super("Upload paused because Graph could not confirm its progress. Keep this page open and use Resume upload when the service is available.");
    this.name = "DriveUploadInterruptedError";
    this.cause = cause;
  }

  resume(): Promise<GraphDriveItem> {
    return this.resumeUpload();
  }
}

/** Identify an interrupted upload across duplicated package bundles or realms. */
export function isDriveUploadInterruptedError(error: unknown): error is DriveUploadInterruptedError {
  if (!error || typeof error !== "object") return false;
  const candidate = error as { code?: unknown; resume?: unknown };
  return candidate.code === "SPFX_FAST_DRIVE_UPLOAD_INTERRUPTED" && typeof candidate.resume === "function";
}

export interface DriveSyncResult {
  pages: number;
  changed: number;
  complete: boolean;
  nextCursor?: string;
  deltaLink?: string;
}

interface DriveDeltaCheckpoint {
  nextLink: string;
  deletedFolderIds: string[];
}

function isDriveDeltaCheckpoint(value: unknown): value is DriveDeltaCheckpoint {
  if (!value || typeof value !== "object") return false;
  const checkpoint = value as Partial<DriveDeltaCheckpoint>;
  if (typeof checkpoint.nextLink !== "string" || !Array.isArray(checkpoint.deletedFolderIds)) return false;
  try {
    const url = new URL(checkpoint.nextLink);
    return url.origin === "https://graph.microsoft.com" && url.pathname.endsWith("/root/delta") && checkpoint.deletedFolderIds.every((id) => typeof id === "string");
  } catch { return false; }
}

export interface GraphListItemVersion {
  id: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: { user?: { id?: string; displayName?: string }; application?: { id?: string; displayName?: string }; [key: string]: unknown };
  fields?: Record<string, unknown>;
  published?: { level?: string; versionId?: number; [key: string]: unknown };
  [key: string]: unknown;
}

export interface GraphDriveItemVersion {
  id: string;
  lastModifiedDateTime?: string;
  lastModifiedBy?: { user?: { id?: string; displayName?: string }; application?: { id?: string; displayName?: string }; [key: string]: unknown };
  size?: number;
  [key: string]: unknown;
}

/** Effective sharing permissions visible to the signed-in Graph caller. */
export interface GraphDriveItemPermission {
  id: string;
  roles?: string[];
  grantedToV2?: { user?: { id?: string; displayName?: string; email?: string }; siteUser?: { id?: string; displayName?: string; email?: string }; [key: string]: unknown };
  grantedToIdentitiesV2?: Array<{ user?: { id?: string; displayName?: string; email?: string }; siteUser?: { id?: string; displayName?: string; email?: string }; [key: string]: unknown }>;
  /** @deprecated Use grantedToV2. */
  grantedTo?: Record<string, unknown>;
  /** @deprecated Use grantedToIdentitiesV2. */
  grantedToIdentities?: Array<Record<string, unknown>>;
  link?: { type?: string; scope?: string; webUrl?: string; [key: string]: unknown };
  invitation?: { email?: string; signInRequired?: boolean; [key: string]: unknown };
  inheritedFrom?: { driveId?: string; id?: string; path?: string; [key: string]: unknown };
  expirationDateTime?: string;
  hasPassword?: boolean;
  [key: string]: unknown;
}

export interface DriveItemSharingLinkOptions {
  type: "view" | "edit";
  scope: "anonymous" | "organization" | "users";
  expirationDateTime?: string;
  signal?: AbortSignal;
}

export interface CopyDriveItemOptions {
  targetDriveId: string;
  targetFolderId: string;
  name?: string;
  conflictBehavior?: "fail" | "replace" | "rename";
  includeAllVersionHistory?: boolean;
  signal?: AbortSignal;
}

export interface GraphRequestOptions {
  method?: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
  retryable?: boolean;
  headers?: Record<string, string>;
  body?: unknown;
  /** Use for file content or other non-JSON payloads. */
  rawBody?: BodyInit;
  responseType?: "json" | "text" | "blob" | "raw";
  signal?: AbortSignal;
}

/** Return parsed JSON for reads and writes. The host adapter owns auth and HTTP error parsing. */
export type GraphRequest = (url: string, options?: GraphRequestOptions) => Promise<unknown>;
export type GraphGet = (url: string, signal?: AbortSignal) => Promise<unknown>;

export interface CacheStore {
  get<T>(key: string): Promise<T | undefined>;
  set<T>(key: string, value: T): Promise<void>;
}

export interface StoredItem {
  key: string;
  listKey: string;
  scope: string;
  itemId: string;
  item: GraphListItem;
  updatedAt: number;
}

export interface LocalStore extends CacheStore {
  getItems(scope: string, siteId: string, listId: string): Promise<GraphListItem[]>;
  queryItems?(scope: string, siteId: string, listId: string, query: ItemQuery): Promise<GraphListItem[]>;
  putItems(scope: string, siteId: string, listId: string, items: GraphListItem[]): Promise<void>;
  deleteItems(scope: string, siteId: string, listId: string, ids: string[]): Promise<void>;
  clearItems(scope: string, siteId: string, listId: string): Promise<void>;
  clearScope(scope: string): Promise<void>;
  getCursor(scope: string, siteId: string, listId: string): Promise<string | undefined>;
  setCursor(scope: string, siteId: string, listId: string, cursor: string): Promise<void>;
  clearCursor(scope: string, siteId: string, listId: string): Promise<void>;
  invalidateListPages(scope: string, siteId: string, listId: string): Promise<void>;
  invalidateDrivePages?(scope: string, driveId: string): Promise<void>;
  subscribe(scope: string, siteId: string, listId: string, listener: () => void): () => void;
}

export interface ReadPage<T> {
  data: GraphCollection<T>;
  source: "cache" | "network";
  refresh: Promise<GraphCollection<T>>;
}

export interface CachedValue<T> {
  data: T;
  source: "cache" | "network";
  refresh: Promise<T>;
}

export type ClientMetric = {
  type: "cache-hit" | "cache-miss" | "local-query" | "network-success" | "network-error" | "delta-sync" | "storage-error" | "mutation" | "retry";
  operation: string;
  durationMs?: number;
  delayMs?: number;
  count?: number;
  error?: unknown;
};

export interface ClientOptions {
  /** Stable tenant and signed-in-user scope. Never use a caller-controlled identity as authorization. */
  scope: string;
  get: GraphGet;
  request?: GraphRequest;
  store?: LocalStore;
  onStorageError?: (error: unknown) => void;
  onMetric?: (metric: ClientMetric) => void;
}

export interface ItemQuery {
  equals?: Record<string, unknown>;
  contains?: { field: string; value: string };
  sort?: { field: string; direction?: "asc" | "desc" };
  limit?: number;
}

export interface GraphListItemFilter {
  field: string;
  operator: "eq" | "ne" | "lt" | "le" | "gt" | "ge" | "startswith";
  value: string | number | boolean;
}

export interface GraphBatchRequest {
  id: string;
  method: "GET" | "POST" | "PUT" | "PATCH" | "DELETE";
  /** Resource path relative to the Graph version root, such as `/sites/{id}/lists`. */
  url: string;
  /** Content-Type defaults to application/json when a write request has a body. */
  headers?: Record<string, string>;
  body?: unknown;
}

const databaseName = "spfx-fast-cache";
const stores = { responses: "responses", items: "items", cursors: "cursors" } as const;
const subscribers = new Map<string, Set<() => void>>();
const instanceId = () => Math.random().toString(36).slice(2);

function listKey(scope: string, siteId: string, listId: string): string {
  return JSON.stringify([scope, siteId, listId]);
}

const driveStoreSiteId = "__spfx_fast_drive_items__";

function driveCheckpointKey(scope: string, driveId: string): string {
  return JSON.stringify([scope, "drive-delta-checkpoint", driveId]);
}

function itemKey(key: string, id: string): string {
  return `${key}\u0000${id}`;
}

function graphListItemFilter(filter?: GraphListItemFilter): string | undefined {
  if (!filter) return undefined;
  if (!/^[A-Za-z0-9_]{1,128}$/.test(filter.field)) throw new TypeError("Filter field must be a SharePoint internal field name.");
  if (!("eq ne lt le gt ge startswith" as const).split(" ").includes(filter.operator)) throw new TypeError("Unsupported Graph list filter operator.");
  if (!["string", "number", "boolean"].includes(typeof filter.value)) throw new TypeError("Filter value must be a string, number, or boolean.");
  if (typeof filter.value === "number" && !Number.isFinite(filter.value)) throw new TypeError("Filter number must be finite.");
  if (filter.operator === "startswith" && typeof filter.value !== "string") throw new TypeError("The startswith filter requires a string value.");
  const value = typeof filter.value === "string" ? `'${filter.value.replaceAll("'", "''")}'` : String(filter.value);
  return filter.operator === "startswith"
    ? `startswith(fields/${filter.field},${value})`
    : `fields/${filter.field} ${filter.operator} ${value}`;
}

function responseKey(scope: string, url: string): string {
  return JSON.stringify([scope, url]);
}

function reportStorageError(options: ClientOptions, error: unknown): void {
  try { options.onStorageError?.(error); } catch { /* Telemetry must not interrupt data access. */ }
  try { options.onMetric?.({ type: "storage-error", operation: "indexeddb", error }); } catch { /* Telemetry must not interrupt data access. */ }
}

export function createIndexedDbStore(name = databaseName): LocalStore {
  let database: Promise<IDBDatabase> | undefined;
  const id = instanceId();
  const channel = typeof BroadcastChannel === "undefined" || typeof window === "undefined"
    ? undefined
    : new BroadcastChannel(`${name}:updates`);

  const notify = (key: string) => {
    subscribers.get(key)?.forEach((listener) => listener());
    channel?.postMessage({ source: id, key });
  };

  if (channel) {
    channel.onmessage = ({ data }: MessageEvent<{ source?: string; key?: string; scope?: string }>) => {
      if (data?.source === id) return;
      if (data?.key) subscribers.get(data.key)?.forEach((listener) => listener());
      if (data?.scope) notifyScopeSubscribers(data.scope);
    };
  }

  const open = (): Promise<IDBDatabase> => {
    if (typeof indexedDB === "undefined") {
      return Promise.reject(new Error("IndexedDB is unavailable in this environment."));
    }

    database ??= new Promise((resolve, reject) => {
      const request = indexedDB.open(name, 3);
      request.onupgradeneeded = (event) => {
        const db = request.result;
        if (!db.objectStoreNames.contains(stores.responses)) db.createObjectStore(stores.responses);
        else if (event.oldVersion < 3) request.transaction!.objectStore(stores.responses).clear();
        if (!db.objectStoreNames.contains(stores.items)) {
          const items = db.createObjectStore(stores.items, { keyPath: "key" });
          items.createIndex("listKey", "listKey", { unique: false });
          items.createIndex("scope", "scope", { unique: false });
        } else {
          const items = request.transaction!.objectStore(stores.items);
          if (!items.indexNames.contains("scope")) {
            items.createIndex("scope", "scope", { unique: false });
            const migration = items.openCursor();
            migration.onsuccess = () => {
              const cursor = migration.result;
              if (cursor) {
                const record = cursor.value as StoredItem;
                try { record.scope = JSON.parse(record.listKey)[0]; cursor.update(record); } catch { /* Ignore malformed old records. */ }
                cursor.continue();
              }
            };
          }
        }
        if (!db.objectStoreNames.contains(stores.cursors)) db.createObjectStore(stores.cursors);
      };
      request.onsuccess = () => {
        request.result.onversionchange = () => request.result.close();
        resolve(request.result);
      };
      request.onerror = () => reject(request.error ?? new Error("Could not open IndexedDB."));
      request.onblocked = () => reject(new Error("IndexedDB upgrade is blocked by another tab."));
    });

    return database;
  };

  const transact = async <T>(storeNames: string[], mode: IDBTransactionMode, run: (tx: IDBTransaction) => T): Promise<T> => {
    const db = await open();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(storeNames, mode);
      let result: T;
      try {
        result = run(tx);
      } catch (error) {
        tx.abort();
        reject(error);
        return;
      }
      tx.oncomplete = () => resolve(result);
      tx.onerror = () => reject(tx.error ?? new Error("IndexedDB transaction failed."));
      tx.onabort = () => reject(tx.error ?? new Error("IndexedDB transaction was aborted."));
    });
  };

  return {
    async get<T>(key: string): Promise<T | undefined> {
      const db = await open();
      return new Promise((resolve, reject) => {
        const request = db.transaction(stores.responses, "readonly").objectStore(stores.responses).get(key);
        request.onsuccess = () => resolve(request.result as T | undefined);
        request.onerror = () => reject(request.error ?? new Error("Could not read cached data."));
      });
    },
    async set<T>(key: string, value: T): Promise<void> {
      await transact([stores.responses], "readwrite", (tx) => tx.objectStore(stores.responses).put(value, key));
    },
    async getItems(scope, siteId, listId) {
      const key = listKey(scope, siteId, listId);
      const db = await open();
      return new Promise((resolve, reject) => {
        const request = db.transaction(stores.items, "readonly").objectStore(stores.items).index("listKey").getAll(key);
        request.onsuccess = () => resolve((request.result as StoredItem[]).map(({ item }) => item));
        request.onerror = () => reject(request.error ?? new Error("Could not read local items."));
      });
    },
    async queryItems(scope, siteId, listId, query) {
      if (query.limit !== undefined && query.limit <= 0) return [];
      const key = listKey(scope, siteId, listId);
      const db = await open();
      return new Promise((resolve, reject) => {
        const tx = db.transaction(stores.items, "readonly");
        const index = tx.objectStore(stores.items).index("listKey");
        if (query.sort) {
          const request = index.getAll(key);
          request.onsuccess = () => resolve(applyItemQuery((request.result as StoredItem[]).map(({ item }) => item), query));
          request.onerror = () => reject(request.error ?? new Error("Could not query local items."));
          return;
        }
        const request = index.openCursor(key);
        const matched: GraphListItem[] = [];
        request.onsuccess = () => {
          const cursor = request.result;
          if (!cursor) return;
          const item = (cursor.value as StoredItem).item;
          if (matchesItemQuery(item, query)) {
            matched.push(item);
            if (!query.sort && query.limit !== undefined && matched.length >= query.limit) return;
          }
          cursor.continue();
        };
        tx.oncomplete = () => resolve(applyItemQuery(matched, { sort: query.sort, limit: query.limit }));
        tx.onerror = () => reject(tx.error ?? new Error("Could not query local items."));
        tx.onabort = () => reject(tx.error ?? new Error("Local item query was aborted."));
      });
    },
    async putItems(scope, siteId, listId, items) {
      const key = listKey(scope, siteId, listId);
      await transact([stores.items], "readwrite", (tx) => {
        const store = tx.objectStore(stores.items);
        items.forEach((item) => store.put({ key: itemKey(key, item.id), listKey: key, scope, itemId: item.id, item, updatedAt: Date.now() } satisfies StoredItem));
      });
      notify(key);
    },
    async deleteItems(scope, siteId, listId, ids) {
      const key = listKey(scope, siteId, listId);
      await transact([stores.items], "readwrite", (tx) => {
        const store = tx.objectStore(stores.items);
        ids.forEach((id) => store.delete(itemKey(key, id)));
      });
      notify(key);
    },
    async clearItems(scope, siteId, listId) {
      const key = listKey(scope, siteId, listId);
      const db = await open();
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction(stores.items, "readwrite");
        const request = tx.objectStore(stores.items).index("listKey").openKeyCursor(key);
        request.onsuccess = () => {
          const cursor = request.result;
          if (cursor) {
            tx.objectStore(stores.items).delete(cursor.primaryKey);
            cursor.continue();
          }
        };
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error ?? new Error("Could not clear local items."));
        tx.onabort = () => reject(tx.error ?? new Error("Clear operation was aborted."));
      });
      notify(key);
    },
    async clearScope(scope) {
      const db = await open();
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction([stores.responses, stores.items, stores.cursors], "readwrite");
        const responseStore = tx.objectStore(stores.responses);
        const responseCursor = responseStore.openKeyCursor();
        responseCursor.onsuccess = () => {
          const cursor = responseCursor.result;
          if (cursor) {
            try {
              if (JSON.parse(String(cursor.key))[0] === scope) cursor.delete();
            } catch { /* Ignore pre-scope cache entries. */ }
            cursor.continue();
          }
        };
        const itemStore = tx.objectStore(stores.items);
        const itemCursor = itemStore.index("scope").openKeyCursor(scope);
        itemCursor.onsuccess = () => {
          const cursor = itemCursor.result;
          if (cursor) { itemStore.delete(cursor.primaryKey); cursor.continue(); }
        };
        const cursorStore = tx.objectStore(stores.cursors);
        const cursorRequest = cursorStore.openCursor();
        cursorRequest.onsuccess = () => {
          const cursor = cursorRequest.result;
          if (cursor) {
            try {
              if (JSON.parse(String(cursor.key))[0] === scope) cursor.delete();
            } catch { /* Ignore keys from older cache versions. */ }
            cursor.continue();
          }
        };
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error ?? new Error("Could not clear cached user data."));
        tx.onabort = () => reject(tx.error ?? new Error("Cache clear was aborted."));
      });
      notifyScopeSubscribers(scope);
      channel?.postMessage({ source: id, scope });
    },
    async getCursor(scope, siteId, listId) {
      const db = await open();
      const key = listKey(scope, siteId, listId);
      return new Promise((resolve, reject) => {
        const request = db.transaction(stores.cursors, "readonly").objectStore(stores.cursors).get(key);
        request.onsuccess = () => resolve(request.result as string | undefined);
        request.onerror = () => reject(request.error ?? new Error("Could not read delta cursor."));
      });
    },
    async setCursor(scope, siteId, listId, cursor) {
      const key = listKey(scope, siteId, listId);
      await transact([stores.cursors], "readwrite", (tx) => tx.objectStore(stores.cursors).put(cursor, key));
    },
    async clearCursor(scope, siteId, listId) {
      const key = listKey(scope, siteId, listId);
      await transact([stores.cursors], "readwrite", (tx) => tx.objectStore(stores.cursors).delete(key));
    },
    async invalidateListPages(scope, siteId, listId) {
      const prefix = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`;
      const db = await open();
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction(stores.responses, "readwrite");
        const objectStore = tx.objectStore(stores.responses);
        const request = objectStore.openKeyCursor();
        request.onsuccess = () => {
          const cursor = request.result;
          if (cursor) {
            try {
              const [cachedScope, url] = JSON.parse(String(cursor.key)) as [string, string];
              if (cachedScope === scope && url.startsWith(prefix)) cursor.delete();
            } catch { /* Ignore non-standard keys. */ }
            cursor.continue();
          }
        };
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error ?? new Error("Could not invalidate list pages."));
        tx.onabort = () => reject(tx.error ?? new Error("Page invalidation was aborted."));
      });
    },
    async invalidateDrivePages(scope, driveId) {
      const prefix = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/`;
      const db = await open();
      await new Promise<void>((resolve, reject) => {
        const tx = db.transaction(stores.responses, "readwrite");
        const request = tx.objectStore(stores.responses).openKeyCursor();
        request.onsuccess = () => {
          const cursor = request.result;
          if (cursor) {
            try {
              const [cachedScope, url] = JSON.parse(String(cursor.key)) as [string, string];
              if (cachedScope === scope && url.startsWith(prefix)) cursor.delete();
            } catch { /* Ignore non-standard cache keys. */ }
            cursor.continue();
          }
        };
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error ?? new Error("Could not invalidate drive pages."));
        tx.onabort = () => reject(tx.error ?? new Error("Drive page invalidation was aborted."));
      });
    },
    subscribe(scope, siteId, listId, listener) {
      const key = listKey(scope, siteId, listId);
      const bucket = subscribers.get(key) ?? new Set<() => void>();
      bucket.add(listener);
      subscribers.set(key, bucket);
      return () => {
        bucket.delete(listener);
        if (bucket.size === 0) subscribers.delete(key);
      };
    },
  };
}

/** Volatile store used when IndexedDB is unavailable or fails during a session. */
export function createMemoryStore(): LocalStore {
  const responses = new Map<string, unknown>();
  const items = new Map<string, Map<string, GraphListItem>>();
  const cursors = new Map<string, string>();
  const listeners = new Map<string, Set<() => void>>();
  const notify = (key: string) => listeners.get(key)?.forEach((listener) => listener());

  return {
    async get<T>(key: string) { return responses.get(key) as T | undefined; },
    async set<T>(key: string, value: T) { responses.set(key, value); },
    async getItems(scope, siteId, listId) { return [...(items.get(listKey(scope, siteId, listId))?.values() ?? [])]; },
    async queryItems(scope, siteId, listId, query) { return applyItemQuery(await this.getItems(scope, siteId, listId), query); },
    async putItems(scope, siteId, listId, values) {
      const key = listKey(scope, siteId, listId);
      const list = items.get(key) ?? new Map<string, GraphListItem>();
      values.forEach((item) => list.set(item.id, item));
      items.set(key, list);
      notify(key);
    },
    async deleteItems(scope, siteId, listId, ids) {
      const key = listKey(scope, siteId, listId);
      const list = items.get(key);
      ids.forEach((id) => list?.delete(id));
      notify(key);
    },
    async clearItems(scope, siteId, listId) {
      const key = listKey(scope, siteId, listId);
      items.delete(key);
      notify(key);
    },
    async clearScope(scope) {
      for (const key of responses.keys()) {
        try { if (JSON.parse(key)[0] === scope) responses.delete(key); } catch { /* Ignore non-scoped response keys. */ }
      }
      for (const key of [...items.keys()]) if (JSON.parse(key)[0] === scope) { items.delete(key); notify(key); }
      for (const key of cursors.keys()) if (JSON.parse(key)[0] === scope) cursors.delete(key);
      for (const key of listeners.keys()) if (JSON.parse(key)[0] === scope) notify(key);
    },
    async getCursor(scope, siteId, listId) { return cursors.get(listKey(scope, siteId, listId)); },
    async setCursor(scope, siteId, listId, cursor) { cursors.set(listKey(scope, siteId, listId), cursor); },
    async clearCursor(scope, siteId, listId) { cursors.delete(listKey(scope, siteId, listId)); },
    async invalidateListPages(scope, siteId, listId) {
      const prefix = `/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`;
      for (const key of responses.keys()) {
        try {
          const [cachedScope, url] = JSON.parse(key) as [string, string];
          if (cachedScope === scope && new URL(url).pathname.startsWith(prefix)) responses.delete(key);
        } catch { /* Ignore non-standard response keys. */ }
      }
    },
    async invalidateDrivePages(scope, driveId) {
      const prefix = `/v1.0/drives/${encodeURIComponent(driveId)}/`;
      for (const key of responses.keys()) {
        try {
          const [cachedScope, url] = JSON.parse(key) as [string, string];
          if (cachedScope === scope && new URL(url).pathname.startsWith(prefix)) responses.delete(key);
        } catch { /* Ignore non-standard cache keys. */ }
      }
    },
    subscribe(scope, siteId, listId, listener) {
      const key = listKey(scope, siteId, listId);
      const bucket = listeners.get(key) ?? new Set<() => void>();
      bucket.add(listener);
      listeners.set(key, bucket);
      return () => { bucket.delete(listener); if (!bucket.size) listeners.delete(key); };
    },
  };
}

/** Switches permanently to volatile memory after the first persistent-store failure. */
export function createResilientLocalStore(primary: LocalStore, fallback: LocalStore, onPrimaryError?: (error: unknown) => void): LocalStore {
  let useFallback = false;
  return new Proxy(primary, {
    get(target, property) {
      if (property === "subscribe") {
        return (scope: string, siteId: string, listId: string, listener: () => void) => {
          const unsubscribePrimary = target.subscribe(scope, siteId, listId, listener);
          const unsubscribeFallback = fallback.subscribe(scope, siteId, listId, listener);
          return () => { unsubscribePrimary(); unsubscribeFallback(); };
        };
      }
      const primaryMethod = Reflect.get(target, property) as unknown;
      if (typeof primaryMethod !== "function") return primaryMethod;
      return async (...args: unknown[]) => {
        if (useFallback) return Reflect.apply(Reflect.get(fallback, property), fallback, args);
        try {
          return await Reflect.apply(primaryMethod, target, args);
        } catch (error) {
          useFallback = true;
          try { onPrimaryError?.(error); } catch { /* Error reporting must not block fallback. */ }
          return Reflect.apply(Reflect.get(fallback, property), fallback, args);
        }
      };
    },
  }) as LocalStore;
}

function notifyScopeSubscribers(scope: string): void {
  for (const [key, listeners] of subscribers) {
    try {
      if (JSON.parse(key)[0] === scope) listeners.forEach((listener) => listener());
    } catch { /* Ignore cache keys from prior versions. */ }
  }
}

export async function getCacheStorageStatus(): Promise<{ usage?: number; quota?: number; persistent?: boolean }> {
  if (typeof navigator === "undefined" || !navigator.storage) return {};
  const [estimate, persistent] = await Promise.all([
    navigator.storage.estimate().catch(() => undefined),
    navigator.storage.persisted().catch(() => false),
  ]);
  return { usage: estimate?.usage, quota: estimate?.quota, persistent };
}

/** Ask the browser to preserve this origin's cache; approval is browser-controlled. */
export async function requestPersistentCacheStorage(): Promise<boolean> {
  if (typeof navigator === "undefined" || !navigator.storage?.persist) return false;
  return navigator.storage.persist().catch(() => false);
}

function assertGraphUrl(url: string): void {
  // Graph request URLs are absolute. Avoid using an opaque host origin ("null")
  // as a URL base in workbenches, tests, and embedded documents.
  const parsed = new URL(url);
  if (parsed.protocol !== "https:" || parsed.hostname !== "graph.microsoft.com") {
    throw new TypeError("Only HTTPS Microsoft Graph URLs are accepted by this reader.");
  }
}

function itemValue(item: GraphListItem, field: string): unknown {
  return item.fields?.[field] ?? item[field];
}

function matchesItemQuery(item: GraphListItem, query: ItemQuery): boolean {
  if (query.equals && !Object.entries(query.equals).every(([field, value]) => itemValue(item, field) === value)) return false;
  if (query.contains) {
    const needle = query.contains.value.toLocaleLowerCase();
    return String(itemValue(item, query.contains.field) ?? "").toLocaleLowerCase().includes(needle);
  }
  return true;
}

export function applyItemQuery(items: GraphListItem[], query: ItemQuery = {}): GraphListItem[] {
  if (query.limit !== undefined && query.limit <= 0) return [];
  if (query.sort && Number.isInteger(query.limit) && query.limit! <= 50) {
    const { field, direction = "asc" } = query.sort;
    const compare = (a: GraphListItem, b: GraphListItem) => {
      const left = itemValue(a, field);
      const right = itemValue(b, field);
      const order = typeof left === "number" && typeof right === "number"
        ? left - right
        : String(left ?? "").localeCompare(String(right ?? ""));
      return direction === "desc" ? -order : order;
    };
    // Keep only the smallest bounded result instead of sorting every match.
    const top: GraphListItem[] = [];
    for (const item of items) {
      if (!matchesItemQuery(item, query)) continue;
      let low = 0;
      let high = top.length;
      while (low < high) {
        const middle = (low + high) >>> 1;
        if (compare(item, top[middle]) < 0) high = middle;
        else low = middle + 1;
      }
      if (low < query.limit!) {
        top.splice(low, 0, item);
        if (top.length > query.limit!) top.pop();
      }
    }
    return top;
  }

  let result = query.equals || query.contains ? items.filter((item) => matchesItemQuery(item, query)) : items;
  if (query.sort) {
    const { field, direction = "asc" } = query.sort;
    const compare = (a: GraphListItem, b: GraphListItem) => {
      const left = itemValue(a, field);
      const right = itemValue(b, field);
      const order = typeof left === "number" && typeof right === "number"
        ? left - right
        : String(left ?? "").localeCompare(String(right ?? ""));
      return direction === "desc" ? -order : order;
    };
    result = [...result].sort(compare);
  }
  return query.limit === undefined ? result : result.slice(0, Math.max(0, query.limit));
}

function cloneListItems(items: GraphListItem[]): GraphListItem[] {
  if (typeof globalThis.structuredClone === "function") return globalThis.structuredClone(items);
  return items.map((item) => ({ ...item, ...(item.fields ? { fields: { ...item.fields } } : {}) }));
}

function isCacheableItemQuery(query: ItemQuery): boolean {
  return Number.isInteger(query.limit) && query.limit! > 0 && query.limit! <= 100 &&
    (!query.equals || Object.values(query.equals).every((value) => value === null || ["string", "number", "boolean"].includes(typeof value)));
}

function errorStatus(error: unknown): number | undefined {
  if (!error || typeof error !== "object") return undefined;
  const value = error as { status?: number; statusCode?: number; code?: number };
  return value.status ?? value.statusCode ?? value.code;
}

function errorHeader(error: unknown, name: string): string | undefined {
  if (!error || typeof error !== "object") return undefined;
  const headers = (error as { headers?: { get?: (header: string) => string | null } }).headers;
  if (typeof headers?.get !== "function") return undefined;
  try { return headers.get(name) ?? undefined; }
  catch { return undefined; }
}

export function createClient(options: ClientOptions) {
  if (!options.scope.trim()) throw new TypeError("A tenant and signed-in-user cache scope is required.");

  const store = options.store ?? createResilientLocalStore(
    createIndexedDbStore(),
    createMemoryStore(),
    (error) => reportStorageError(options, error),
  );
  const inFlight = new Map<string, Promise<GraphCollection<unknown>>>();
  const cachedValueInFlight = new Map<string, Promise<unknown>>();
  const localQueryCache = new Map<string, { listKey: string; expiresAt: number; items: GraphListItem[] }>();
  const localQueryCacheSubscriptions = new Map<string, () => void>();
  const localQueryCacheVersions = new Map<string, number>();
  const pending = new Set<string>();
  const metric = (event: ClientMetric) => {
    try { options.onMetric?.(event); } catch { /* Observability must not interrupt data access. */ }
  };

  function releaseQueryCacheSubscription(scopedListKey: string): void {
    if ([...localQueryCache.values()].some((entry) => entry.listKey === scopedListKey)) return;
    localQueryCacheSubscriptions.get(scopedListKey)?.();
    localQueryCacheSubscriptions.delete(scopedListKey);
  }

  function invalidateLocalQueryCache(scopedListKey: string): void {
    localQueryCacheVersions.set(scopedListKey, (localQueryCacheVersions.get(scopedListKey) ?? 0) + 1);
    for (const [key, entry] of localQueryCache) if (entry.listKey === scopedListKey) localQueryCache.delete(key);
    releaseQueryCacheSubscription(scopedListKey);
  }

  function ensureQueryCacheSubscription(scopedListKey: string, siteId: string, listId: string): void {
    if (localQueryCacheSubscriptions.has(scopedListKey)) return;
    localQueryCacheSubscriptions.set(scopedListKey, store.subscribe(options.scope, siteId, listId, () => invalidateLocalQueryCache(scopedListKey)));
  }

  async function fetchPage<T>(url: string, signal?: AbortSignal): Promise<GraphCollection<T>> {
    assertGraphUrl(url);
    const key = responseKey(options.scope, url);
    let request = signal ? undefined : inFlight.get(key);
    if (!request) {
      const started = performance.now();
      const network = options.get(url, signal).then((response) => {
        const data = response as GraphCollection<unknown>;
        if (!data || !Array.isArray(data.value)) throw new TypeError("Graph returned an invalid collection response.");
        void store.set(key, data).catch((error: unknown) => reportStorageError(options, error));
        metric({ type: "network-success", operation: "graph.get", durationMs: performance.now() - started, count: data.value.length });
        return data;
      }).catch((error: unknown) => {
        metric({ type: "network-error", operation: "graph.get", durationMs: performance.now() - started, error });
        throw error;
      });
      request = signal ? network : network.finally(() => inFlight.delete(key));
      if (!signal) inFlight.set(key, request);
    }
    return request as Promise<GraphCollection<T>>;
  }

  async function readPage<T>(url: string, signal?: AbortSignal): Promise<ReadPage<T>> {
    assertGraphUrl(url);
    const key = responseKey(options.scope, url);
    let cached: GraphCollection<T> | undefined;
    try {
      cached = await store.get<GraphCollection<T>>(key);
    } catch (error) {
      reportStorageError(options, error);
    }
    if (cached) {
      metric({ type: "cache-hit", operation: "graph.page", count: cached.value.length });
      return { data: cached, source: "cache", refresh: fetchPage<T>(url, signal) };
    }
    metric({ type: "cache-miss", operation: "graph.page" });
    const data = await fetchPage<T>(url, signal);
    return { data, source: "network", refresh: Promise.resolve(data) };
  }

  async function getResource(url: string): Promise<unknown> {
    assertGraphUrl(url);
    return options.get(url);
  }

  async function readCachedValue<T>(url: string, cacheId: string, fetch: () => Promise<T>, cacheMaxBytes = 1_000_000): Promise<CachedValue<T>> {
    assertGraphUrl(url);
    const key = responseKey(options.scope, `spfx-fast:${cacheId}:${url}`);
    const ttlMs = 5 * 60_000;
    let cached: { data: T; expiresAt: number } | undefined;
    try { cached = await store.get<typeof cached>(key); }
    catch (error) { reportStorageError(options, error); }
    const update = () => {
      let request = cachedValueInFlight.get(key) as Promise<T> | undefined;
      if (!request) {
        request = (async () => {
          const data = await fetch();
          const size = data instanceof Blob ? data.size : typeof data === "string" ? new Blob([data]).size : 0;
          if (size <= cacheMaxBytes) {
            try { await store.set(key, { data, expiresAt: Date.now() + ttlMs }); }
            catch (error) { reportStorageError(options, error); }
          }
          return data;
        })();
        cachedValueInFlight.set(key, request);
        void request.finally(() => cachedValueInFlight.delete(key)).catch(() => undefined);
      }
      return request;
    };
    if (cached && cached.expiresAt > Date.now()) {
      metric({ type: "cache-hit", operation: "graph.file", count: 1 });
      return { data: cached.data, source: "cache", refresh: update() };
    }
    metric({ type: "cache-miss", operation: "graph.file" });
    const data = await update();
    return { data, source: "network", refresh: Promise.resolve(data) };
  }

  function getDriveItem(driveId: string, itemId: string, config: { cacheMaxBytes?: number } = {}) {
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}?$select=id,name,size,file,folder,webUrl,eTag,lastModifiedDateTime,parentReference`;
    return readCachedValue(url, "drive-item", () => getResource(url), config.cacheMaxBytes);
  }

  async function getDriveItemPage(driveId: string, itemId: string, config: { cursor?: string; pageSize?: number } = {}) {
    const pageSize = Math.min(Math.max(config.pageSize ?? 50, 1), 100);
    const path = `/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/children`;
    let target = config.cursor;
    if (!target) {
      const url = new URL(path, "https://graph.microsoft.com");
      url.searchParams.set("$top", String(pageSize));
      url.searchParams.set("$select", "id,name,size,file,folder,webUrl,eTag,lastModifiedDateTime,parentReference");
      target = url.toString();
    } else {
      const cursorUrl = new URL(target);
      if (cursorUrl.origin !== "https://graph.microsoft.com" || cursorUrl.pathname !== path) throw new TypeError("Continuation cursors must belong to the requested drive children path.");
    }
    const page = await readPage<GraphDriveItem>(target);
    return { items: page.data.value, nextCursor: page.data["@odata.nextLink"], source: page.source, refresh: page.refresh };
  }

  async function getDriveItemPermissions(driveId: string, itemId: string, config: { maxPages?: number } = {}): Promise<GraphDriveItemPermission[]> {
    const path = `/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/permissions`;
    const maxPages = config.maxPages ?? 20;
    if (!Number.isInteger(maxPages) || maxPages < 1 || maxPages > 100) throw new RangeError("maxPages must be an integer from 1 to 100.");
    const result: GraphDriveItemPermission[] = [];
    let next: string | undefined = new URL(path, "https://graph.microsoft.com").toString();
    let pages = 0;
    while (next) {
      if (++pages > maxPages) throw new Error(`Drive item permissions exceeded ${maxPages} pages.`);
      const target = new URL(next);
      if (target.origin !== "https://graph.microsoft.com" || target.pathname !== path) throw new TypeError("Continuation cursors must belong to the requested drive item permissions path.");
      const page = await getResource(target.toString()) as GraphCollection<GraphDriveItemPermission>;
      if (!page || !Array.isArray(page.value)) throw new TypeError("Graph permissions response must contain a value array.");
      result.push(...page.value);
      next = page["@odata.nextLink"];
    }
    return result;
  }

  /** Create a link with an explicit audience; inherited permissions are kept. Confirm sharing in the host UI first. */
  async function createDriveItemSharingLink(driveId: string, itemId: string, config: DriveItemSharingLinkOptions): Promise<GraphDriveItemPermission> {
    if (!config || !["view", "edit"].includes(config.type)) throw new TypeError("Sharing link type must be view or edit.");
    if (!["anonymous", "organization", "users"].includes(config.scope)) throw new TypeError("Sharing link scope must be explicit: anonymous, organization, or users.");
    if (config.expirationDateTime !== undefined && (!Number.isFinite(Date.parse(config.expirationDateTime)) || !/^\d{4}-\d\d-\d\dT/.test(config.expirationDateTime))) {
      throw new TypeError("Sharing link expiration must be an ISO date-time.");
    }
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/createLink`;
    const permission = await requestGraph(url, {
      method: "POST",
      body: { type: config.type, scope: config.scope, retainInheritedPermissions: true, ...(config.expirationDateTime ? { expirationDateTime: config.expirationDateTime } : {}) },
      signal: config.signal,
    }) as GraphDriveItemPermission;
    if (!permission || typeof permission.id !== "string" || typeof permission.link?.webUrl !== "string") throw new TypeError("Graph sharing link response did not contain a permission ID and URL.");
    metric({ type: "mutation", operation: "drive-create-sharing-link", count: 1 });
    return permission;
  }

  /** Remove a visible, non-inherited permission after the user confirms the access change. */
  async function deleteDriveItemPermission(driveId: string, itemId: string, permissionId: string, config: { signal?: AbortSignal } = {}): Promise<void> {
    if (!permissionId.trim()) throw new TypeError("Drive item permission ID must not be empty.");
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/permissions/${encodeURIComponent(permissionId)}`;
    const permission = await requestGraph(url, { signal: config.signal }) as GraphDriveItemPermission;
    if (!permission || permission.id !== permissionId) throw new TypeError("Graph permission response did not match the requested permission.");
    if (permission.inheritedFrom !== null) throw new Error("Only a verified non-inherited permission can be removed.");
    await requestGraph(url, { method: "DELETE", signal: config.signal });
    metric({ type: "mutation", operation: "drive-delete-permission", count: 1 });
  }

  /** Start an asynchronous copy. Poll the returned monitorUrl with the same-origin SharePoint REST transport. */
  async function copyDriveItem(driveId: string, itemId: string, config: CopyDriveItemOptions): Promise<{ monitorUrl: string }> {
    if (!config?.targetDriveId || !config.targetFolderId) throw new TypeError("Copy requires a destination drive and folder ID.");
    if (config.name !== undefined && (!config.name.trim() || config.name.includes("/") || config.name.includes("\\"))) throw new TypeError("Copy name must be a non-empty leaf name.");
    if (config.includeAllVersionHistory && config.name) throw new TypeError("Copy version history without a new name, then rename the completed copy.");
    const url = new URL(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/copy`);
    if (config.conflictBehavior) url.searchParams.set("@microsoft.graph.conflictBehavior", config.conflictBehavior);
    const response = await requestGraph(url.toString(), {
      method: "POST",
      body: {
        parentReference: { driveId: config.targetDriveId, id: config.targetFolderId },
        ...(config.name ? { name: config.name } : {}),
        ...(config.includeAllVersionHistory ? { includeAllVersionHistory: true } : {}),
      },
      responseType: "raw",
      signal: config.signal,
    }) as { status?: number; headers?: { get(name: string): string | null } };
    if (response?.status !== 202) throw new Error(`Graph did not accept the copy operation (HTTP ${response?.status ?? "unknown"}).`);
    const location = response.headers?.get("Location");
    if (!location) throw new TypeError("Graph accepted the copy but did not return an operation monitor URL.");
    const monitor = new URL(location);
    if (monitor.protocol !== "https:" || monitor.username || monitor.password || !/_api\/v2\.0\/monitor\//.test(monitor.pathname)) {
      throw new TypeError("Graph returned an invalid copy operation monitor URL.");
    }
    metric({ type: "mutation", operation: "drive-copy-start", count: 1 });
    return { monitorUrl: monitor.toString() };
  }

  function readDriveItemContent(driveId: string, itemId: string, config: { as?: "text" | "blob"; cacheMaxBytes?: number } = {}) {
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`;
    const as = config.as ?? "blob";
    return readCachedValue(url, `drive-content:${as}`, () => requestGraph(url, { responseType: as }) as Promise<Blob | string>, config.cacheMaxBytes);
  }

  /** Fetch current file bytes through the authenticated Graph transport without persisting them. */
  async function downloadDriveItem(driveId: string, itemId: string): Promise<Blob> {
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`;
    const content = await requestGraph(url, { responseType: "blob" });
    if (!(content instanceof Blob)) throw new TypeError("Graph file download did not return a Blob.");
    return content;
  }

  /** Fetch a historical file version without persisting its bytes in the cache. */
  async function downloadDriveItemVersion(driveId: string, itemId: string, versionId: string, config: { signal?: AbortSignal } = {}): Promise<Blob> {
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/versions/${encodeURIComponent(versionId)}/content`;
    const content = await requestGraph(url, { responseType: "blob", signal: config.signal });
    if (!(content instanceof Blob)) throw new TypeError("Graph file version download did not return a Blob.");
    return content;
  }

  /** Restore a file version as a new current version. Confirm the action in the calling UI first. */
  function restoreDriveItemVersion(driveId: string, itemId: string, versionId: string, config: { signal?: AbortSignal } = {}): Promise<unknown> {
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/versions/${encodeURIComponent(versionId)}/restoreVersion`;
    return requestGraph(url, { method: "POST", signal: config.signal });
  }

  async function uploadDriveItem(
    driveId: string,
    parentItemId: string,
    name: string,
    file: Blob,
    config: { chunkSize?: number; conflictBehavior?: "fail" | "replace" | "rename"; signal?: AbortSignal; onProgress?: (uploadedBytes: number, totalBytes: number) => void } = {},
  ): Promise<GraphDriveItem> {
    if (config.signal?.aborted) throw config.signal.reason ?? new DOMException("Aborted", "AbortError");
    if (!name.trim() || name.includes("/") || name.includes("\\")) throw new TypeError("File name must be a non-empty leaf name.");
    if (!(file instanceof Blob)) throw new TypeError("File content must be a Blob or File.");
    const unit = 320 * 1024;
    const chunkSize = config.chunkSize ?? 10 * 1024 * 1024;
    if (!Number.isInteger(chunkSize) || chunkSize < unit || chunkSize >= 60 * 1024 * 1024 || chunkSize % unit !== 0) {
      throw new RangeError("chunkSize must be a multiple of 320 KiB, at least 320 KiB, and less than 60 MiB.");
    }
    const drivePath = `/drives/${encodeURIComponent(driveId)}`;
    const itemPath = `/items/${encodeURIComponent(parentItemId)}:/${encodeURIComponent(name)}:`;
    const conflictBehavior = config.conflictBehavior ?? "fail";
    if (file.size <= 10 * 1024 * 1024) {
      const path = `https://graph.microsoft.com/v1.0${drivePath}${itemPath}/content?@microsoft.graph.conflictBehavior=${encodeURIComponent(conflictBehavior)}`;
      return await requestGraph(path, { method: "PUT", headers: { "Content-Type": file.type || "application/octet-stream" }, rawBody: file, signal: config.signal }) as GraphDriveItem;
    }
    let uploadParentId = parentItemId;
    if (parentItemId === "root") {
      const root = await requestGraph(`https://graph.microsoft.com/v1.0${drivePath}/root`, { signal: config.signal }) as { id?: unknown };
      if (typeof root?.id !== "string" || !root.id) throw new TypeError("Graph did not return the document library root ID.");
      uploadParentId = root.id;
    }
    const uploadItemPath = `/items/${encodeURIComponent(uploadParentId)}:/${encodeURIComponent(name)}:`;
    const cacheUploadedItem = async (result: unknown) => {
      const id = (result as { id?: string } | null)?.id;
      const itemUrl = `https://graph.microsoft.com/v1.0${drivePath}/items/${encodeURIComponent(id ?? "")}?$select=id,name,size,file,folder,webUrl,eTag,lastModifiedDateTime,parentReference`;
      await store.set(responseKey(options.scope, `spfx-fast:drive-item:${itemUrl}`), { data: result, expiresAt: Date.now() + 300_000 }).catch((error: unknown) => reportStorageError(options, error));
    };
    const matchesFileBytes = async (candidate: Blob) => {
      if (candidate.size !== file.size) return false;
      for (let offset = 0; offset < file.size; offset += 1024 * 1024) {
        const end = Math.min(offset + 1024 * 1024, file.size);
        const [expected, actual] = await Promise.all([
          file.slice(offset, end).arrayBuffer(),
          candidate.slice(offset, end).arrayBuffer(),
        ]);
        const a = new Uint8Array(expected);
        const b = new Uint8Array(actual);
        if (a.some((byte, index) => byte !== b[index])) return false;
      }
      return true;
    };
    const verifyDestination = async (): Promise<GraphDriveItem | undefined> => {
      const itemUrl = `https://graph.microsoft.com/v1.0${drivePath}${uploadItemPath}?$select=id,name,size,file,folder,webUrl,eTag,lastModifiedDateTime,parentReference`;
      const item = await requestGraph(itemUrl, { signal: config.signal }) as GraphDriveItem;
      if (typeof item?.id !== "string" || item.name !== name || item.size !== file.size) return undefined;
      const content = await requestGraph(`https://graph.microsoft.com/v1.0${drivePath}/items/${encodeURIComponent(item.id)}/content`, { responseType: "blob", signal: config.signal });
      return content instanceof Blob && await matchesFileBytes(content) ? item : undefined;
    };
    const verifyRenamedDestination = async (): Promise<GraphDriveItem | undefined> => {
      const extensionIndex = name.lastIndexOf(".");
      const stem = extensionIndex > 0 ? name.slice(0, extensionIndex) : name;
      const extension = extensionIndex > 0 ? name.slice(extensionIndex) : "";
      const candidates = Array.from({ length: 5 }, (_, index) => index + 1).flatMap((index) => [`${stem} ${index}${extension}`, `${stem} (${index})${extension}`]);
      for (const delay of [0, 250, 750, 1500]) {
        if (delay) await new Promise((resolve) => setTimeout(resolve, delay));
        for (const candidateName of candidates) {
          let item: GraphDriveItem;
          try {
            item = await requestGraph(`https://graph.microsoft.com/v1.0${drivePath}/items/${encodeURIComponent(uploadParentId)}:/${encodeURIComponent(candidateName)}:?$select=id,name,size,file,folder,parentReference`, { signal: config.signal }) as GraphDriveItem;
          } catch (error) {
            if (errorStatus(error) === 404) continue;
            throw error;
          }
          if (typeof item?.id !== "string" || item.id === existingItemId || item.name !== candidateName || item.size !== file.size || item.folder || (item.parentReference?.id && item.parentReference.id !== uploadParentId)) continue;
          const content = await requestGraph(`https://graph.microsoft.com/v1.0${drivePath}/items/${encodeURIComponent(item.id)}/content`, { responseType: "blob", signal: config.signal });
          if (content instanceof Blob && await matchesFileBytes(content)) return item;
        }
      }
      return undefined;
    };
    const reconcileDestination = () => conflictBehavior === "rename" ? verifyRenamedDestination() : verifyDestination();
    let existingItemId: string | undefined;
    if (conflictBehavior === "replace") {
      try {
        const existing = await requestGraph(`https://graph.microsoft.com/v1.0${drivePath}${uploadItemPath}?$select=id,name,folder`, { signal: config.signal }) as GraphDriveItem;
        if (typeof existing?.id !== "string" || !existing.id) throw new TypeError("Graph did not return the existing destination item ID.");
        if (existing.folder) throw new TypeError("Cannot replace a folder with a file.");
        existingItemId = existing.id;
      } catch (error) {
        if (errorStatus(error) !== 404) throw error;
      }
    }
    const sessionItemPath = existingItemId ? `/items/${encodeURIComponent(existingItemId)}` : uploadItemPath;
    const url = `https://graph.microsoft.com/v1.0${drivePath}${sessionItemPath}/createUploadSession`;
    const sessionOptions: GraphRequestOptions = { method: "POST", signal: config.signal };
    if (conflictBehavior !== "fail" && !existingItemId) {
      sessionOptions.body = { item: { "@microsoft.graph.conflictBehavior": conflictBehavior } };
    }
    const session = await requestGraph(url, sessionOptions) as { uploadUrl?: string; expirationDateTime?: string };
    if (!session?.uploadUrl) throw new TypeError("Graph did not return an upload session URL.");
    const uploadUrl = new URL(session.uploadUrl);
    if (uploadUrl.protocol !== "https:" || uploadUrl.username || uploadUrl.password) throw new TypeError("Graph returned an invalid HTTPS upload session URL.");
    let sessionCancelled = false;
    const cancelSession = () => {
      if (sessionCancelled) return;
      sessionCancelled = true;
      void fetch(uploadUrl, { method: "DELETE", credentials: "omit", redirect: "error" }).catch(() => undefined);
    };
    const getUploadSession = async () => {
      for (let attempt = 0; ; attempt++) {
        try {
          return await fetch(uploadUrl, { credentials: "omit", redirect: "error", signal: config.signal });
        } catch (error) {
          if (config.signal?.aborted || attempt >= 5) throw error;
          await new Promise((resolve) => setTimeout(resolve, Math.min(500 * 2 ** attempt, 8000)));
        }
      }
    };
    config.signal?.addEventListener("abort", cancelSession, { once: true });
    let completed = false;
    let preserveSession = false;
    let resumeInProgress = false;
    let uploadFrom: (startOffset: number) => Promise<GraphDriveItem>;
    let resumeUpload: () => Promise<GraphDriveItem>;
    const interruptedError = (cause?: unknown) => new DriveUploadInterruptedError(resumeUpload, cause);
    uploadFrom = async (startOffset) => {
      let offset = startOffset;
      let retriesAtOffset = 0;
      let reconciledFinalAttempt = false;
      while (offset < file.size) {
        if (sessionCancelled) throw new Error("Upload session was cancelled.");
        if (config.signal?.aborted) throw config.signal.reason ?? new DOMException("Aborted", "AbortError");
        const end = Math.min(offset + chunkSize, file.size);
        let response: Response;
        try {
          response = await fetch(uploadUrl, {
            method: "PUT",
            headers: { "Content-Range": `bytes ${offset}-${end - 1}/${file.size}` },
            body: file.slice(offset, end),
            credentials: "omit",
            redirect: "error",
            signal: config.signal,
          });
        } catch (error) {
          if (config.signal?.aborted) throw error;
          let status: Response;
          try {
            status = await getUploadSession();
          } catch (statusError) {
            if (config.signal?.aborted) throw statusError;
            if (end !== file.size) {
              preserveSession = true;
              throw statusError;
            }
            try {
              const item = await reconcileDestination();
              if (item) {
                await cacheUploadedItem(item);
                completed = true;
                config.onProgress?.(file.size, file.size);
                return item;
              }
            } catch (reconcileError) {
              if (config.signal?.aborted) throw reconcileError;
            }
            preserveSession = true;
            throw Object.assign(new Error("Upload may have completed, but its final response and session status could not be confirmed."), { cause: statusError });
          }
          if (!status.ok && end === file.size && !reconciledFinalAttempt && (status.status === 404 || status.status === 410)) {
            reconciledFinalAttempt = true;
            try {
              const item = await reconcileDestination();
              if (item) {
                await cacheUploadedItem(item);
                completed = true;
                config.onProgress?.(file.size, file.size);
                return item;
              }
            } catch (reconcileError) {
              if (config.signal?.aborted) throw reconcileError;
            }
            throw Object.assign(new Error("Upload may have completed, but its final response was lost and the session expired. Refresh the destination folder and inspect it before retrying."), { cause: error });
          }
          if (!status.ok) {
            retriesAtOffset++;
            if (retriesAtOffset > 4) {
              if (status.status >= 500) preserveSession = true;
              throw error;
            }
            await new Promise((resolve) => setTimeout(resolve, Math.min(500 * 2 ** (retriesAtOffset - 1), 8000)));
            continue;
          }
          const state = await status.json() as { nextExpectedRanges?: string[] };
          const next = Number(state.nextExpectedRanges?.[0]?.split("-")[0]);
          if (!Number.isInteger(next) || next < 0 || next > file.size) throw error;
          retriesAtOffset = next === offset ? retriesAtOffset + 1 : 0;
          if (retriesAtOffset > 4) throw error;
          offset = next;
          continue;
        }
        if (response.status === 202) {
          const state = await response.json() as { nextExpectedRanges?: string[] };
          const next = Number(state.nextExpectedRanges?.[0]?.split("-")[0] ?? end);
          if (!Number.isInteger(next) || next <= offset || next > file.size) throw new Error("Upload session returned an invalid nextExpectedRanges value.");
          offset = next;
          retriesAtOffset = 0;
          config.onProgress?.(offset, file.size);
          continue;
        }
        if (response.status === 200 || response.status === 201) {
          config.onProgress?.(file.size, file.size);
          const result = await response.json() as GraphDriveItem;
          await cacheUploadedItem(result);
          completed = true;
          return result;
        }
        if (response.status === 416 || response.status >= 500) {
          let status: Response;
          try {
            status = await getUploadSession();
          } catch (statusError) {
            preserveSession = true;
            throw statusError;
          }
          if (status.ok) {
            const state = await status.json() as { nextExpectedRanges?: string[] };
            const next = Number(state.nextExpectedRanges?.[0]?.split("-")[0]);
            if (Number.isInteger(next) && next >= 0 && next <= file.size) {
              retriesAtOffset = next === offset ? retriesAtOffset + 1 : 0;
              if (retriesAtOffset > 4) throw new Error("Upload session repeatedly rejected the current byte range.");
              offset = next;
              continue;
            }
          }
          retriesAtOffset++;
          if (retriesAtOffset <= 4) {
            await new Promise((resolve) => setTimeout(resolve, Math.min(500 * 2 ** (retriesAtOffset - 1), 8000)));
            continue;
          }
          if (status.status >= 500) preserveSession = true;
          throw new Error(`Upload session status check failed with HTTP ${status.status}.`);
        }
        throw Object.assign(new Error(`Upload session failed with HTTP ${response.status}.`), { status: response.status });
      }
      throw new Error("Upload session ended without returning the completed drive item.");
    };
    resumeUpload = async () => {
      if (resumeInProgress) throw new Error("Upload resume is already running.");
      if (completed) throw new Error("Upload has already completed.");
      if (sessionCancelled) throw new Error("Upload session was cancelled or expired.");
      if (config.signal?.aborted) throw config.signal.reason ?? new DOMException("Aborted", "AbortError");
      resumeInProgress = true;
      try {
        let status: Response;
        try {
          status = await getUploadSession();
        } catch (error) {
          preserveSession = true;
          throw interruptedError(error);
        }
        if (!status.ok) {
          if (status.status === 404 || status.status === 410) {
            try {
              const item = await reconcileDestination();
              if (item) {
                await cacheUploadedItem(item);
                completed = true;
                config.onProgress?.(file.size, file.size);
                return item;
              }
            } catch (reconcileError) {
              if (config.signal?.aborted) throw reconcileError;
            }
          }
          preserveSession = status.status >= 500;
          if (preserveSession) throw interruptedError();
          config.signal?.removeEventListener("abort", cancelSession);
          throw new Error(`Upload session is no longer available (HTTP ${status.status}). Inspect the destination before retrying.`);
        }
        let next: number;
        try {
          const state = await status.json() as { nextExpectedRanges?: string[] };
          next = Number(state.nextExpectedRanges?.[0]?.split("-")[0]);
        } catch (error) {
          preserveSession = true;
          throw interruptedError(error);
        }
        if (!Number.isInteger(next) || next < 0 || next > file.size) {
          preserveSession = true;
          throw interruptedError(new Error("Graph returned an invalid resume offset."));
        }
        preserveSession = false;
        return await uploadFrom(next);
      } catch (error) {
        if (!completed && preserveSession && !(error instanceof DriveUploadInterruptedError)) throw interruptedError(error);
        if (!completed && !preserveSession) cancelSession();
        throw error;
      } finally {
        resumeInProgress = false;
        if (completed || !preserveSession) config.signal?.removeEventListener("abort", cancelSession);
      }
    };
    try {
      return await uploadFrom(0);
    } catch (error) {
      if (!completed && preserveSession) throw interruptedError(error);
      if (!completed) cancelSession();
      throw error;
    } finally {
      if (completed || !preserveSession) config.signal?.removeEventListener("abort", cancelSession);
    }
  }

  function listItems<T extends GraphListItem>(
    siteId: string,
    listId: string,
    config: {
      select?: string[];
      fields?: string[];
      pageSize?: number;
      maxItems?: number;
      filter?: GraphListItemFilter;
      signal?: AbortSignal;
      onBackgroundError?: (error: unknown) => void;
    } = {},
  ): AsyncIterable<T> {
    const pageSize = Math.min(Math.max(config.pageSize ?? 100, 1), 200);
    const maxItems = Math.max(config.maxItems ?? 500, 0);
    const url = new URL(`/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`, "https://graph.microsoft.com");
    url.searchParams.set("$top", String(pageSize));
    url.searchParams.set("$select", config.select?.length ? config.select.join(",") : "id,eTag,fields");
    url.searchParams.set("$expand", config.fields?.length ? `fields($select=${config.fields.join(",")})` : "fields");
    const filter = graphListItemFilter(config.filter);
    if (filter) url.searchParams.set("$filter", filter);

    return {
      async *[Symbol.asyncIterator]() {
        let next: string | undefined = url.toString();
        let count = 0;
        while (next && count < maxItems) {
          if (config.signal?.aborted) throw config.signal.reason ?? new DOMException("Aborted", "AbortError");
          const page: ReadPage<T> = await readPage<T>(next, config.signal);
          if (page.source === "cache") void page.refresh.catch((error: unknown) => config.onBackgroundError?.(error));
          const data: GraphCollection<T> = page.data;
          for (const item of data.value) {
            if (count++ >= maxItems) return;
            yield item;
          }
          next = data["@odata.nextLink"];
        }
      },
    };
  }

  async function getItemPage<T extends GraphListItem>(
    siteId: string,
    listId: string,
    config: { cursor?: string; fields?: string[]; pageSize?: number; filter?: GraphListItemFilter; signal?: AbortSignal; onBackgroundError?: (error: unknown) => void } = {},
  ) {
    const pageSize = Math.min(Math.max(config.pageSize ?? 50, 1), 100);
    const itemPath = `/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`;
    const filter = graphListItemFilter(config.filter);
    let target = config.cursor;
    if (!target) {
      const url = new URL(itemPath, "https://graph.microsoft.com");
      url.searchParams.set("$top", String(pageSize));
      url.searchParams.set("$select", "id,eTag,fields");
      url.searchParams.set("$expand", config.fields?.length ? `fields($select=${config.fields.join(",")})` : "fields");
      if (filter) url.searchParams.set("$filter", filter);
      target = url.toString();
    } else {
      const cursorUrl = new URL(target);
      if (cursorUrl.origin !== "https://graph.microsoft.com" || cursorUrl.pathname !== itemPath) {
        throw new TypeError("Continuation cursors must belong to the requested Microsoft Graph list items path.");
      }
      if ((cursorUrl.searchParams.get("$filter") ?? undefined) !== filter) throw new TypeError("Continuation cursors must use the same Graph filter as the initial page.");
    }
    const page = await readPage<T>(target, config.signal);
    if (page.source === "cache") void page.refresh.catch((error: unknown) => config.onBackgroundError?.(error));
    return { items: page.data.value, nextCursor: page.data["@odata.nextLink"], source: page.source, refresh: page.refresh };
  }

  /** Query synchronized local rows; exact queries with limits up to 100 cache briefly and clear on list updates. */
  async function queryItems(siteId: string, listId: string, query: ItemQuery = {}): Promise<GraphListItem[]> {
    const started = performance.now();
    let queryCacheKey: string | undefined;
    const scopedListKey = listKey(options.scope, siteId, listId);
    let cacheVersion: number | undefined;
    if (isCacheableItemQuery(query)) {
      queryCacheKey = JSON.stringify([scopedListKey, query]);
      const cached = localQueryCache.get(queryCacheKey);
      if (cached && cached.expiresAt > performance.now()) {
        localQueryCache.delete(queryCacheKey);
        localQueryCache.set(queryCacheKey, cached);
        const result = cloneListItems(cached.items);
        metric({ type: "cache-hit", operation: "items.query", durationMs: performance.now() - started, count: result.length });
        return result;
      }
      if (cached) {
        localQueryCache.delete(queryCacheKey);
        releaseQueryCacheSubscription(scopedListKey);
      }
      ensureQueryCacheSubscription(scopedListKey, siteId, listId);
      cacheVersion = localQueryCacheVersions.get(scopedListKey) ?? 0;
    }
    let items: GraphListItem[];
    try {
      items = store.queryItems
        ? await store.queryItems(options.scope, siteId, listId, query)
        : await store.getItems(options.scope, siteId, listId);
    } catch (error) {
      reportStorageError(options, error);
      metric({ type: "local-query", operation: "indexeddb.items", durationMs: performance.now() - started, error });
      if (queryCacheKey) releaseQueryCacheSubscription(scopedListKey);
      return [];
    }
    const result = store.queryItems ? items : applyItemQuery(items, query);
    if (queryCacheKey && cacheVersion === (localQueryCacheVersions.get(scopedListKey) ?? 0)) {
      localQueryCache.set(queryCacheKey, { listKey: scopedListKey, expiresAt: performance.now() + 1000, items: cloneListItems(result) });
      if (localQueryCache.size > 64) {
        const oldestKey = localQueryCache.keys().next().value;
        if (oldestKey) {
          const oldest = localQueryCache.get(oldestKey)!;
          localQueryCache.delete(oldestKey);
          releaseQueryCacheSubscription(oldest.listKey);
        }
      }
    } else if (queryCacheKey) {
      releaseQueryCacheSubscription(scopedListKey);
    }
    metric({ type: "local-query", operation: "indexeddb.items", durationMs: performance.now() - started, count: result.length });
    return result;
  }

  function subscribeItems(siteId: string, listId: string, listener: () => void): () => void {
    return store.subscribe(options.scope, siteId, listId, listener);
  }

  async function getDriveItems(driveId: string): Promise<GraphDriveItem[]> {
    try { return await store.getItems(options.scope, driveStoreSiteId, driveId) as unknown as GraphDriveItem[]; }
    catch (error) { reportStorageError(options, error); return []; }
  }

  async function invalidateDrivePages(driveId: string): Promise<void> {
    try { await store.invalidateDrivePages?.(options.scope, driveId); }
    catch (error) { reportStorageError(options, error); }
  }

  async function createDriveFolder(
    driveId: string,
    parentItemId: string,
    name: string,
    config: { conflictBehavior?: "fail" | "replace" | "rename"; signal?: AbortSignal } = {},
  ): Promise<GraphDriveItem> {
    if (!name.trim() || name.includes("/") || name.includes("\\")) throw new TypeError("Folder name must be a non-empty leaf name.");
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(parentItemId)}/children`;
    const result = await requestGraph(url, {
      method: "POST",
      body: { name, folder: {}, "@microsoft.graph.conflictBehavior": config.conflictBehavior ?? "fail" },
      signal: config.signal,
    }) as GraphDriveItem;
    if (!result || typeof result.id !== "string") throw new TypeError("Graph folder creation returned an invalid drive item.");
    await store.putItems(options.scope, driveStoreSiteId, driveId, [result as unknown as GraphListItem]).catch((error: unknown) => reportStorageError(options, error));
    await invalidateDrivePages(driveId);
    metric({ type: "mutation", operation: "drive-create-folder", count: 1 });
    return result;
  }

  async function moveDriveItem(
    driveId: string,
    itemId: string,
    parentItemId: string,
    config: { name?: string; eTag?: string; signal?: AbortSignal } = {},
  ): Promise<GraphDriveItem> {
    if (config.name !== undefined && (!config.name.trim() || config.name.includes("/") || config.name.includes("\\"))) {
      throw new TypeError("Item name must be a non-empty leaf name.");
    }
    const cached = (await getDriveItems(driveId)).find((item) => item.id === itemId);
    const eTag = config.eTag ?? cached?.eTag;
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`;
    const result = await requestGraph(url, {
      method: "PATCH",
      headers: eTag ? { "If-Match": eTag } : undefined,
      body: { parentReference: { id: parentItemId }, ...(config.name === undefined ? {} : { name: config.name }) },
      signal: config.signal,
    }) as GraphDriveItem;
    if (!result || typeof result.id !== "string") throw new TypeError("Graph move returned an invalid drive item.");
    await store.putItems(options.scope, driveStoreSiteId, driveId, [result as unknown as GraphListItem]).catch((error: unknown) => reportStorageError(options, error));
    await invalidateDrivePages(driveId);
    metric({ type: "mutation", operation: "drive-move-item", count: 1 });
    return result;
  }

  async function deleteDriveItem(driveId: string, itemId: string, config: { eTag?: string; signal?: AbortSignal } = {}): Promise<void> {
    const items = await getDriveItems(driveId);
    const target = items.find((item) => item.id === itemId);
    const eTag = config.eTag ?? target?.eTag;
    const descendants = new Set([itemId]);
    if (target?.folder) {
      const children = new Map<string, string[]>();
      for (const item of items) {
        const parentId = item.parentReference?.id;
        if (parentId) children.set(parentId, [...(children.get(parentId) ?? []), item.id]);
      }
      const pending = [itemId];
      while (pending.length) for (const childId of children.get(pending.pop()!) ?? []) if (!descendants.has(childId)) {
        descendants.add(childId);
        pending.push(childId);
      }
    }
    const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}`;
    await requestGraph(url, {
      method: "DELETE",
      headers: eTag ? { "If-Match": eTag } : undefined,
      signal: config.signal,
    });
    await store.deleteItems(options.scope, driveStoreSiteId, driveId, [...descendants]).catch((error: unknown) => reportStorageError(options, error));
    await invalidateDrivePages(driveId);
    metric({ type: "mutation", operation: "drive-delete-item", count: descendants.size });
  }

  function subscribeDriveItems(driveId: string, listener: () => void): () => void {
    return store.subscribe(options.scope, driveStoreSiteId, driveId, listener);
  }

  async function syncList(siteId: string, listId: string, config: { fields?: string[]; maxPages?: number; signal?: AbortSignal } = {}) {
    const started = performance.now();
    const cursor = await store.getCursor(options.scope, siteId, listId).catch((error: unknown) => {
      reportStorageError(options, error);
      return undefined;
    });
    const firstUrl = new URL(`/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/delta`, "https://graph.microsoft.com");
    firstUrl.searchParams.set("$select", "id,eTag,fields");
    firstUrl.searchParams.set("$expand", config.fields?.length ? `fields($select=${config.fields.join(",")})` : "fields");
    const normalizedPath = (pathname: string): string | undefined => {
      const segments: string[] = [];
      for (const segment of pathname.split("/")) {
        if (/%(?:2f|5c)/i.test(segment)) return undefined;
        try { segments.push(decodeURIComponent(segment)); } catch { return undefined; }
      }
      return segments.join("/");
    };
    const normalizeListDeltaUrl = (url: string): string | undefined => {
      try {
        const target = new URL(url, firstUrl);
        const actualPath = normalizedPath(target.pathname);
        const expectedPath = normalizedPath(firstUrl.pathname);
        if (target.origin !== firstUrl.origin || !actualPath || !expectedPath || actualPath !== expectedPath || target.username || target.password) return undefined;
        return target.toString();
      } catch { return undefined; }
    };
    const assertListDeltaUrl = (url: string) => {
      if (!normalizeListDeltaUrl(url)) throw new TypeError("List delta continuation must stay on the requested list items delta endpoint.");
    };
    let startUrl = cursor ?? firstUrl.toString();
    let resetAttempted = false;
    let pages = 0;
    let changed = 0;

    while (true) {
      try {
        let next: string | undefined = startUrl;
        let deltaLink: string | undefined;
        pages = 0;
        changed = 0;
        while (next) {
          assertListDeltaUrl(next);
          if (config.signal?.aborted) throw config.signal.reason ?? new DOMException("Aborted", "AbortError");
          if (++pages > (config.maxPages ?? 1000)) throw new Error("Delta sync exceeded maxPages; cursor was not advanced.");
          const page: GraphCollection<GraphListItem> = await fetchPage<GraphListItem>(next, config.signal);
          const latest = new Map<string, GraphListItem>();
          page.value.forEach((item) => latest.set(item.id, item));
          const deltaItems = [...latest.values()];
          const deleted = deltaItems.filter((item) => item.deleted?.state === "deleted").map((item) => item.id);
          const changedItems = deltaItems.filter((item) => item.deleted?.state !== "deleted");
          if (changedItems.length) await store.putItems(options.scope, siteId, listId, changedItems);
          if (deleted.length) await store.deleteItems(options.scope, siteId, listId, deleted);
          changed += page.value.length;
          next = page["@odata.nextLink"];
          deltaLink = page["@odata.deltaLink"];
          if (next) assertListDeltaUrl(next);
        }
        if (!deltaLink) throw new Error("Graph delta sync completed without a deltaLink.");
        assertListDeltaUrl(deltaLink);
        await store.setCursor(options.scope, siteId, listId, deltaLink);
        if (changed) await store.invalidateListPages(options.scope, siteId, listId);
        metric({ type: "delta-sync", operation: "graph.list.delta", durationMs: performance.now() - started, count: changed });
        return { pages, changed, deltaLink };
      } catch (error) {
        if (!resetAttempted && cursor && errorStatus(error) === 410) {
          resetAttempted = true;
          const location = errorHeader(error, "Location");
          await store.clearCursor(options.scope, siteId, listId);
          await store.clearItems(options.scope, siteId, listId);
          await store.invalidateListPages(options.scope, siteId, listId);
          startUrl = (location && normalizeListDeltaUrl(location)) ?? firstUrl.toString();
          continue;
        }
        metric({ type: "network-error", operation: "graph.list.delta", durationMs: performance.now() - started, error });
        throw error;
      }
    }
  }

  async function syncDrive(driveId: string, config: { maxPages?: number; signal?: AbortSignal } = {}): Promise<DriveSyncResult> {
    const started = performance.now();
    const maxPages = config.maxPages ?? 100;
    if (!Number.isInteger(maxPages) || maxPages < 1 || maxPages > 1000) throw new RangeError("maxPages must be an integer from 1 to 1000.");
    const path = `/v1.0/drives/${encodeURIComponent(driveId)}/root/delta`;
    const firstUrl = new URL(path, "https://graph.microsoft.com");
    firstUrl.searchParams.set("$top", "200");
    const normalizeDriveDeltaUrl = (url: string): string | undefined => {
      try {
        const target = new URL(url, firstUrl);
        if (target.origin !== firstUrl.origin || target.pathname !== path || target.username || target.password) return undefined;
        return target.toString();
      } catch { return undefined; }
    };
    const checkpointKey = driveCheckpointKey(options.scope, driveId);
    let cursor = await store.getCursor(options.scope, driveStoreSiteId, driveId).catch((error: unknown) => {
      reportStorageError(options, error);
      return undefined;
    });
    const storedCheckpoint = await store.get<unknown>(checkpointKey).catch((error: unknown) => {
      reportStorageError(options, error);
      return undefined;
    });
    let checkpoint = isDriveDeltaCheckpoint(storedCheckpoint) ? storedCheckpoint : null;
    if (storedCheckpoint && !checkpoint) await store.set<DriveDeltaCheckpoint | null>(checkpointKey, null);
    const knownItems = await getDriveItems(driveId);
    let resetAttempted = false;
    let hasResumeState = Boolean(cursor || checkpoint);
    let next: string | undefined = checkpoint?.nextLink ?? cursor ?? firstUrl.toString();
    let deferredDeletedFolders = new Set(checkpoint?.deletedFolderIds ?? []);
    let current = new Map(knownItems.map((item) => [item.id, item]));
    if (!cursor && !checkpoint) {
      await store.clearItems(options.scope, driveStoreSiteId, driveId);
      current.clear();
    }

    while (true) {
      let pages = 0;
      let changed = 0;
      try {
        while (next) {
          if (config.signal?.aborted) throw config.signal.reason ?? new DOMException("Aborted", "AbortError");
          const target = new URL(next);
          if (target.origin !== "https://graph.microsoft.com" || target.pathname !== path || target.username || target.password) {
            throw new TypeError("Drive delta continuation must stay on the requested drive root delta endpoint.");
          }
          const response = await options.get(target.toString(), config.signal) as GraphCollection<GraphDriveItem>;
          if (!response || !Array.isArray(response.value)) throw new TypeError("Graph drive delta response must contain a value array.");
          const latest = new Map<string, GraphDriveItem>();
          response.value.forEach((item) => latest.set(item.id, item));
          const changedItems: GraphDriveItem[] = [];
          const deletedIds: string[] = [];
          for (const item of latest.values()) {
            if (item.deleted?.state === "deleted") {
              if (current.get(item.id)?.folder) deferredDeletedFolders.add(item.id);
              else deletedIds.push(item.id);
            } else {
              changedItems.push(item);
              deferredDeletedFolders.delete(item.id);
            }
          }
          if (changedItems.length) {
            await store.putItems(options.scope, driveStoreSiteId, driveId, changedItems as unknown as GraphListItem[]);
            changedItems.forEach((item) => current.set(item.id, item));
            await invalidateDrivePages(driveId);
          }
          if (deletedIds.length) {
            await store.deleteItems(options.scope, driveStoreSiteId, driveId, deletedIds);
            deletedIds.forEach((id) => current.delete(id));
            await invalidateDrivePages(driveId);
          }
          changed += response.value.length;
          pages++;
          next = response["@odata.nextLink"];
          const deltaLink = response["@odata.deltaLink"];
          if (next) {
            const continuation = new URL(next);
            if (continuation.origin !== "https://graph.microsoft.com" || continuation.pathname !== path || continuation.username || continuation.password) {
              throw new TypeError("Drive delta continuation must stay on the requested drive root delta endpoint.");
            }
            await store.set<DriveDeltaCheckpoint>(checkpointKey, { nextLink: next, deletedFolderIds: [...deferredDeletedFolders] });
            hasResumeState = true;
            if (pages >= maxPages) {
              metric({ type: "delta-sync", operation: "graph.drive.delta", durationMs: performance.now() - started, count: changed });
              return { pages, changed, complete: false, nextCursor: next };
            }
          } else {
            if (!deltaLink) throw new Error("Graph drive delta sync completed without a deltaLink.");
            const delta = new URL(deltaLink);
            if (delta.origin !== "https://graph.microsoft.com" || delta.pathname !== path || delta.username || delta.password) {
              throw new TypeError("Drive delta link must stay on the requested drive root delta endpoint.");
            }
            const parentsWithChildren = new Set([...current.values()].map((item) => item.parentReference?.id).filter((id): id is string => Boolean(id)));
            const emptyDeletedFolders = [...deferredDeletedFolders].filter((folderId) => !parentsWithChildren.has(folderId));
            if (emptyDeletedFolders.length) await store.deleteItems(options.scope, driveStoreSiteId, driveId, emptyDeletedFolders);
            if (emptyDeletedFolders.length) await invalidateDrivePages(driveId);
            await store.setCursor(options.scope, driveStoreSiteId, driveId, deltaLink);
            await store.set<DriveDeltaCheckpoint | null>(checkpointKey, null);
            metric({ type: "delta-sync", operation: "graph.drive.delta", durationMs: performance.now() - started, count: changed });
            return { pages, changed, complete: true, deltaLink };
          }
        }
      } catch (error) {
        if (!resetAttempted && hasResumeState && errorStatus(error) === 410) {
          resetAttempted = true;
          const location = errorHeader(error, "Location");
          await store.clearCursor(options.scope, driveStoreSiteId, driveId);
          await store.clearItems(options.scope, driveStoreSiteId, driveId);
          await store.set<DriveDeltaCheckpoint | null>(checkpointKey, null);
          cursor = undefined;
          checkpoint = null;
          deferredDeletedFolders.clear();
          current.clear();
          hasResumeState = false;
          next = (location && normalizeDriveDeltaUrl(location)) ?? firstUrl.toString();
          continue;
        }
        metric({ type: "network-error", operation: "graph.drive.delta", error });
        throw error;
      }
    }
  }

  function requireRequest(): GraphRequest {
    if (!options.request) throw new Error("A Graph request transport is required for writes.");
    return options.request;
  }

  function requestGraph(url: string, config: GraphRequestOptions = {}): Promise<unknown> {
    assertGraphUrl(url);
    return requireRequest()(url, config);
  }

  async function batch(requests: GraphBatchRequest[]) {
    if (requests.length < 1 || requests.length > 20) throw new RangeError("Microsoft Graph batches must contain between 1 and 20 requests.");
    const ids = new Set<string>();
    for (const item of requests) {
      if (!item || typeof item.id !== "string" || !item.id.trim()) throw new TypeError("Batch request IDs must be non-empty and unique.");
      const normalizedId = item.id.toLowerCase();
      if (ids.has(normalizedId)) throw new TypeError("Batch request IDs must be non-empty and unique.");
      ids.add(normalizedId);
      if (typeof item.url !== "string") throw new TypeError("Batch URLs must be Graph resource paths such as /sites/{id}/lists; omit the version and $batch segment.");
      const target = new URL(item.url, "https://graph.microsoft.com");
      if (!item.url.startsWith("/") || item.url.startsWith("//") || target.origin !== "https://graph.microsoft.com" || target.hash || /^\/(?:v1\.0|beta)(?:\/|$)/i.test(target.pathname) || /\/$batch$/i.test(target.pathname)) {
        throw new TypeError("Batch URLs must be Graph resource paths such as /sites/{id}/lists; omit the version and $batch segment.");
      }
    }
    const batchRequests = requests.map((item) => {
      const hasBody = item.body !== undefined && ["POST", "PUT", "PATCH"].includes(item.method);
      const hasContentType = Object.keys(item.headers ?? {}).some((name) => name.toLowerCase() === "content-type");
      return hasBody && !hasContentType
        ? { ...item, headers: { ...item.headers, "Content-Type": "application/json" } }
        : item;
    });
    const started = performance.now();
    try {
      const response = await requireRequest()("https://graph.microsoft.com/v1.0/$batch", { method: "POST", body: { requests: batchRequests } }) as { responses?: unknown };
      if (!response || !Array.isArray(response.responses)) throw new TypeError("Graph returned an invalid batch response.");
      metric({ type: "network-success", operation: "graph.batch", durationMs: performance.now() - started, count: requests.length });
      return response.responses as Array<{ id: string; status: number; headers?: Record<string, string>; body?: unknown }>;
    } catch (error) {
      metric({ type: "network-error", operation: "graph.batch", durationMs: performance.now() - started, error });
      throw error;
    }
  }

  async function createItem(siteId: string, listId: string, fields: Record<string, unknown>): Promise<GraphListItem> {
    const tempId = `local-${crypto.randomUUID()}`;
    const optimistic: GraphListItem = { id: tempId, fields };
    const mutationKey = listKey(options.scope, siteId, listId) + tempId;
    pending.add(mutationKey);
    try {
      await store.putItems(options.scope, siteId, listId, [optimistic]);
      await store.invalidateListPages(options.scope, siteId, listId);
      const result = await requireRequest()(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items`, { method: "POST", body: { fields } }) as GraphListItem;
      await store.deleteItems(options.scope, siteId, listId, [tempId]);
      await store.putItems(options.scope, siteId, listId, [result]);
      metric({ type: "mutation", operation: "create-item", count: 1 });
      return result;
    } catch (error) {
      metric({ type: "network-error", operation: "create-item", error });
      await store.deleteItems(options.scope, siteId, listId, [tempId]).catch((storageError: unknown) => reportStorageError(options, storageError));
      throw error;
    } finally {
      pending.delete(mutationKey);
    }
  }

  async function updateItem(siteId: string, listId: string, itemId: string, fields: Record<string, unknown>, config: { eTag?: string } = {}): Promise<GraphListItem> {
    const current = (await store.getItems(options.scope, siteId, listId)).find((item) => item.id === itemId);
    const eTag = config.eTag ?? current?.eTag;
    const optimistic = current ? { ...current, fields: { ...current.fields, ...fields } } : undefined;
    if (!current && !eTag) throw new Error(`Item ${itemId} is not cached; pass its ETag to update it safely.`);
    const mutationKey = listKey(options.scope, siteId, listId) + itemId;
    if (pending.has(mutationKey)) throw new Error(`Item ${itemId} already has a pending write.`);
    pending.add(mutationKey);
    try {
      if (optimistic) await store.putItems(options.scope, siteId, listId, [optimistic]);
      await store.invalidateListPages(options.scope, siteId, listId);
      const result = await requireRequest()(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}`, {
        method: "PATCH",
        headers: eTag ? { "If-Match": eTag } : undefined,
        body: { fields },
      }) as GraphListItem;
      const updated = optimistic ? { ...optimistic, ...result, fields: { ...optimistic.fields, ...result.fields } } : result;
      await store.putItems(options.scope, siteId, listId, [updated]);
      metric({ type: "mutation", operation: "update-item", count: 1 });
      return updated;
    } catch (error) {
      metric({ type: "network-error", operation: "update-item", error });
      if (current) await store.putItems(options.scope, siteId, listId, [current]).catch((storageError: unknown) => reportStorageError(options, storageError));
      throw error;
    } finally {
      pending.delete(mutationKey);
    }
  }

  async function deleteItem(siteId: string, listId: string, itemId: string): Promise<void> {
    const current = (await store.getItems(options.scope, siteId, listId)).find((item) => item.id === itemId);
    const mutationKey = listKey(options.scope, siteId, listId) + itemId;
    if (pending.has(mutationKey)) throw new Error(`Item ${itemId} already has a pending write.`);
    pending.add(mutationKey);
    try {
      await store.deleteItems(options.scope, siteId, listId, [itemId]);
      await store.invalidateListPages(options.scope, siteId, listId);
      await requireRequest()(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}`, {
        method: "DELETE",
        headers: current?.eTag ? { "If-Match": current.eTag } : undefined,
      });
      metric({ type: "mutation", operation: "delete-item", count: 1 });
    } catch (error) {
      metric({ type: "network-error", operation: "delete-item", error });
      if (current) await store.putItems(options.scope, siteId, listId, [current]).catch((storageError: unknown) => reportStorageError(options, storageError));
      throw error;
    } finally {
      pending.delete(mutationKey);
    }
  }

  async function inspectList(siteId: string, listId: string) {
    const [list, columns] = await Promise.all([
      getResource(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}`),
      collectPages(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/columns`),
    ]);
    return { list, columns };
  }

  async function getListContentTypes(siteId: string, listId: string): Promise<GraphContentType[]> {
    return collectPages<GraphContentType>(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/contentTypes`);
  }

  function getTermStoreGroups(siteId: string): Promise<GraphTermStoreGroup[]> {
    return collectPages<GraphTermStoreGroup>(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/termStore/groups`);
  }

  function getTermStoreSets(siteId: string, groupId: string): Promise<GraphTermStoreSet[]> {
    return collectPages<GraphTermStoreSet>(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/termStore/groups/${encodeURIComponent(groupId)}/sets`);
  }

  function getTermSetChildren(siteId: string, setId: string): Promise<GraphTermStoreTerm[]> {
    return collectPages<GraphTermStoreTerm>(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/termStore/sets/${encodeURIComponent(setId)}/children`);
  }

  function getTermChildren(siteId: string, setId: string, termId: string): Promise<GraphTermStoreTerm[]> {
    return collectPages<GraphTermStoreTerm>(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/termStore/sets/${encodeURIComponent(setId)}/terms/${encodeURIComponent(termId)}/children`);
  }

  async function collectPages<T>(firstUrl: string, maxPages = 100): Promise<T[]> {
    const result: T[] = [];
    let next: string | undefined = firstUrl;
    let pages = 0;
    while (next) {
      if (++pages > maxPages) throw new Error(`Graph collection exceeded ${maxPages} pages.`);
      const page: GraphCollection<T> = await fetchPage<T>(next);
      result.push(...page.value);
      next = page["@odata.nextLink"];
    }
    return result;
  }

  async function collectBoundedVersions<T>(firstUrl: string, config: { maxItems?: number; maxPages?: number; signal?: AbortSignal } = {}): Promise<T[]> {
    const maxItems = config.maxItems ?? 100;
    const maxPages = config.maxPages ?? 20;
    if (!Number.isInteger(maxItems) || maxItems < 1 || maxItems > 10_000) throw new RangeError("maxItems must be an integer from 1 to 10000.");
    if (!Number.isInteger(maxPages) || maxPages < 1 || maxPages > 100) throw new RangeError("maxPages must be an integer from 1 to 100.");
    assertGraphUrl(firstUrl);
    const endpoint = new URL(firstUrl);
    const result: T[] = [];
    let next: string | undefined = firstUrl;
    let pages = 0;
    while (next && result.length < maxItems) {
      if (config.signal?.aborted) throw config.signal.reason ?? new DOMException("Aborted", "AbortError");
      if (++pages > maxPages) throw new Error(`Graph version history exceeded ${maxPages} pages.`);
      const target = new URL(next);
      if (target.origin !== endpoint.origin || target.pathname !== endpoint.pathname || target.username || target.password) throw new TypeError("Graph version continuation must stay on the requested item versions endpoint.");
      const response = await options.get(target.toString(), config.signal) as GraphCollection<T>;
      if (!response || !Array.isArray(response.value)) throw new TypeError("Graph version response must contain a value array.");
      result.push(...response.value.slice(0, maxItems - result.length));
      next = response["@odata.nextLink"];
    }
    return result;
  }

  async function getSite(hostname: string, serverRelativePath: string) {
    const path = serverRelativePath.replace(/^\/+|\/+$/g, "");
    const encodedPath = path.split("/").map(encodeURIComponent).join("/");
    return getResource(`https://graph.microsoft.com/v1.0/sites/${hostname}:/${encodedPath}`);
  }

  /** Read a bounded, uncached Graph calendar window. Use Calendars.ReadBasic for the least-privileged delegated access. */
  async function getCalendarView(startDateTime: string, endDateTime: string, config: CalendarViewOptions = {}): Promise<CalendarViewResult> {
    const datePattern = /^\d{4}-\d\d-\d\dT\d\d:\d\d(?::\d\d(?:\.\d+)?)?(?:Z|[+-]\d\d:\d\d)$/i;
    if (!datePattern.test(startDateTime) || !Number.isFinite(Date.parse(startDateTime))) throw new TypeError("startDateTime must be an ISO 8601 date-time with a timezone.");
    if (!datePattern.test(endDateTime) || !Number.isFinite(Date.parse(endDateTime)) || Date.parse(startDateTime) >= Date.parse(endDateTime)) throw new TypeError("endDateTime must be after startDateTime and include a valid ISO 8601 timezone.");
    const maxItems = config.maxItems ?? 500;
    const maxPages = config.maxPages ?? 10;
    if (!Number.isInteger(maxItems) || maxItems < 1 || maxItems > 5000) throw new RangeError("maxItems must be an integer from 1 to 5000.");
    if (!Number.isInteger(maxPages) || maxPages < 1 || maxPages > 100) throw new RangeError("maxPages must be an integer from 1 to 100.");
    const path = config.calendarId
      ? `/v1.0/me/calendars/${encodeURIComponent(config.calendarId)}/calendarView`
      : "/v1.0/me/calendar/calendarView";
    const validateCursor = (value: string) => {
      const target = new URL(value);
      if (target.origin !== "https://graph.microsoft.com" || target.pathname !== path || target.username || target.password ||
        target.searchParams.get("startDateTime") !== startDateTime || target.searchParams.get("endDateTime") !== endDateTime) {
        throw new TypeError("Calendar continuation must stay on the requested calendar window.");
      }
      return target.toString();
    };
    let next = config.cursor ? validateCursor(config.cursor) : undefined;
    if (!next) {
      const initial = new URL(path, "https://graph.microsoft.com");
      initial.searchParams.set("startDateTime", startDateTime);
      initial.searchParams.set("endDateTime", endDateTime);
      initial.searchParams.set("$select", "id,subject,start,end,isAllDay,location,webLink,organizer");
      initial.searchParams.set("$orderby", "start/dateTime");
      initial.searchParams.set("$top", String(Math.min(1000, maxItems)));
      next = initial.toString();
    }
    const items: GraphCalendarEvent[] = [];
    let pages = 0;
    while (next && items.length < maxItems && pages < maxPages) {
      if (config.signal?.aborted) throw config.signal.reason ?? new DOMException("Aborted", "AbortError");
      const requested = validateCursor(next);
      const requestedPageSize = Number(new URL(requested).searchParams.get("$top") ?? Math.min(1000, maxItems));
      if (!Number.isInteger(requestedPageSize) || requestedPageSize < 1 || requestedPageSize > 1000) throw new TypeError("Calendar continuation has an invalid page size.");
      if (requestedPageSize > maxItems) throw new RangeError("Resume the calendar cursor with maxItems at least as large as the original page size.");
      const response = await requestGraph(requested, { signal: config.signal }) as GraphCollection<GraphCalendarEvent>;
      if (!response || !Array.isArray(response.value)) throw new TypeError("Graph calendar response must contain a value array.");
      if (response.value.length > requestedPageSize) throw new TypeError("Graph calendar page exceeded its requested page size.");
      if (items.length + response.value.length > maxItems) {
        next = requested;
        break;
      }
      items.push(...response.value);
      next = response["@odata.nextLink"] ? validateCursor(response["@odata.nextLink"]!) : undefined;
      pages++;
    }
    return { items, nextCursor: next, complete: !next };
  }

  async function getLists(siteId: string): Promise<unknown[]> {
    return collectPages(`https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists?$select=id,displayName,name,webUrl,list`);
  }

  async function findList(siteId: string, title: string) {
    const lists = await getLists(siteId) as Array<{ id: string; displayName?: string; name?: string }>;
    const match = lists.find((list) => (list.displayName ?? list.name)?.toLocaleLowerCase() === title.toLocaleLowerCase());
    if (!match) throw new Error(`SharePoint list "${title}" was not found.`);
    return match;
  }

  return {
    readPage,
    listItems,
    getItemPage,
    queryItems,
    subscribeItems,
    syncList,
    getSite,
    getCalendarView,
    getDriveItem,
    getDriveItemPage,
    getDriveItemPermissions,
    createDriveItemSharingLink,
    deleteDriveItemPermission,
    copyDriveItem,
    downloadDriveItemVersion,
    restoreDriveItemVersion,
    readDriveItemContent,
    downloadDriveItem,
    uploadDriveItem,
    createDriveFolder,
    moveDriveItem,
    deleteDriveItem,
    getDriveItems,
    subscribeDriveItems,
    syncDrive,
    getLists,
    findList,
    inspectList,
    getListContentTypes,
    getTermStoreGroups,
    getTermStoreSets,
    getTermSetChildren,
    getTermChildren,
    getListItemVersions(siteId: string, listId: string, itemId: string, config: { maxItems?: number; maxPages?: number; signal?: AbortSignal } = {}) {
      const url = `https://graph.microsoft.com/v1.0/sites/${encodeURIComponent(siteId)}/lists/${encodeURIComponent(listId)}/items/${encodeURIComponent(itemId)}/versions`;
      return collectBoundedVersions<GraphListItemVersion>(url, config);
    },
    getDriveItemVersions(driveId: string, itemId: string, config: { maxItems?: number; maxPages?: number; signal?: AbortSignal } = {}) {
      const url = `https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/versions`;
      return collectBoundedVersions<GraphDriveItemVersion>(url, config);
    },
    batch,
    requestGraph,
    createItem,
    updateItem,
    deleteItem,
    isPending: (siteId: string, listId: string, itemId: string) => pending.has(listKey(options.scope, siteId, listId) + itemId),
    clearCache: async () => {
      try {
        await store.clearScope(options.scope);
      } catch (error) {
        reportStorageError(options, error);
        throw error;
      }
    },
  };
}
