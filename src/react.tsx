import { useCallback, useEffect, useMemo, useRef, useState, type FormEvent, type ReactElement, type ReactNode } from "react";
import {
  Button,
  Checkbox,
  DataGrid,
  DataGridBody,
  DataGridCell,
  DataGridHeader,
  DataGridHeaderCell,
  DataGridRow,
  Dropdown,
  Field,
  Input,
  makeStyles,
  MessageBar,
  MessageBarBody,
  Option,
  SearchBox,
  Spinner,
  TableColumnDefinition,
  Textarea,
  createTableColumn,
  tokens,
} from "@fluentui/react-components";
import { applyItemQuery, createClient, type GraphDriveItem, type GraphListItem, type GraphListItemFilter, type ItemQuery } from "./index.js";

type Client = ReturnType<typeof createClient>;

const useLayoutStyles = makeStyles({
  stack: { display: "flex", flexDirection: "column", rowGap: tokens.spacingVerticalM },
  tableWrap: { maxWidth: "100%", overflowX: "auto" },
  detailList: { display: "grid", rowGap: tokens.spacingVerticalM },
  detail: { display: "grid", rowGap: tokens.spacingVerticalXS },
});

export interface UseItemsResult {
  items: GraphListItem[];
  source?: "cache" | "network" | "local";
  loading: boolean;
  error?: unknown;
  refresh(): Promise<void>;
}

export interface UseListItemsResult {
  items: GraphListItem[];
  source?: "cache" | "network";
  loading: boolean;
  error?: unknown;
  hasMore: boolean;
  refresh(): Promise<void>;
  loadMore(): Promise<void>;
}

export interface UseDriveItemsResult {
  items: GraphDriveItem[];
  source?: "cache" | "network" | "local";
  loading: boolean;
  error?: unknown;
  complete?: boolean;
  refresh(): Promise<void>;
}

export interface UseDriveItemContentResult {
  content?: Blob | string;
  source?: "cache" | "network";
  loading: boolean;
  error?: unknown;
  refresh(): Promise<void>;
}

export interface UseMarkdownFileResult extends UseDriveItemContentResult {
  markdown?: string;
}

/** Cache-aware authenticated file content hook. Markdown remains data; rendering policy belongs to the app. */
export function useDriveItemContent(
  client: Client,
  driveId: string,
  itemId: string,
  options: { as?: "text" | "blob"; enabled?: boolean; cacheMaxBytes?: number } = {},
): UseDriveItemContentResult {
  const as = options.as ?? "blob";
  const enabled = options.enabled ?? true;
  const identity = JSON.stringify([driveId, itemId, as]);
  const [state, setState] = useState<{ identity: string; content?: Blob | string; source?: "cache" | "network"; loading: boolean; error?: unknown }>();
  const requestId = useRef(0);
  const refresh = useCallback(async () => {
    if (!enabled || !driveId || !itemId) return;
    const currentRequest = ++requestId.current;
    setState((previous) => previous?.identity === identity
      ? { ...previous, loading: true, error: undefined }
      : { identity, loading: true });
    try {
      const result = await client.readDriveItemContent(driveId, itemId, { as, cacheMaxBytes: options.cacheMaxBytes });
      if (currentRequest !== requestId.current) return;
      setState({ identity, content: result.data, source: result.source, loading: false });
      try {
        const fresh = await result.refresh;
        if (currentRequest !== requestId.current) return;
        setState((previous) => previous?.identity === identity ? { ...previous, content: fresh, source: "network" } : previous);
      } catch (reason) {
        if (currentRequest !== requestId.current) return;
        setState((previous) => previous?.identity === identity ? { ...previous, error: reason } : previous);
      }
    } catch (reason) {
      if (currentRequest !== requestId.current) return;
      setState((previous) => previous?.identity === identity
        ? { ...previous, loading: false, error: reason }
        : { identity, loading: false, error: reason });
    }
  }, [client, driveId, itemId, as, enabled, identity, options.cacheMaxBytes]);

  useEffect(() => {
    void refresh();
    return () => { requestId.current++; };
  }, [refresh]);
  const current = state?.identity === identity ? state : undefined;
  return { content: current?.content, source: current?.source, loading: current?.loading ?? enabled, error: current?.error, refresh };
}

export type FloorPlanLocation = { id: string; label: string; x: number; y: number; detail?: string };

const useFloorPlanStyles = makeStyles({
  layout: {
    display: "grid",
    gridTemplateColumns: "minmax(0, 1.45fr) minmax(200px, 0.75fr)",
    gap: tokens.spacingHorizontalL,
    alignItems: "start",
    "@media (max-width: 680px)": { gridTemplateColumns: "1fr" },
  },
});

/** Displays caller-supplied normalized (0–100) locations over a local or authenticated visual. */
export function FloorPlanMap({
  locations,
  plan,
  title = "Locations",
  onSelect,
}: {
  locations: readonly FloorPlanLocation[];
  /** Decorative image/SVG; load private assets through the authenticated host. */
  plan?: ReactNode;
  title?: string;
  onSelect?: (location: FloorPlanLocation) => void;
}): ReactElement {
  const styles = useFloorPlanStyles();
  const visible = locations.filter(({ id, label, x, y }) =>
    Boolean(id && label.trim()) && Number.isFinite(x) && Number.isFinite(y) && x >= 0 && x <= 100 && y >= 0 && y <= 100,
  );
  const [selectedId, setSelectedId] = useState(visible[0]?.id);
  const selected = visible.find(({ id }) => id === selectedId) ?? visible[0];

  useEffect(() => {
    if (selected?.id !== selectedId) setSelectedId(selected?.id);
  }, [selected?.id, selectedId]);

  function select(location: FloorPlanLocation) {
    setSelectedId(location.id);
    onSelect?.(location);
  }

  return <section aria-label={title} style={{ display: "grid", gap: tokens.spacingVerticalL }}>
    <h2 style={{ margin: 0, color: tokens.colorNeutralForeground1, fontSize: tokens.fontSizeBase600, lineHeight: tokens.lineHeightBase600 }}>{title}</h2>
    {visible.length === 0 ? <p>No locations with valid map coordinates.</p> : <>
      <div className={styles.layout}>
        <div role="group" aria-label={`${title} map`} style={{ position: "relative", overflow: "hidden", aspectRatio: "4 / 3", minHeight: 240, border: `1px solid ${tokens.colorNeutralStroke2}`, borderRadius: tokens.borderRadiusMedium, background: tokens.colorNeutralBackground2 }}>
          <div aria-hidden="true" style={{ position: "absolute", inset: 0, display: "grid", placeItems: "center", color: tokens.colorNeutralForeground1 }}>
            {plan ?? <DefaultFloorPlan />}
          </div>
          {visible.map((location) => <Button
            key={location.id}
            type="button"
            aria-pressed={selected?.id === location.id}
            appearance={selected?.id === location.id ? "primary" : "secondary"}
            size="small"
            shape="circular"
            onClick={() => select(location)}
            style={{ position: "absolute", left: `${location.x}%`, top: `${location.y}%`, transform: "translate(-50%, -50%)", minWidth: 0, boxShadow: tokens.shadow4 }}
          >{location.label}</Button>)}
        </div>
        <nav aria-label={`${title} location list`} style={{ display: "grid", gap: tokens.spacingVerticalXS, maxWidth: 400, width: "100%" }}>
          {visible.map((location) => <Button
            key={location.id}
            type="button"
            aria-pressed={selected?.id === location.id}
            appearance={selected?.id === location.id ? "primary" : "subtle"}
            onClick={() => select(location)}
            style={{ justifyContent: "flex-start", textAlign: "left", width: "100%" }}
          >{location.label}</Button>)}
        </nav>
      </div>
      {selected ? <div aria-live="polite" aria-atomic="true">
        <strong>{selected.label}</strong>{selected.detail ? <span>: {selected.detail}</span> : null}
      </div> : null}
    </>}
  </section>;
}

const supportedFloorPlanImageTypes = new Set(["image/png", "image/jpeg", "image/webp"]);

/** Loads a private plan image through the signed-in Graph client and cleans up its object URL. */
export function DriveFloorPlanMap({
  client,
  driveId,
  itemId,
  locations,
  title = "Locations",
  cacheMaxBytes = 8 * 1024 * 1024,
  onSelect,
}: {
  client: Client;
  driveId: string;
  itemId: string;
  locations: readonly FloorPlanLocation[];
  title?: string;
  cacheMaxBytes?: number;
  onSelect?: (location: FloorPlanLocation) => void;
}): ReactElement {
  const { content, loading, error } = useDriveItemContent(client, driveId, itemId, { as: "blob", cacheMaxBytes });
  const image = content instanceof Blob && supportedFloorPlanImageTypes.has(content.type.toLowerCase()) ? content : undefined;
  const [preview, setPreview] = useState<{ image: Blob; url: string }>();

  useEffect(() => {
    if (!image) {
      setPreview(undefined);
      return;
    }
    const url = URL.createObjectURL(image);
    setPreview({ image, url });
    return () => URL.revokeObjectURL(url);
  }, [image]);

  const imageUrl = preview && image && preview.image === image ? preview.url : undefined;
  return <section aria-label={`${title} floor plan`}>
    {loading ? <p role="status">Loading floor plan…</p> : null}
    {error ? <p role="alert">The floor plan could not be loaded. Check file access and try again.</p> : null}
    {!loading && !error && content && !image ? <p role="status">Use a PNG, JPEG, or WebP floor-plan image.</p> : null}
    <FloorPlanMap
      locations={locations}
      title={title}
      onSelect={onSelect}
      plan={imageUrl ? <img src={imageUrl} alt="" style={{ width: "100%", height: "100%", objectFit: "contain" }} /> : undefined}
    />
  </section>;
}

function DefaultFloorPlan(): ReactElement {
  return <svg viewBox="0 0 800 600" role="img" aria-label="Office plan with four rooms and a central corridor" style={{ width: "92%", height: "90%" }}>
    <rect x="22" y="22" width="756" height="556" rx="12" fill={tokens.colorNeutralBackground1} stroke={tokens.colorNeutralStroke1} strokeWidth="10" />
    <rect x="38" y="38" width="328" height="238" rx="4" fill={tokens.colorPaletteBlueBackground2} stroke={tokens.colorNeutralStroke2} strokeWidth="5" />
    <rect x="378" y="38" width="384" height="238" rx="4" fill={tokens.colorPaletteYellowBackground2} stroke={tokens.colorNeutralStroke2} strokeWidth="5" />
    <rect x="38" y="324" width="328" height="238" rx="4" fill={tokens.colorPaletteGreenBackground2} stroke={tokens.colorNeutralStroke2} strokeWidth="5" />
    <rect x="378" y="324" width="384" height="238" rx="4" fill={tokens.colorPaletteLavenderBackground2} stroke={tokens.colorNeutralStroke2} strokeWidth="5" />
    <path d="M38 300H762" stroke={tokens.colorNeutralBackground2} strokeWidth="22" />
    <g fill={tokens.colorNeutralForeground1} fontFamily={tokens.fontFamilyBase} fontWeight={tokens.fontWeightBold} fontSize={tokens.fontSizeBase400} letterSpacing="1">
      <text x="62" y="78">OPEN WORKSPACE</text>
      <text x="402" y="78">COLLABORATION</text>
      <text x="62" y="364">FOCUS ROOMS</text>
      <text x="402" y="364">MEETING ROOM 04</text>
    </g>
    <g fill={tokens.colorNeutralBackground1} stroke={tokens.colorNeutralStroke1} strokeWidth="3">
      <rect x="68" y="106" width="82" height="50" rx="6" /><rect x="164" y="106" width="82" height="50" rx="6" />
      <rect x="68" y="176" width="82" height="50" rx="6" /><rect x="164" y="176" width="82" height="50" rx="6" />
      <rect x="448" y="110" width="250" height="74" rx="14" />
      <rect x="78" y="402" width="106" height="116" rx="8" /><rect x="214" y="402" width="106" height="116" rx="8" />
      <rect x="452" y="408" width="238" height="98" rx="10" />
    </g>
    <g fill={tokens.colorNeutralForeground2} fontFamily={tokens.fontFamilyBase} fontSize={tokens.fontSizeBase200} fontWeight={tokens.fontWeightSemibold}>
      <text x="72" y="250">DESK BANK A</text>
      <text x="490" y="210">6 SEATS · DISPLAY</text>
      <text x="82" y="540">QUIET ZONE</text>
      <text x="476" y="532">ROOM 04 · 8 SEATS</text>
    </g>
  </svg>;
}

/** Cached-first Markdown text hook for files read inside an authenticated SPFx host. */
export function useMarkdownFile(
  client: Client,
  driveId: string,
  itemId: string,
  options: { enabled?: boolean; cacheMaxBytes?: number } = {},
): UseMarkdownFileResult {
  const result = useDriveItemContent(client, driveId, itemId, { ...options, as: "text" });
  return { ...result, markdown: typeof result.content === "string" ? result.content : undefined };
}

/** Local-first hook: reads IndexedDB first, then incrementally syncs the list in the background. */
export function useItems(
  client: Client,
  siteId: string,
  listId: string,
  options: { query?: ItemQuery; fields?: string[]; sync?: boolean; syncIntervalMs?: number; syncOnVisible?: boolean; maxItems?: number; syncFn?: () => Promise<unknown> } = {},
): UseItemsResult {
  const [items, setItems] = useState<GraphListItem[]>([]);
  const [source, setSource] = useState<UseItemsResult["source"]>();
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<unknown>();
  const hasLocalData = useRef(false);
  const queryJson = JSON.stringify(options.query ?? {});
  const fieldsJson = JSON.stringify(options.fields ?? []);

  const refresh = useCallback(async () => {
    setError(undefined);
    try {
      await (options.syncFn?.() ?? client.syncList(siteId, listId, { fields: JSON.parse(fieldsJson) as string[] }));
      setItems(await client.queryItems(siteId, listId, JSON.parse(queryJson) as ItemQuery));
      setSource("network");
    } catch (reason) {
      setError(reason);
    } finally {
      setLoading(false);
    }
  }, [client, fieldsJson, listId, options.syncFn, queryJson, siteId]);

  useEffect(() => {
    let active = true;
    let lastSync = 0;
    const controller = new AbortController();
    const update = async () => {
      try {
        const next = await client.queryItems(siteId, listId, JSON.parse(queryJson) as ItemQuery);
        if (active) { hasLocalData.current = next.length > 0; setItems(next); setSource(next.length ? "cache" : "local"); }
      } catch (reason) {
        if (active) setError(reason);
      } finally {
        if (active && options.sync === false) setLoading(false);
      }
    };
    const unsubscribe = client.subscribeItems(siteId, listId, update);
    const sync = async () => {
      if (options.sync === false || Date.now() - lastSync < (options.syncIntervalMs ?? 30_000)) return;
      lastSync = Date.now();
      try {
        if (options.syncFn) await options.syncFn();
        else await client.syncList(siteId, listId, { fields: JSON.parse(fieldsJson) as string[], signal: controller.signal });
        if (active) { setSource("network"); setLoading(false); }
      } catch (reason) {
        if (controller.signal.aborted) return;
        if (options.syncFn || hasLocalData.current) {
          if (active) { setError(reason); setLoading(false); }
          return;
        }
        try {
          const fallback: GraphListItem[] = [];
          for await (const item of client.listItems<GraphListItem>(siteId, listId, { fields: JSON.parse(fieldsJson) as string[], maxItems: options.maxItems ?? 5000, signal: controller.signal })) fallback.push(item);
          if (active) {
            const result = applyItemQuery(fallback, JSON.parse(queryJson) as ItemQuery);
            hasLocalData.current = result.length > 0;
            setItems(result);
            setSource("network");
            setError(undefined);
            setLoading(false);
          }
        } catch (fallbackError) {
          if (active) { setError(fallbackError ?? reason); setLoading(false); }
        }
      }
    };
    const onVisibility = () => { if (options.syncOnVisible !== false && document.visibilityState === "visible") void sync(); };
    void update();
    void sync();
    document.addEventListener("visibilitychange", onVisibility);
    window.addEventListener("focus", onVisibility);
    return () => {
      active = false;
      controller.abort();
      document.removeEventListener("visibilitychange", onVisibility);
      window.removeEventListener("focus", onVisibility);
      unsubscribe();
    };
  }, [client, fieldsJson, listId, options.maxItems, options.sync, options.syncIntervalMs, options.syncOnVisible, options.syncFn, queryJson, siteId]);

  return useMemo(() => ({ items, source, loading, error, refresh }), [error, items, loading, refresh, source]);
}

/** Cached-first drive hierarchy hook backed by resumable Graph delta synchronization. */
export function useDriveItems(
  client: Client,
  driveId: string,
  options: { enabled?: boolean; maxPages?: number; syncIntervalMs?: number; syncOnVisible?: boolean } = {},
): UseDriveItemsResult {
  const [items, setItems] = useState<GraphDriveItem[]>([]);
  const [source, setSource] = useState<UseDriveItemsResult["source"]>();
  const [loading, setLoading] = useState(Boolean(options.enabled ?? true));
  const [error, setError] = useState<unknown>();
  const [complete, setComplete] = useState<boolean>();
  const controller = useRef<AbortController>();
  const inFlight = useRef<{ client: Client; driveId: string; maxPages: number; promise: Promise<void> }>();
  const lastSync = useRef(0);
  const enabled = options.enabled !== false;
  const maxPages = options.maxPages ?? 100;

  const refresh = useCallback((): Promise<void> => {
    if (inFlight.current?.client === client && inFlight.current.driveId === driveId && inFlight.current.maxPages === maxPages) return inFlight.current.promise;
    const currentController = controller.current;
    const operation = (async () => {
      setError(undefined);
      setLoading(true);
      lastSync.current = Date.now();
      try {
        const result = await client.syncDrive(driveId, { maxPages, signal: currentController?.signal });
        const updated = await client.getDriveItems(driveId);
        if (currentController?.signal.aborted) return;
        setItems(updated);
        setSource("network");
        setComplete(result.complete);
      } catch (reason) {
        if (!currentController?.signal.aborted) setError(reason);
      } finally {
        if (!currentController?.signal.aborted) setLoading(false);
      }
    })().finally(() => {
      if (inFlight.current?.promise === operation) inFlight.current = undefined;
    });
    inFlight.current = { client, driveId, maxPages, promise: operation };
    return operation;
  }, [client, driveId, maxPages]);

  useEffect(() => {
    let active = true;
    const currentController = new AbortController();
    controller.current = currentController;
    const update = async () => {
      try {
        const cached = await client.getDriveItems(driveId);
        if (active) {
          setItems(cached);
          setSource((current) => current === "network" ? current : cached.length ? "cache" : "local");
        }
      } catch (reason) {
        if (active) setError(reason);
      } finally {
        if (active && !enabled) setLoading(false);
      }
    };
    const unsubscribe = client.subscribeDriveItems(driveId, () => { void update(); });
    void update();
    if (enabled && driveId) void refresh();
    const onVisible = () => {
      if (enabled && options.syncOnVisible !== false && document.visibilityState === "visible" && Date.now() - lastSync.current >= (options.syncIntervalMs ?? 30_000)) void refresh();
    };
    document.addEventListener("visibilitychange", onVisible);
    window.addEventListener("focus", onVisible);
    return () => {
      active = false;
      currentController.abort();
      if (controller.current === currentController) controller.current = undefined;
      unsubscribe();
      document.removeEventListener("visibilitychange", onVisible);
      window.removeEventListener("focus", onVisible);
    };
  }, [client, driveId, enabled, options.syncIntervalMs, options.syncOnVisible, refresh]);

  return useMemo(() => ({ items, source, loading, error, complete, refresh }), [complete, error, items, loading, refresh, source]);
}

/** Cached paged list hook with safe server-side OData filtering and explicit load-more control. */
export function useListItems(
  client: Client,
  siteId: string,
  listId: string,
  options: { fields?: string[]; filter?: GraphListItemFilter; pageSize?: number; maxItems?: number; enabled?: boolean } = {},
): UseListItemsResult {
  const [pages, setPages] = useState<GraphListItem[][]>([]);
  const [cursor, setCursor] = useState<string>();
  const [source, setSource] = useState<UseListItemsResult["source"]>();
  const [loading, setLoading] = useState(options.enabled !== false);
  const [error, setError] = useState<unknown>();
  const generation = useRef(0);
  const loadingRef = useRef(false);
  const pagesRef = useRef<GraphListItem[][]>([]);
  const fieldsJson = JSON.stringify(options.fields ?? []);
  const filterJson = JSON.stringify(options.filter ?? null);
  const requestedPageSize = options.pageSize ?? 50;
  const pageSize = Number.isFinite(requestedPageSize) ? Math.min(Math.max(Math.floor(requestedPageSize), 1), 100) : 50;
  const requestedMaxItems = options.maxItems ?? 500;
  const maxItems = Number.isFinite(requestedMaxItems) ? Math.max(0, Math.floor(requestedMaxItems)) : 500;
  const enabled = options.enabled !== false;

  const fetchPage = useCallback(async (nextCursor: string | undefined, replace: boolean, requestGeneration: number) => {
    if (!enabled || requestGeneration !== generation.current || loadingRef.current) return;
    if (maxItems === 0) { setLoading(false); return; }
    loadingRef.current = true;
    setLoading(true);
    setError(undefined);
    const pageIndex = replace ? 0 : pagesRef.current.length;
    try {
      const result = await client.getItemPage(siteId, listId, {
        cursor: nextCursor,
        fields: JSON.parse(fieldsJson) as string[],
        filter: (JSON.parse(filterJson) as GraphListItemFilter | null) ?? undefined,
        pageSize,
      });
      if (requestGeneration !== generation.current) return;
      const alreadyLoaded = replace ? 0 : pagesRef.current.reduce((count, page) => count + page.length, 0);
      const rows = result.items.slice(0, Math.max(maxItems - alreadyLoaded, 0));
      pagesRef.current = replace ? [rows] : [...pagesRef.current, rows];
      setPages([...pagesRef.current]);
      setCursor(result.nextCursor);
      setSource(result.source);
      void result.refresh.then((fresh) => {
        if (requestGeneration !== generation.current) return;
        const refreshed = fresh.value.slice(0, Math.max(maxItems - pageIndex * pageSize, 0));
        pagesRef.current = pagesRef.current.map((page, index) => index === pageIndex ? refreshed : page);
        setPages([...pagesRef.current]);
        setSource("network");
      }).catch((reason: unknown) => {
        if (requestGeneration === generation.current) setError(reason);
      });
    } catch (reason) {
      if (requestGeneration === generation.current) setError(reason);
    } finally {
      if (requestGeneration === generation.current) {
        loadingRef.current = false;
        setLoading(false);
      }
    }
  }, [client, enabled, fieldsJson, filterJson, listId, maxItems, pageSize, siteId]);

  const refresh = useCallback(async () => {
    const requestGeneration = ++generation.current;
    loadingRef.current = false;
    pagesRef.current = [];
    setPages([]);
    setCursor(undefined);
    setSource(undefined);
    setError(undefined);
    if (!enabled) { setLoading(false); return; }
    await fetchPage(undefined, true, requestGeneration);
  }, [enabled, fetchPage]);

  useEffect(() => {
    const requestGeneration = ++generation.current;
    loadingRef.current = false;
    pagesRef.current = [];
    setPages([]);
    setCursor(undefined);
    setSource(undefined);
    setError(undefined);
    if (enabled) void fetchPage(undefined, true, requestGeneration);
    else setLoading(false);
    return () => {
      generation.current++;
      loadingRef.current = false;
    };
  }, [enabled, fetchPage]);

  const loadMore = useCallback(async () => {
    if (cursor && pagesRef.current.reduce((count, page) => count + page.length, 0) < maxItems) {
      await fetchPage(cursor, false, generation.current);
    }
  }, [cursor, fetchPage, maxItems]);
  const items = useMemo(() => pages.flat().slice(0, maxItems), [maxItems, pages]);
  return { items, source, loading, error, hasMore: Boolean(cursor) && items.length < maxItems, refresh, loadMore };
}

export interface FieldDefinition {
  name: string;
  title: string;
  type?: "text" | "multiline" | "number" | "date" | "boolean" | "choice" | "multiChoice" | "person" | "lookup";
  required?: boolean;
  choices?: string[];
  /** Preloaded options from the authenticated host; values are SharePoint user or target-item IDs. */
  options?: Array<{ value: string | number; label: string }>;
}

/** Map form values to Graph fieldValueSet keys for Person and Lookup columns. */
export function toGraphFieldValues(fields: FieldDefinition[], values: Record<string, unknown>): Record<string, unknown> {
  const result: Record<string, unknown> = { ...values };
  for (const field of fields) {
    if (field.type !== "person" && field.type !== "lookup") continue;
    if (!(field.name in values)) continue;
    const value = values[field.name];
    delete result[field.name];
    result[`${field.name}LookupId`] = value;
  }
  return result;
}

function toFormValues(fields: FieldDefinition[], values: Record<string, unknown>): Record<string, unknown> {
  const result = { ...values };
  for (const field of fields) {
    if ((field.type === "person" || field.type === "lookup") && result[field.name] === undefined) {
      result[field.name] = values[`${field.name}LookupId`];
    }
  }
  return result;
}

/** Required fields with no value; false and zero remain valid values. */
export function getMissingRequiredFieldNames(fields: FieldDefinition[], values: Record<string, unknown>): string[] {
  return fields
    .filter((field) => field.required)
    .filter((field) => {
      const value = values[field.name];
      return value === undefined || value === null || (typeof value === "string" && value.trim() === "") || (Array.isArray(value) && value.length === 0);
    })
    .map((field) => field.name);
}

export function ListView(props: {
  items: GraphListItem[];
  fields: FieldDefinition[];
  loading?: boolean;
  error?: unknown;
  pageSize?: number;
  renderItem?: (item: GraphListItem) => ReactNode;
  renderField?: (item: GraphListItem, field: FieldDefinition) => ReactNode | undefined;
  onSelect?: (item: GraphListItem) => void;
}) {
  const styles = useLayoutStyles();
  const [search, setSearch] = useState("");
  const [pageIndex, setPageIndex] = useState(0);
  const pageSize = Math.max(1, Math.floor(props.pageSize ?? 100));
  const visible = useMemo(() => {
    const needle = search.toLocaleLowerCase();
    return props.items.filter((item) => props.fields.some((field) => String(item.fields?.[field.name] ?? "").toLocaleLowerCase().includes(needle)));
  }, [props.fields, props.items, search]);
  const pageCount = Math.ceil(visible.length / pageSize);
  const page = visible.slice(pageIndex * pageSize, (pageIndex + 1) * pageSize);
  const columns = useMemo<TableColumnDefinition<GraphListItem>[]>(() => props.fields.map((field) => createTableColumn<GraphListItem>({
    columnId: field.name,
    renderHeaderCell: () => field.title,
    renderCell: (item) => props.renderItem ? props.renderItem(item) : props.renderField?.(item, field) ?? formatFieldValue(item.fields?.[field.name], field),
  })), [props.fields, props.renderField, props.renderItem]);

  useEffect(() => setPageIndex(0), [search, props.items, pageSize]);

  return (
    <div className={styles.stack} aria-busy={props.loading}>
      <SearchBox aria-label="Search list items" placeholder="Search this list" value={search} onChange={(_, data) => setSearch(data.value)} />
      {props.error ? <MessageBar intent="error"><MessageBarBody>Could not load this list.</MessageBarBody></MessageBar> : null}
      {props.loading && props.items.length === 0 ? <Spinner size="tiny" label="Loading items" /> : null}
      {!props.loading && visible.length === 0 ? <MessageBar intent="info"><MessageBarBody>No matching items.</MessageBarBody></MessageBar> : null}
      {page.length > 0 ? (
        <div className={styles.tableWrap}>
          <DataGrid
            items={page}
            columns={columns}
            getRowId={(item) => item.id}
            selectionMode={props.onSelect ? "single" : undefined}
            onSelectionChange={props.onSelect ? (_, data) => {
              const selectedId = data.selectedItems.values().next().value;
              const item = page.find((candidate) => candidate.id === selectedId);
              if (item) props.onSelect?.(item);
            } : undefined}
          >
            <DataGridHeader>
              <DataGridRow>
                {({ renderHeaderCell }) => <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>}
              </DataGridRow>
            </DataGridHeader>
            <DataGridBody<GraphListItem>>
              {({ item, rowId }) => (
                <DataGridRow<GraphListItem> key={rowId}>
                  {({ renderCell }) => <DataGridCell>{renderCell(item)}</DataGridCell>}
                </DataGridRow>
              )}
            </DataGridBody>
          </DataGrid>
        </div>
      ) : null}
      {pageCount > 1 ? <nav aria-label="List pagination" style={{ display: "flex", alignItems: "center", gap: 8 }}>
        <Button appearance="secondary" disabled={pageIndex === 0} onClick={() => setPageIndex((index) => Math.max(0, index - 1))}>Previous page</Button>
        <span aria-live="polite">Page {pageIndex + 1} of {pageCount} · Showing {pageIndex * pageSize + 1}–{Math.min((pageIndex + 1) * pageSize, visible.length)} of {visible.length}</span>
        <Button appearance="secondary" disabled={pageIndex + 1 >= pageCount} onClick={() => setPageIndex((index) => Math.min(pageCount - 1, index + 1))}>Next page</Button>
      </nav> : visible.length > 0 ? <span>Showing {visible.length} {visible.length === 1 ? "item" : "items"}</span> : null}
    </div>
  );
}

export function DetailView(props: { item: GraphListItem; fields: FieldDefinition[] }) {
  const styles = useLayoutStyles();
  return (
    <dl className={styles.detailList}>
      {props.fields.map((field) => (
        <div className={styles.detail} key={field.name}>
          <dt><strong>{field.title}</strong></dt>
          <dd>{formatFieldValue(props.item.fields?.[field.name] ?? ((field.type === "person" || field.type === "lookup") ? props.item.fields?.[`${field.name}LookupId`] : undefined), field)}</dd>
        </div>
      ))}
    </dl>
  );
}

export function ItemForm(props: {
  fields: FieldDefinition[];
  initial?: Record<string, unknown>;
  submitLabel?: string;
  attachments?: ReactNode;
  onSubmit(values: Record<string, unknown>): Promise<void> | void;
}) {
  const styles = useLayoutStyles();
  const [values, setValues] = useState<Record<string, unknown>>(() => toFormValues(props.fields, props.initial ?? {}));
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<unknown>();
  const [missingRequiredFields, setMissingRequiredFields] = useState<string[]>([]);
  const invalidChoiceRef = useRef<HTMLButtonElement>(null);
  const missingChoiceName = props.fields.find((field) => ["choice", "multiChoice", "person", "lookup"].includes(field.type ?? "") && missingRequiredFields.includes(field.name))?.name;
  const initialJson = JSON.stringify(props.initial ?? {});
  useEffect(() => {
    setValues(toFormValues(props.fields, JSON.parse(initialJson) as Record<string, unknown>));
    setMissingRequiredFields([]);
    setError(undefined);
  }, [initialJson, props.fields]);
  useEffect(() => {
    if (missingChoiceName) invalidChoiceRef.current?.focus();
  }, [missingChoiceName]);
  const submit = async (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();
    const missing = getMissingRequiredFieldNames(props.fields, values);
    if (missing.length > 0) {
      setMissingRequiredFields(missing);
      setError(undefined);
      return;
    }
    setMissingRequiredFields([]);
    setSaving(true);
    setError(undefined);
    try {
      await props.onSubmit(toGraphFieldValues(props.fields, values));
    } catch (reason) {
      setError(reason);
    } finally {
      setSaving(false);
    }
  };

  return (
    <>
      <form className={styles.stack} onSubmit={submit}>
        {props.fields.map((field) => {
        const value = values[field.name];
        const onChange = (next: unknown) => {
          setValues((current) => ({ ...current, [field.name]: next }));
          setMissingRequiredFields((current) => current.filter((name) => name !== field.name));
        };
        if (field.type === "boolean") {
          return <Checkbox key={field.name} label={field.title} checked={Boolean(value)} onChange={(_, data) => onChange(data.checked)} />;
        }
        if (field.type === "choice" || field.type === "multiChoice" || field.type === "person" || field.type === "lookup") {
          const isReference = field.type === "person" || field.type === "lookup";
          const options = isReference
            ? (field.options ?? []).map((option) => ({ value: String(option.value), label: option.label }))
            : (field.choices ?? []).map((choice) => ({ value: choice, label: choice }));
          const selectedOptions = field.type === "multiChoice"
            ? Array.isArray(value) ? value.map(String) : []
            : value === undefined ? [] : [String(value)];
          const selectedLabel = options.find((option) => option.value === String(value))?.label ?? String(value ?? "");
          return (
            <Field
              key={field.name}
              label={field.title}
              required={field.required}
              validationState={missingRequiredFields.includes(field.name) ? "error" : "none"}
              validationMessage={missingRequiredFields.includes(field.name) ? "Choose an option." : undefined}
            >
              <Dropdown
                ref={missingChoiceName === field.name ? invalidChoiceRef : undefined}
                placeholder="Select an option"
                value={selectedLabel}
                multiselect={field.type === "multiChoice"}
                selectedOptions={selectedOptions}
                onOptionSelect={(_, data) => onChange(field.type === "multiChoice"
                  ? data.selectedOptions
                  : data.optionValue)}
              >
                {options.map((option) => <Option key={option.value} value={option.value}>{option.label}</Option>)}
              </Dropdown>
            </Field>
          );
        }
        return (
          <Field key={field.name} label={field.title} required={field.required}>
            {field.type === "multiline" ? (
              <Textarea required={field.required} value={value === undefined ? "" : String(value)} onChange={(_, data) => onChange(data.value)} />
            ) : (
              <Input
                required={field.required}
                type={field.type === "number" ? "number" : field.type === "date" ? "date" : "text"}
                value={value === undefined ? "" : String(value)}
                onChange={(_, data) => onChange(field.type === "number" && data.value !== "" ? Number(data.value) : data.value)}
              />
            )}
          </Field>
        );
        })}
        {missingRequiredFields.length > 0 ? <MessageBar intent="error" role="status"><MessageBarBody>Complete the required fields: {missingRequiredFields.map((name) => props.fields.find((field) => field.name === name)?.title ?? name).join(", ")}.</MessageBarBody></MessageBar> : null}
        {error ? <MessageBar intent="error" role="alert"><MessageBarBody>Could not save this item. Review the item and try again.</MessageBarBody></MessageBar> : null}
        <Button type="submit" appearance="primary" disabled={saving}>{saving ? "Saving…" : props.submitLabel ?? "Save"}</Button>
      </form>
      {props.attachments}
    </>
  );
}

function formatValue(value: unknown): string {
  if (value === null || value === undefined || value === "") return "—";
  if (typeof value === "object") return JSON.stringify(value);
  return String(value);
}

function formatFieldValue(value: unknown, field: FieldDefinition): string {
  if (field.type === "person" || field.type === "lookup") {
    const option = (field.options ?? []).find((candidate) => String(candidate.value) === String(value));
    if (option) return option.label;
  }
  return formatValue(value);
}
