import type { ClientMetric, GraphGet, GraphRequest, GraphRequestOptions } from "./index.js";
import { createClient, type ClientOptions } from "./index.js";
import { createLegacySharePointApi } from "./sharepoint.js";

interface GraphBuilder {
  header?(name: string, value: string): GraphBuilder;
  responseType?(value: string): GraphBuilder;
  get(): Promise<unknown>;
  post(body?: unknown): Promise<unknown>;
  put(body?: unknown): Promise<unknown>;
  patch(body?: unknown): Promise<unknown>;
  delete(): Promise<unknown>;
}

/** Structural subset of SPFx MSGraphClientV3; avoids taking a runtime dependency on SPFx. */
export interface SPFxGraphClient {
  api(url: string): GraphBuilder;
}

export function createSPFxGraphTransport(
  client: SPFxGraphClient,
  transportOptions: { onRetry?: (delayMs: number, error: unknown) => void; onMetric?: (metric: ClientMetric) => void; maxRetries?: number; maxBackoffMs?: number } = {},
): { get: GraphGet; request: GraphRequest } {
  const request: GraphRequest = async (url, options: GraphRequestOptions = {}) => {
    const send = async () => {
      if (options.signal?.aborted) throw options.signal.reason ?? new DOMException("Aborted", "AbortError");
      let builder = client.api(url);
      for (const [name, value] of Object.entries(options.headers ?? {})) builder = builder.header?.(name, value) ?? builder;
      if (options.responseType) builder = builder.responseType?.(options.responseType) ?? builder;
      switch (options.method ?? "GET") {
        case "POST": return builder.post(options.rawBody ?? options.body);
        case "PUT": return builder.put(options.rawBody ?? options.body);
        case "PATCH": return builder.patch(options.rawBody ?? options.body);
        case "DELETE": return builder.delete();
        default: return builder.get();
      }
    };
    const onRetry = (delayMs: number, error: unknown) => {
      try { transportOptions.onRetry?.(delayMs, error); } catch { /* Retry instrumentation must not interrupt the request. */ }
      try { transportOptions.onMetric?.({ type: "retry", operation: "graph.request", delayMs, error }); } catch { /* Retry instrumentation must not interrupt the request. */ }
    };
    return options.method === undefined || options.method === "GET"
      ? withSharePointReadRetry(send, transportOptions.maxRetries ?? 3, onRetry, transportOptions.maxBackoffMs ?? 8000)
      : send();
  };
  return { get: (url) => request(url), request };
}

export interface SPFxHttpResponse {
  ok: boolean;
  status: number;
  statusText?: string;
  headers: { get(name: string): string | null };
  json(): Promise<unknown>;
  text(): Promise<string>;
  blob(): Promise<Blob>;
}

export interface SPFxHttpClient {
  fetch(url: string, configuration: any, options: RequestInit & { webUrl?: string }): Promise<SPFxHttpResponse>;
}

export interface SharePointRestTransport {
  /** Hosting SharePoint web URL used to validate continuation links. */
  siteUrl?: string;
  get(url: string, options?: Pick<GraphRequestOptions, "signal">): Promise<unknown>;
  request(url: string, options?: GraphRequestOptions): Promise<unknown>;
  raw(url: string, options?: GraphRequestOptions): Promise<SPFxHttpResponse>;
}

export interface DriveCopyOperation {
  monitorUrl: string;
}

export interface DriveCopyStatus {
  id?: string;
  status: "inProgress" | "completed" | "failed" | string;
  percentComplete?: number;
  percentageComplete?: number;
  resourceId?: string;
  resourceLocation?: string;
  error?: unknown;
}

/** Structural subset of WebPartContext needed to create authenticated services. */
export interface SPFxServiceContext {
  msGraphClientFactory: { getClient(version: "3"): Promise<SPFxGraphClient> };
  spHttpClient: SPFxHttpClient;
  pageContext: {
    site: { absoluteUrl: string };
    web: { absoluteUrl: string };
    user: { loginName: string };
  };
}

/**
 * Compose the signed-in SPFx Graph and REST clients once and share the resulting
 * cache, typed client, and long-tail SharePoint REST helpers across a web part.
 * Pass SPHttpClient.configurations.v1 from @microsoft/sp-http explicitly.
 */
export async function createSPFxServices(
  context: SPFxServiceContext,
  restConfiguration: unknown,
  options: Omit<ClientOptions, "scope" | "get" | "request"> & {
    scope?: string;
    transport?: { onRetry?: (delayMs: number, error: unknown) => void; maxRetries?: number; maxBackoffMs?: number };
  } = {},
) {
  const graph = await context.msGraphClientFactory.getClient("3");
  const scope = options.scope ?? `${new URL(context.pageContext.site.absoluteUrl).hostname}:${context.pageContext.user.loginName}`;
  const { transport, ...clientOptions } = options;
  const graphTransport = createSPFxGraphTransport(graph, { ...transport, onMetric: options.onMetric });
  const rest = createSPFxRestTransport(context.spHttpClient, restConfiguration, context.pageContext.web.absoluteUrl, {
    ...transport,
    onMetric: options.onMetric,
  });
  const client = createClient({ ...clientOptions, scope, ...graphTransport });
  return {
    client,
    graph,
    rest,
    sharepoint: createLegacySharePointApi(rest, context.pageContext.web.absoluteUrl),
    scope,
    async waitForDriveCopy(operation: DriveCopyOperation, config: {
      pollIntervalMs?: number;
      timeoutMs?: number;
      signal?: AbortSignal;
      onProgress?: (status: DriveCopyStatus) => void;
    } = {}): Promise<DriveCopyStatus> {
      const site = new URL(rest.siteUrl ?? context.pageContext.web.absoluteUrl);
      const monitor = new URL(operation.monitorUrl);
      if (monitor.origin !== site.origin || monitor.username || monitor.password || monitor.hash || !/_api\/v2\.0\/monitor\//.test(monitor.pathname)) throw new TypeError("Drive copy monitor URL must be a same-origin SharePoint monitor endpoint.");
      const pollIntervalMs = config.pollIntervalMs ?? 1000;
      const timeoutMs = config.timeoutMs ?? 120_000;
      if (!Number.isInteger(pollIntervalMs) || pollIntervalMs < 50 || pollIntervalMs > 10_000) throw new RangeError("pollIntervalMs must be an integer from 50 to 10000.");
      if (!Number.isInteger(timeoutMs) || timeoutMs < 1000 || timeoutMs > 600_000) throw new RangeError("timeoutMs must be an integer from 1000 to 600000.");
      const started = Date.now();
      while (true) {
        if (config.signal?.aborted) throw config.signal.reason ?? new DOMException("Aborted", "AbortError");
        if (Date.now() - started >= timeoutMs) throw new Error(`Drive copy did not finish within ${timeoutMs} ms.`);
        const status = await rest.get(monitor.toString(), { signal: config.signal }) as DriveCopyStatus;
        if (!status || typeof status.status !== "string") throw new TypeError("SharePoint copy monitor returned an invalid status.");
        try { config.onProgress?.(status); } catch { /* Progress UI must not interrupt polling. */ }
        if (status.status === "completed") return status;
        if (status.status === "failed") throw Object.assign(new Error("SharePoint file copy failed."), { operation: status });
        await new Promise<void>((resolve, reject) => {
          if (config.signal?.aborted) {
            reject(config.signal.reason ?? new DOMException("Aborted", "AbortError"));
            return;
          }
          const timer = setTimeout(() => {
            config.signal?.removeEventListener("abort", onAbort);
            resolve();
          }, pollIntervalMs);
          const onAbort = () => {
            clearTimeout(timer);
            config.signal?.removeEventListener("abort", onAbort);
            reject(config.signal?.reason ?? new DOMException("Aborted", "AbortError"));
          };
          config.signal?.addEventListener("abort", onAbort, { once: true });
        });
      }
    },
  };
}

function httpError(response: SPFxHttpResponse): Error & { status: number; statusCode: number; retryAfter?: string | null } {
  return Object.assign(new Error(`SharePoint request failed with HTTP ${response.status}${response.statusText ? ` ${response.statusText}` : ""}.`), {
    status: response.status,
    statusCode: response.status,
    retryAfter: response.headers.get("Retry-After") ?? response.headers.get("RateLimit-Reset"),
  });
}

function retryDelay(error: unknown, attempt: number, maxBackoffMs: number): number | undefined {
  if (error instanceof TypeError && /failed to fetch|networkerror|network request failed|load failed/i.test(error.message)) {
    return Math.min(500 * 2 ** attempt, maxBackoffMs);
  }
  if (!error || typeof error !== "object") return undefined;
  const value = error as { status?: number; statusCode?: number; retryAfter?: string | null; headers?: { get(name: string): string | null } };
  if (![429, 503].includes(Number(value.status ?? value.statusCode ?? 0))) return undefined;
  const header = value.retryAfter ?? value.headers?.get("Retry-After");
  if (header) {
    const seconds = Number(header);
    if (Number.isFinite(seconds)) return Math.max(0, seconds * 1000);
    const date = Date.parse(header);
    if (Number.isFinite(date)) return Math.max(0, date - Date.now());
  }
  return Math.min(1000 * 2 ** attempt, maxBackoffMs);
}

/** Retries only idempotent reads and honors Retry-After on SharePoint throttling responses. */
export async function withSharePointReadRetry<T>(
  operation: () => Promise<T>,
  maxRetries = 3,
  onRetry?: (delayMs: number, error: unknown) => void,
  maxBackoffMs = 8000,
): Promise<T> {
  if (!Number.isInteger(maxRetries) || maxRetries < 0) throw new RangeError("maxRetries must be a non-negative integer.");
  if (!Number.isFinite(maxBackoffMs) || maxBackoffMs < 0) throw new RangeError("maxBackoffMs must be a non-negative number.");
  for (let attempt = 0; ; attempt++) {
    try {
      return await operation();
    } catch (error) {
      const delay = retryDelay(error, attempt, maxBackoffMs);
      if (delay === undefined || attempt >= maxRetries) throw error;
      try { onRetry?.(delay, error); } catch { /* Retry instrumentation must not interrupt the request. */ }
      await new Promise((resolve) => setTimeout(resolve, delay));
    }
  }
}

/** Uses the SPFx host's cookie-authenticated client; configuration is SPHttpClient.configurations.v1. */
export function createSPFxRestTransport(
  client: SPFxHttpClient,
  configuration: unknown,
  siteUrl: string,
  transportOptions: { onRetry?: (delayMs: number, error: unknown) => void; onMetric?: (metric: ClientMetric) => void; maxRetries?: number; maxBackoffMs?: number } = {},
): SharePointRestTransport {
  const site = new URL(siteUrl);
  const origin = site.origin;
  site.pathname = `${site.pathname.replace(/\/+$/g, "")}/`;
  site.search = "";
  site.hash = "";

  const raw = async (url: string, options: GraphRequestOptions = {}): Promise<SPFxHttpResponse> => {
    // SPFx REST paths are relative to the current web. Resolving `/_api/...`
    // against the origin would silently drop a subsite path such as `/sites/team`.
    const target = /^https:\/\//i.test(url) ? new URL(url) : new URL(url.replace(/^\/+/, ""), site);
    if (target.protocol !== "https:" || target.origin !== origin) {
      throw new TypeError("SharePoint REST requests must remain on the configured site origin.");
    }
    const headers: Record<string, string> = { Accept: options.responseType === "blob" ? "*/*" : "application/json;odata=nometadata", ...options.headers };
    if (options.body !== undefined && options.rawBody === undefined && !Object.keys(headers).some((key) => key.toLowerCase() === "content-type")) {
      headers["Content-Type"] = "application/json;odata=nometadata";
    }
    const onRetry = (delayMs: number, error: unknown) => {
      try { transportOptions.onRetry?.(delayMs, error); } catch { /* Retry instrumentation must not interrupt the request. */ }
      try { transportOptions.onMetric?.({ type: "retry", operation: "sharepoint.rest", delayMs, error }); } catch { /* Retry instrumentation must not interrupt the request. */ }
    };
    return withSharePointReadRetry(async () => {
      const response = await client.fetch(target.toString(), configuration, {
        method: options.method ?? "GET",
        headers,
        body: options.rawBody ?? (options.body === undefined ? undefined : JSON.stringify(options.body)),
        credentials: "same-origin",
        signal: options.signal,
      });
      if (!response.ok) throw httpError(response);
      return response;
    }, options.method && options.method !== "GET" && !options.retryable ? 0 : transportOptions.maxRetries ?? 3, onRetry, transportOptions.maxBackoffMs ?? 8000);
  };

  return {
    siteUrl: site.toString().replace(/\/$/, ""),
    raw,
    async get(url, options) {
      const response = await raw(url, options);
      return response.status === 204 ? undefined : response.json();
    },
    async request(url, options) {
      const response = await raw(url, options);
      return response.status === 204 ? undefined : response.json();
    },
  };
}
