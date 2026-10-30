import type { GraphRequestOptions } from "./index.js";
import type { GraphListItem, LocalStore } from "./index.js";
import type { SharePointRestTransport } from "./spfx.js";

const listPath = (listId: string) => {
  const id = listId.replace(/^\{|\}$/g, "");
  if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) {
    throw new TypeError("SharePoint REST list IDs must be GUIDs.");
  }
  return `/_api/web/lists(guid'${id}')`;
};

export interface SharePointSearchOptions {
  rowLimit?: number;
  startRow?: number;
  selectProperties?: string[];
  refiners?: string[];
  refinementFilters?: string[];
  sortList?: string;
  sourceId?: string;
  queryTemplate?: string;
  clientType?: string;
  enableFql?: boolean;
  trimDuplicates?: boolean;
}

export interface SharePointRoleAssignment {
  PrincipalId?: number;
  Member?: {
    Id?: number;
    Title?: string;
    LoginName?: string;
    Email?: string;
    PrincipalType?: number;
    [key: string]: unknown;
  };
  RoleDefinitionBindings?: Array<{
    Id?: number;
    Name?: string;
    Description?: string;
    RoleTypeKind?: number;
    [key: string]: unknown;
  }>;
  [key: string]: unknown;
}

export interface SharePointRoleDefinition {
  Id: number;
  Name: string;
  Description?: string;
  RoleTypeKind?: number;
  [key: string]: unknown;
}

export interface SharePointPermissionChangeOptions {
  /** Set only after the app has shown the target list and proposed access change to the user. */
  confirmed: boolean;
  signal?: AbortSignal;
}

export interface SharePointListPermissionScope {
  Id?: string;
  Title?: string;
  HasUniqueRoleAssignments: boolean;
}

export interface SharePointUser {
  Id: number;
  Title?: string;
  Email?: string;
  LoginName?: string;
  PrincipalType?: number;
  IsSiteAdmin?: boolean;
  [key: string]: unknown;
}

export interface SharePointGroup {
  Id: number;
  Title?: string;
  Description?: string;
  OwnerTitle?: string;
  OnlyAllowMembersViewMembership?: boolean;
  AllowMembersEditMembership?: boolean;
  [key: string]: unknown;
}

export interface SharePointPersonProperties {
  AccountName?: string;
  DisplayName?: string;
  Email?: string;
  IsFollowed?: boolean;
  PictureUrl?: string;
  UserProfileProperties?: Array<{ Key: string; Value: string; [key: string]: unknown }>;
  [key: string]: unknown;
}

export interface SharePointListView {
  Id?: string;
  Title?: string;
  ViewType?: string;
  PersonalView?: boolean;
  DefaultView?: boolean;
  Hidden?: boolean;
  ServerRelativeUrl?: string;
  ViewQuery?: string;
  RowLimit?: number;
  Paged?: boolean;
  [key: string]: unknown;
}

export interface SharePointNavigationNode {
  Id: number;
  Title?: string;
  Url?: string;
  IsExternal?: boolean;
  IsVisible?: boolean;
  Children?: SharePointNavigationNode[];
  [key: string]: unknown;
}

export interface SharePointRecycleBinItem {
  Id: string;
  Title?: string;
  ItemType?: number;
  ItemState?: number;
  LeafName?: string;
  DirName?: string;
  DeletedDate?: string;
  DeletedByName?: string;
  AuthorName?: string;
  Size?: number;
  [key: string]: unknown;
}

export interface SharePointRegionalSettings {
  LocaleId?: number;
  CalendarType?: number;
  Time24?: boolean;
  FirstDayOfWeek?: number;
  FirstWeekOfYear?: number;
  WorkDays?: number;
  WorkDayStartHour?: number;
  WorkDayEndHour?: number;
  DateFormat?: number;
  DateSeparator?: string;
  TimeZone?: { Id?: number; Description?: string; [key: string]: unknown };
  [key: string]: unknown;
}

export interface SharePointChangeQuery {
  Activity?: boolean;
  Add?: boolean;
  Alert?: boolean;
  ChangeTokenEnd?: { StringValue: string };
  ChangeTokenStart?: { StringValue: string };
  ContentType?: boolean;
  DeleteObject?: boolean;
  FetchLimit?: number | string;
  Field?: boolean;
  File?: boolean;
  Folder?: boolean;
  Group?: boolean;
  GroupMembershipAdd?: boolean;
  GroupMembershipDelete?: boolean;
  Item?: boolean;
  LatestFirst?: boolean;
  List?: boolean;
  Move?: boolean;
  Navigation?: boolean;
  RecursiveAll?: boolean;
  Rename?: boolean;
  RequireSecurityTrim?: boolean;
  Restore?: boolean;
  RoleAssignmentAdd?: boolean;
  RoleAssignmentDelete?: boolean;
  RoleDefinitionAdd?: boolean;
  RoleDefinitionDelete?: boolean;
  RoleDefinitionUpdate?: boolean;
  SecurityPolicy?: boolean;
  Site?: boolean;
  SystemUpdate?: boolean;
  Update?: boolean;
  User?: boolean;
  View?: boolean;
  Web?: boolean;
}

export interface SharePointPage<T> {
  items: T[];
  nextCursor?: string;
}

function quoteSearchValue(value: string): string {
  return `'${value.replaceAll("'", "''")}'`;
}

function xmlEscape(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function parseChangeXml(xml: string) {
  if (typeof DOMParser === "undefined") throw new Error("The SharePoint change-token adapter requires a browser DOMParser.");
  const document = new DOMParser().parseFromString(xml, "application/xml");
  if (document.querySelector("parsererror")) throw new SyntaxError("SharePoint returned invalid change-token XML.");
  const rows = Array.from(document.getElementsByTagName("z:row"));
  const items: GraphListItem[] = rows.flatMap((row) => {
    const id = row.getAttribute("ows_ID") ?? row.getAttribute("ows_Id");
    if (!id) return [];
    const fields: Record<string, unknown> = {};
    for (const attribute of Array.from(row.attributes)) {
      if (attribute.name.startsWith("ows_") && attribute.name.toLocaleLowerCase() !== "ows_id") {
        fields[attribute.name.slice(4)] = attribute.value;
      }
    }
    return [{ id, fields }];
  });
  const changes = document.getElementsByTagName("Changes")[0];
  const deletedIds = Array.from(document.getElementsByTagName("Id"))
    .filter((node) => node.getAttribute("ChangeType") === "Delete")
    .map((node) => node.textContent?.trim())
    .filter((id): id is string => Boolean(id));
  const reset = Array.from(document.getElementsByTagName("Id")).some((node) => ["InvalidToken", "Restore"].includes(node.getAttribute("ChangeType") ?? ""));
  const position = document.querySelector("[ListItemCollectionPositionNext]")?.getAttribute("ListItemCollectionPositionNext") ?? undefined;
  return { items, deletedIds, reset, token: changes?.getAttribute("LastChangeToken") ?? undefined, position };
}

/** Uses SharePoint's legacy XML change token as a fallback when Graph delta is unavailable. */
export async function syncListWithChangeToken(options: {
  rest: SharePointRestTransport;
  store: LocalStore;
  scope: string;
  siteId: string;
  listId: string;
  fields?: string[];
  rowLimit?: number;
  maxPages?: number;
}): Promise<{ pages: number; changed: number; token: string }> {
  const cursorId = `rest:${options.listId}`;
  const endpoint = `${listPath(options.listId)}/GetListItemChangesSinceToken`;
  const rowLimit = Math.min(Math.max(options.rowLimit ?? 500, 1), 2000);
  const viewFields = ["ID", ...(options.fields ?? [])].map((name) => `<FieldRef Name="${xmlEscape(name)}" />`).join("");
  let cursor = await options.store.getCursor(options.scope, options.siteId, cursorId);
  let resetAttempted = false;

  while (true) {
    let pages = 0;
    let changed = 0;
    let nextPosition: string | undefined;
    let finalToken: string | undefined;
    try {
      do {
        if (++pages > (options.maxPages ?? 1000)) throw new Error("Change-token sync exceeded maxPages; cursor was not advanced.");
        const queryOptions = nextPosition
          ? `<QueryOptions><Paging ListItemCollectionPositionNext="${xmlEscape(nextPosition)}" /></QueryOptions>`
          : "<QueryOptions><DateInUtc>TRUE</DateInUtc></QueryOptions>";
        const response = await options.rest.raw(endpoint, {
          method: "POST",
          retryable: true,
          headers: { Accept: "application/xml", "Content-Type": "application/json;odata=verbose" },
          body: { query: {
            __metadata: { type: "SP.ChangeLogItemQuery" },
            ChangeToken: cursor ?? null,
            ViewName: "",
            Query: "<Query />",
            ViewFields: `<ViewFields>${viewFields}</ViewFields>`,
            RowLimit: String(rowLimit),
            QueryOptions: queryOptions,
          } },
        });
        const parsed = parseChangeXml(await response.text());
        if (!resetAttempted && cursor && parsed.reset) throw Object.assign(new Error("SharePoint change token is invalid and requires a full refresh."), { resetToken: true });
        const latest = new Map<string, GraphListItem>();
        parsed.items.forEach((item) => latest.set(item.id, item));
        const items = [...latest.values()];
        if (items.length) await options.store.putItems(options.scope, options.siteId, options.listId, items);
        if (parsed.deletedIds.length) await options.store.deleteItems(options.scope, options.siteId, options.listId, parsed.deletedIds);
        changed += items.length + parsed.deletedIds.length;
        finalToken = parsed.token ?? finalToken;
        nextPosition = parsed.position;
      } while (nextPosition);

      if (!finalToken) throw new Error("SharePoint change response did not include LastChangeToken.");
      await options.store.setCursor(options.scope, options.siteId, cursorId, finalToken);
      return { pages, changed, token: finalToken };
    } catch (error) {
      if (!resetAttempted && cursor && error && typeof error === "object" && "resetToken" in error) {
        resetAttempted = true;
        cursor = undefined;
        await options.store.clearCursor(options.scope, options.siteId, cursorId);
        await options.store.clearItems(options.scope, options.siteId, options.listId);
        continue;
      }
      throw error;
    }
  }
}

/** Narrow wrappers for older REST/CAML surfaces; raw REST remains available for uncovered endpoints. */
export function createLegacySharePointApi(rest: SharePointRestTransport, webUrl = rest.siteUrl) {
  const collectionPage = async <T>(endpoint: string, select: string, cursor?: string, pageSize = 50, expand?: string): Promise<SharePointPage<T>> => {
    if (!Number.isSafeInteger(pageSize) || pageSize < 1) throw new RangeError("SharePoint pageSize must be a positive integer.");
    const size = Math.min(pageSize, 100);
    const base = new URL(endpoint.replace(/^\/+/, ""), `${webUrl?.replace(/\/+$/, "") ?? "https://invalid.local"}/`);
    if (cursor) {
      if (!webUrl) throw new TypeError("A hosting web URL is required to validate SharePoint continuation links.");
      const next = new URL(cursor, `${webUrl.replace(/\/+$/, "")}/`);
      if (next.protocol !== "https:" || next.origin !== base.origin || next.pathname !== base.pathname) {
        throw new TypeError("SharePoint continuation links must remain on the same collection endpoint and web.");
      }
      base.search = next.search;
    } else {
      base.searchParams.set("$top", String(size));
      base.searchParams.set("$select", select);
      if (expand) base.searchParams.set("$expand", expand);
    }
    const requestUrl = webUrl ? base.toString() : `${base.pathname}${base.search}`;
    const response = await rest.get(requestUrl) as {
      value?: T[]; "@odata.nextLink"?: string;
      d?: { results?: T[]; __next?: string };
    };
    const items = response?.value ?? response?.d?.results;
    if (!Array.isArray(items)) throw new TypeError("SharePoint collection response must contain a value collection.");
    return { items, nextCursor: response["@odata.nextLink"] ?? response.d?.__next };
  };

  const listRoleAssignmentsPage = (listId: string, options: { cursor?: string; pageSize?: number } = {}) => collectionPage<SharePointRoleAssignment>(
    `${listPath(listId)}/roleassignments`,
    "PrincipalId,Member/Id,Member/Title,Member/LoginName,Member/Email,Member/PrincipalType,RoleDefinitionBindings/Id,RoleDefinitionBindings/Name,RoleDefinitionBindings/Description,RoleDefinitionBindings/RoleTypeKind",
    options.cursor,
    options.pageSize,
    "Member,RoleDefinitionBindings",
  );

  const viewPath = (listId: string, viewId: string) => {
    const id = viewId.replace(/^\{|\}$/g, "");
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)) throw new TypeError("SharePoint view IDs must be GUIDs.");
    return `${listPath(listId)}/views(guid'${id}')`;
  };

  const recycleBinPath = (scope: "web" | "site" = "web") => `/_api/${scope}/RecycleBin`;
  const recycleBinItemPath = (id: string, scope: "web" | "site" = "web") => {
    const itemId = id.replace(/^\{|\}$/g, "");
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(itemId)) {
      throw new TypeError("SharePoint Recycle Bin item IDs must be GUIDs.");
    }
    return `${recycleBinPath(scope)}('${itemId}')`;
  };

  const numericGroupId = (groupId: number) => {
    if (!Number.isSafeInteger(groupId) || groupId < 1) throw new TypeError("SharePoint group IDs must be positive integers.");
    return groupId;
  };
  const positiveId = (id: number, label: string) => {
    if (!Number.isSafeInteger(id) || id < 1) throw new TypeError(`${label} must be a positive integer.`);
    return id;
  };
  const requirePermissionConfirmation = (options: SharePointPermissionChangeOptions) => {
    if (options.confirmed !== true) throw new Error("Confirm this SharePoint permission change with the user before calling it.");
  };
  const profileAccountAlias = (accountName: string) => {
    if (!accountName.trim() || accountName.length > 512) throw new TypeError("SharePoint profile account name must be a non-empty value up to 512 characters.");
    return new URLSearchParams({ "@v": quoteSearchValue(accountName) }).toString();
  };
  const profileProperties = (value: unknown): SharePointPersonProperties => {
    const person = (value as { d?: unknown } | undefined)?.d ?? value;
    if (!person || typeof person !== "object" || Array.isArray(person)) throw new TypeError("SharePoint profile response must be an object.");
    return person as SharePointPersonProperties;
  };
  const navigationNodes = (value: unknown, includeChildren: boolean): SharePointNavigationNode[] => {
    const collection = (value as { d?: unknown } | undefined)?.d ?? value;
    const items = Array.isArray(collection)
      ? collection
      : (collection as { results?: unknown; value?: unknown } | undefined)?.results
        ?? (collection as { value?: unknown } | undefined)?.value;
    if (!Array.isArray(items)) throw new TypeError("SharePoint navigation response must contain a node collection.");
    return items.map((item) => {
      if (!item || typeof item !== "object" || Array.isArray(item) || !Number.isSafeInteger((item as { Id?: unknown }).Id)) {
        throw new TypeError("SharePoint navigation node must include a numeric Id.");
      }
      const node = item as SharePointNavigationNode & { Children?: unknown };
      if (includeChildren) return { ...node, Children: navigationNodes(node.Children ?? [], false) };
      const { Children: _children, ...withoutChildren } = node;
      void _children;
      return withoutChildren;
    });
  };
  const updateListProperties = (listId: string, properties: Record<string, unknown>, etag = "*") => rest.request(listPath(listId), {
    method: "POST",
    headers: { "Content-Type": "application/json;odata=verbose", "X-HTTP-Method": "MERGE", "If-Match": etag },
    body: { __metadata: { type: "SP.List" }, ...properties },
  });

  const attachmentPath = (listId: string, itemId: string | number) => {
    const id = String(itemId);
    if (!/^\d+$/.test(id)) throw new TypeError("SharePoint list item IDs must be numeric.");
    return `${listPath(listId)}/items(${id})/AttachmentFiles`;
  };
  const attachmentNamePath = (listId: string, itemId: string | number, fileName: string) => {
    if (!fileName.trim() || fileName.includes("/") || fileName.includes("\\")) throw new TypeError("Attachment name must be a non-empty leaf name.");
    const escaped = encodeURIComponent(fileName.replaceAll("'", "''"));
    return `${attachmentPath(listId, itemId)}/getByFileName('${escaped}')`;
  };

  return {
    updateListProperties,

    /** Read regional settings for the hosting web; the values are not cached. */
    async getRegionalSettings(): Promise<SharePointRegionalSettings> {
      const response = await rest.get("/_api/web/RegionalSettings") as SharePointRegionalSettings | { d?: SharePointRegionalSettings };
      const settings = response && typeof response === "object" && "d" in response ? response.d : response;
      if (!settings || typeof settings !== "object" || Array.isArray(settings)) throw new TypeError("SharePoint regional settings response must be an object.");
      return settings as SharePointRegionalSettings;
    },

    /** Read bounded pages from the hosting web or site collection recycle bin. */
    getRecycleBinItemsPage(options: { scope?: "web" | "site"; cursor?: string; pageSize?: number } = {}) {
      const select = "Id,Title,ItemType,ItemState,LeafName,DirName,DeletedDate,DeletedByName,AuthorName,Size";
      return collectionPage<SharePointRecycleBinItem>(recycleBinPath(options.scope), select, options.cursor, options.pageSize);
    },

    /** Restore one item to its original location. */
    restoreRecycleBinItem(id: string, options: { scope?: "web" | "site"; signal?: AbortSignal } = {}) {
      return rest.request(`${recycleBinItemPath(id, options.scope)}/restore()`, { method: "POST", signal: options.signal });
    },

    /** Permanently remove one recycle-bin item. Confirm with the user before calling. */
    permanentlyDeleteRecycleBinItem(id: string, options: { scope?: "web" | "site"; signal?: AbortSignal } = {}) {
      return rest.request(`${recycleBinItemPath(id, options.scope)}/deleteObject()`, { method: "POST", signal: options.signal });
    },

    /** Current signed-in principal as represented by the hosting SharePoint web. */
    async getCurrentUser(): Promise<SharePointUser> {
      const response = await rest.get("/_api/web/currentuser?$select=Id,Title,Email,LoginName,PrincipalType,IsSiteAdmin") as SharePointUser;
      if (!response || !Number.isSafeInteger(response.Id)) throw new TypeError("SharePoint current-user response did not include a numeric Id.");
      return response;
    },

    getSiteUsersPage(options: { cursor?: string; pageSize?: number } = {}) {
      return collectionPage<SharePointUser>("/_api/web/siteusers", "Id,Title,Email,LoginName,PrincipalType,IsSiteAdmin", options.cursor, options.pageSize);
    },

    getSiteGroupsPage(options: { cursor?: string; pageSize?: number } = {}) {
      return collectionPage<SharePointGroup>("/_api/web/sitegroups", "Id,Title,Description,OwnerTitle,OnlyAllowMembersViewMembership,AllowMembersEditMembership", options.cursor, options.pageSize);
    },

    getSiteGroupMembersPage(groupId: number, options: { cursor?: string; pageSize?: number } = {}) {
      return collectionPage<SharePointUser>(`/_api/web/sitegroups(${numericGroupId(groupId)})/users`, "Id,Title,Email,LoginName,PrincipalType,IsSiteAdmin", options.cursor, options.pageSize);
    },

    /** Read the signed-in user's SharePoint profile; profile values are not cached. */
    async getMyProperties(): Promise<SharePointPersonProperties> {
      return profileProperties(await rest.get("/_api/SP.UserProfiles.PeopleManager/GetMyProperties"));
    },

    /** Read a profile visible to the current user; accountName is a SharePoint claims login, not an Entra object ID. */
    async getPropertiesFor(accountName: string): Promise<SharePointPersonProperties> {
      return profileProperties(await rest.get(`/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?${profileAccountAlias(accountName)}`));
    },

    /** Read one visible SharePoint profile property without fetching or caching the full profile. */
    async getUserProfilePropertyFor(accountName: string, propertyName: string): Promise<string> {
      if (!/^[A-Za-z0-9_.-]{1,128}$/.test(propertyName)) throw new TypeError("SharePoint profile property name must contain only letters, numbers, dots, underscores, or hyphens.");
      const query = profileAccountAlias(accountName);
      const property = encodeURIComponent(quoteSearchValue(propertyName));
      const response = await rest.get(`/_api/SP.UserProfiles.PeopleManager/GetUserProfilePropertyFor(accountName=@v,propertyName=${property})?${query}`) as string | { d?: unknown };
      const value = typeof response === "string" ? response : response?.d;
      if (typeof value !== "string") throw new TypeError("SharePoint profile property response must be a string.");
      return value;
    },

    /** Read a bounded page of public list views and the current user's personal views. */
    getListViewsPage(listId: string, options: { cursor?: string; pageSize?: number } = {}) {
      return collectionPage<SharePointListView>(`${listPath(listId)}/views`, "Id,Title,ViewType,PersonalView,DefaultView,Hidden,ServerRelativeUrl,ViewQuery,RowLimit,Paged", options.cursor, options.pageSize);
    },

    /** Read one page of web or root-web navigation nodes. Child nodes expand one level when requested. */
    async getNavigationNodesPage(options: {
      scope?: "web" | "site";
      bar?: "quickLaunch" | "topNavigationBar";
      includeChildren?: boolean;
      cursor?: string;
      pageSize?: number;
    } = {}): Promise<SharePointPage<SharePointNavigationNode>> {
      const scope = options.scope ?? "web";
      const bar = options.bar ?? "quickLaunch";
      if (scope !== "web" && scope !== "site") throw new TypeError("Navigation scope must be web or site.");
      if (bar !== "quickLaunch" && bar !== "topNavigationBar") throw new TypeError("Navigation bar must be quickLaunch or topNavigationBar.");
      const root = scope === "web" ? "/_api/web/Navigation" : "/_api/site/RootWeb/Navigation";
      const collection = bar === "quickLaunch" ? "QuickLaunch" : "TopNavigationBar";
      const endpoint = `${root}/${collection}`;
      const includeChildren = options.includeChildren ?? false;
      const select = includeChildren
        ? "Id,Title,Url,IsExternal,IsVisible,Children/Id,Children/Title,Children/Url,Children/IsExternal,Children/IsVisible"
        : "Id,Title,Url,IsExternal,IsVisible";
      const page = await collectionPage<unknown>(endpoint, select, options.cursor, options.pageSize, includeChildren ? "Children" : undefined);
      return { items: navigationNodes(page.items, includeChildren), nextCursor: page.nextCursor };
    },

    /** Read the selected internal field names for one list view. */
    async getListViewFields(listId: string, viewId: string): Promise<string[]> {
      const response = await rest.get(`${viewPath(listId, viewId)}/ViewFields`) as {
        Items?: unknown; value?: unknown; d?: { Items?: unknown };
      };
      const fields = response?.Items ?? response?.value ?? response?.d?.Items;
      if (!Array.isArray(fields) || fields.some((field) => typeof field !== "string")) {
        throw new TypeError("SharePoint view fields response must contain an array of field names.");
      }
      return fields;
    },

    /** Read one bounded page of role bindings exposed on this list. */
    getListRoleAssignmentsPage(listId: string, options: { cursor?: string; pageSize?: number } = {}) {
      return listRoleAssignmentsPage(listId, options);
    },

    /** Read all role bindings on one list, following at most 100 pages by default. Inherited and item-level access is not included. */
    async getListRoleAssignments(listId: string, options: { maxPages?: number } = {}): Promise<SharePointRoleAssignment[]> {
      const maxPages = options.maxPages ?? 100;
      if (!Number.isSafeInteger(maxPages) || maxPages < 1 || maxPages > 1000) throw new RangeError("maxPages must be an integer from 1 to 1000.");
      const assignments: SharePointRoleAssignment[] = [];
      let cursor: string | undefined;
      for (let pageNumber = 0; pageNumber < maxPages; pageNumber++) {
        const page = await listRoleAssignmentsPage(listId, { cursor, pageSize: 100 });
        assignments.push(...page.items);
        cursor = page.nextCursor;
        if (!cursor) return assignments;
      }
      throw new Error("SharePoint role assignment read exceeded maxPages.");
    },

    async getListPermissionScope(listId: string): Promise<SharePointListPermissionScope> {
      const response = await rest.get(`${listPath(listId)}?$select=Id,Title,HasUniqueRoleAssignments`) as {
        Id?: string;
        Title?: string;
        HasUniqueRoleAssignments?: boolean;
      };
      if (typeof response?.HasUniqueRoleAssignments !== "boolean") throw new TypeError("SharePoint permission scope response did not include HasUniqueRoleAssignments.");
      return { Id: response.Id, Title: response.Title, HasUniqueRoleAssignments: response.HasUniqueRoleAssignments };
    },

    /** Read role definitions available in the hosting web. */
    getRoleDefinitionsPage(options: { cursor?: string; pageSize?: number } = {}) {
      return collectionPage<SharePointRoleDefinition>("/_api/web/roledefinitions", "Id,Name,Description,RoleTypeKind", options.cursor, options.pageSize);
    },

    /** Break inheritance on one list, keeping child-item scopes unchanged. Confirm with the user first. */
    async breakListRoleInheritance(listId: string, options: SharePointPermissionChangeOptions & { copyRoleAssignments?: boolean }) {
      requirePermissionConfirmation(options);
      const copy = options.copyRoleAssignments ?? true;
      return rest.request(`${listPath(listId)}/breakroleinheritance(copyRoleAssignments=${copy},clearSubscopes=false)`, { method: "POST", signal: options.signal });
    },

    /** Restore parent inheritance for one list, removing its unique assignments. Confirm with the user first. */
    async resetListRoleInheritance(listId: string, options: SharePointPermissionChangeOptions) {
      requirePermissionConfirmation(options);
      return rest.request(`${listPath(listId)}/resetroleinheritance()`, { method: "POST", signal: options.signal });
    },

    /** Add one principal's role binding to a list. Confirm the list, principal, and role with the user first. */
    async addListRoleAssignment(listId: string, principalId: number, roleDefinitionId: number, options: SharePointPermissionChangeOptions) {
      requirePermissionConfirmation(options);
      const principal = positiveId(principalId, "SharePoint principal ID");
      const role = positiveId(roleDefinitionId, "SharePoint role definition ID");
      return rest.request(`${listPath(listId)}/roleassignments/addroleassignment(principalid=${principal},roledefid=${role})`, { method: "POST", signal: options.signal });
    },

    /** Remove all role bindings for one principal from a list. Confirm the resulting access with the user first. */
    async removeListRoleAssignment(listId: string, principalId: number, options: SharePointPermissionChangeOptions) {
      requirePermissionConfirmation(options);
      const principal = positiveId(principalId, "SharePoint principal ID");
      return rest.request(`${listPath(listId)}/roleassignments/getbyprincipalid(${principal})`, {
        method: "POST", headers: { "X-HTTP-Method": "DELETE" }, signal: options.signal,
      });
    },

    search(queryText: string, options: SharePointSearchOptions = {}) {
      if (!queryText.trim()) throw new TypeError("SharePoint Search query text must not be empty.");
      const rowLimit = options.rowLimit ?? 50;
      const startRow = options.startRow ?? 0;
      if (!Number.isInteger(rowLimit) || rowLimit < 1 || rowLimit > 500) throw new RangeError("Search rowLimit must be an integer from 1 to 500.");
      if (!Number.isInteger(startRow) || startRow < 0 || startRow > 50_000) throw new RangeError("Search startRow must be an integer from 0 to 50000.");
      const parameters = new URLSearchParams({ querytext: quoteSearchValue(queryText), rowlimit: String(rowLimit), startrow: String(startRow) });
      if (options.selectProperties?.length) parameters.set("selectproperties", quoteSearchValue(options.selectProperties.join(",")));
      if (options.refiners?.length) parameters.set("refiners", quoteSearchValue(options.refiners.join(",")));
      if (options.refinementFilters?.length) parameters.set("refinementfilters", quoteSearchValue(options.refinementFilters.join(",")));
      if (options.sortList) parameters.set("sortlist", quoteSearchValue(options.sortList));
      if (options.sourceId) parameters.set("sourceid", quoteSearchValue(options.sourceId));
      if (options.queryTemplate) parameters.set("querytemplate", quoteSearchValue(options.queryTemplate));
      if (options.clientType) parameters.set("clienttype", quoteSearchValue(options.clientType));
      if (options.enableFql !== undefined) parameters.set("enablefql", String(options.enableFql));
      if (options.trimDuplicates !== undefined) parameters.set("trimduplicates", String(options.trimDuplicates));
      return rest.get(`/_api/search/query?${parameters.toString()}`);
    },

    listItemAttachments(listId: string, itemId: string | number) {
      return rest.get(attachmentPath(listId, itemId));
    },

    async readItemAttachment(listId: string, itemId: string | number, fileName: string) {
      const response = await rest.raw(`${attachmentNamePath(listId, itemId, fileName)}/$value`, { responseType: "blob" });
      return response.blob();
    },

    async addItemAttachment(listId: string, itemId: string | number, fileName: string, content: BodyInit) {
      const path = `${attachmentPath(listId, itemId)}/add(FileName='${encodeURIComponent(fileName.replaceAll("'", "''"))}')`;
      if (!fileName.trim() || fileName.includes("/") || fileName.includes("\\")) throw new TypeError("Attachment name must be a non-empty leaf name.");
      const response = await rest.raw(path, { method: "POST", rawBody: content, headers: { Accept: "application/json;odata=nometadata", "Content-Type": "application/octet-stream" } });
      return response.status === 204 ? undefined : response.json();
    },

    async updateItemAttachment(listId: string, itemId: string | number, fileName: string, content: BodyInit, etag = "*") {
      const response = await rest.raw(`${attachmentNamePath(listId, itemId, fileName)}/$value`, {
        method: "POST",
        rawBody: content,
        headers: { "Content-Type": "application/octet-stream", "X-HTTP-Method": "PUT", "If-Match": etag },
      });
      return response.status === 204 ? undefined : response.json();
    },

    deleteItemAttachment(listId: string, itemId: string | number, fileName: string, etag = "*") {
      return rest.request(attachmentNamePath(listId, itemId, fileName), { method: "DELETE", headers: { "If-Match": etag } });
    },

    /** Adds or removes SharePoint's standard TargetTo field; EnableAudienceTargeting is not writable on SP.List via REST. */
    async setAudienceTargeting(listId: string, enabled: boolean, etag = "*") {
      const fieldsPath = `${listPath(listId)}/fields`;
      const internalName = "Target_x0020_Audiences";
      const response = await rest.get(`${fieldsPath}?$select=Id,InternalName,TypeAsString&$filter=InternalName eq '${internalName}'`) as {
        value?: Array<{ Id?: string; InternalName?: string; TypeAsString?: string }>;
        d?: { results?: Array<{ Id?: string; InternalName?: string; TypeAsString?: string }> };
      };
      const existing = response?.value ?? response?.d?.results ?? [];
      const audienceField = existing.find((field) => field.InternalName === internalName);
      if (audienceField && audienceField.TypeAsString !== "TargetTo") {
        throw new TypeError("The Target Audiences internal name exists but is not a SharePoint TargetTo field.");
      }
      if (enabled) {
        if (audienceField) return audienceField;
        const schema = `<Field ID="{61CBB965-1E04-4273-B658-EEDAA662F48D}" Type="TargetTo" Name="${internalName}" StaticName="${internalName}" DisplayName="Target Audiences" Required="FALSE" />`;
        return rest.request(`${fieldsPath}/CreateFieldAsXml`, {
          method: "POST",
          headers: { "Content-Type": "application/json;odata=verbose" },
          body: { parameters: {
            __metadata: { type: "SP.XmlSchemaFieldCreationInformation" },
            SchemaXml: schema,
            Options: 0,
          } },
        });
      }
      if (!audienceField) return undefined;
      if (!audienceField.Id || !/^[0-9a-f-]{36}$/i.test(audienceField.Id)) throw new TypeError("The Target Audiences field response did not include a valid field ID.");
      return rest.request(`${fieldsPath}(guid'${audienceField.Id}')`, { method: "DELETE", headers: { "If-Match": etag } });
    },

    getItemsByCaml(listId: string, viewXml: string) {
      return rest.request(`${listPath(listId)}/GetItems`, {
        method: "POST",
        retryable: true,
        headers: { "Content-Type": "application/json;odata=verbose" },
        body: {
          query: {
            __metadata: { type: "SP.CamlQuery" },
            ViewXml: viewXml,
          },
        },
      });
    },

    renderListDataAsStream(listId: string, viewXml: string, parameters?: Record<string, unknown>) {
      return rest.request(`${listPath(listId)}/RenderListDataAsStream`, {
        method: "POST",
        retryable: true,
        headers: { "Content-Type": "application/json;odata=verbose" },
        body: { parameters: { ViewXml: viewXml, ...parameters } },
      });
    },

    async getListItemChangesSinceToken(listId: string, query: Record<string, unknown> = {}) {
      const response = await rest.raw(`${listPath(listId)}/GetListItemChangesSinceToken`, {
        method: "POST",
        retryable: true,
        headers: { Accept: "application/xml", "Content-Type": "application/json;odata=verbose" },
        body: query,
      });
      const xml = await response.text();
      if (typeof DOMParser === "undefined") return xml;
      const document = new DOMParser().parseFromString(xml, "application/xml");
      if (document.querySelector("parsererror")) throw new SyntaxError("SharePoint returned invalid change-token XML.");
      return document;
    },

    getChanges(listId: string, changeQuery: SharePointChangeQuery) {
      const fetchLimit = changeQuery.FetchLimit;
      if (fetchLimit !== undefined && !(typeof fetchLimit === "string" && /^\d+$/.test(fetchLimit)) && !(typeof fetchLimit === "number" && Number.isSafeInteger(fetchLimit) && fetchLimit > 0)) {
        throw new RangeError("SharePoint FetchLimit must be a positive integer.");
      }
      return rest.request(`${listPath(listId)}/GetChanges`, {
        method: "POST",
        retryable: true,
        headers: { "Content-Type": "application/json;odata=verbose" },
        body: { query: { __metadata: { type: "SP.ChangeQuery" }, ...changeQuery, ...(fetchLimit === undefined ? {} : { FetchLimit: String(fetchLimit) }) } },
      });
    },

    syncListWithChangeToken,

    /** Use the raw method for CSOM-only operations; there is no CSOM runtime in this browser package. */
    raw(url: string, options?: GraphRequestOptions) {
      return rest.raw(url, options);
    },
  };
}

export { listPath as sharePointListPath };
