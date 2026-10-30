import { createClient, type GraphListItem } from "./index.js";

type Client = ReturnType<typeof createClient>;

/** Define a dependency-free, accessible list element for vanilla DOM/SPFx hosts. */
export function defineSPFxFastListElement(tagName = "spfx-fast-list") {
  if (!/^[a-z][a-z0-9]*(?:-[a-z0-9]+)+$/.test(tagName)) throw new TypeError("Custom element names must be lowercase and contain a hyphen.");
  if (typeof customElements === "undefined" || typeof HTMLElement === "undefined") throw new Error("Custom elements are available only in a browser.");
  const existing = customElements.get(tagName);
  if (existing) return existing;

  class SPFxFastListElement extends HTMLElement {
    static observedAttributes = ["site-id", "list-id", "title", "max-items", "fields"];
    #client?: Client;
    #generation = 0;
    #root: ShadowRoot;
    #style: HTMLStyleElement;

    get client(): Client | undefined { return this.#client; }

    set client(value: Client | undefined) {
      if (this.#client === value) return;
      this.#client = value;
      if (this.isConnected) void this.refresh();
    }

    constructor() {
      super();
      this.#root = this.attachShadow({ mode: "open" });
      this.#style = document.createElement("style");
      this.#style.textContent = `:host{display:block;color:var(--colorNeutralForeground1,#242424);font:14px/1.5 system-ui,sans-serif}h2{font-size:1.25rem;margin:0 0 12px}button{font:inherit;padding:6px 12px;border:1px solid var(--colorNeutralStroke1,#d1d1d1);border-radius:4px;background:var(--colorNeutralBackground1,#fff);color:inherit;cursor:pointer}button:focus-visible,a:focus-visible{outline:2px solid var(--colorBrandStroke1,#0f6cbd);outline-offset:2px}table{border-collapse:collapse;width:100%;margin-top:12px}th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--colorNeutralStroke2,#e0e0e0)}th{font-weight:600}a{color:var(--colorBrandForegroundLink,#115ea3)}`;
      this.#root.append(this.#style);
    }

    connectedCallback() { void this.refresh(); }
    attributeChangedCallback() { if (this.isConnected) void this.refresh(); }

    async refresh() {
      const generation = ++this.#generation;
      this.#root.replaceChildren(this.#style);
      const heading = document.createElement("h2");
      heading.textContent = this.getAttribute("title") ?? "SharePoint list";
      const status = document.createElement("p");
      status.setAttribute("role", "status");
      status.textContent = "Loading list items…";
      const content = document.createElement("div");
      const refreshButton = document.createElement("button");
      refreshButton.type = "button";
      refreshButton.textContent = "Refresh";
      refreshButton.addEventListener("click", () => void this.refresh());
      this.#root.append(heading, refreshButton, status, content);

      const siteId = this.getAttribute("site-id");
      const listId = this.getAttribute("list-id");
      if (!this.client || !siteId || !listId) {
        status.textContent = "Set the authenticated client property and site-id/list-id attributes.";
        return;
      }

      try {
        const maxItems = Math.min(Math.max(Number(this.getAttribute("max-items") ?? 100), 1), 500);
        const items: GraphListItem[] = [];
        for await (const item of this.client.listItems(siteId, listId, { maxItems, pageSize: 100 })) items.push(item);
        if (generation !== this.#generation) return;
        if (!items.length) {
          status.textContent = "No list items found.";
          return;
        }
        status.textContent = `${items.length} items loaded.`;
        const fields = (this.getAttribute("fields") ?? "Title,Status").split(",").map((field) => field.trim()).filter(Boolean);
        if (!fields.length) fields.push("Title");
        const table = document.createElement("table");
        const header = table.createTHead().insertRow();
        for (const name of fields) {
          const cell = document.createElement("th");
          cell.scope = "col";
          cell.textContent = name;
          header.append(cell);
        }
        const body = table.createTBody();
        for (const item of items) {
          const row = body.insertRow();
          for (const name of fields) {
            const cell = row.insertCell();
            const value = item.fields?.[name];
            appendSafeValue(cell, value);
          }
        }
        content.append(table);
      } catch {
        if (generation === this.#generation) status.textContent = "Could not load this list. Check list access and the SPFx host connection.";
      }
    }
  }

  customElements.define(tagName, SPFxFastListElement);
  return SPFxFastListElement;
}

const richTextTags = new Set(["A", "B", "BLOCKQUOTE", "BR", "EM", "H1", "H2", "H3", "H4", "H5", "H6", "I", "LI", "OL", "P", "SPAN", "STRONG", "U", "UL"]);
const discardedTags = new Set(["IFRAME", "OBJECT", "SCRIPT", "STYLE", "SVG", "TEMPLATE"]);

/** Preserve common SharePoint rich-text formatting while rebuilding nodes from a safe allow-list. */
function appendSafeValue(target: HTMLElement, value: unknown): void {
  if (value === null || value === undefined) return;
  const text = typeof value === "object" ? JSON.stringify(value) : String(value);
  if (!/<[a-z][^>]*>/i.test(text)) {
    target.textContent = text;
    return;
  }
  const parsed = new DOMParser().parseFromString(text, "text/html");
  for (const child of Array.from(parsed.body.childNodes)) {
    const safe = copySafeNode(child);
    if (safe) target.append(safe);
  }
}

function copySafeNode(node: Node): Node | undefined {
  if (node.nodeType === Node.TEXT_NODE) return document.createTextNode(node.textContent ?? "");
  if (!(node instanceof Element)) return undefined;
  if (discardedTags.has(node.tagName)) return undefined;
  if (!richTextTags.has(node.tagName)) {
    const fragment = document.createDocumentFragment();
    for (const child of Array.from(node.childNodes)) {
      const safe = copySafeNode(child);
      if (safe) fragment.append(safe);
    }
    return fragment;
  }
  const output = document.createElement(node.tagName.toLowerCase());
  if (node.tagName === "A") {
    const href = node.getAttribute("href");
    if (href && isSafeLink(href)) {
      output.setAttribute("href", href);
      output.setAttribute("rel", "noopener noreferrer");
    }
  }
  for (const child of Array.from(node.childNodes)) {
    const safe = copySafeNode(child);
    if (safe) output.append(safe);
  }
  return output;
}

function isSafeLink(href: string): boolean {
  try {
    const url = new URL(href, document.baseURI);
    return !url.username && !url.password && (url.protocol === "https:" || url.protocol === "http:" || url.protocol === "mailto:");
  } catch {
    return false;
  }
}
