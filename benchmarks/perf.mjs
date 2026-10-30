import "fake-indexeddb/auto";
import { createClient, createIndexedDbStore, createMemoryStore } from "../dist/index.js";

const samples = 5;
const now = () => performance.now();
async function measure(label, operation) {
  const runs = [];
  for (let index = 0; index < samples; index++) {
    const start = now();
    await operation();
    runs.push(now() - start);
  }
  runs.sort((a, b) => a - b);
  return { label, samples, medianMs: Number(runs[Math.floor(runs.length / 2)].toFixed(2)), minMs: Number(runs[0].toFixed(2)) };
}

async function measureUncachedQuery(label, createSampleClient, siteId, listId, query) {
  const clients = Array.from({ length: samples }, createSampleClient);
  return measure(label, async () => {
    const items = await clients.pop().queryItems(siteId, listId, query);
    if (items.length !== 50) throw new Error(`Expected 50 query results, received ${items.length}.`);
  });
}

const store = createMemoryStore();
let networkCalls = 0;
const client = createClient({
  scope: "synthetic-tenant:synthetic-user",
  store,
  get: async (url) => {
    networkCalls++;
    const page = Number(new URL(url).searchParams.get("page") ?? 0);
    return { value: Array.from({ length: 100 }, (_, index) => ({ id: `${page}-${index}`, fields: { Title: `Item ${page * 100 + index}` } })), "@odata.nextLink": page < 9 ? `https://graph.microsoft.com/v1.0/sites/site/lists/list/items?page=${page + 1}` : undefined };
  },
});

await client.readPage("https://graph.microsoft.com/v1.0/sites/site/lists/list/items?page=0");
const results = [];
results.push(await measure("warm cached page read", async () => {
  const page = await client.readPage("https://graph.microsoft.com/v1.0/sites/site/lists/list/items?page=0");
  await page.refresh;
}));
results.push(await measure("bounded pagination (max 1,000 items)", async () => {
  let count = 0;
  for await (const _item of client.listItems("site", "list", { pageSize: 100, maxItems: 1000 })) count++;
  if (count !== 1000) throw new Error(`Expected 1,000 paged items, received ${count}.`);
}));

const indexedDbName = `spfx-fast-benchmark-${Date.now()}`;
const indexedDbStore = createIndexedDbStore(indexedDbName);
const indexedDbClient = createClient({ scope: "synthetic-tenant:synthetic-user", store: indexedDbStore, get: async () => ({ value: [] }) });
for (const size of [10_000, 50_000, 100_000]) {
  const rows = Array.from({ length: size }, (_, index) => ({ id: String(index), fields: { Title: `Item ${index}`, Status: index % 2 ? "Open" : "Closed" } }));
  const memoryList = `memory-benchmark-${size}`;
  const indexedList = `indexeddb-benchmark-${size}`;
  await store.putItems("synthetic-tenant:synthetic-user", "site", memoryList, rows);
  await indexedDbStore.putItems("synthetic-tenant:synthetic-user", "site", indexedList, rows);
  console.log(`Loaded ${size.toLocaleString()} synthetic rows into Memory and fake IndexedDB.`);
  const query = { equals: { Status: "Open" }, sort: { field: "Title" }, limit: 50 };
  results.push(await measureUncachedQuery(`memory uncached local query (${size.toLocaleString()} rows, Status=Open, sort Title, limit 50)`, () => createClient({ scope: "synthetic-tenant:synthetic-user", store, get: async () => ({ value: [] }) }), "site", memoryList, query));
  results.push(await measureUncachedQuery(`IndexedDB uncached local query (${size.toLocaleString()} rows, Status=Open, sort Title, limit 50)`, () => createClient({ scope: "synthetic-tenant:synthetic-user", store: indexedDbStore, get: async () => ({ value: [] }) }), "site", indexedList, query));
  const boundedQuery = { equals: { Status: "Open" }, limit: 50 };
  results.push(await measureUncachedQuery(`memory uncached local query (${size.toLocaleString()} rows, Status=Open, limit 50)`, () => createClient({ scope: "synthetic-tenant:synthetic-user", store, get: async () => ({ value: [] }) }), "site", memoryList, boundedQuery));
  results.push(await measureUncachedQuery(`IndexedDB uncached cursor query (${size.toLocaleString()} rows, Status=Open, limit 50)`, () => createClient({ scope: "synthetic-tenant:synthetic-user", store: indexedDbStore, get: async () => ({ value: [] }) }), "site", indexedList, boundedQuery));
  console.log(`Measured local queries at ${size.toLocaleString()} rows.`);
}

let deltaBatch = 0;
const deltaClient = createClient({
  scope: "synthetic-tenant:synthetic-user",
  store: createMemoryStore(),
  get: async () => {
    deltaBatch++;
    return {
      value: Array.from({ length: 100 }, (_, index) => ({ id: `${deltaBatch}-${index}`, fields: { Title: `Delta item ${index}` } })),
      "@odata.deltaLink": `https://graph.microsoft.com/v1.0/sites/site/lists/delta-benchmark/items/delta?token=${deltaBatch}`,
    };
  },
});
await deltaClient.syncList("site", "delta-benchmark");
results.push(await measure("Graph delta sync batch (100 changed rows)", async () => {
  const result = await deltaClient.syncList("site", "delta-benchmark");
  if (result.changed !== 100 || result.pages !== 1) throw new Error("Expected one delta page with 100 changed rows.");
}));

console.log(JSON.stringify({ kind: "synthetic-only; fake IndexedDB is not a browser benchmark and mock Graph is not a tenant benchmark", timestamp: new Date().toISOString(), networkCalls, deltaBatches: deltaBatch, results }, null, 2));
