import { applyItemQuery, createClient, createIndexedDbStore, createMemoryStore } from "../../dist/index.js";

const samples = 5;

async function measure(label, operation) {
  const values = [];
  for (let index = 0; index < samples; index++) {
    const start = performance.now();
    const items = await operation();
    if (items.length !== 50) throw new Error(`${label} returned ${items.length} items, expected 50.`);
    values.push(performance.now() - start);
  }
  values.sort((a, b) => a - b);
  return { label, samples, medianMs: Number(values[Math.floor(samples / 2)].toFixed(3)), minMs: Number(values[0].toFixed(3)) };
}

async function measureWarm(label, operation) {
  const samples = 100;
  const start = performance.now();
  for (let index = 0; index < samples; index++) {
    const items = await operation();
    if (items.length !== 50) throw new Error(`${label} returned ${items.length} items, expected 50.`);
  }
  const totalMs = Number((performance.now() - start).toFixed(3));
  return { label, samples, totalMs, meanMs: Number((totalMs / samples).toFixed(4)) };
}

function boundedSort(items, limit = 50) {
  const top = [];
  for (const item of items) {
    if (item.fields?.Status !== "Open") continue;
    let low = 0;
    let high = top.length;
    while (low < high) {
      const middle = (low + high) >>> 1;
      if (item.fields.Title.localeCompare(top[middle].fields.Title) < 0) high = middle;
      else low = middle + 1;
    }
    if (low < limit) {
      top.splice(low, 0, item);
      if (top.length > limit) top.pop();
    }
  }
  return top;
}

async function cursorTopNSort(databaseName, key, limit = 50) {
  const openRequest = indexedDB.open(databaseName, 3);
  const db = await new Promise((resolve, reject) => {
    openRequest.onsuccess = () => resolve(openRequest.result);
    openRequest.onerror = () => reject(openRequest.error);
  });
  return new Promise((resolve, reject) => {
    const tx = db.transaction("items", "readonly");
    const request = tx.objectStore("items").index("listKey").openCursor(IDBKeyRange.only(key));
    const top = [];
    request.onsuccess = () => {
      const cursor = request.result;
      if (!cursor) return;
      const item = cursor.value.item;
      if (item.fields?.Status === "Open") {
        let low = 0;
        let high = top.length;
        while (low < high) {
          const middle = (low + high) >>> 1;
          if (String(item.fields.Title ?? "").localeCompare(String(top[middle].fields?.Title ?? "")) < 0) high = middle;
          else low = middle + 1;
        }
        if (low < limit) {
          top.splice(low, 0, item);
          if (top.length > limit) top.pop();
        }
      }
      cursor.continue();
    };
    request.onerror = () => reject(request.error);
    tx.oncomplete = () => resolve(top);
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error ?? new Error("Cursor query aborted."));
  });
}

globalThis.runSPFxFastBrowserBenchmark = async () => {
  const scope = "browser-benchmark:synthetic-user";
  const memoryStore = createMemoryStore();
  const databaseName = `spfx-fast-browser-benchmark-${crypto.randomUUID()}`;
  const indexedDbStore = createIndexedDbStore(databaseName);
  const memoryClient = createClient({ scope, store: memoryStore, get: async () => ({ value: [] }) });
  const indexedDbClient = createClient({ scope, store: indexedDbStore, get: async () => ({ value: [] }) });
  const makeUncachedClients = (store) => Array.from({ length: samples }, () => createClient({ scope, store, get: async () => ({ value: [] }) }));
  const results = [];

  for (const size of [10_000, 50_000, 100_000]) {
    const rows = Array.from({ length: size }, (_, index) => ({
      id: String(index),
      fields: { Title: `Item ${String(index).padStart(6, "0")}`, Status: index % 2 ? "Open" : "Closed" },
    }));
    const memoryList = `memory-${size}`;
    const indexedList = `indexeddb-${size}`;
    await memoryStore.putItems(scope, "site", memoryList, rows);
    await indexedDbStore.putItems(scope, "site", indexedList, rows);

    const sorted = { equals: { Status: "Open" }, sort: { field: "Title" }, limit: 50 };
    const bounded = { equals: { Status: "Open" }, limit: 50 };
    const uncachedMemorySorted = makeUncachedClients(memoryStore);
    const uncachedIndexedSorted = makeUncachedClients(indexedDbStore);
    const uncachedMemoryBounded = makeUncachedClients(memoryStore);
    const uncachedIndexedBounded = makeUncachedClients(indexedDbStore);
    results.push(await measure(`Memory uncached sorted query ${size} rows (store already populated)`, () => uncachedMemorySorted.pop().queryItems("site", memoryList, sorted)));
    await memoryClient.queryItems("site", memoryList, sorted);
    results.push(await measureWarm(`Memory warm cached sorted ${size} rows`, () => memoryClient.queryItems("site", memoryList, sorted)));
    results.push(await measure(`IndexedDB uncached sorted query ${size} rows (store already populated)`, () => uncachedIndexedSorted.pop().queryItems("site", indexedList, sorted)));
    await indexedDbClient.queryItems("site", indexedList, sorted);
    results.push(await measureWarm(`IndexedDB warm cached sorted ${size} rows`, () => indexedDbClient.queryItems("site", indexedList, sorted)));
    results.push(await measure(`Memory uncached bounded query ${size} rows (store already populated)`, () => uncachedMemoryBounded.pop().queryItems("site", memoryList, bounded)));
    await memoryClient.queryItems("site", memoryList, bounded);
    results.push(await measureWarm(`Memory warm cached bounded ${size} rows`, () => memoryClient.queryItems("site", memoryList, bounded)));
    results.push(await measure(`IndexedDB uncached cursor query ${size} rows (store already populated)`, () => uncachedIndexedBounded.pop().queryItems("site", indexedList, bounded)));
    await indexedDbClient.queryItems("site", indexedList, bounded);
    results.push(await measureWarm(`IndexedDB warm cached bounded ${size} rows`, () => indexedDbClient.queryItems("site", indexedList, bounded)));
    if (size === 100_000) {
      const indexedRows = await indexedDbStore.getItems(scope, "site", indexedList);
      const cursorSorted = await cursorTopNSort(databaseName, JSON.stringify([scope, "site", indexedList]));
      const expectedSorted = await indexedDbClient.queryItems("site", indexedList, sorted);
      if (cursorSorted.some((item, index) => item.id !== expectedSorted[index]?.id)) throw new Error("Cursor top-50 sorting differs from the library's sorted query.");
      results.push(await measure("IndexedDB cursor filter and top-50 sort 100000 rows", () => cursorTopNSort(databaseName, JSON.stringify([scope, "site", indexedList]))));
      results.push(await measure("IndexedDB bulk read 100000 rows", async () => (await indexedDbStore.getItems(scope, "site", indexedList)).slice(0, 50)));
      results.push(await measure("JavaScript filter and full sort 100000 rows", () => applyItemQuery(indexedRows, sorted)));
      results.push(await measure("JavaScript bounded top-50 sort 100000 rows", () => boundedSort(indexedRows)));
    }
  }

  const freshnessList = "cross-instance-cache-invalidation";
  const freshnessQuery = { equals: { Status: "Open" }, limit: 10 };
  const freshnessClient = createClient({ scope, store: indexedDbStore, get: async () => ({ value: [] }) });
  const writerStore = createIndexedDbStore(databaseName);
  await indexedDbStore.putItems(scope, "site", freshnessList, [{ id: "first", fields: { Status: "Open" } }]);
  await freshnessClient.queryItems("site", freshnessList, freshnessQuery);
  await writerStore.putItems(scope, "site", freshnessList, [{ id: "second", fields: { Status: "Open" } }]);
  await new Promise((resolve) => setTimeout(resolve, 20));
  const freshRows = await freshnessClient.queryItems("site", freshnessList, freshnessQuery);
  if (freshRows.length !== 2) throw new Error("IndexedDB changes from a second store instance did not invalidate the local query cache.");

  return {
    kind: "real Chromium IndexedDB on localhost; synthetic records; no tenant calls",
    browser: navigator.userAgent,
    origin: location.origin,
    checks: ["second-store query-cache invalidation passed"],
    results,
  };
};
