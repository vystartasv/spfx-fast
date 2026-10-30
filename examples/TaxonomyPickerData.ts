import type { GraphTermStoreGroup, GraphTermStoreSet, GraphTermStoreTerm, createClient } from "../src/index.js";

type Client = ReturnType<typeof createClient>;

/** Load one level of a term set for an authenticated taxonomy picker. */
export async function loadTaxonomyPickerData(
  client: Client,
  siteId: string,
): Promise<{ groups: GraphTermStoreGroup[]; sets: GraphTermStoreSet[]; terms: GraphTermStoreTerm[] }> {
  const groups = await client.getTermStoreGroups(siteId);
  const firstGroup = groups[0];
  if (!firstGroup) return { groups, sets: [], terms: [] };

  const sets = await client.getTermStoreSets(siteId, firstGroup.id);
  const firstSet = sets[0];
  const terms = firstSet ? await client.getTermSetChildren(siteId, firstSet.id) : [];
  return { groups, sets, terms };
}

// This read-only walkthrough requests only one level at a time. Graph requires
// delegated TermStore.Read.All; results reflect the current user's term access.
