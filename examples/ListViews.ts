import type { createLegacySharePointApi } from "../src/sharepoint.js";

type SharePointApi = ReturnType<typeof createLegacySharePointApi>;

/** Load a bounded view catalog and the selected view's internal field names. */
export async function inspectListViews(sharepoint: SharePointApi, listId: string) {
  const page = await sharepoint.getListViewsPage(listId, { pageSize: 50 });
  const selected = page.items.find((view) => view.DefaultView) ?? page.items[0];
  return {
    views: page.items.map(({ Id, Title, ViewType, DefaultView, PersonalView, RowLimit, Paged }) => ({
      id: Id,
      title: Title,
      type: ViewType,
      isDefault: DefaultView,
      isPersonal: PersonalView,
      rowLimit: RowLimit,
      paged: Paged,
    })),
    selectedViewFields: selected?.Id ? await sharepoint.getListViewFields(listId, selected.Id) : [],
    nextCursor: page.nextCursor,
  };
}
