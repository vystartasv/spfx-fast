import type { createLegacySharePointApi } from "../src/sharepoint.js";

type SharePointApi = ReturnType<typeof createLegacySharePointApi>;

/** Load the current principal and one bounded page for a Fluent UI person picker. */
export async function loadSitePeople(sharepoint: SharePointApi) {
  const [currentUser, users, groups] = await Promise.all([
    sharepoint.getCurrentUser(),
    sharepoint.getSiteUsersPage({ pageSize: 50 }),
    sharepoint.getSiteGroupsPage({ pageSize: 50 }),
  ]);
  return {
    currentUser,
    people: users.items.map((user) => ({ id: user.Id, text: user.Title ?? user.LoginName ?? `User ${user.Id}`, secondaryText: user.Email })),
    groups: groups.items.map((group) => ({ id: group.Id, text: group.Title ?? `Group ${group.Id}` })),
    loadMorePeople: users.nextCursor ? () => sharepoint.getSiteUsersPage({ cursor: users.nextCursor, pageSize: 50 }) : undefined,
    loadMoreGroups: groups.nextCursor ? () => sharepoint.getSiteGroupsPage({ cursor: groups.nextCursor, pageSize: 50 }) : undefined,
  };
}

/** Resolve the members of a selected SharePoint group one bounded page at a time. */
export function loadSiteGroupMembers(sharepoint: SharePointApi, groupId: number) {
  return sharepoint.getSiteGroupMembersPage(groupId, { pageSize: 50 });
}
