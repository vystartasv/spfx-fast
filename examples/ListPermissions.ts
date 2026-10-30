import type { createLegacySharePointApi } from "../src/sharepoint.js";

type SharePointApi = ReturnType<typeof createLegacySharePointApi>;

/** Inspect one list's role bindings and whether that list has unique permissions. */
export async function inspectListPermissionScope(sharepoint: SharePointApi, listId: string) {
  const [scope, assignments] = await Promise.all([
    sharepoint.getListPermissionScope(listId),
    sharepoint.getListRoleAssignments(listId),
  ]);
  return {
    scope,
    assignments: assignments.map(({ Member, RoleDefinitionBindings }) => ({
      principal: Member?.Title ?? Member?.LoginName ?? `Principal ${Member?.Id ?? "unknown"}`,
      roles: (RoleDefinitionBindings ?? []).map(({ Name }) => Name).filter((name): name is string => Boolean(name)),
    })),
  };
}

/** Add a principal role after the calling UI has confirmed the exact target and access. */
export async function addConfirmedListRole(
  sharepoint: SharePointApi,
  listId: string,
  principalId: number,
  roleDefinitionId: number,
  confirmed: boolean,
) {
  if (!confirmed) throw new Error("Show the list, principal, and role to the user before changing permissions.");
  const scope = await sharepoint.getListPermissionScope(listId);
  if (!scope.HasUniqueRoleAssignments) {
    await sharepoint.breakListRoleInheritance(listId, { confirmed: true, copyRoleAssignments: true });
  }
  await sharepoint.addListRoleAssignment(listId, principalId, roleDefinitionId, { confirmed: true });
}

/** Remove every role binding for a principal from one list after the UI confirms the impact. */
export async function removeConfirmedListRoles(
  sharepoint: SharePointApi,
  listId: string,
  principalId: number,
  confirmed: boolean,
) {
  if (!confirmed) throw new Error("Show the list and resulting access change to the user before removing permissions.");
  await sharepoint.removeListRoleAssignment(listId, principalId, { confirmed: true });
}

// This inspects one list's REST permission scope. Inherited parent assignments,
// Entra group membership, sharing links, and item-level unique scopes are not
// resolved, so this output must not be presented as the complete effective ACL.
