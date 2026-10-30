import { createClient } from "../src/index.js";

/** Read sharing permissions Graph exposes to the current signed-in user. */
export async function readVisibleDriveItemPermissions(
  client: ReturnType<typeof createClient>,
  driveId: string,
  itemId: string,
) {
  const permissions = await client.getDriveItemPermissions(driveId, itemId);
  return permissions.map((permission) => ({
    roles: permission.roles ?? [],
    linkType: permission.link?.type,
    linkScope: permission.link?.scope,
    principal: permission.grantedToV2?.user?.displayName
      ?? permission.grantedToV2?.siteUser?.displayName
      ?? permission.invitation?.email,
    inheritedFrom: permission.inheritedFrom?.path,
    expiresAt: permission.expirationDateTime,
  }));
}

// Avoid logging link.webUrl, shareId, or invitation secrets. Graph returns only
// caller-visible effective permissions; this is not a complete access audit.
