import type { createSPFxServices } from "@spfx-fast/core/spfx";

type Services = Awaited<ReturnType<typeof createSPFxServices>>;

/** Read SharePoint profile data through the hosting web's signed-in SPHttpClient. */
export async function loadCurrentProfile(services: Services) {
  const profile = await services.sharepoint.getMyProperties();
  return {
    displayName: profile.DisplayName,
    email: profile.Email,
    department: profile.UserProfileProperties?.find(({ Key }) => Key === "Department")?.Value,
  };
}

/** Profile lookup uses a SharePoint claims login and remains subject to profile visibility. */
export async function loadColleagueDepartment(services: Services, claimsLogin: string) {
  return services.sharepoint.getUserProfilePropertyFor(claimsLogin, "SPS-Department");
}
