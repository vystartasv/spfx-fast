/**
 * Properties schema for this Copilot Component.
 *
 * This schema is defined with Zod and exported as JSON Schema via
 * `zod-to-json-schema`. The manifest references the compiled `.js` default
 * export, which the Copilot host uses to validate and describe the tool
 * arguments that Copilot passes when invoking this component.
 *
 * To add more properties, extend the `z.object({...})` below — they will
 * automatically appear as tool parameters in the Copilot UI.
 */
// The Copilot manifest generator currently types its schema helper against
// Zod v3. Zod v4 ships the compatible v3 API at this subpath, with the security
// fixes from the current v4 line.
import { z } from 'zod/v3';
import zodToJsonSchema from 'zod-to-json-schema';

const propertiesSchema = z.object({
  siteUrl: z.string().url().optional().describe('Absolute HTTPS URL of the SharePoint site to inspect. Use the site URL only, without a page or list path. Omit it to use the current SharePoint site; never infer a URL from a site name.'),
  listTitle: z.string().trim().min(1).max(128).describe('Exact display title of the SharePoint list the user named. Do not guess or substitute a list title.')
});

export type IWorkItemsCopilotComponentProperties = z.infer<typeof propertiesSchema>;

export default zodToJsonSchema(propertiesSchema);
