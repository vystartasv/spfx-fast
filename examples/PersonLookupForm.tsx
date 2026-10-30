import * as React from "react";
import { FluentProvider, MessageBar, MessageBarBody, Spinner, webLightTheme } from "@fluentui/react-components";
import { ItemForm, type FieldDefinition } from "@spfx-fast/core/react";
import type { GraphListItem } from "@spfx-fast/core";
import type { createSPFxServices } from "@spfx-fast/core/spfx";

type Services = Awaited<ReturnType<typeof createSPFxServices>>;

/** ItemForm wired to authenticated SharePoint users and a bounded Graph lookup list. */
export function PersonLookupForm({
  services,
  siteId,
  lookupListId,
  onSubmit,
}: {
  services: Services;
  siteId: string;
  lookupListId: string;
  onSubmit(values: Record<string, unknown>): Promise<void> | void;
}): React.ReactElement {
  const [users, setUsers] = React.useState<Array<{ value: number; label: string }>>([]);
  const [projects, setProjects] = React.useState<Array<{ value: string; label: string }>>([]);
  const [loading, setLoading] = React.useState(true);
  const [error, setError] = React.useState(false);

  React.useEffect(() => {
    let active = true;
    void (async () => {
      try {
        const people: Array<{ value: number; label: string }> = [];
        let cursor: string | undefined;
        do {
          const page = await services.sharepoint.getSiteUsersPage({ cursor, pageSize: 100 });
          people.push(...page.items.map((user) => ({ value: user.Id, label: user.Title ?? user.Email ?? String(user.Id) })));
          cursor = page.nextCursor;
        } while (cursor && people.length < 1000);

        const lookupItems: GraphListItem[] = [];
        for await (const item of services.client.listItems<GraphListItem>(siteId, lookupListId, { fields: ["Title"], maxItems: 1000 })) {
          lookupItems.push(item);
        }
        if (!active) return;
        setUsers(people);
        setProjects(lookupItems.map((item) => ({ value: item.id, label: String(item.fields?.Title ?? item.id) })));
      } catch {
        if (active) setError(true);
      } finally {
        if (active) setLoading(false);
      }
    })();
    return () => { active = false; };
  }, [services, siteId, lookupListId]);

  const fields: FieldDefinition[] = [
    { name: "Title", title: "Request", type: "text", required: true },
    { name: "Owner", title: "Owner", type: "person", required: true, options: users },
    { name: "Project", title: "Project", type: "lookup", options: projects },
  ];

  return (
    <FluentProvider theme={webLightTheme}>
      {loading ? <Spinner label="Loading people and projects" /> : null}
      {error ? <MessageBar intent="error"><MessageBarBody>Could not load people or projects. Check access and the lookup list.</MessageBarBody></MessageBar> : null}
      {!loading && !error ? <ItemForm fields={fields} onSubmit={onSubmit} /> : null}
    </FluentProvider>
  );
}
