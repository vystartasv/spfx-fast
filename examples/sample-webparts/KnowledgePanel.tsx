import * as React from "react";
import { Card, CardHeader, FluentProvider, MessageBar, MessageBarBody, Spinner } from "@fluentui/react-components";
import type { createClient } from "@spfx-fast/core";
import { useMarkdownFile } from "@spfx-fast/core/react";
import { MarkdownContent } from "@spfx-fast/core/markdown";
import { sampleDriveId, sampleMarkdownItemId } from "./sampleData.js";
import { SampleSection } from "./SampleSection.js";
import { sampleLightTheme } from "./sampleTheme.js";

type Client = ReturnType<typeof createClient>;

/** Cached-first safe Markdown rendering. Source files never execute as HTML or MDX. */
export function KnowledgePanel({ client, driveId = sampleDriveId, itemId = sampleMarkdownItemId }: { client: Client; driveId?: string; itemId?: string }): React.ReactElement {
  const file = useMarkdownFile(client, driveId, itemId);
  return (
    <FluentProvider theme={sampleLightTheme}>
      <SampleSection label="Team knowledge" title="Facilities and IT guide" description="Quick answers for equipment, meeting rooms, and service desk support.">
      <Card>
        <CardHeader header="North Campus handbook" description="Applies to staff and visitors" />
        {file.loading && file.markdown === undefined ? <Spinner label="Loading handbook" /> : null}
        {file.error ? <MessageBar intent="warning"><MessageBarBody>Refresh failed. Any cached handbook text remains available.</MessageBarBody></MessageBar> : null}
        {file.markdown !== undefined ? <MarkdownContent>{file.markdown}</MarkdownContent> : null}
      </Card>
      </SampleSection>
    </FluentProvider>
  );
}
