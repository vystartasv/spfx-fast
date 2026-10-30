import { Card, CardHeader, MessageBar, MessageBarBody, Spinner, Text } from "@fluentui/react-components";
import { useMarkdownFile, type UseMarkdownFileResult } from "@spfx-fast/core/react";
import type { createClient } from "@spfx-fast/core";
import { MarkdownContent } from "@spfx-fast/core/markdown";

type Client = ReturnType<typeof createClient>;

export function MarkdownFilePanel({ client, driveId, itemId }: { client: Client; driveId: string; itemId: string }) {
  const file: UseMarkdownFileResult = useMarkdownFile(client, driveId, itemId);

  return (
    <Card>
      <CardHeader header={<Text>Markdown file</Text>} />
      {file.loading && <Spinner label="Loading Markdown" />}
      {file.error && <MessageBar intent="warning"><MessageBarBody>Could not refresh this file. Cached content may still be shown.</MessageBarBody></MessageBar>}
      {file.markdown !== undefined && <MarkdownContent>{file.markdown}</MarkdownContent>}
      <Text size={200}>Source: {file.source ?? "loading"}</Text>
    </Card>
  );
}
