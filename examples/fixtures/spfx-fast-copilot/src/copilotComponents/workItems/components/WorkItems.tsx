import * as React from 'react';
import {
  FluentProvider,
  IdPrefixProvider,
  webLightTheme,
  webDarkTheme,
  Title3,
  Body1,
  Badge,
  Button,
  Card,
  CardHeader,
  MessageBar,
  MessageBarBody,
  Spinner,
  Text,
  makeStyles,
  tokens
} from '@fluentui/react-components';
import { ArrowExpand24Regular, Chat24Regular, Open24Regular } from '@fluentui/react-icons';
import { createCopilotTextContent } from '@microsoft/sp-copilot-component';
import { ListView, useListItems } from '@spfx-fast/core/react';

import type { IWorkItemsProps } from './IWorkItemsProps';
import type { GraphListItem } from '@spfx-fast/core';

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalM, padding: tokens.spacingHorizontalM },
  header: { display: 'flex', alignItems: 'flex-start', justifyContent: 'space-between', gap: tokens.spacingHorizontalM, flexWrap: 'wrap' },
  actions: { display: 'flex', gap: tokens.spacingHorizontalS, flexWrap: 'wrap' }
});

type Client = NonNullable<IWorkItemsProps['client']>;

export default function WorkItems(props: IWorkItemsProps): React.ReactElement {
  const styles = useStyles();
  const theme = props.hostContext.theme === 'dark' ? webDarkTheme : webLightTheme;
  const sendSummary = async (items: GraphListItem[]): Promise<void> => {
    const titles = items.slice(0, 10).map((item) => String(item.fields?.Title ?? '(untitled)'));
    const context = `SharePoint list “${props.listTitle}” at ${props.siteUrl}. First ${titles.length} loaded item titles (of at most 50 fetched):\n${titles.map((title) => `- ${title}`).join('\n')}`;
    await props.bridge.sendFollowUpMessageAsync([createCopilotTextContent(context.slice(0, 9000))]);
  };

  return (
    <IdPrefixProvider value="spfx-fast-copilot-">
      <FluentProvider theme={theme} targetDocument={props.targetDocument} style={{ minHeight: '100%' }}>
        <main className={styles.root}>
          <div className={styles.header}>
            <div>
              <Title3>{props.listTitle || 'SharePoint list'}</Title3>
              <Body1>{props.siteUrl || 'Resolving the requested SharePoint site…'}</Body1>
            </div>
            <div className={styles.actions}>
              <Button appearance="primary" icon={<ArrowExpand24Regular />} onClick={() => props.onRequestDisplayMode('fullscreen')}>
                {props.strings.ExpandButtonLabel}
              </Button>
              {props.siteUrl ? <Button appearance="secondary" icon={<Open24Regular />} onClick={() => props.bridge.openLinkAsync(props.siteUrl)}>
                {props.strings.OpenSiteButtonLabel}
              </Button> : null}
            </div>
          </div>
          <div>
            <Badge appearance="outline" color="informative">{props.strings.ThemeBadgePrefix} {props.hostContext.theme || props.strings.UnknownTheme}</Badge>{' '}
            <Badge appearance="outline" color="subtle">{props.strings.ModeBadgePrefix} {props.hostContext.displayMode || props.strings.DefaultDisplayMode}</Badge>
          </div>
          {props.loading ? <Spinner label="Resolving the requested SharePoint list" /> : null}
          {props.error ? <MessageBar intent="error"><MessageBarBody>{props.error}</MessageBarBody></MessageBar> : null}
          {!props.loading && !props.error && props.client && props.siteId && props.listId
            ? <ListItemsPanel client={props.client} siteId={props.siteId} listId={props.listId} listTitle={props.listTitle} onSendSummary={sendSummary} />
            : null}
          <Text size={200}>Read-only view. Results are permission-trimmed and bounded to 50 list items.</Text>
        </main>
      </FluentProvider>
    </IdPrefixProvider>
  );
}

function ListItemsPanel({ client, siteId, listId, listTitle, onSendSummary }: {
  client: Client;
  siteId: string;
  listId: string;
  listTitle: string;
  onSendSummary: (items: GraphListItem[]) => Promise<void>;
}): React.ReactElement {
  const items = useListItems(client, siteId, listId, { fields: ['Title', 'Modified'], pageSize: 20, maxItems: 50 });
  return (
    <Card>
      <CardHeader header={<Text>{listTitle}</Text>} description={<Text>{items.items.length} of up to 50 items loaded · {items.source ?? 'loading'}</Text>} />
      <ListView items={items.items} fields={[{ name: 'Title', title: 'Title' }, { name: 'Modified', title: 'Modified' }]} loading={items.loading} error={items.error} />
      <div>
        {items.hasMore ? <Button disabled={items.loading} onClick={items.loadMore}>Load more</Button> : null}{' '}
        <Button disabled={items.items.length === 0} icon={<Chat24Regular />} onClick={() => onSendSummary(items.items)}>Share first 10 titles with Copilot</Button>
      </div>
    </Card>
  );
}
