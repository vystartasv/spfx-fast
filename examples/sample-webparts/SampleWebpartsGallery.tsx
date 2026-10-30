import * as React from "react";
import { FluentProvider, makeStyles, Text, tokens } from "@fluentui/react-components";
import { createSampleClient } from "./createSampleClient.js";
import { KnowledgePanel } from "./KnowledgePanel.js";
import { ApprovalInbox } from "./ApprovalInbox.js";
import { RequestDashboard } from "./RequestDashboard.js";
import { TeamTaskQueue } from "./TeamTaskQueue.js";
import { sampleLightTheme } from "./sampleTheme.js";

const useStyles = makeStyles({
  root: {
    display: "grid",
    rowGap: tokens.spacingVerticalXXL,
    minHeight: "100vh",
    padding: `${tokens.spacingVerticalXXL} ${tokens.spacingHorizontalL}`,
    backgroundColor: tokens.colorNeutralBackground3,
    boxSizing: "border-box",
  },
  content: {
    display: "grid",
    gridTemplateColumns: "repeat(2, minmax(0, 1fr))",
    gap: tokens.spacingVerticalXXL,
    width: "100%",
    maxWidth: "1160px",
    marginInline: "auto",
    "@media (max-width: 760px)": { gridTemplateColumns: "1fr" },
  },
  full: { gridColumn: "1 / -1", "@media (max-width: 760px)": { gridColumn: "auto" } },
  header: { display: "grid", gridColumn: "1 / -1", rowGap: tokens.spacingVerticalS, paddingBlockEnd: tokens.spacingVerticalL, paddingInlineStart: tokens.spacingHorizontalL, borderBottom: `1px solid ${tokens.colorNeutralStroke2}`, borderLeft: `4px solid ${tokens.colorBrandBackground}` },
  title: { margin: 0, color: tokens.colorNeutralForeground1, fontSize: tokens.fontSizeHero700, lineHeight: tokens.lineHeightHero700, fontWeight: tokens.fontWeightSemibold },
  description: { color: tokens.colorNeutralForeground2, maxWidth: "70ch" },
});

/** Local gallery for trying the samples with synthetic in-memory data and no tenant connection. */
export function SampleWebpartsGallery(): React.ReactElement {
  const client = React.useMemo(() => createSampleClient(), []);
  const styles = useStyles();
  return <FluentProvider theme={sampleLightTheme}>
    <main aria-label="Workplace services" className={styles.root}>
      <div className={styles.content}>
        <header className={styles.header}>
          <h1 className={styles.title}>Workplace services</h1>
          <Text className={styles.description}>North Campus requests, team tasks, facilities guidance, and manager approvals.</Text>
        </header>
        <div className={styles.full}><RequestDashboard client={client} /></div>
        <div className={styles.full}><TeamTaskQueue client={client} /></div>
        <KnowledgePanel client={client} />
        <ApprovalInbox client={client} />
      </div>
    </main>
  </FluentProvider>;
}
