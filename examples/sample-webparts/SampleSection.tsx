import * as React from "react";
import { makeStyles, Text, tokens } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    display: "grid",
    rowGap: tokens.spacingVerticalL,
    width: "100%",
    maxWidth: "1120px",
    minWidth: 0,
    marginInline: "auto",
    padding: `${tokens.spacingVerticalL} ${tokens.spacingHorizontalL}`,
    backgroundColor: tokens.colorNeutralBackground1,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusLarge,
    boxShadow: tokens.shadow2,
    boxSizing: "border-box",
  },
  header: {
    display: "grid",
    rowGap: tokens.spacingVerticalXS,
    paddingBlockEnd: tokens.spacingVerticalM,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  title: {
    margin: 0,
    color: tokens.colorNeutralForeground1,
    fontSize: tokens.fontSizeBase600,
    lineHeight: tokens.lineHeightBase600,
    fontWeight: tokens.fontWeightSemibold,
  },
  description: { color: tokens.colorNeutralForeground2 },
});

export function SampleSection({ label, title, description, children }: {
  label: string;
  title: string;
  description: string;
  children: React.ReactNode;
}): React.ReactElement {
  const styles = useStyles();
  return <section aria-label={label} className={styles.root}>
    <header className={styles.header}>
      <h2 className={styles.title}>{title}</h2>
      <Text className={styles.description}>{description}</Text>
    </header>
    {children}
  </section>;
}
