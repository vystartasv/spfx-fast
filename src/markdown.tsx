import * as React from "react";
import ReactMarkdown from "react-markdown";
import rehypeSanitize from "rehype-sanitize";
import remarkGfm from "remark-gfm";

export interface MarkdownContentProps {
  /** Untrusted Markdown text, such as content returned by SharePoint or Graph. */
  children: string;
  className?: string;
}

/** Render Markdown as accessible content. Raw HTML is never parsed or executed. */
export function MarkdownContent({ children, className }: MarkdownContentProps): React.ReactElement {
  return (
    <div className={className} style={{ color: "var(--colorNeutralForeground1, #242424)", overflowWrap: "anywhere", lineHeight: 1.5 }}>
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        rehypePlugins={[rehypeSanitize]}
        components={{
          h1: ({ node: _node, ...props }) => <h1 {...props} style={{ fontSize: "1.7em", lineHeight: 1.2, margin: "0 0 .7em" }} />,
          h2: ({ node: _node, ...props }) => <h2 {...props} style={{ fontSize: "1.35em", lineHeight: 1.25, margin: "1.2em 0 .5em" }} />,
          h3: ({ node: _node, ...props }) => <h3 {...props} style={{ fontSize: "1.1em", margin: "1em 0 .4em" }} />,
          p: ({ node: _node, ...props }) => <p {...props} style={{ margin: "0 0 .8em", maxWidth: "75ch" }} />,
          a: ({ node: _node, ...props }) => <a {...props} style={{ color: "var(--colorBrandForegroundLink, #115ea3)", textUnderlineOffset: "2px" }} />,
          pre: ({ node: _node, ...props }) => <pre {...props} style={{ overflowX: "auto", padding: "12px 16px", borderRadius: "4px", background: "var(--colorNeutralBackground3, #f0f0f0)" }} />,
          code: ({ node: _node, ...props }) => <code {...props} style={{ fontFamily: "var(--fontFamilyMonospace, Consolas, monospace)", fontSize: ".92em" }} />,
          blockquote: ({ node: _node, ...props }) => <blockquote {...props} style={{ margin: "0 0 .8em", paddingLeft: "1em", borderLeft: "2px solid var(--colorNeutralStroke1, #d1d1d1)", color: "var(--colorNeutralForeground2, #424242)" }} />,
          ul: ({ node: _node, ...props }) => <ul {...props} style={{ paddingLeft: "1.5em", margin: "0 0 .8em" }} />,
          ol: ({ node: _node, ...props }) => <ol {...props} style={{ paddingLeft: "1.5em", margin: "0 0 .8em" }} />,
          table: ({ node: _node, ...props }) => <table {...props} style={{ borderCollapse: "collapse", width: "100%", margin: "0 0 .8em" }} />,
          th: ({ node: _node, ...props }) => <th {...props} style={{ padding: "6px 10px", borderBottom: "1px solid var(--colorNeutralStroke1, #d1d1d1)", textAlign: "left" }} />,
          td: ({ node: _node, ...props }) => <td {...props} style={{ padding: "6px 10px", borderBottom: "1px solid var(--colorNeutralStroke2, #e0e0e0)" }} />,
          hr: ({ node: _node, ...props }) => <hr {...props} style={{ border: 0, borderTop: "1px solid var(--colorNeutralStroke1, #d1d1d1)", margin: "1em 0" }} />,
        }}
      >
        {children}
      </ReactMarkdown>
    </div>
  );
}
