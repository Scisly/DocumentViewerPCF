import * as React from "react";
import {
  Body1,
  Button,
  Link,
  makeStyles,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  MessageBarTitle,
  Subtitle2,
  tokens,
} from "@fluentui/react-components";
import { FolderIcon } from "./FolderIcon";
import { isIframeEmbeddingLikelyBlocked } from "../utils/hostEnvironment";

export interface ISharePointFrameProps {
  src: string | null;
  placeholderTitle?: string;
  placeholderSub?: string;
}

const useStyles = makeStyles({
  host: {
    flex: "1 1 auto",
    display: "flex",
    flexDirection: "column",
    minHeight: 0,
    minWidth: 0,
  },
  frameWrap: {
    flex: "1 1 auto",
    display: "flex",
    flexDirection: "column",
    minHeight: 0,
    minWidth: 0,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  frame: {
    flex: "1 1 auto",
    width: "100%",
    minHeight: 0,
    border: "none",
    display: "block",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  empty: {
    flex: "1 1 auto",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: tokens.spacingVerticalL,
    minHeight: 0,
    padding: tokens.spacingVerticalXXL,
    textAlign: "center",
  },
  emptyIcon: {
    width: "54px",
    height: "54px",
    borderRadius: tokens.borderRadiusXLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground3,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontSize: "24px",
    color: tokens.colorNeutralForeground3,
  },
  emptyText: {
    maxWidth: "420px",
    color: tokens.colorNeutralForeground3,
  },
  messageBar: {
    flexShrink: 0,
  },
});

export const SharePointFrame: React.FC<ISharePointFrameProps> = ({
  src,
  placeholderTitle = "SharePoint not configured",
  placeholderSub = "Bind a SharePoint folder URL field or paste a folder URL in the test harness.",
}) => {
  const styles = useStyles();
  const embedBlocked = isIframeEmbeddingLikelyBlocked();

  const openLibrary = React.useCallback(() => {
    if (src) {
      window.open(src, "_blank", "noopener,noreferrer");
    }
  }, [src]);

  if (!src) {
    return (
      <div className={styles.host}>
        <MessageBar intent="warning" className={styles.messageBar}>
          <MessageBarBody>
            <MessageBarTitle>{placeholderTitle}</MessageBarTitle>
            {placeholderSub}
          </MessageBarBody>
        </MessageBar>
        <div className={styles.empty}>
          <div className={styles.emptyIcon}>
            <FolderIcon />
          </div>
          <Subtitle2>{placeholderTitle}</Subtitle2>
          <Body1 className={styles.emptyText}>{placeholderSub}</Body1>
        </div>
      </div>
    );
  }

  if (embedBlocked) {
    return (
      <div className={styles.host}>
        <MessageBar intent="info" className={styles.messageBar}>
          <MessageBarBody>
            <MessageBarTitle>Embedded preview unavailable on this host</MessageBarTitle>
            SharePoint blocks iframe embedding from localhost and most non-Microsoft domains.
            Being signed in on another tab does not change that. Use the button below, or deploy
            this control to a model-driven app / Power Apps page (*.dynamics.com, *.powerapps.com)
            where Microsoft allows framing.
          </MessageBarBody>
          <MessageBarActions>
            <Button appearance="primary" onClick={openLibrary}>
              Open folder in new tab
            </Button>
          </MessageBarActions>
        </MessageBar>
        <div className={styles.empty}>
          <div className={styles.emptyIcon}>
            <FolderIcon />
          </div>
          <Subtitle2>Preview blocked by SharePoint security policy</Subtitle2>
          <Body1 className={styles.emptyText}>
            Your folder URL is configured correctly. Open it directly to verify access:{" "}
            <Link href={src} target="_blank" rel="noopener noreferrer">
              {src}
            </Link>
          </Body1>
        </div>
      </div>
    );
  }

  return (
    <div className={styles.host}>
      <div className={styles.frameWrap}>
        <iframe
          className={styles.frame}
          src={src}
          title="SharePoint document library"
          allow="clipboard-read; clipboard-write"
        />
      </div>
    </div>
  );
};
