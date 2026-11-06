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
  Spinner,
  Subtitle2,
  Title3,
  tokens,
} from "@fluentui/react-components";
import { FolderIcon } from "./FolderIcon";
import { isIframeEmbeddingLikelyBlocked } from "../utils/hostEnvironment";
import { FolderPresence } from "../utils/folderProvision";

export interface ISharePointFrameProps {
  src: string | null;
  placeholderTitle?: string;
  placeholderSub?: string;
  creating?: boolean;
  createError?: string | null;
  onCreate?: () => void;
  presence?: FolderPresence;
  folderCode?: string;
  frameKey?: string;
}

type FolderView = "checking" | "missing" | "preview" | "preview-with-button";

function resolveFolderView(presence: FolderPresence, canCreate: boolean): FolderView {
  switch (presence) {
    case "checking":
      return "checking";
    case "missing":
      return "missing";
    case "unknown":
      return canCreate ? "preview-with-button" : "preview";
    case "idle":
    case "present":
      return "preview";
    default: {
      const unreachable: never = presence;
      return unreachable;
    }
  }
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
  notice: {
    width: "min(440px, 100%)",
    boxSizing: "border-box",
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    gap: tokens.spacingVerticalL,
    padding: "28px 28px 24px",
    textAlign: "left",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusXLarge,
    backgroundColor: tokens.colorNeutralBackground1,
  },
  noticeMark: {
    width: "40px",
    height: "40px",
    borderRadius: tokens.borderRadiusLarge,
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground2,
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
  },
  noticeCopy: {
    display: "flex",
    flexDirection: "column",
    alignItems: "flex-start",
    gap: tokens.spacingVerticalXS,
  },
  noticeEyebrow: {
    color: tokens.colorNeutralForeground3,
  },
  noticeBody: {
    color: tokens.colorNeutralForeground2,
  },
  messageBar: {
    flexShrink: 0,
  },
  toolbar: {
    flexShrink: 0,
    display: "flex",
    justifyContent: "flex-end",
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalM}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
});

export const SharePointFrame: React.FC<ISharePointFrameProps> = ({
  src,
  placeholderTitle = "SharePoint not configured",
  placeholderSub = "Bind a SharePoint folder URL field or paste a folder URL in the test harness.",
  creating = false,
  createError = null,
  onCreate,
  presence = "idle",
  folderCode = "",
  frameKey,
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

  const folderView = resolveFolderView(presence, Boolean(onCreate));

  if (src && folderView === "checking") {
    return (
      <div className={styles.host}>
        <div className={styles.empty}>
          <Spinner size="large" label="Sprawdzanie folderu" />
        </div>
      </div>
    );
  }

  if (src && folderView === "missing" && onCreate) {
    return (
      <div className={styles.host}>
        {createError ? (
          <MessageBar intent="error" className={styles.messageBar}>
            <MessageBarBody>
              <MessageBarTitle>Nie udało się utworzyć struktury</MessageBarTitle>
              {createError}
            </MessageBarBody>
          </MessageBar>
        ) : null}
        <div className={styles.empty}>
          <div className={styles.notice}>
            <div className={styles.noticeMark}>
              <FolderIcon />
            </div>
            <div className={styles.noticeCopy}>
              <Subtitle2 className={styles.noticeEyebrow}>Folder nie istnieje</Subtitle2>
              {folderCode ? <Title3>{folderCode}</Title3> : null}
              <Body1 className={styles.noticeBody}>
                Ten rekord nie ma jeszcze folderu w bibliotece dokumentów. Przycisk utworzy folder rekordu razem z podfolderami.
              </Body1>
            </div>
            <Button appearance="primary" disabled={creating} onClick={onCreate}>
              {creating ? "Tworzenie struktury…" : "Utwórz strukturę"}
            </Button>
          </div>
        </div>
      </div>
    );
  }

  const createBar = folderView === "preview-with-button" && onCreate ? (
    <>
      {createError ? (
        <MessageBar intent="error" className={styles.messageBar}>
          <MessageBarBody>
            <MessageBarTitle>Nie udało się utworzyć struktury</MessageBarTitle>
            {createError}
          </MessageBarBody>
        </MessageBar>
      ) : null}
      <div className={styles.toolbar}>
        <Button appearance="primary" disabled={creating} onClick={onCreate}>
          {creating ? "Tworzenie struktury…" : "Utwórz strukturę"}
        </Button>
      </div>
    </>
  ) : null;

  if (embedBlocked) {
    return (
      <div className={styles.host}>
        {createBar}
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
      {createBar}
      <div className={styles.frameWrap}>
        <iframe
          key={frameKey ?? src}
          className={styles.frame}
          src={src}
          title="SharePoint document library"
          allow="clipboard-read; clipboard-write"
        />
      </div>
    </div>
  );
};
