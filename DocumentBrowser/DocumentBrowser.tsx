import * as React from "react";
import { FluentProvider, makeStyles, tokens, webLightTheme } from "@fluentui/react-components";
import { SharePointFrame } from "./components/SharePointFrame";
import {
  buildCheckFlowUrl,
  buildProvisionFlowUrl,
  checkFolderExists,
  createFolderStructure,
  FolderPresence,
} from "./utils/folderProvision";
import { getSharePointConfig, parseFolderProvisionTarget } from "./utils/sharePointUrl";

export interface IDocumentBrowserProps {
  sharePointFolderUrl: string;
  provisionFlowUrl: string;
  containerStyle?: React.CSSProperties;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    flex: "1 1 auto",
    alignSelf: "stretch",
    width: "100%",
    height: "100%",
    minHeight: 0,
    minWidth: 0,
    overflow: "hidden",
    backgroundColor: tokens.colorNeutralBackground1,
    boxSizing: "border-box",
  },
  body: {
    flex: "1 1 auto",
    display: "flex",
    flexDirection: "column",
    minHeight: 0,
    minWidth: 0,
  },
});

export const DocumentBrowser: React.FC<IDocumentBrowserProps> = ({
  sharePointFolderUrl,
  provisionFlowUrl,
  containerStyle,
}) => {
  const styles = useStyles();
  const sharePointConfig = React.useMemo(
    () => getSharePointConfig(sharePointFolderUrl),
    [sharePointFolderUrl]
  );
  const provisionTarget = React.useMemo(
    () => parseFolderProvisionTarget(sharePointFolderUrl),
    [sharePointFolderUrl]
  );
  const createUrl = React.useMemo(
    () => (provisionTarget ? buildProvisionFlowUrl(provisionFlowUrl, provisionTarget) : null),
    [provisionFlowUrl, provisionTarget]
  );
  const createUrlRef = React.useRef(createUrl);
  createUrlRef.current = createUrl;
  const [creating, setCreating] = React.useState(false);
  const [createError, setCreateError] = React.useState<string | null>(null);
  const [frameGeneration, setFrameGeneration] = React.useState(0);
  const [presence, setPresence] = React.useState<FolderPresence>("idle");

  React.useEffect(() => {
    if (!createUrl) {
      setPresence("idle");
      setCreating(false);
      setCreateError(null);
      return;
    }

    let active = true;
    const requestedUrl = createUrl;
    setPresence("checking");
    setCreating(false);
    setCreateError(null);

    const run = async (): Promise<void> => {
      const result = await checkFolderExists(buildCheckFlowUrl(requestedUrl));
      if (!active || createUrlRef.current !== requestedUrl) {
        return;
      }

      setPresence(result);
    };

    void run();

    return () => {
      active = false;
    };
  }, [createUrl]);

  const onCreate = React.useCallback(() => {
    if (!createUrl || creating) {
      return;
    }

    const requestedUrl = createUrl;
    setCreating(true);
    setCreateError(null);

    const run = async (): Promise<void> => {
      const result = await createFolderStructure(requestedUrl);
      if (createUrlRef.current !== requestedUrl) {
        return;
      }

      setCreating(false);
      if (result.ok) {
        setPresence("present");
        setFrameGeneration((current) => current + 1);
        return;
      }

      setCreateError(result.message);
    };

    void run();
  }, [createUrl, creating]);

  return (
    <FluentProvider theme={webLightTheme} className={styles.root} style={containerStyle}>
      <div className={styles.body}>
        <SharePointFrame
          src={sharePointConfig.src}
          placeholderTitle={sharePointConfig.placeholderTitle}
          placeholderSub={sharePointConfig.placeholderSub}
          creating={creating}
          createError={createError}
          onCreate={createUrl ? onCreate : undefined}
          presence={presence}
          folderCode={provisionTarget?.kod ?? ""}
          frameKey={`${sharePointConfig.src ?? ""}:${frameGeneration}`}
        />
      </div>
    </FluentProvider>
  );
};
