import * as React from "react";
import { FluentProvider, makeStyles, tokens, webLightTheme } from "@fluentui/react-components";
import { SharePointFrame } from "./components/SharePointFrame";
import { getSharePointConfig } from "./utils/sharePointUrl";

export interface IDocumentBrowserProps {
  sharePointFolderUrl: string;
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
  containerStyle,
}) => {
  const styles = useStyles();

  const sharePointConfig = React.useMemo(
    () => getSharePointConfig(sharePointFolderUrl),
    [sharePointFolderUrl]
  );

  return (
    <FluentProvider theme={webLightTheme} className={styles.root} style={containerStyle}>
      <div className={styles.body}>
        <SharePointFrame
          src={sharePointConfig.src}
          placeholderTitle={sharePointConfig.placeholderTitle}
          placeholderSub={sharePointConfig.placeholderSub}
        />
      </div>
    </FluentProvider>
  );
};
