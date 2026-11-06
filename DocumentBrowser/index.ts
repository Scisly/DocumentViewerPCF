import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { DocumentBrowser as DocumentBrowserView, IDocumentBrowserProps } from "./DocumentBrowser";
import * as React from "react";

function resolveHeight(
  configuredHeight: number | null | undefined,
  allocatedHeight: number
): number | undefined {
  const configured = configuredHeight ?? 0;

  if (configured > 0 && allocatedHeight > 0) {
    return Math.max(configured, allocatedHeight);
  }

  if (configured > 0) {
    return configured;
  }

  if (allocatedHeight > 0) {
    return allocatedHeight;
  }

  return undefined;
}

function getContainerStyle(context: ComponentFramework.Context<IInputs>): React.CSSProperties {
  const style: React.CSSProperties = {
    display: "flex",
    flexDirection: "column",
    boxSizing: "border-box",
  };
  const { allocatedHeight, allocatedWidth } = context.mode;
  const height = resolveHeight(context.parameters.viewerHeight.raw, allocatedHeight);

  if (height !== undefined) {
    style.height = height;
    style.minHeight = height;
  }

  if (allocatedWidth > 0) {
    style.width = allocatedWidth;
    style.minWidth = allocatedWidth;
  }

  return style;
}

export class DocumentBrowser implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private notifyOutputChanged: () => void;

  constructor() {
    // Empty
  }

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    context.mode.trackContainerResize(true);
  }

  public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
    const props: IDocumentBrowserProps = {
      sharePointFolderUrl: context.parameters.sharePointFolderUrl.raw ?? "",
      provisionFlowUrl: context.parameters.provisionFlowUrl.raw ?? "",
      containerStyle: getContainerStyle(context),
    };

    return React.createElement(DocumentBrowserView, props);
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    // No cleanup required
  }
}
