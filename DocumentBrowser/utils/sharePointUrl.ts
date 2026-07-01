/** Placeholder string injected by the PCF local test harness for unset properties. */
const HARNESS_PLACEHOLDER = "val";

export interface ISharePointConfig {
  src: string | null;
  placeholderTitle: string;
  placeholderSub: string;
}

function normalizeInput(value: string | null | undefined): string {
  const trimmed = (value ?? "").trim();
  return trimmed === HARNESS_PLACEHOLDER ? "" : trimmed;
}

export function isValidSharePointFolderUrl(url: string): boolean {
  if (!url) {
    return false;
  }

  try {
    const parsed = new URL(url);
    return (
      parsed.protocol === "https:" &&
      /\.sharepoint\.(com|us|de|cn|mil)$/i.test(parsed.hostname)
    );
  } catch {
    return false;
  }
}

function isSharePointHost(hostname: string): boolean {
  return /\.sharepoint\.(com|us|de|cn|mil)$/i.test(hostname);
}

/**
 * Normalizes SharePoint folder URLs from Dataverse, including AllItems.aspx links
 * that store the folder path in the `id` query parameter.
 */
export function normalizeSharePointFolderUrl(raw: string | null | undefined): string | null {
  const value = normalizeInput(raw);
  if (!value) {
    return null;
  }

  try {
    const parsed = new URL(value);

    if (parsed.protocol !== "https:" || !isSharePointHost(parsed.hostname)) {
      return null;
    }

    if (parsed.pathname.includes("/Forms/AllItems.aspx")) {
      const idParam = parsed.searchParams.get("id");
      if (idParam) {
        const folderPath = decodeURIComponent(idParam).replace(/\/$/, "");
        const folderUrl = `${parsed.origin}${folderPath.startsWith("/") ? folderPath : `/${folderPath}`}`;
        return isValidSharePointFolderUrl(folderUrl) ? folderUrl : null;
      }

      const libraryUrl = value.split("/Forms/AllItems.aspx")[0]?.replace(/\/$/, "") ?? "";
      return isValidSharePointFolderUrl(libraryUrl) ? libraryUrl : null;
    }

    parsed.search = "";
    parsed.hash = "";
    const directUrl = parsed.toString().replace(/\/$/, "");
    return isValidSharePointFolderUrl(directUrl) ? directUrl : null;
  } catch {
    return null;
  }
}

export function getSharePointConfig(folderUrl: string): ISharePointConfig {
  const normalized = normalizeSharePointFolderUrl(folderUrl);

  if (!normalizeInput(folderUrl)) {
    return {
      src: null,
      placeholderTitle: "SharePoint folder not configured",
      placeholderSub:
        "Bind this control to the Dataverse URL column, or paste a folder URL in the test harness Value field.",
    };
  }

  if (!normalized) {
    return {
      src: null,
      placeholderTitle: "Invalid SharePoint folder URL",
      placeholderSub:
        "Use a direct folder URL or an AllItems.aspx link with an id query parameter pointing to the folder.",
    };
  }

  return {
    src: normalized,
    placeholderTitle: "",
    placeholderSub: "",
  };
}
