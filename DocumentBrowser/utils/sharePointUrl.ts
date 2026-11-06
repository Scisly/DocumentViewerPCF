/** Placeholder string injected by the PCF local test harness for unset properties. */
const HARNESS_PLACEHOLDER = "val";

export interface ISharePointConfig {
  src: string | null;
  placeholderTitle: string;
  placeholderSub: string;
}

export function readControlText(value: string | null | undefined): string {
  const trimmed = (value ?? "").trim();
  return trimmed === HARNESS_PLACEHOLDER ? "" : trimmed;
}

function normalizeInput(value: string | null | undefined): string {
  return readControlText(value);
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

const FOLDER_LIBRARIES = ["szansasprzedazy", "kontrakty", "projekty"] as const;

export type FolderLibrary = (typeof FOLDER_LIBRARIES)[number];

export interface IFolderProvisionTarget {
  encja: FolderLibrary;
  rok: string;
  miesiac: string;
  kod: string;
}

const SHARING_LINK_PATH = new RegExp(
  `^/:f:/r/(?:sites|teams)/[^/]+/(${FOLDER_LIBRARIES.join("|")})/(\\d{4})/(\\d{2})/([^/]+)/?$`
);

function isFolderLibrary(value: string): value is FolderLibrary {
  return FOLDER_LIBRARIES.some((library) => library === value);
}

function decodePathSegment(segment: string): string | null {
  try {
    const decoded = decodeURIComponent(segment).trim();
    return decoded ? decoded : null;
  } catch {
    return null;
  }
}

/**
 * Reads Encja, Rok, Miesiac and Kod from a stored `:f:/r/` sharing link.
 * Returns null when the link is not one of the document libraries.
 */
export function parseFolderProvisionTarget(
  raw: string | null | undefined
): IFolderProvisionTarget | null {
  const value = readControlText(raw);
  if (!value) {
    return null;
  }

  let pathname: string;
  try {
    pathname = new URL(value).pathname;
  } catch {
    return null;
  }

  const match = SHARING_LINK_PATH.exec(pathname);
  if (!match) {
    return null;
  }

  const library = match[1] ?? "";
  const rok = match[2] ?? "";
  const miesiac = match[3] ?? "";
  const kod = decodePathSegment(match[4] ?? "");

  if (!isFolderLibrary(library) || !kod) {
    return null;
  }

  return {
    encja: library,
    rok,
    miesiac,
    kod,
  };
}
