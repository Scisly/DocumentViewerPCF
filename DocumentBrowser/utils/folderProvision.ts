import {
  IFolderProvisionTarget,
  readControlText,
} from "./sharePointUrl";

export type FolderCreateResult =
  | { ok: true }
  | { ok: false; message: string };

export type FolderPresence = "idle" | "checking" | "present" | "missing" | "unknown";

export type FolderCheckResult = "present" | "missing" | "unknown";

const FLOW_PLACEHOLDERS = ["{Encja}", "{Rok}", "{Miesiac}", "{Kod}"] as const;

/**
 * POST with no body and no custom headers, so the browser treats it as a simple
 * request and does not send a CORS preflight.
 */
const FLOW_REQUEST: RequestInit = { method: "POST" };

export function buildProvisionFlowUrl(
  template: string | null | undefined,
  target: IFolderProvisionTarget
): string | null {
  const value = readControlText(template);
  if (!value || !FLOW_PLACEHOLDERS.every((placeholder) => value.includes(placeholder))) {
    return null;
  }

  const substituted = value
    .replace("{Encja}", encodeURIComponent(target.encja))
    .replace("{Rok}", encodeURIComponent(target.rok))
    .replace("{Miesiac}", encodeURIComponent(target.miesiac))
    .replace("{Kod}", encodeURIComponent(target.kod));

  try {
    if (new URL(substituted).protocol !== "https:") {
      return null;
    }
  } catch {
    return null;
  }

  return substituted;
}

export function buildCheckFlowUrl(createUrl: string): string {
  const separator = createUrl.includes("?") ? "&" : "?";
  return `${createUrl}${separator}mode=check`;
}

function readExistsFlag(value: unknown): boolean | null {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value !== "string") {
    return null;
  }

  const normalized = value.trim().toLowerCase();
  if (normalized === "true") {
    return true;
  }

  if (normalized === "false") {
    return false;
  }

  return null;
}

export async function checkFolderExists(flowUrl: string): Promise<FolderCheckResult> {
  try {
    const response = await fetch(flowUrl, FLOW_REQUEST);
    if (!response.ok) {
      return "unknown";
    }

    const body = (await response.text()).trim();
    let parsed: unknown;
    try {
      parsed = JSON.parse(body) as unknown;
    } catch {
      return "unknown";
    }

    if (typeof parsed !== "object" || parsed === null || !("exists" in parsed)) {
      return "unknown";
    }

    const exists = readExistsFlag((parsed as { exists: unknown }).exists);
    if (exists === true) {
      return "present";
    }

    if (exists === false) {
      return "missing";
    }

    return "unknown";
  } catch {
    return "unknown";
  }
}

export async function createFolderStructure(flowUrl: string): Promise<FolderCreateResult> {
  try {
    const response = await fetch(flowUrl, FLOW_REQUEST);
    if (response.ok) {
      return { ok: true };
    }

    const message = (await response.text()).trim().slice(0, 400);
    return {
      ok: false,
      message: message || "Nie udało się utworzyć struktury folderów.",
    };
  } catch {
    return { ok: false, message: "Nie udało się wywołać przepływu." };
  }
}
