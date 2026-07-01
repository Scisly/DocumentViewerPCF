/** Hosts not on SharePoint's frame-ancestors allowlist (e.g. local PCF harness). */
export function isIframeEmbeddingLikelyBlocked(): boolean {
  const host = window.location.hostname.toLowerCase();
  return host === "localhost" || host === "127.0.0.1";
}
