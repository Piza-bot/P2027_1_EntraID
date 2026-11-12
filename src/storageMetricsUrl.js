const sharePointSuffixes = [
  ".sharepoint.com",
  ".sharepoint.us",
  ".sharepoint.de",
  ".sharepoint.cn",
];

export function getStorageMetricsUrl(webUrl) {
  if (typeof webUrl !== "string") return null;

  let driveUrl;
  try {
    driveUrl = new URL(webUrl);
  } catch {
    return null;
  }

  if (
    driveUrl.protocol !== "https:"
    || !sharePointSuffixes.some((suffix) => driveUrl.hostname.endsWith(suffix))
  ) {
    return null;
  }

  const personalSite = driveUrl.pathname.match(/^\/personal\/([^/]+)(?:\/|$)/i);
  if (!personalSite) return null;

  let siteName;
  try {
    siteName = decodeURIComponent(personalSite[1]);
  } catch {
    return null;
  }
  if (!/^[\w.-]+$/.test(siteName)) return null;

  return new URL(
    `/personal/${encodeURIComponent(siteName)}/_layouts/15/storman.aspx`,
    driveUrl.origin,
  ).href;
}
