export function compareVersions(left, right) {
  const leftParts = String(left || "").split(".").map(Number);
  const rightParts = String(right || "").split(".").map(Number);
  const length = Math.max(leftParts.length, rightParts.length);
  for (let index = 0; index < length; index += 1) {
    const difference = (leftParts[index] || 0) - (rightParts[index] || 0);
    if (difference !== 0) return Math.sign(difference);
  }
  return 0;
}

export async function redirectToLatestVersion(
  currentVersion,
  {
    baseUrl = globalThis.location?.href,
    fetchImpl = globalThis.fetch,
    replace = globalThis.location?.replace?.bind(globalThis.location),
    now = Date.now,
  } = {},
) {
  if (!baseUrl || !fetchImpl || !replace) return false;

  try {
    const versionUrl = new URL("./version.json", baseUrl);
    versionUrl.searchParams.set("_", String(now()));
    const response = await fetchImpl(versionUrl, { cache: "no-store" });
    if (!response.ok) return false;

    const latestVersion = String((await response.json())?.version || "");
    if (compareVersions(latestVersion, currentVersion) <= 0) return false;

    const applicationUrl = new URL("./index.html", baseUrl);
    applicationUrl.searchParams.set("v", latestVersion);
    applicationUrl.searchParams.set("_", String(now()));
    replace(applicationUrl.toString());
    return true;
  } catch (error) {
    console.warn("Vérification de version indisponible.", error);
    return false;
  }
}
