function clean(value) {
  return String(value ?? "").trim().replace(/\/$/, "");
}

export function getConfig() {
  const config = {
    clientId: clean(import.meta.env.VITE_CLIENT_ID),
    tenantId: clean(import.meta.env.VITE_TENANT_ID),
    dataverseUrl: clean(import.meta.env.VITE_DATAVERSE_URL),
  };

  const missing = [];
  if (!config.clientId || config.clientId.startsWith("00000000")) missing.push("VITE_CLIENT_ID");
  if (!config.tenantId || config.tenantId.startsWith("00000000")) missing.push("VITE_TENANT_ID");
  if (!config.dataverseUrl || config.dataverseUrl.includes("YOURORG")) missing.push("VITE_DATAVERSE_URL");

  if (missing.length) {
    throw new Error(`Configuration manquante : ${missing.join(", ")}`);
  }

  return config;
}
