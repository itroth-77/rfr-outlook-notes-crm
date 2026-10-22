import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

function readEnvFile() {
  const envPath = path.join(projectRoot, ".env");
  if (!fs.existsSync(envPath)) return {};
  return Object.fromEntries(
    fs
      .readFileSync(envPath, "utf8")
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter((line) => line && !line.startsWith("#") && line.includes("="))
      .map((line) => {
        const separator = line.indexOf("=");
        return [line.slice(0, separator).trim(), line.slice(separator + 1).trim()];
      }),
  );
}

const localEnv = readEnvFile();
const packageMetadata = JSON.parse(
  fs.readFileSync(path.join(projectRoot, "package.json"), "utf8"),
);
const appVersion = packageMetadata.version;
const manifestVersion = `${appVersion}.0`;
const appUrl = (process.env.APP_URL || localEnv.APP_URL || "").replace(/\/$/, "");
const dataverseUrl = (
  process.env.VITE_DATAVERSE_URL ||
  localEnv.VITE_DATAVERSE_URL ||
  ""
).replace(/\/$/, "");

if (!appUrl.startsWith("https://") || appUrl.includes("YOUR-HOST")) {
  throw new Error("APP_URL doit contenir l’URL HTTPS publiée de l’application.");
}
if (!dataverseUrl.startsWith("https://") || dataverseUrl.includes("YOURORG")) {
  throw new Error("VITE_DATAVERSE_URL doit contenir l’URL de l’environnement Dataverse.");
}

const appOrigin = new URL(appUrl).origin;

const template = fs.readFileSync(path.join(projectRoot, "manifest.template.xml"), "utf8");
const manifest = template
  .replaceAll("{{APP_VERSION}}", appVersion)
  .replaceAll("{{MANIFEST_VERSION}}", manifestVersion)
  .replaceAll("{{APP_URL}}", appUrl)
  .replaceAll("{{APP_ORIGIN}}", appOrigin)
  .replaceAll("{{DATAVERSE_URL}}", dataverseUrl);

fs.writeFileSync(path.join(projectRoot, "manifest.xml"), manifest, "utf8");
console.log("manifest.xml généré.");
