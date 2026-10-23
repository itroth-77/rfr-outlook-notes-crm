import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const packageMetadata = JSON.parse(
  fs.readFileSync(path.join(projectRoot, "package.json"), "utf8"),
);
const target = path.join(projectRoot, "public", "version.json");

fs.writeFileSync(
  target,
  `${JSON.stringify({ version: packageMetadata.version })}\n`,
  "utf8",
);
console.log(`version.json généré pour ${packageMetadata.version}.`);
