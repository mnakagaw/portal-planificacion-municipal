import fs from "node:fs";
import path from "node:path";
import process from "node:process";

const manifestPath = process.argv[2];

if (!manifestPath) {
  throw new Error(
    "Usage: node scripts/apply-generated-drafts.mjs <absolute-or-relative-manifest.json>",
  );
}

const resolvedManifest = path.resolve(manifestPath);
const repoRoot = path.resolve(import.meta.dirname, "..");
const dataPath = path.join(repoRoot, "app", "data", "municipios.json");
const publicManifestPath = path.join(
  repoRoot,
  "public",
  "data",
  "generated-pmd-drafts.json",
);

const manifest = JSON.parse(fs.readFileSync(resolvedManifest, "utf8"));
const municipalities = JSON.parse(fs.readFileSync(dataPath, "utf8"));
const byId = new Map(
  manifest.municipalities.map((item) => [Number(item.id), item]),
);

if (manifest.expected_target_count !== 104 || byId.size !== 104) {
  throw new Error(
    `Expected a complete 104-municipality manifest; received ${byId.size}`,
  );
}

let updated = 0;
for (const municipality of municipalities) {
  const generated = byId.get(Number(municipality.id));
  if (!generated) {
    delete municipality.pmd.generatedDraftUrl;
    delete municipality.pmd.generatedDraftPeriod;
    delete municipality.pmd.generatedDraftGeneratedAt;
    continue;
  }
  municipality.pmd.generatedDraftUrl = generated.relative_url;
  municipality.pmd.generatedDraftPeriod = manifest.period;
  municipality.pmd.generatedDraftGeneratedAt = manifest.generated_at;
  updated += 1;
}

if (updated !== 104) {
  throw new Error(`Applied ${updated} drafts instead of 104`);
}

fs.writeFileSync(dataPath, `${JSON.stringify(municipalities, null, 2)}\n`);
fs.mkdirSync(path.dirname(publicManifestPath), { recursive: true });
fs.writeFileSync(publicManifestPath, `${JSON.stringify(manifest, null, 2)}\n`);

console.log(
  JSON.stringify({
    updated,
    dataPath,
    publicManifestPath,
  }),
);
