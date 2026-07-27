import { readFile, writeFile } from "node:fs/promises";

const [municipalitiesPath, sismapPath] = process.argv.slice(2);

if (!municipalitiesPath || !sismapPath) {
  throw new Error(
    "Usage: node scripts/reclassify-pmd.mjs <municipios.json> <sismap_2_02_raw.json>",
  );
}

const [municipalities, sismapRows] = await Promise.all([
  readFile(municipalitiesPath, "utf8").then(JSON.parse),
  readFile(sismapPath, "utf8").then(JSON.parse),
]);

const criterionPattern = (criterion) =>
  new RegExp(
    `(?:^|[^0-9])0?${criterion}\\s*[-_./]\\s*12(?:[^0-9]|$)`,
    "i",
  );

const evidenceFor = (row, criterion) =>
  row.evidences.filter((evidence) =>
    criterionPattern(criterion).test(evidence.title),
  );

const sismapByUrl = new Map(
  sismapRows.map((row) => [row.organismUrl.toLowerCase(), row]),
);

let matched = 0;
let official = 0;
let borradorOnly = 0;

for (const municipality of municipalities) {
  const row = sismapByUrl.get(municipality.sismapUrl.toLowerCase());
  if (!row) {
    municipality.pmd.hasOfficialEvidence = false;
    municipality.pmd.has7_12 = false;
    municipality.pmd.has8_12 = false;
    municipality.pmd.officialEvidenceCount = 0;
    municipality.pmd.officialEvidenceTitles = [];
    municipality.pmd.officialEvidences = [];
    municipality.pmd.officialUrl = "";
    municipality.pmd.officialReason = "";
    continue;
  }

  matched += 1;
  const score = Number.parseFloat(row.score) || 0;
  const seven = evidenceFor(row, 7);
  const eight = evidenceFor(row, 8);
  const isOfficial = score >= 100 || eight.length > 0;

  municipality.pmd.score = score;
  municipality.pmd.expiry = row.expiry || municipality.pmd.expiry;
  municipality.pmd.has7_12 = seven.length > 0;
  municipality.pmd.has8_12 = eight.length > 0;
  municipality.pmd.hasOfficialEvidence = isOfficial;
  municipality.pmd.officialEvidenceCount = eight.length;
  municipality.pmd.officialEvidenceTitles = eight.map(
    (evidence) => evidence.title,
  );
  municipality.pmd.officialEvidences = eight;
  municipality.pmd.officialReason =
    score >= 100 && eight.length > 0
      ? "SISMAP 100% + evidencia 8-12"
      : score >= 100
        ? "SISMAP 100%"
        : "Evidencia 8-12";

  const preferredEight =
    eight.find((evidence) =>
      /plan municipal|plan de desarrollo|libro pmd|pmd 20/i.test(
        evidence.title,
      ),
    ) ??
    eight.find((evidence) => /publicaci[oó]n/i.test(evidence.title)) ??
    eight.find((evidence) => /resoluci[oó]n|aprueba/i.test(evidence.title)) ??
    eight[0];

  municipality.pmd.officialUrl = preferredEight?.href ?? "";

  if (isOfficial) {
    official += 1;
    municipality.pmd.label = "PMD oficial";
    municipality.pmd.level = "complete";
    municipality.action = "ver";
    municipality.actionLabel = "Ver PMD oficial";
  } else if (seven.length > 0) {
    borradorOnly += 1;
    municipality.pmd.label = "Borrador disponible";
    municipality.pmd.level = "progress";
    municipality.action = "continuar";
    municipality.actionLabel = "Continuar elaboración";
  }
}

await writeFile(
  municipalitiesPath,
  `${JSON.stringify(municipalities, null, 2)}\n`,
  "utf8",
);

console.log(
  JSON.stringify(
    {
      municipalities: municipalities.length,
      sismapRows: sismapRows.length,
      matched,
      official,
      borradorOnly,
    },
    null,
    2,
  ),
);
