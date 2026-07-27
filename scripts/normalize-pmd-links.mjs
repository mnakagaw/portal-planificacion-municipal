import { readFile, writeFile } from "node:fs/promises";

const dataPath = new URL("../app/data/municipios.json", import.meta.url);
const municipalities = JSON.parse(await readFile(dataPath, "utf8"));

function normalize(value = "") {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function isAdministrativeEvidence(value) {
  const text = normalize(value);
  const combinesResolutionAndPlan = text.includes(
    "resolucionyplanmunicipaldedesarrollo",
  );

  if (combinesResolutionAndPlan) return false;

  return [
    "resoluc",
    "aprobac",
    "publicac",
    "portal",
    "paginaweb",
    "captura",
    "certific",
    "constancia",
    "remision",
    "solicitud",
  ].some((term) => text.includes(term));
}

function preferredPmdDocument(pmd) {
  const candidates = [];
  const hasCurrentDocumentContext =
    pmd.hasCurrent || pmd.has7_12 || pmd.has8_12;

  if (pmd.pdfUrl && hasCurrentDocumentContext) {
    candidates.push({
      href: pmd.pdfUrl,
      title: pmd.pdfTitle || pmd.pdfUrl,
      source: "inventory",
    });
  }

  const matchingOfficialEvidence = (pmd.officialEvidences ?? []).find(
    (evidence) => evidence?.href === pmd.officialUrl,
  );

  if (pmd.officialUrl && matchingOfficialEvidence) {
    candidates.push({
      href: pmd.officialUrl,
      title: matchingOfficialEvidence?.title || pmd.officialUrl,
      source: "official",
    });
  }

  for (const evidence of pmd.officialEvidences ?? []) {
    if (evidence?.href) {
      candidates.push({
        href: evidence.href,
        title: evidence.title || evidence.href,
        source: "evidence",
      });
    }
  }

  const uniqueCandidates = [
    ...new Map(candidates.map((candidate) => [candidate.href, candidate])).values(),
  ];

  return (
    uniqueCandidates
      .filter(
        (candidate) =>
          !isAdministrativeEvidence(candidate.title) &&
          !isAdministrativeEvidence(candidate.href),
      )
      .map((candidate) => {
        const title = normalize(candidate.title);
        let score =
          candidate.source === "official"
            ? 50
            : candidate.source === "inventory"
              ? 30
              : 0;
        if (title.includes("202812")) score += 40;
        if (title.includes("2020712")) score -= 10;
        if (
          title.includes("planmunicipal") ||
          title.includes("plandedesarrollo") ||
          title.includes("pmd") ||
          title.includes("libro") ||
          title.includes("documento")
        ) {
          score += 20;
        }
        return { ...candidate, score };
      })
      .sort((left, right) => right.score - left.score)[0] ?? null
  );
}

const changes = [];

for (const municipality of municipalities) {
  if (!municipality.pmd?.hasOfficialEvidence) continue;

  const preferred = preferredPmdDocument(municipality.pmd);
  const nextUrl = preferred?.href ?? "";

  if (municipality.pmd.officialUrl !== nextUrl) {
    changes.push({
      municipio: municipality.municipio,
      from: municipality.pmd.officialUrl,
      to: nextUrl,
    });
    municipality.pmd.officialUrl = nextUrl;
  }
}

await writeFile(dataPath, `${JSON.stringify(municipalities, null, 2)}\n`, "utf8");

console.log(`Enlaces PMD actualizados: ${changes.length}`);
for (const change of changes) {
  console.log(`- ${change.municipio}`);
}
