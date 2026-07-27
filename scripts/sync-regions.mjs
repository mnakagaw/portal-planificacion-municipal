import { readFile, writeFile } from "node:fs/promises";

const [municipalitiesPath, regionsPath] = process.argv.slice(2);

if (!municipalitiesPath || !regionsPath) {
  throw new Error(
    "Usage: node scripts/sync-regions.mjs <municipios.json> <regions_index.json>",
  );
}

const [municipalities, regions] = await Promise.all([
  readFile(municipalitiesPath, "utf8").then(JSON.parse),
  readFile(regionsPath, "utf8").then(JSON.parse),
]);

const normalize = (value) =>
  value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

const provinceAliases = {
  bahoruco: "baoruco",
};

const canonicalProvince = (value) => {
  const normalized = normalize(value);
  return provinceAliases[normalized] ?? normalized;
};

const regionByProvince = new Map();
for (const region of regions) {
  for (const province of region.provincias) {
    regionByProvince.set(canonicalProvince(province), region.name);
  }
}

const changes = [];
for (const municipality of municipalities) {
  const expectedRegion = regionByProvince.get(
    canonicalProvince(municipality.provincia),
  );
  if (!expectedRegion) {
    throw new Error(`No region found for province: ${municipality.provincia}`);
  }
  if (municipality.region !== expectedRegion) {
    changes.push({
      municipio: municipality.municipio,
      provincia: municipality.provincia,
      from: municipality.region,
      to: expectedRegion,
    });
    municipality.region = expectedRegion;
  }
}

await writeFile(
  municipalitiesPath,
  `${JSON.stringify(municipalities, null, 2)}\n`,
  "utf8",
);

const totals = Object.fromEntries(
  regions.map((region) => [
    region.name,
    municipalities.filter((municipality) => municipality.region === region.name)
      .length,
  ]),
);

console.log(JSON.stringify({ changes, totals }, null, 2));
