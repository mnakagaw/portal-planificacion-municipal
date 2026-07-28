import assert from "node:assert/strict";
import { readFile, stat } from "node:fs/promises";
import test from "node:test";

async function render() {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);

  return worker.fetch(
    new Request("http://localhost/", {
      headers: { accept: "text/html" },
    }),
    {
      ASSETS: {
        fetch: async () => new Response("Not found", { status: 404 }),
      },
    },
    {
      waitUntil() {},
      passThroughOnException() {},
    },
  );
}

test("server-renders the municipal portal", async () => {
  const response = await render();
  assert.equal(response.status, 200);
  assert.match(response.headers.get("content-type") ?? "", /^text\/html\b/i);

  const html = await response.text();
  assert.match(html, /<title>Portal de Planificación Municipal<\/title>/i);
  assert.match(html, /Seleccione un municipio/);
  assert.match(html, />Región</);
  assert.match(html, />Provincia</);
  assert.match(html, />Municipio</);
  assert.match(html, />Todos</);
  assert.match(html, /PMD Oficial/);
  assert.match(html, /PMD Borrador/);
  assert.match(html, />CDM</);
  assert.match(html, />OMPP</);
  assert.doesNotMatch(html, /Building your site|codex-preview|loading skeleton/i);
});

test("keeps the four status definitions aligned with the source data", async () => {
  const [source, rawData] = await Promise.all([
    readFile(new URL("../app/PortalApp.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/data/municipios.json", import.meta.url), "utf8"),
  ]);
  const data = JSON.parse(rawData);

  assert.equal(data.length, 162);
  assert.equal(data.filter((item) => item.pmd.hasOfficialEvidence).length, 42);
  assert.equal(
    data.filter(
      (item) =>
        item.pmd.hasOfficialEvidence ||
        item.pmd.has7_12 ||
        item.pmd.hasDraft,
    ).length,
    58,
  );
  assert.equal(
    data.filter(
      (item) =>
        item.pmd.hasOfficialEvidence ||
        item.pmd.has7_12 ||
        item.pmd.hasDraft ||
        item.cdm.level === "complete",
    ).length,
    81,
  );
  assert.equal(
    data.filter(
      (item) =>
        item.pmd.hasOfficialEvidence ||
        item.pmd.has7_12 ||
        item.pmd.hasDraft ||
        item.ompp.level === "complete",
    ).length,
    99,
  );

  assert.match(source, /useState<MapLayer>\("all"\)/);
  assert.match(source, /function matchesLayer/);
  assert.match(source, /function overviewLayer/);
  assert.match(source, /Fuente SISMAP 2\.02/);
  assert.match(source, /setSelectedRegions\(\[item\.region\]\)/);
  assert.match(
    source,
    /const chooseMunicipality[\s\S]*?setProvince\("Todas"\)[\s\S]*?setSelected\(item\)/,
  );
  assert.match(source, /onClick=\{\(\) => item && chooseMunicipality\(item\)\}/);
  assert.match(source, /import\.meta\.env\.BASE_URL.*data\/adm2\.geojson/);
  assert.match(source, /function RegionMultiSelect/);
  assert.match(source, /type="checkbox"/);
  assert.match(source, /viewBox=\{mapViewBox\}/);
  assert.match(source, /function boundsToViewBox/);
  assert.match(source, /https:\/\/es\.wikipedia\.org\/w\/api\.php/);
  assert.match(source, /href=\{wikipediaUrl\}/);
  assert.match(source, /Ver municipio en Wikipedia/);
});

test("keeps the map palette aligned with the color specification", async () => {
  const [source, css] = await Promise.all([
    readFile(new URL("../app/PortalApp.tsx", import.meta.url), "utf8"),
    readFile(new URL("../app/globals.css", import.meta.url), "utf8"),
  ]);
  const styles = `${source}\n${css}`;

  for (const color of [
    "#286A50",
    "#75A67E",
    "#527E91",
    "#8FADBA",
    "#F2E5BF",
    "#EEF1F0",
  ]) {
    assert.match(styles, new RegExp(color, "i"));
  }

  for (const oldColor of ["347c6c", "5878a3", "806d91", "b48150", "d9dfdd"].map(
    (value) => `#${value}`,
  )) {
    assert.doesNotMatch(styles, new RegExp(oldColor, "i"));
  }

  assert.match(source, /strokeWidth=\{isSelected \? 2 : 0\.9\}/);
  assert.match(css, /\.map-canvas\s*\{\s*background:\s*#f6f8f6/i);
  assert.match(css, /stroke:\s*rgba\(255,\s*255,\s*255,\s*0\.95\)/i);
});

test("uses PMD documents instead of administrative evidence links", async () => {
  const rawData = await readFile(
    new URL("../app/data/municipios.json", import.meta.url),
    "utf8",
  );
  const data = JSON.parse(rawData);
  const administrativeLink =
    /resoluc|aprobac|publicac|portal|pagina-web|captura|certific/i;
  const permittedCombinedDocument = /resolucion-y-plan-municipal/i;

  const incorrect = data
    .filter((item) => item.pmd.hasOfficialEvidence && item.pmd.officialUrl)
    .filter(
      (item) =>
        administrativeLink.test(item.pmd.officialUrl) &&
        !permittedCombinedDocument.test(item.pmd.officialUrl),
    )
    .map((item) => item.municipio);

  assert.deepEqual(incorrect, []);

  const mao = data.find((item) => item.municipio === "Mao");
  assert.match(mao.pmd.officialUrl, /Plan-Municipal-de-Desarrollo-2024-2028-Mao/);
  assert.doesNotMatch(mao.pmd.officialUrl, /Resoluc/i);
});

test("maps all 162 municipalities, including the four former districts", async () => {
  const rawGeojson = await readFile(
    new URL("../public/data/adm2.geojson", import.meta.url),
    "utf8",
  );
  const geojson = JSON.parse(rawGeojson);
  const mappedNames = new Set(
    geojson.features.map((feature) => feature.properties.municipio),
  );

  assert.equal(geojson.features.length, 162);
  for (const name of ["Villa Central", "Tireo", "La Caleta", "La Victoria"]) {
    assert.equal(mappedNames.has(name), true, `${name} is missing from the map`);
  }
  assert.ok(rawGeojson.length < 2_000_000, "map GeoJSON should remain lightweight");
});

test("publishes one traceable Word draft for each of the 104 target municipalities", async () => {
  const [rawData, rawManifest] = await Promise.all([
    readFile(new URL("../app/data/municipios.json", import.meta.url), "utf8"),
    readFile(
      new URL("../public/data/generated-pmd-drafts.json", import.meta.url),
      "utf8",
    ),
  ]);
  const data = JSON.parse(rawData);
  const manifest = JSON.parse(rawManifest);
  const generated = data.filter((item) => item.pmd.generatedDraftUrl);

  assert.equal(generated.length, 104);
  assert.equal(manifest.generated_count, 104);
  assert.equal(manifest.municipalities.length, 104);
  assert.equal(
    generated.filter(
      (item) =>
        item.pmd.hasOfficialEvidence ||
        item.pmd.has7_12 ||
        item.pmd.hasDraft,
    ).length,
    0,
  );

  await Promise.all(
    generated.map(async (item) => {
      assert.match(
        item.pmd.generatedDraftUrl,
        /^downloads\/pmd-borradores\/.+\.docx$/,
      );
      const fileUrl = new URL(`../public/${item.pmd.generatedDraftUrl}`, import.meta.url);
      const details = await stat(fileUrl);
      assert.ok(details.size > 30_000, `${item.municipio} draft is unexpectedly small`);
    }),
  );
});
