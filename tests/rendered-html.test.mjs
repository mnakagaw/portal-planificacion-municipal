import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
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
        !item.pmd.hasOfficialEvidence &&
        (item.pmd.has7_12 || item.pmd.hasDraft),
    ).length,
    16,
  );
  assert.equal(data.filter((item) => item.cdm.level === "complete").length, 71);
  assert.equal(data.filter((item) => item.ompp.level === "complete").length, 83);

  assert.match(source, /setRegion\(item\.region\)/);
  assert.match(source, /setProvince\(item\.provincia\)/);
  assert.match(source, /onClick=\{\(\) => item && chooseMunicipality\(item\)\}/);
  assert.match(source, /fetch\("\/data\/adm2\.geojson"\)/);
});
