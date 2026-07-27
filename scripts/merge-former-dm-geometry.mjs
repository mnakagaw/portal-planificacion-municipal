import { readFile, writeFile } from "node:fs/promises";

const geojsonPath = new URL("../public/data/adm2.geojson", import.meta.url);
const geojson = JSON.parse(await readFile(geojsonPath, "utf8"));

const sourceUrl = new URL(
  "https://services3.arcgis.com/1WxBjaucfKL4MjiW/arcgis/rest/services/DPA_Distritos_Municipales_(vista)_/FeatureServer/29/query",
);
sourceUrl.search = new URLSearchParams({
  objectIds: "111,123,124,197",
  outFields: "*",
  outSR: "4326",
  f: "geojson",
}).toString();

const response = await fetch(sourceUrl);
if (!response.ok) {
  throw new Error(`No fue posible descargar la cartografía: ${response.status}`);
}

const source = await response.json();
const targetByObjectId = new Map([
  [111, { municipio: "Tireo", provincia: "La Vega" }],
  [123, { municipio: "La Victoria", provincia: "Santo Domingo" }],
  [124, { municipio: "La Caleta", provincia: "Santo Domingo" }],
  [197, { municipio: "Villa Central", provincia: "Barahona" }],
]);

const additions = source.features
  .map((feature) => {
    const target = targetByObjectId.get(feature.properties.OBJECTID);
    if (!target) return null;

    return {
      type: "Feature",
      id: feature.properties.ENLACE,
      properties: {
        adm2_code: feature.properties.ENLACE,
        municipio: target.municipio,
        provincia: target.provincia,
        region: feature.properties.RUP,
        shapeName: target.municipio,
        source_name: feature.properties.TOPO2,
        prov_code: feature.properties.PROV,
        mun_code: feature.properties.MUN,
        dm_code: feature.properties.DM,
        rup: feature.properties.RUP,
        km2: feature.properties.KM2,
        geometry_source: "IGN/ONE DPA Distritos Municipales, actualización 2022",
      },
      geometry: feature.geometry,
    };
  })
  .filter(Boolean);

if (additions.length !== targetByObjectId.size) {
  throw new Error(
    `Se esperaban ${targetByObjectId.size} geometrías y se recibieron ${additions.length}.`,
  );
}

const addedNames = new Set(additions.map((feature) => feature.properties.municipio));
geojson.features = geojson.features
  .filter((feature) => !addedNames.has(feature.properties.municipio))
  .concat(additions);

await writeFile(geojsonPath, JSON.stringify(geojson), "utf8");

console.log(`Geometrías incorporadas: ${[...addedNames].join(", ")}`);
