import { readFile, writeFile } from "node:fs/promises";

const geojsonPath = new URL("../public/data/adm2.geojson", import.meta.url);
const geojson = JSON.parse(await readFile(geojsonPath, "utf8"));
const decimalPlaces = 3;

let pointsBefore = 0;
let pointsAfter = 0;

function samePoint(left, right) {
  return left[0] === right[0] && left[1] === right[1];
}

function removeCollinearPoints(points) {
  let current = points;
  let changed = true;

  while (changed && current.length > 3) {
    changed = false;
    let remaining = current.length;
    const next = [];

    for (let index = 0; index < current.length; index += 1) {
      const previous = current[(index - 1 + current.length) % current.length];
      const point = current[index];
      const following = current[(index + 1) % current.length];
      const crossProduct =
        (point[0] - previous[0]) * (following[1] - point[1]) -
        (point[1] - previous[1]) * (following[0] - point[0]);

      if (crossProduct === 0 && remaining > 3) {
        remaining -= 1;
        changed = true;
      } else {
        next.push(point);
      }
    }

    current = next;
  }

  return current;
}

function optimizeRing(ring) {
  pointsBefore += ring.length;

  const quantized = [];
  for (const [longitude, latitude] of ring) {
    const point = [
      Number(longitude.toFixed(decimalPlaces)),
      Number(latitude.toFixed(decimalPlaces)),
    ];
    if (!quantized.length || !samePoint(quantized.at(-1), point)) {
      quantized.push(point);
    }
  }

  if (
    quantized.length > 1 &&
    samePoint(quantized[0], quantized.at(-1))
  ) {
    quantized.pop();
  }

  const optimized =
    quantized.length >= 3 ? removeCollinearPoints(quantized) : quantized;
  const closed =
    optimized.length >= 3 ? [...optimized, [...optimized[0]]] : ring;

  pointsAfter += closed.length;
  return closed;
}

function optimizePolygon(polygon) {
  return polygon.map(optimizeRing);
}

for (const feature of geojson.features) {
  if (feature.geometry.type === "Polygon") {
    feature.geometry.coordinates = optimizePolygon(feature.geometry.coordinates);
  } else if (feature.geometry.type === "MultiPolygon") {
    feature.geometry.coordinates =
      feature.geometry.coordinates.map(optimizePolygon);
  }
}

await writeFile(geojsonPath, JSON.stringify(geojson), "utf8");

console.log(
  `GeoJSON optimizado: ${pointsBefore.toLocaleString()} → ${pointsAfter.toLocaleString()} puntos`,
);
