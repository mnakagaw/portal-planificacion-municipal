"use client";

import { useEffect, useMemo, useState } from "react";
import municipalitiesData from "./data/municipios.json";

type Level = "complete" | "progress" | "warning" | "none" | "unknown";
type MapLayer = "all" | "pmdOfficial" | "pmdDraft" | "cdm" | "ompp";
type StatusLayer = Exclude<MapLayer, "all">;

type Municipality = {
  id: number;
  municipio: string;
  provincia: string;
  region: string;
  sismapUrl: string;
  ompp: { label: string; level: Level; evidenceCount: number };
  cdm: {
    label: string;
    level: Level;
    score: number | null;
    expiry: string;
    validity: string;
    evidenceCount: number;
  };
  pmd: {
    label: string;
    level: Level;
    score: number | null;
    expiry: string;
    validity: string;
    period: string;
    pdfUrl: string;
    pdfTitle: string;
    documentCount: number;
    hasDraft: boolean;
    hasHistorical: boolean;
    hasCurrent: boolean;
    has7_12: boolean;
    has8_12: boolean;
    hasOfficialEvidence: boolean;
    officialUrl: string;
    officialReason: string;
  };
  checkedAt: string;
};

type GeoGeometry = {
  type: "Polygon" | "MultiPolygon";
  coordinates: number[][][] | number[][][][];
};

type GeoFeature = {
  properties: {
    adm2_code: string;
    municipio: string;
    provincia: string;
  };
  geometry: GeoGeometry;
};

type GeoCollection = {
  type: "FeatureCollection";
  features: GeoFeature[];
};

type MapShape = {
  adm2Code: string;
  municipalityKey: string;
  name: string;
  path: string;
};

const municipalities = municipalitiesData as Municipality[];

const regionOrder = [
  "Cibao Noroeste",
  "Cibao Norte",
  "Cibao Nordeste",
  "Cibao Sur",
  "Valdesia",
  "El Valle",
  "Enriquillo",
  "Ozama",
  "Higuamo",
  "Yuma",
];

const layerMeta: Record<
  MapLayer,
  {
    label: string;
    shortLabel: string;
    color: string;
    softColor: string;
    description: string;
  }
> = {
  all: {
    label: "Todos los estados",
    shortLabel: "Todos",
    color: "#31453f",
    softColor: "#edf2f0",
    description: "Color principal según el avance confirmado de cada municipio.",
  },
  pmdOfficial: {
    label: "PMD oficial",
    shortLabel: "PMD Oficial",
    color: "#347c6c",
    softColor: "#e9f2ef",
    description: "SISMAP 2.02 al 100% o evidencia 8-12 confirmada.",
  },
  pmdDraft: {
    label: "PMD borrador",
    shortLabel: "PMD Borrador",
    color: "#5878a3",
    softColor: "#ebf0f6",
    description: "Borrador 7-12 disponible, sin condición oficial confirmada.",
  },
  cdm: {
    label: "CDM constituido",
    shortLabel: "CDM",
    color: "#806d91",
    softColor: "#f0edf2",
    description: "El CDM figura como constituido o institucionalizado.",
  },
  ompp: {
    label: "OMPP establecida",
    shortLabel: "OMPP",
    color: "#b48150",
    softColor: "#f5efe9",
    description: "La OMPP figura como establecida en la evidencia revisada.",
  },
};

const municipalityAliases: Record<string, string> = {
  azuadecompostela: "azua",
  concepciondelavega: "lavega",
  distritonacional: "distritonacional",
  hatomayordelrey: "hatomayor",
  lamata: "villalamata",
  neiba: "neyba",
  sanfelipedepuertoplata: "puertoplata",
  sanjosedelosllanos: "losllanos",
  sanjuandelamaguana: "sanjuan",
  santabarbaradesamana: "samana",
  santodomingodeguzman: "distritonacional",
  santiagodeloscaballeros: "santiago",
  villabisononavarrete: "bisono",
};

function cleanName(value: string) {
  return value
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("es")
    .replace(/[^a-z0-9]/g, "");
}

function canonicalMunicipality(value: string) {
  const cleaned = cleanName(value);
  return municipalityAliases[cleaned] ?? cleaned;
}

function canonicalProvince(value: string) {
  const cleaned = cleanName(value);
  return cleaned === "bahoruco" ? "baoruco" : cleaned;
}

function territoryKey(municipality: string, province: string) {
  return `${canonicalMunicipality(municipality)}|${canonicalProvince(province)}`;
}

function hasLayerStatus(item: Municipality, layer: MapLayer) {
  switch (layer) {
    case "all":
      return true;
    case "pmdOfficial":
      return item.pmd.hasOfficialEvidence;
    case "pmdDraft":
      return (
        item.pmd.hasOfficialEvidence || item.pmd.has7_12 || item.pmd.hasDraft
      );
    case "cdm":
      return item.pmd.hasOfficialEvidence || item.cdm.level === "complete";
    case "ompp":
      return item.pmd.hasOfficialEvidence || item.ompp.level === "complete";
  }
}

function overviewLayer(item: Municipality): StatusLayer | null {
  if (item.pmd.hasOfficialEvidence) return "pmdOfficial";
  if (item.pmd.has7_12 || item.pmd.hasDraft) return "pmdDraft";
  if (item.cdm.level === "complete") return "cdm";
  if (item.ompp.level === "complete") return "ompp";
  return null;
}

function overviewLabel(item: Municipality) {
  const layer = overviewLayer(item);
  return layer ? layerMeta[layer].label : "Sin estado confirmado";
}

function projectPoint([longitude, latitude]: number[]) {
  const x = ((longitude + 72.05) / 3.95) * 1000;
  const y = ((19.95 - latitude) / 2.65) * 670;
  return [x, y];
}

function ringToPath(ring: number[][]) {
  return ring
    .map((point, index) => {
      const [x, y] = projectPoint(point);
      return `${index === 0 ? "M" : "L"}${x.toFixed(2)},${y.toFixed(2)}`;
    })
    .join(" ")
    .concat(" Z");
}

function geometryToPath(geometry: GeoGeometry) {
  const polygons =
    geometry.type === "Polygon"
      ? [geometry.coordinates as number[][][]]
      : (geometry.coordinates as number[][][][]);
  return polygons.flatMap((polygon) => polygon.map(ringToPath)).join(" ");
}

function documentInfo(item: Municipality) {
  if (item.pmd.hasOfficialEvidence) {
    return {
      label: "Abrir PMD oficial",
      note: item.pmd.period || "Período por confirmar",
      url: item.pmd.officialUrl || item.pmd.pdfUrl || item.sismapUrl,
      tone: "official",
    };
  }
  if (item.pmd.has7_12 || item.pmd.hasDraft) {
    return {
      label: "Abrir borrador existente",
      note: "Documento de trabajo para revisión",
      url: item.pmd.pdfUrl || item.sismapUrl,
      tone: "draft",
    };
  }
  return {
    label: "Borrador Word en preparación",
    note: "Se publicará aquí cuando esté listo",
    url: "",
    tone: "pending",
  };
}

function StatusItem({
  item,
  layer,
}: {
  item: Municipality;
  layer: StatusLayer;
}) {
  const active = hasLayerStatus(item, layer);
  const meta = layerMeta[layer];
  const inheritedFromOfficial =
    item.pmd.hasOfficialEvidence &&
    ((layer === "pmdDraft" && !item.pmd.has7_12 && !item.pmd.hasDraft) ||
      (layer === "cdm" && item.cdm.level !== "complete") ||
      (layer === "ompp" && item.ompp.level !== "complete"));
  return (
    <div
      className={`municipal-status ${active ? "is-active" : ""}`}
      style={
        {
          "--status-color": meta.color,
          "--status-soft": meta.softColor,
        } as React.CSSProperties
      }
    >
      <span className="status-symbol" aria-hidden="true">
        {active ? "✓" : "−"}
      </span>
      <span>
        <strong>{meta.shortLabel}</strong>
        <small>
          {active
            ? inheritedFromOfficial
              ? "Incluido en PMD oficial"
              : layer === "cdm"
                ? item.cdm.label
                : layer === "ompp"
                  ? item.ompp.label
                  : layer === "pmdOfficial"
                    ? "Oficial"
                    : "Disponible"
            : "No confirmado"}
        </small>
      </span>
    </div>
  );
}

export function PortalApp() {
  const [region, setRegion] = useState("Todas");
  const [province, setProvince] = useState("Todas");
  const [selected, setSelected] = useState<Municipality | null>(null);
  const [hovered, setHovered] = useState<Municipality | null>(null);
  const [activeLayer, setActiveLayer] = useState<MapLayer>("all");
  const [mapShapes, setMapShapes] = useState<MapShape[]>([]);
  const [mapError, setMapError] = useState(false);

  const municipalityLookup = useMemo(
    () =>
      new Map(
        municipalities.map((item) => [
          territoryKey(item.municipio, item.provincia),
          item,
        ]),
      ),
    [],
  );

  useEffect(() => {
    let active = true;
    fetch("/data/adm2.geojson")
      .then((response) => {
        if (!response.ok) throw new Error(`GeoJSON ${response.status}`);
        return response.json() as Promise<GeoCollection>;
      })
      .then((collection) => {
        if (!active) return;
        setMapShapes(
          collection.features.map((feature) => ({
            adm2Code: feature.properties.adm2_code,
            municipalityKey: territoryKey(
              feature.properties.municipio,
              feature.properties.provincia,
            ),
            name: feature.properties.municipio,
            path: geometryToPath(feature.geometry),
          })),
        );
      })
      .catch(() => {
        if (active) setMapError(true);
      });
    return () => {
      active = false;
    };
  }, []);

  const regions = useMemo(
    () => [
      "Todas",
      ...regionOrder.filter((name) =>
        municipalities.some((item) => item.region === name),
      ),
    ],
    [],
  );

  const provinces = useMemo(() => {
    const items =
      region === "Todas"
        ? municipalities
        : municipalities.filter((item) => item.region === region);
    return [
      "Todas",
      ...Array.from(new Set(items.map((item) => item.provincia))).sort((a, b) =>
        a.localeCompare(b, "es"),
      ),
    ];
  }, [region]);

  const municipalityOptions = useMemo(
    () =>
      municipalities
        .filter((item) => region === "Todas" || item.region === region)
        .filter((item) => province === "Todas" || item.provincia === province)
        .sort((a, b) => a.municipio.localeCompare(b.municipio, "es")),
    [province, region],
  );

  const filteredIds = useMemo(
    () => new Set(municipalityOptions.map((item) => item.id)),
    [municipalityOptions],
  );

  const mappedIds = useMemo(
    () =>
      new Set(
        mapShapes
          .map((shape) => municipalityLookup.get(shape.municipalityKey)?.id)
          .filter((id): id is number => typeof id === "number"),
      ),
    [mapShapes, municipalityLookup],
  );

  const unmappedCount = municipalities.filter(
    (item) => !mappedIds.has(item.id),
  ).length;

  const layerCounts = useMemo(
    () =>
      Object.fromEntries(
        (Object.keys(layerMeta) as MapLayer[]).map((layer) => [
          layer,
          municipalities.filter((item) => hasLayerStatus(item, layer)).length,
        ]),
      ) as Record<MapLayer, number>,
    [],
  );

  const displayMunicipality = hovered ?? selected;
  const activeMeta = layerMeta[activeLayer];
  const selectedDocument = selected ? documentInfo(selected) : null;

  const chooseMunicipality = (item: Municipality) => {
    setRegion(item.region);
    setProvince(item.provincia);
    setSelected(item);
  };

  return (
    <main className="portal">
      <header className="app-header">
        <div className="brand">
          <span className="brand-mark" aria-hidden="true">
            PM
          </span>
          <span>
            <strong>Planificación Municipal</strong>
            <small>República Dominicana</small>
          </span>
        </div>
        <p>Estado de 162 municipios · 27 jul 2026</p>
      </header>

      <section className="page-heading">
        <div>
          <span>Portal municipal</span>
          <h1>Seleccione un municipio</h1>
        </div>
        <p>
          Consulte el PMD, el borrador disponible y la situación del CDM y la
          OMPP.
        </p>
      </section>

      <section className="layer-switcher" aria-label="Información mostrada en el mapa">
        {(Object.keys(layerMeta) as MapLayer[]).map((layer) => {
          const meta = layerMeta[layer];
          return (
            <button
              key={layer}
                className={`${activeLayer === layer ? "is-selected" : ""} ${
                  layer === "all" ? "layer-all" : ""
                }`}
              onClick={() => setActiveLayer(layer)}
              style={
                {
                  "--layer-color": meta.color,
                  "--layer-soft": meta.softColor,
                } as React.CSSProperties
              }
            >
              <span className="layer-dot" aria-hidden="true" />
              <span>
                <small>{meta.shortLabel}</small>
                <strong>{layerCounts[layer]}</strong>
              </span>
            </button>
          );
        })}
      </section>

      <section className="selection-bar" aria-label="Selección territorial">
        <label>
          <span>Región</span>
          <select
            value={region}
            onChange={(event) => {
              setRegion(event.target.value);
              setProvince("Todas");
              setSelected(null);
            }}
          >
            {regions.map((item) => (
              <option key={item} value={item}>
                {item}
              </option>
            ))}
          </select>
        </label>

        <span className="selection-arrow" aria-hidden="true">
          →
        </span>

        <label>
          <span>Provincia</span>
          <select
            value={province}
            onChange={(event) => {
              setProvince(event.target.value);
              setSelected(null);
            }}
          >
            {provinces.map((item) => (
              <option key={item} value={item}>
                {item}
              </option>
            ))}
          </select>
        </label>

        <span className="selection-arrow" aria-hidden="true">
          →
        </span>

        <label>
          <span>Municipio</span>
          <select
            value={selected?.id ?? ""}
            onChange={(event) => {
              const id = Number(event.target.value);
              const item = municipalities.find((municipality) => municipality.id === id);
              if (item) chooseMunicipality(item);
              else setSelected(null);
            }}
          >
            <option value="">Seleccione un municipio</option>
            {municipalityOptions.map((item) => (
              <option key={item.id} value={item.id}>
                {item.municipio}
              </option>
            ))}
          </select>
        </label>
      </section>

      <section className="workspace">
        <article className="map-panel">
          <div className="map-header">
            <div>
              <span
                className={`map-layer-dot ${
                  activeLayer === "all" ? "layer-all-dot" : ""
                }`}
                style={
                  activeLayer === "all"
                    ? undefined
                    : { background: activeMeta.color }
                }
                aria-hidden="true"
              />
              <span>
                <strong>{activeMeta.label}</strong>
                <small>{activeMeta.description}</small>
              </span>
            </div>
            <p>Haga clic en el mapa para elegir un municipio</p>
          </div>

          <div className="map-canvas">
            {mapError ? (
              <div className="map-message">
                No fue posible cargar la cartografía.
              </div>
            ) : mapShapes.length === 0 ? (
              <div className="map-message">Cargando mapa…</div>
            ) : (
              <svg
                className="dominican-map"
                viewBox="0 0 1000 670"
                role="img"
                aria-label="Mapa de municipios de la República Dominicana"
              >
                <g>
                  {mapShapes.map((shape) => {
                    const item = municipalityLookup.get(shape.municipalityKey);
                    const included = item ? filteredIds.has(item.id) : false;
                    const active = item ? hasLayerStatus(item, activeLayer) : false;
                    const isSelected = item?.id === selected?.id;
                    const overview = item ? overviewLayer(item) : null;
                    const fill = !included
                      ? "#eef1f0"
                      : activeLayer === "all"
                        ? overview
                          ? layerMeta[overview].color
                          : "#d9dfdd"
                        : active
                          ? activeMeta.color
                          : "#d9dfdd";
                    return (
                      <path
                        key={shape.adm2Code}
                        d={shape.path}
                        fill={fill}
                        fillOpacity={included ? 1 : 0.58}
                        fillRule="evenodd"
                        stroke={isSelected ? "#152f2a" : "#ffffff"}
                        strokeWidth={isSelected ? 3.6 : 1.15}
                        vectorEffect="non-scaling-stroke"
                        className={item ? "map-shape" : "map-shape is-muted"}
                        onMouseEnter={() => item && setHovered(item)}
                        onMouseLeave={() => setHovered(null)}
                        onFocus={() => item && setHovered(item)}
                        onBlur={() => setHovered(null)}
                        onClick={() => item && chooseMunicipality(item)}
                        tabIndex={item ? 0 : -1}
                        aria-label={
                          item
                            ? `${item.municipio}, ${
                                activeLayer === "all"
                                  ? overviewLabel(item)
                                  : hasLayerStatus(item, activeLayer)
                                    ? activeMeta.label
                                    : "no confirmado"
                              }`
                            : shape.name
                        }
                      >
                        <title>
                          {item
                            ? `${item.municipio} · ${
                                activeLayer === "all"
                                  ? overviewLabel(item)
                                  : hasLayerStatus(item, activeLayer)
                                    ? activeMeta.label
                                    : "No confirmado"
                              }`
                            : shape.name}
                        </title>
                      </path>
                    );
                  })}
                </g>
              </svg>
            )}

            {displayMunicipality && (
              <div className="map-tooltip" aria-live="polite">
                <small>{displayMunicipality.provincia}</small>
                <strong>{displayMunicipality.municipio}</strong>
                <span>
                  {activeLayer === "all"
                    ? overviewLabel(displayMunicipality)
                    : hasLayerStatus(displayMunicipality, activeLayer)
                      ? activeMeta.label
                      : "No confirmado"}
                </span>
              </div>
            )}

            <div className="map-legend">
              {activeLayer === "all" ? (
                <>
                  {(Object.keys(layerMeta) as MapLayer[])
                    .filter((layer): layer is StatusLayer => layer !== "all")
                    .map((layer) => (
                      <span key={layer}>
                        <i style={{ background: layerMeta[layer].color }} />
                        {layerMeta[layer].shortLabel}
                      </span>
                    ))}
                  <span>
                    <i className="legend-no" />
                    Sin confirmar
                  </span>
                </>
              ) : (
                <>
                  <span>
                    <i style={{ background: activeMeta.color }} />
                    Sí
                  </span>
                  <span>
                    <i className="legend-no" />
                    No confirmado
                  </span>
                </>
              )}
            </div>
          </div>

          <div className="map-footer">
            <span>Cartografía disponible para 158 municipios</span>
            {unmappedCount > 0 && (
              <span>{unmappedCount} municipios nuevos disponibles en el selector</span>
            )}
          </div>
        </article>

        <aside className="detail-panel">
          {selected ? (
            <>
              <div className="detail-heading">
                <span>{selected.region}</span>
                <h2>{selected.municipio}</h2>
                <p>{selected.provincia}</p>
              </div>

              <div
                className={`sismap-score ${
                  selected.pmd.score === 100 ? "is-complete" : ""
                }`}
                aria-label={
                  selected.pmd.score === null
                    ? "Puntuación del PMD en SISMAP no disponible"
                    : `Puntuación del PMD en SISMAP: ${selected.pmd.score}%`
                }
                style={
                  {
                    "--sismap-score": `${selected.pmd.score ?? 0}%`,
                  } as React.CSSProperties
                }
              >
                <span>
                  <small>Puntuación PMD en SISMAP</small>
                  <strong>
                    {selected.pmd.score === null
                      ? "No disponible"
                      : `${selected.pmd.score}%`}
                  </strong>
                </span>
                <span className="sismap-score-track" aria-hidden="true">
                  <i />
                </span>
              </div>

              <div className="status-grid">
                {(Object.keys(layerMeta) as MapLayer[])
                  .filter((layer): layer is StatusLayer => layer !== "all")
                  .map((layer) => (
                    <StatusItem key={layer} item={selected} layer={layer} />
                  ))}
              </div>

              <section className="document-card">
                <span>Documento</span>
                <strong>{selectedDocument?.label}</strong>
                <small>{selectedDocument?.note}</small>
                {selectedDocument?.url ? (
                  <a
                    href={selectedDocument.url}
                    target="_blank"
                    rel="noreferrer"
                    className={`document-action action-${selectedDocument.tone}`}
                  >
                    {selectedDocument.label}
                    <span aria-hidden="true">↗</span>
                  </a>
                ) : (
                  <button className="document-action" disabled>
                    Próximamente
                  </button>
                )}
              </section>

              <a
                className="sismap-link"
                href={selected.sismapUrl}
                target="_blank"
                rel="noreferrer"
              >
                Ver evidencias en SISMAP
              </a>
              <p className="checked-date">
                Verificado: {selected.checkedAt || "fecha no disponible"}
              </p>
            </>
          ) : (
            <div className="empty-detail">
              <span aria-hidden="true">⌖</span>
              <h2>Elija un municipio</h2>
              <p>
                Use los selectores o haga clic en el mapa para ver sus cuatro
                estados y el documento disponible.
              </p>
            </div>
          )}
        </aside>
      </section>

      <footer>
        <p>
          La condición de PMD oficial y la vigencia del período se revisan por
          separado. Los borradores Word se incorporarán de forma progresiva.
        </p>
        <span>Fuente de estado: SISMAP Municipal</span>
      </footer>
    </main>
  );
}
