"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import municipalLinksData from "./data/municipal-links.json";
import municipalitiesData from "./data/municipios.json";

type Level = "complete" | "progress" | "warning" | "none" | "unknown";
type MapLayer = "all" | "pmdOfficial" | "pmdDraft" | "cdm" | "ompp";
type StatusLayer = Exclude<MapLayer, "all">;

const SISMAP_PMD_URL =
  "https://www.sismap.gob.do/municipal/ranking/listaevidenciasorganismos/16?catchall=2.02-Plan-de-De&tipoId=17";

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
    generatedDraftUrl?: string;
    generatedDraftPeriod?: string;
    generatedDraftGeneratedAt?: string;
  };
  checkedAt: string;
};

type MunicipalLink = {
  id: number;
  officialWebsiteUrl: string;
  wikipediaUrl: string;
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
  bounds: MapBounds;
};

type MapBounds = {
  minX: number;
  minY: number;
  maxX: number;
  maxY: number;
};

const municipalities = municipalitiesData as Municipality[];
const municipalLinks = municipalLinksData as MunicipalLink[];
const municipalLinksLookup = new Map(
  municipalLinks.map((item) => [item.id, item]),
);

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
    textColor: string;
    description: string;
  }
> = {
  all: {
    label: "Todos los estados",
    shortLabel: "Todos",
    color: "#31453f",
    softColor: "#edf2f0",
    textColor: "#31453f",
    description: "Color principal según el avance confirmado de cada municipio.",
  },
  pmdOfficial: {
    label: "PMD oficial",
    shortLabel: "PMD Oficial",
    color: "#286A50",
    softColor: "#E8F0EB",
    textColor: "#286A50",
    description: "SISMAP 2.02 al 100% o evidencia 8-12 confirmada.",
  },
  pmdDraft: {
    label: "PMD borrador en SISMAP",
    shortLabel: "Borrador SISMAP",
    color: "#75A67E",
    softColor: "#EEF4EF",
    textColor: "#486D50",
    description: "Borrador 7-12 disponible, sin condición oficial confirmada.",
  },
  cdm: {
    label: "CDM constituido",
    shortLabel: "CDM",
    color: "#527E91",
    softColor: "#ECF1F3",
    textColor: "#456B7B",
    description: "El CDM figura como constituido o institucionalizado.",
  },
  ompp: {
    label: "OMPP establecida",
    shortLabel: "OMPP",
    color: "#8FADBA",
    softColor: "#F1F5F6",
    textColor: "#53727E",
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

function hasFormalPmd(item: Municipality) {
  return (
    item.pmd.hasOfficialEvidence || item.pmd.has7_12 || item.pmd.hasDraft
  );
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
      return hasFormalPmd(item) || item.cdm.level === "complete";
    case "ompp":
      return hasFormalPmd(item) || item.ompp.level === "complete";
  }
}

function matchesLayer(item: Municipality, layer: MapLayer) {
  switch (layer) {
    case "all":
      return true;
    case "pmdOfficial":
      return item.pmd.hasOfficialEvidence;
    case "pmdDraft":
      return hasLayerStatus(item, "pmdDraft");
    case "cdm":
      return hasLayerStatus(item, "cdm");
    case "ompp":
      return hasLayerStatus(item, "ompp");
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

function geometryBounds(geometry: GeoGeometry): MapBounds {
  const polygons =
    geometry.type === "Polygon"
      ? [geometry.coordinates as number[][][]]
      : (geometry.coordinates as number[][][][]);
  const bounds: MapBounds = {
    minX: Number.POSITIVE_INFINITY,
    minY: Number.POSITIVE_INFINITY,
    maxX: Number.NEGATIVE_INFINITY,
    maxY: Number.NEGATIVE_INFINITY,
  };

  for (const polygon of polygons) {
    for (const ring of polygon) {
      for (const point of ring) {
        const [x, y] = projectPoint(point);
        bounds.minX = Math.min(bounds.minX, x);
        bounds.minY = Math.min(bounds.minY, y);
        bounds.maxX = Math.max(bounds.maxX, x);
        bounds.maxY = Math.max(bounds.maxY, y);
      }
    }
  }

  return bounds;
}

function boundsToViewBox(bounds: MapBounds) {
  const mapRatio = 1000 / 670;
  const horizontalPadding = Math.max((bounds.maxX - bounds.minX) * 0.08, 18);
  const verticalPadding = Math.max((bounds.maxY - bounds.minY) * 0.08, 18);
  let minX = bounds.minX - horizontalPadding;
  let minY = bounds.minY - verticalPadding;
  let width = bounds.maxX - bounds.minX + horizontalPadding * 2;
  let height = bounds.maxY - bounds.minY + verticalPadding * 2;

  if (width / height < mapRatio) {
    const expandedWidth = height * mapRatio;
    minX -= (expandedWidth - width) / 2;
    width = expandedWidth;
  } else {
    const expandedHeight = width / mapRatio;
    minY -= (expandedHeight - height) / 2;
    height = expandedHeight;
  }

  return [minX, minY, width, height]
    .map((value) => value.toFixed(2))
    .join(" ");
}

function documentInfo(item: Municipality) {
  if (item.pmd.hasOfficialEvidence) {
    if (!item.pmd.officialUrl) {
      return {
        heading: "Evidencias del PMD",
        label: "Ver evidencias en SISMAP",
        note: "PMD completo no identificado en los archivos disponibles",
        url: item.sismapUrl,
        tone: "official",
        download: false,
      };
    }
    return {
      heading: "PMD oficial",
      label: "Abrir PMD oficial",
      note: item.pmd.period || "Período por confirmar",
      url: item.pmd.officialUrl,
      tone: "official",
      download: false,
    };
  }
  if (item.pmd.has7_12 || item.pmd.hasDraft) {
    return {
      heading: "PMD borrador en SISMAP",
      label: "Abrir PMD borrador en SISMAP",
      note: "Documento de trabajo para revisión",
      url: item.pmd.pdfUrl || item.sismapUrl,
      tone: "draft",
      download: false,
    };
  }
  if (item.pmd.generatedDraftUrl) {
    return {
      heading: "Documento base para elaborar el PMD",
      label: "Descargar documento base del PMD (Word)",
      note: `Información general y diagnóstico preelaborados para revisión de la OMPP y el CDM · ${item.pmd.generatedDraftPeriod || "2025-2028"}`,
      url: item.pmd.generatedDraftUrl,
      tone: "base",
      download: true,
    };
  }
  return {
    heading: "Documento base para elaborar el PMD",
    label: "Documento base en preparación",
    note: "Se publicará aquí cuando esté listo",
    url: "",
    tone: "pending",
    download: false,
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
  const inheritedFromPmd =
    (item.pmd.hasOfficialEvidence &&
      layer === "pmdDraft" &&
      !item.pmd.has7_12 &&
      !item.pmd.hasDraft) ||
    (hasFormalPmd(item) &&
      ((layer === "cdm" && item.cdm.level !== "complete") ||
        (layer === "ompp" && item.ompp.level !== "complete")));
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
            ? inheritedFromPmd
              ? layer === "pmdDraft"
                ? "Incluido en PMD oficial"
                : "Requisito del PMD · evidencia por verificar"
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

function RegionMultiSelect({
  regions,
  selectedRegions,
  onChange,
}: {
  regions: string[];
  selectedRegions: string[];
  onChange: (regions: string[]) => void;
}) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const closeOnOutsideClick = (event: PointerEvent) => {
      if (!rootRef.current?.contains(event.target as Node)) setOpen(false);
    };
    const closeOnEscape = (event: KeyboardEvent) => {
      if (event.key === "Escape") setOpen(false);
    };

    document.addEventListener("pointerdown", closeOnOutsideClick);
    document.addEventListener("keydown", closeOnEscape);
    return () => {
      document.removeEventListener("pointerdown", closeOnOutsideClick);
      document.removeEventListener("keydown", closeOnEscape);
    };
  }, []);

  const label =
    selectedRegions.length === 0
      ? "Todas las regiones"
      : selectedRegions.length === 1
        ? selectedRegions[0]
        : `${selectedRegions.length} regiones seleccionadas`;

  const toggleRegion = (region: string, checked: boolean) => {
    const next = checked
      ? [...selectedRegions, region]
      : selectedRegions.filter((item) => item !== region);
    onChange(regions.filter((item) => next.includes(item)));
  };

  return (
    <div className="region-multiselect" ref={rootRef}>
      <button
        type="button"
        className="region-multiselect-button"
        aria-expanded={open}
        aria-haspopup="true"
        onClick={() => setOpen((current) => !current)}
      >
        <span>{label}</span>
        <span aria-hidden="true">⌄</span>
      </button>
      {open && (
        <div className="region-multiselect-menu" role="group" aria-label="Regiones">
          <label className="region-option region-option-all">
            <input
              type="checkbox"
              checked={selectedRegions.length === 0}
              onChange={() => onChange([])}
            />
            <span>Todas las regiones</span>
          </label>
          {regions.map((region) => (
            <label className="region-option" key={region}>
              <input
                type="checkbox"
                checked={selectedRegions.includes(region)}
                onChange={(event) => toggleRegion(region, event.target.checked)}
              />
              <span>{region}</span>
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

export function PortalApp() {
  const [selectedRegions, setSelectedRegions] = useState<string[]>([]);
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
    fetch(`${import.meta.env.BASE_URL}data/adm2.geojson`)
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
            bounds: geometryBounds(feature.geometry),
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
    () =>
      regionOrder.filter((name) =>
        municipalities.some((item) => item.region === name),
      ),
    [],
  );

  const provinces = useMemo(() => {
    const items =
      selectedRegions.length === 0
        ? municipalities
        : municipalities.filter((item) => selectedRegions.includes(item.region));
    return [
      "Todas",
      ...Array.from(new Set(items.map((item) => item.provincia))).sort((a, b) =>
        a.localeCompare(b, "es"),
      ),
    ];
  }, [selectedRegions]);

  const municipalityOptions = useMemo(
    () =>
      municipalities
        .filter(
          (item) =>
            selectedRegions.length === 0 || selectedRegions.includes(item.region),
        )
        .filter((item) => province === "Todas" || item.provincia === province)
        .sort((a, b) => a.municipio.localeCompare(b.municipio, "es")),
    [province, selectedRegions],
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

  const mapViewBox = useMemo(() => {
    let viewportRegions = selected ? [selected.region] : selectedRegions;

    if (!selected && province !== "Todas") {
      viewportRegions = Array.from(
        new Set(
          municipalities
            .filter((item) => item.provincia === province)
            .map((item) => item.region),
        ),
      );
    }

    if (viewportRegions.length === 0) return "0 0 1000 670";

    const regionSet = new Set(viewportRegions);
    const targetShapes = mapShapes.filter((shape) => {
      const item = municipalityLookup.get(shape.municipalityKey);
      return item ? regionSet.has(item.region) : false;
    });
    if (targetShapes.length === 0) return "0 0 1000 670";

    const bounds = targetShapes.reduce<MapBounds>(
      (combined, shape) => ({
        minX: Math.min(combined.minX, shape.bounds.minX),
        minY: Math.min(combined.minY, shape.bounds.minY),
        maxX: Math.max(combined.maxX, shape.bounds.maxX),
        maxY: Math.max(combined.maxY, shape.bounds.maxY),
      }),
      {
        minX: Number.POSITIVE_INFINITY,
        minY: Number.POSITIVE_INFINITY,
        maxX: Number.NEGATIVE_INFINITY,
        maxY: Number.NEGATIVE_INFINITY,
      },
    );

    return boundsToViewBox(bounds);
  }, [
    mapShapes,
    municipalityLookup,
    province,
    selected,
    selectedRegions,
  ]);

  const layerCounts = useMemo(
    () =>
      Object.fromEntries(
        (Object.keys(layerMeta) as MapLayer[]).map((layer) => [
          layer,
          municipalities.filter((item) => matchesLayer(item, layer)).length,
        ]),
      ) as Record<MapLayer, number>,
    [],
  );

  const displayMunicipality = hovered ?? selected;
  const activeMeta = layerMeta[activeLayer];
  const selectedDocument = selected ? documentInfo(selected) : null;
  const selectedLinks = selected
    ? municipalLinksLookup.get(selected.id)
    : undefined;

  const chooseMunicipality = (item: Municipality) => {
    setSelectedRegions([item.region]);
    setProvince("Todas");
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
        <div className="header-meta">
          <a href={SISMAP_PMD_URL} target="_blank" rel="noreferrer">
            Fuente SISMAP 2.02 ↗
          </a>
          <span>Estado de 162 municipios · 28 jul 2026</span>
        </div>
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
                  "--layer-text": meta.textColor,
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
        <div className="selection-control">
          <span>Región</span>
          <RegionMultiSelect
            regions={regions}
            selectedRegions={selectedRegions}
            onChange={(regions) => {
              setSelectedRegions(regions);
              setProvince("Todas");
              setSelected(null);
            }}
          />
        </div>

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
                viewBox={mapViewBox}
                role="img"
                aria-label="Mapa de municipios de la República Dominicana"
              >
                <g>
                  {mapShapes.map((shape) => {
                    const item = municipalityLookup.get(shape.municipalityKey);
                    const included = item ? filteredIds.has(item.id) : false;
                    const active = item ? matchesLayer(item, activeLayer) : false;
                    const isSelected = item?.id === selected?.id;
                    const overview = item ? overviewLayer(item) : null;
                    const fill = !included
                      ? "#EEF1F0"
                      : activeLayer === "all"
                        ? overview
                          ? layerMeta[overview].color
                          : "#F2E5BF"
                        : active
                          ? activeMeta.color
                          : "#F2E5BF";
                    return (
                      <path
                        key={shape.adm2Code}
                        d={shape.path}
                        fill={fill}
                        fillOpacity={included ? 1 : 0.58}
                        fillRule="evenodd"
                        stroke={isSelected ? "#234A3D" : "rgba(255, 255, 255, 0.68)"}
                        strokeWidth={isSelected ? 2 : 0.9}
                        strokeLinejoin="round"
                        vectorEffect="non-scaling-stroke"
                        className={`${item ? "map-shape" : "map-shape is-muted"} ${
                          isSelected ? "is-selected" : ""
                        }`}
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
                                  : matchesLayer(item, activeLayer)
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
                                  : matchesLayer(item, activeLayer)
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
                    : matchesLayer(displayMunicipality, activeLayer)
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
                      <span
                        key={layer}
                        className={layer === "cdm" ? "legend-group-start" : undefined}
                      >
                        <i style={{ background: layerMeta[layer].color }} />
                        {layerMeta[layer].shortLabel}
                      </span>
                    ))}
                  <span className="legend-group-start">
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
            <span>
              Cartografía IGN/ONE disponible para {mappedIds.size} municipios
            </span>
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
                <span>{selectedDocument?.heading}</span>
                <strong>{selectedDocument?.label}</strong>
                <small>{selectedDocument?.note}</small>
                {selectedDocument?.url ? (
                  <a
                    href={selectedDocument.url}
                    target={selectedDocument.download ? undefined : "_blank"}
                    rel={selectedDocument.download ? undefined : "noreferrer"}
                    download={selectedDocument.download || undefined}
                    className={`document-action action-${selectedDocument.tone}`}
                  >
                    {selectedDocument.label}
                    <span aria-hidden="true">
                      {selectedDocument.download ? "↓" : "↗"}
                    </span>
                  </a>
                ) : (
                  <button className="document-action" disabled>
                    Próximamente
                  </button>
                )}
              </section>

              <div className="external-links">
                {selectedLinks?.officialWebsiteUrl && (
                  <a
                    className="external-link municipal-site-link"
                    href={selectedLinks.officialWebsiteUrl}
                    target="_blank"
                    rel="noreferrer"
                  >
                    Sitio web oficial del municipio
                    <span aria-hidden="true">↗</span>
                  </a>
                )}
                <a
                  className="external-link sismap-link"
                  href={selected.sismapUrl}
                  target="_blank"
                  rel="noreferrer"
                >
                  Ver evidencias en SISMAP
                  <span aria-hidden="true">↗</span>
                </a>
                {selectedLinks?.wikipediaUrl && (
                  <a
                    className="external-link wikipedia-link"
                    href={selectedLinks.wikipediaUrl}
                    target="_blank"
                    rel="noreferrer"
                  >
                    Ver municipio en Wikipedia
                    <span aria-hidden="true">↗</span>
                  </a>
                )}
              </div>
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
          separado. Los municipios sin documento oficial o borrador SISMAP
          disponen de un documento base Word para revisión de la OMPP y el CDM.
        </p>
        <span>Fuente de estado: SISMAP Municipal</span>
      </footer>
    </main>
  );
}
