"use client";

import { useEffect, useMemo, useState } from "react";
import municipalitiesData from "./data/municipios.json";

type Level = "complete" | "progress" | "warning" | "none" | "unknown";
type PmdState = "current" | "draft" | "historical" | "progress" | "none";
type Scope = "all" | PmdState;

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
    hasFormalBody: boolean;
    hasHistorical: boolean;
    hasCurrent: boolean;
  };
  action: "ver" | "continuar" | "actualizar" | "elaborar";
  actionLabel: string;
  checkedAt: string;
  dataQuality: string;
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
    region: string;
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
  province: string;
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

const statusMeta: Record<
  PmdState,
  { label: string; shortLabel: string; color: string; description: string }
> = {
  current: {
    label: "PMD vigente",
    shortLabel: "Vigente",
    color: "#167b67",
    description: "PMD con período actualmente vigente y documento consultable.",
  },
  draft: {
    label: "Borrador disponible",
    shortLabel: "Borrador",
    color: "#3f6fb6",
    description: "Documento de trabajo que aún requiere revisión o validación.",
  },
  historical: {
    label: "PMD histórico / vencido",
    shortLabel: "Histórico",
    color: "#c57632",
    description: "Sirve como insumo para actualizar, pero no cuenta como PMD vigente.",
  },
  progress: {
    label: "En elaboración",
    shortLabel: "En proceso",
    color: "#8c6a24",
    description: "Hay evidencias de avance, sin PMD vigente confirmado.",
  },
  none: {
    label: "Sin PMD confirmado",
    shortLabel: "Sin PMD",
    color: "#aeb9b6",
    description: "No se confirmó un PMD vigente ni un expediente suficiente.",
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
  sanjosédelosllanos: "losllanos",
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

function getPmdState(item: Municipality): PmdState {
  if (item.action === "ver" && item.pmd.hasCurrent && item.pmd.pdfUrl) return "current";
  if (item.pmd.hasDraft) return "draft";
  if (item.action === "actualizar" || item.pmd.hasHistorical) return "historical";
  if (item.action === "continuar" || item.pmd.documentCount > 0) return "progress";
  return "none";
}

function isCurrentPmd(item: Municipality) {
  return getPmdState(item) === "current";
}

function statusFor(item: Municipality) {
  return statusMeta[getPmdState(item)];
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

function StatusBadge({ item }: { item: Municipality }) {
  const state = getPmdState(item);
  const meta = statusMeta[state];
  return (
    <span className={`status-badge status-${state}`}>
      <span style={{ background: meta.color }} aria-hidden="true" />
      {meta.label}
    </span>
  );
}

function SourceCard({
  label,
  title,
  note,
  href,
}: {
  label: string;
  title: string;
  note: string;
  href?: string;
}) {
  const content = (
    <>
      <span>{label}</span>
      <strong>{title}</strong>
      <small>{note}</small>
    </>
  );
  return href ? (
    <a className="source-card" href={href} target="_blank" rel="noreferrer">
      {content}
    </a>
  ) : (
    <article className="source-card">{content}</article>
  );
}

export function PortalApp() {
  const [region, setRegion] = useState("Todas");
  const [province, setProvince] = useState("Todas");
  const [query, setQuery] = useState("");
  const [scope, setScope] = useState<Scope>("all");
  const [selected, setSelected] = useState<Municipality | null>(null);
  const [hovered, setHovered] = useState<Municipality | null>(null);
  const [wizard, setWizard] = useState<Municipality | null>(null);
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
            province: feature.properties.provincia,
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
    const source =
      region === "Todas"
        ? municipalities
        : municipalities.filter((item) => item.region === region);
    return [
      "Todas",
      ...Array.from(new Set(source.map((item) => item.provincia))).sort(),
    ];
  }, [region]);

  const filtered = useMemo(() => {
    const normalized = cleanName(query.trim());
    return municipalities.filter((item) => {
      if (region !== "Todas" && item.region !== region) return false;
      if (province !== "Todas" && item.provincia !== province) return false;
      if (
        normalized &&
        !cleanName(`${item.municipio} ${item.provincia} ${item.region}`).includes(
          normalized,
        )
      )
        return false;
      if (scope !== "all" && getPmdState(item) !== scope) return false;
      return true;
    });
  }, [province, query, region, scope]);

  const filteredIds = useMemo(
    () => new Set(filtered.map((item) => item.id)),
    [filtered],
  );

  const stats = useMemo(
    () => ({
      total: filtered.length,
      current: filtered.filter(isCurrentPmd).length,
      draft: filtered.filter((item) => getPmdState(item) === "draft").length,
      historical: filtered.filter(
        (item) => getPmdState(item) === "historical",
      ).length,
      needsWork: filtered.filter(
        (item) =>
          getPmdState(item) === "progress" || getPmdState(item) === "none",
      ).length,
    }),
    [filtered],
  );

  const mappedMunicipalityIds = useMemo(
    () =>
      new Set(
        mapShapes
          .map((shape) => municipalityLookup.get(shape.municipalityKey)?.id)
          .filter((id): id is number => typeof id === "number"),
      ),
    [mapShapes, municipalityLookup],
  );

  const unmapped = useMemo(
    () => municipalities.filter((item) => !mappedMunicipalityIds.has(item.id)),
    [mappedMunicipalityIds],
  );

  const displayMunicipality = hovered ?? selected;

  const handleAction = (item: Municipality) => {
    if (isCurrentPmd(item) && item.pmd.pdfUrl) {
      window.open(item.pmd.pdfUrl, "_blank", "noopener,noreferrer");
      return;
    }
    setWizard(item);
  };

  const resetFilters = () => {
    setRegion("Todas");
    setProvince("Todas");
    setQuery("");
    setScope("all");
  };

  return (
    <main className="portal">
      <header className="app-header">
        <a className="brand" href="#" aria-label="Inicio">
          <span className="brand-mark" aria-hidden="true">
            PM
          </span>
          <span>
            <strong>Planificación Municipal</strong>
            <small>República Dominicana</small>
          </span>
        </a>
        <div className="header-note">
          <span className="live-dot" aria-hidden="true" />
          Datos revisados · 27 jul 2026
        </div>
      </header>

      <section className="intro">
        <div>
          <span className="eyebrow">162 municipios · OMPP · CDM · PMD</span>
          <h1>El mapa de la planificación municipal.</h1>
        </div>
        <p>
          Explore el territorio y abra el camino correcto: consultar un PMD
          vigente, continuar un borrador, actualizar un plan anterior o iniciar
          el Paquete Mínimo.
        </p>
      </section>

      <section className="summary-row" aria-label="Resumen de la selección">
        <article>
          <span>Municipios</span>
          <strong>{stats.total}</strong>
        </article>
        <article className="summary-current">
          <span>PMD vigente</span>
          <strong>{stats.current}</strong>
          <small>Solo período actual</small>
        </article>
        <article className="summary-draft">
          <span>Borrador</span>
          <strong>{stats.draft}</strong>
        </article>
        <article className="summary-historical">
          <span>Histórico</span>
          <strong>{stats.historical}</strong>
        </article>
        <article>
          <span>Requiere trabajo</span>
          <strong>{stats.needsWork}</strong>
        </article>
      </section>

      <section className="map-workspace">
        <aside className="control-panel">
          <div className="panel-title">
            <span>Explorar</span>
            <button onClick={resetFilters}>Restablecer</button>
          </div>

          <label>
            <span>Buscar municipio</span>
            <input
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="Ej. Moca"
            />
          </label>
          <label>
            <span>Región</span>
            <select
              value={region}
              onChange={(event) => {
                setRegion(event.target.value);
                setProvince("Todas");
              }}
            >
              {regions.map((item) => (
                <option key={item}>{item}</option>
              ))}
            </select>
          </label>
          <label>
            <span>Provincia</span>
            <select
              value={province}
              onChange={(event) => setProvince(event.target.value)}
            >
              {provinces.map((item) => (
                <option key={item}>{item}</option>
              ))}
            </select>
          </label>

          <div className="status-filter" aria-label="Filtrar por estado del PMD">
            <span>Estado del PMD</span>
            <button
              className={scope === "all" ? "active" : ""}
              onClick={() => setScope("all")}
            >
              <i className="all-status" />
              Todos
              <b>{municipalities.length}</b>
            </button>
            {(Object.keys(statusMeta) as PmdState[]).map((state) => (
              <button
                key={state}
                className={scope === state ? "active" : ""}
                onClick={() => setScope(state)}
              >
                <i style={{ background: statusMeta[state].color }} />
                {statusMeta[state].shortLabel}
                <b>
                  {
                    municipalities.filter(
                      (item) => getPmdState(item) === state,
                    ).length
                  }
                </b>
              </button>
            ))}
          </div>
        </aside>

        <article className="map-panel">
          <div className="map-toolbar">
            <div>
              <span>República Dominicana</span>
              <strong>{filtered.length} municipios en la selección</strong>
            </div>
            <div className="map-hint">Seleccione un municipio en el mapa</div>
          </div>

          <div className="map-canvas">
            {mapError ? (
              <div className="map-message">
                No fue posible cargar la cartografía.
              </div>
            ) : mapShapes.length === 0 ? (
              <div className="map-message">Cargando límites municipales…</div>
            ) : (
              <svg
                className="dominican-map"
                viewBox="0 0 1000 670"
                role="img"
                aria-label="Mapa interactivo de municipios de República Dominicana"
              >
                <g>
                  {mapShapes.map((shape) => {
                    const item = municipalityLookup.get(shape.municipalityKey);
                    const included = item ? filteredIds.has(item.id) : false;
                    const isSelected = item?.id === selected?.id;
                    const meta = item ? statusFor(item) : statusMeta.none;
                    return (
                      <path
                        key={shape.adm2Code}
                        d={shape.path}
                        fill={included ? meta.color : "#e8eeec"}
                        fillOpacity={included ? 0.88 : 0.5}
                        fillRule="evenodd"
                        stroke={isSelected ? "#102b27" : "#ffffff"}
                        strokeWidth={isSelected ? 3.4 : 1.15}
                        vectorEffect="non-scaling-stroke"
                        className={item ? "map-shape" : "map-shape map-shape-muted"}
                        onMouseEnter={() => item && setHovered(item)}
                        onMouseLeave={() => setHovered(null)}
                        onFocus={() => item && setHovered(item)}
                        onBlur={() => setHovered(null)}
                        onClick={() => item && setSelected(item)}
                        tabIndex={item ? 0 : -1}
                        aria-label={
                          item
                            ? `${item.municipio}, ${statusFor(item).label}`
                            : shape.name
                        }
                      >
                        <title>
                          {item
                            ? `${item.municipio} · ${statusFor(item).label}`
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
                <span>{displayMunicipality.provincia}</span>
                <strong>{displayMunicipality.municipio}</strong>
                <small>{statusFor(displayMunicipality).label}</small>
              </div>
            )}

            <div className="map-attribution">
              Cartografía ADM2 · 158 límites municipales
            </div>
          </div>

          <div className="map-legend">
            {(Object.keys(statusMeta) as PmdState[]).map((state) => (
              <span key={state}>
                <i style={{ background: statusMeta[state].color }} />
                {statusMeta[state].shortLabel}
              </span>
            ))}
          </div>

          {unmapped.length > 0 && mapShapes.length > 0 && (
            <div className="unmapped-note">
              <strong>{unmapped.length} municipios nuevos sin geometría ADM2:</strong>{" "}
              {unmapped.map((item) => item.municipio).join(", ")}. Se mantienen
              disponibles en la búsqueda y el listado.
            </div>
          )}
        </article>

        <aside className={`detail-panel ${selected ? "has-selection" : ""}`}>
          {selected ? (
            <>
              <button
                className="detail-close"
                onClick={() => setSelected(null)}
                aria-label="Cerrar detalle"
              >
                ×
              </button>
              <span className="detail-region">{selected.region}</span>
              <h2>{selected.municipio}</h2>
              <p>Provincia {selected.provincia}</p>
              <StatusBadge item={selected} />
              <p className="status-description">
                {statusFor(selected).description}
              </p>

              <dl className="detail-facts">
                <div>
                  <dt>OMPP</dt>
                  <dd>{selected.ompp.label}</dd>
                </div>
                <div>
                  <dt>CDM</dt>
                  <dd>
                    {selected.cdm.label}
                    <small>{selected.cdm.score ?? 0}% SISMAP</small>
                  </dd>
                </div>
                <div>
                  <dt>Período PMD</dt>
                  <dd>{selected.pmd.period || "No confirmado"}</dd>
                </div>
                <div>
                  <dt>Documentos</dt>
                  <dd>{selected.pmd.documentCount} PDF</dd>
                </div>
              </dl>

              <div className="detail-links">
                {selected.sismapUrl && (
                  <a href={selected.sismapUrl} target="_blank" rel="noreferrer">
                    SISMAP municipal ↗
                  </a>
                )}
                {selected.pmd.pdfUrl && (
                  <a href={selected.pmd.pdfUrl} target="_blank" rel="noreferrer">
                    Documento PMD ↗
                  </a>
                )}
              </div>

              <button
                className={`primary-action action-${getPmdState(selected)}`}
                onClick={() => handleAction(selected)}
              >
                {isCurrentPmd(selected)
                  ? "Abrir PMD vigente"
                  : selected.actionLabel}
                <span aria-hidden="true">→</span>
              </button>
            </>
          ) : (
            <div className="empty-detail">
              <span className="empty-icon" aria-hidden="true">
                ↖
              </span>
              <h2>Seleccione un municipio</h2>
              <p>
                Verá el estado del PMD, las evidencias OMPP y CDM y la acción
                recomendada.
              </p>
              <div className="current-rule">
                <strong>Regla del mapa</strong>
                <p>
                  Solo los PMD con período actualmente vigente cuentan como
                  “PMD vigente”. Los planes anteriores se muestran como
                  históricos.
                </p>
              </div>
            </div>
          )}
        </aside>
      </section>

      <section className="directory">
        <div className="directory-heading">
          <div>
            <span className="eyebrow">Directorio</span>
            <h2>{filtered.length} municipios</h2>
          </div>
          <p>También incluye municipios nuevos aún no representados en el GeoJSON.</p>
        </div>
        <div className="municipality-list">
          {filtered.map((item) => (
            <button
              key={item.id}
              className={selected?.id === item.id ? "active" : ""}
              onClick={() => setSelected(item)}
            >
              <span
                className="list-status"
                style={{ background: statusFor(item).color }}
                aria-hidden="true"
              />
              <span>
                <strong>{item.municipio}</strong>
                <small>
                  {item.provincia} · {statusFor(item).shortLabel}
                </small>
              </span>
              <i aria-hidden="true">→</i>
            </button>
          ))}
        </div>
      </section>

      <footer>
        <div>
          <strong>Portal de Planificación Municipal</strong>
          <p>
            El estado documental no sustituye la revisión de la OMPP, la
            validación del CDM ni la aprobación del Concejo de Regidores.
          </p>
        </div>
        <div>
          <a
            href="https://github.com/mnakagaw/dashboard-municipal"
            target="_blank"
            rel="noreferrer"
          >
            Fuente cartográfica: dashboard-municipal ↗
          </a>
          <span>Estado de datos: 27 de julio de 2026</span>
        </div>
      </footer>

      {wizard && (
        <div
          className="modal-backdrop"
          role="presentation"
          onMouseDown={() => setWizard(null)}
        >
          <section
            className="workspace-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="workspace-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <button
              className="modal-close"
              onClick={() => setWizard(null)}
              aria-label="Cerrar"
            >
              ×
            </button>
            <header>
              <span>Paquete Mínimo · espacio de trabajo</span>
              <h2 id="workspace-title">{wizard.municipio}</h2>
              <p>
                {wizard.actionLabel} · Provincia {wizard.provincia}
              </p>
            </header>

            <div className="workspace-grid">
              <div className="input-column">
                <span className="column-label">01 · Inputs para la IA</span>
                <h3>Fuentes de partida</h3>
                <p>
                  Se reutilizan para preparar el borrador, nunca para afirmar
                  que un dato antiguo sigue vigente sin verificación.
                </p>
                <div className="source-stack">
                  <SourceCard
                    label="Datos territoriales"
                    title="Dashboard municipal"
                    note="Indicadores estadísticos y comparación nacional."
                    href="https://github.com/mnakagaw/dashboard-municipal"
                  />
                  <SourceCard
                    label="Contexto general"
                    title="Wikipedia"
                    note="Punto de partida; requiere contraste con fuentes oficiales."
                    href={`https://es.wikipedia.org/w/index.php?search=${encodeURIComponent(
                      `${wizard.municipio} República Dominicana`,
                    )}`}
                  />
                  {wizard.pmd.documentCount > 0 ? (
                    <SourceCard
                      label={
                        getPmdState(wizard) === "historical"
                          ? "Antecedente"
                          : "Expediente disponible"
                      }
                      title={
                        getPmdState(wizard) === "historical"
                          ? "PMD anterior"
                          : "Documentación PMD"
                      }
                      note={`${wizard.pmd.period || "Período por confirmar"} · ${
                        wizard.pmd.documentCount
                      } PDF`}
                      href={wizard.pmd.pdfUrl || undefined}
                    />
                  ) : (
                    <SourceCard
                      label="Antecedente"
                      title="PMD anterior no localizado"
                      note="Permite cargar un documento nuevo con sus metadatos."
                    />
                  )}
                </div>
                <div className="source-rule">
                  <strong>Trazabilidad obligatoria</strong>
                  <span>Documento · página · año · URL · fecha base · confianza</span>
                </div>
              </div>

              <div className="draft-column">
                <span className="column-label">02 · Borrador asistido</span>
                <h3>Paquete Mínimo</h3>
                <div className="draft-steps">
                  <article>
                    <b>1</b>
                    <div>
                      <strong>Información General</strong>
                      <p>AI narrativo con fuentes y campos por verificar.</p>
                    </div>
                    <span>OMPP + Coord. CDM</span>
                  </article>
                  <article>
                    <b>2</b>
                    <div>
                      <strong>Diagnóstico Municipal</strong>
                      <p>Datos actuales, brechas y comparación territorial.</p>
                    </div>
                    <span>OMPP + Coord. CDM</span>
                  </article>
                  <article>
                    <b>3</b>
                    <div>
                      <strong>FODA y problemas principales</strong>
                      <p>Propuestas separadas de hechos verificados.</p>
                    </div>
                    <span>OMPP + Coord. CDM</span>
                  </article>
                  <article className="human-step">
                    <b>4</b>
                    <div>
                      <strong>Visión, proyectos y demandas</strong>
                      <p>Construcción participativa; la IA no inventa acuerdos.</p>
                    </div>
                    <span>Elaboración CDM</span>
                  </article>
                </div>
                <div className="validation-gate">
                  <span aria-hidden="true">✓</span>
                  <div>
                    <strong>Revisión y validación por CDM</strong>
                    <p>
                      Ninguna reunión, participación o aprobación se registra
                      como realizada sin evidencia humana.
                    </p>
                  </div>
                </div>
                <button className="primary-action" onClick={() => setWizard(null)}>
                  Preparar fuentes del municipio
                  <span aria-hidden="true">→</span>
                </button>
              </div>
            </div>
          </section>
        </div>
      )}
    </main>
  );
}
