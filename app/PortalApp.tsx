"use client";

import { useMemo, useState } from "react";
import municipalitiesData from "./data/municipios.json";

type Level = "complete" | "progress" | "warning" | "none" | "unknown";

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

const levelLabel: Record<Level, string> = {
  complete: "Completado",
  progress: "En proceso",
  warning: "Requiere actualización",
  none: "Sin iniciar",
  unknown: "No confirmado",
};

function StatusPill({ level, children }: { level: Level; children: React.ReactNode }) {
  return (
    <span className={`status-pill status-${level}`}>
      <span className="status-dot" aria-hidden="true" />
      {children}
    </span>
  );
}

function MetricCard({
  label,
  value,
  note,
  tone,
}: {
  label: string;
  value: number;
  note: string;
  tone: string;
}) {
  return (
    <article className={`metric-card tone-${tone}`}>
      <div className="metric-topline">
        <span>{label}</span>
        <span className="metric-mark" aria-hidden="true" />
      </div>
      <strong>{value}</strong>
      <small>{note}</small>
    </article>
  );
}

export function PortalApp() {
  const [region, setRegion] = useState("Todas");
  const [province, setProvince] = useState("Todas");
  const [query, setQuery] = useState("");
  const [scope, setScope] = useState<"todos" | "con-pmd" | "sin-pmd">("todos");
  const [selected, setSelected] = useState<Municipality | null>(null);
  const [wizard, setWizard] = useState<Municipality | null>(null);

  const regions = useMemo(
    () => ["Todas", ...regionOrder.filter((name) => municipalities.some((item) => item.region === name))],
    [],
  );

  const provinces = useMemo(() => {
    const source =
      region === "Todas"
        ? municipalities
        : municipalities.filter((item) => item.region === region);
    return ["Todas", ...Array.from(new Set(source.map((item) => item.provincia))).sort()];
  }, [region]);

  const filtered = useMemo(() => {
    const normalized = query.trim().toLocaleLowerCase("es");
    return municipalities.filter((item) => {
      if (region !== "Todas" && item.region !== region) return false;
      if (province !== "Todas" && item.provincia !== province) return false;
      if (
        normalized &&
        !`${item.municipio} ${item.provincia} ${item.region}`
          .toLocaleLowerCase("es")
          .includes(normalized)
      )
        return false;
      if (scope === "con-pmd" && item.pmd.documentCount === 0) return false;
      if (scope === "sin-pmd" && item.pmd.documentCount > 0) return false;
      return true;
    });
  }, [province, query, region, scope]);

  const stats = useMemo(() => {
    const source = filtered;
    return {
      total: source.length,
      ompp: source.filter((item) => item.ompp.level === "complete").length,
      cdm: source.filter((item) => (item.cdm.score ?? 0) >= 80).length,
      pmd: source.filter((item) => item.pmd.hasCurrent).length,
      sinPmd: source.filter((item) => item.pmd.documentCount === 0).length,
    };
  }, [filtered]);

  const regionStats = useMemo(
    () =>
      regionOrder.map((name) => {
        const items = municipalities.filter((item) => item.region === name);
        const withPmd = items.filter((item) => item.pmd.documentCount > 0).length;
        return {
          name,
          total: items.length,
          withPmd,
          percentage: items.length ? Math.round((withPmd / items.length) * 100) : 0,
        };
      }),
    [],
  );

  const handleAction = (item: Municipality) => {
    if (item.action === "ver" && item.pmd.pdfUrl) {
      window.open(item.pmd.pdfUrl, "_blank", "noopener,noreferrer");
      return;
    }
    setWizard(item);
  };

  return (
    <main>
      <header className="topbar">
        <div className="brand">
          <div className="brand-seal" aria-hidden="true">
            PM
          </div>
          <div>
            <strong>Portal de Planificación Municipal</strong>
            <span>República Dominicana</span>
          </div>
        </div>
        <div className="header-meta">
          <span>Datos verificados</span>
          <strong>27 JUL 2026</strong>
        </div>
      </header>

      <section className="hero">
        <div className="hero-copy">
          <p className="eyebrow">Planificación territorial · 162 municipios</p>
          <h1>El estado de la planificación municipal, en un solo lugar.</h1>
          <p>
            Consulte la situación de la OMPP, el CDM y el PMD. Identifique qué
            municipio puede visualizar, actualizar o comenzar su plan.
          </p>
        </div>
        <div className="hero-insight">
          <span className="insight-label">Cobertura documental</span>
          <strong>64</strong>
          <p>municipios con al menos un PMD recuperado</p>
          <div className="progress-track">
            <span style={{ width: `${Math.round((64 / 162) * 100)}%` }} />
          </div>
          <small>85 documentos PDF únicos</small>
        </div>
      </section>

      <section className="filter-panel" aria-label="Filtros territoriales">
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
          <select value={province} onChange={(event) => setProvince(event.target.value)}>
            {provinces.map((item) => (
              <option key={item}>{item}</option>
            ))}
          </select>
        </label>
        <label className="search-field">
          <span>Municipio</span>
          <input
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="Buscar municipio…"
          />
        </label>
        <button
          className="clear-button"
          onClick={() => {
            setRegion("Todas");
            setProvince("Todas");
            setQuery("");
            setScope("todos");
          }}
        >
          Limpiar filtros
        </button>
      </section>

      <section className="metrics-grid" aria-label="Indicadores de cobertura">
        <MetricCard label="Municipios" value={stats.total} note="en la selección" tone="navy" />
        <MetricCard label="OMPP establecida" value={stats.ompp} note="evidencia confirmada" tone="teal" />
        <MetricCard label="CDM constituido" value={stats.cdm} note="80 puntos o más" tone="green" />
        <MetricCard label="PMD actual" value={stats.pmd} note="período vigente" tone="blue" />
        <MetricCard label="Sin PDF de PMD" value={stats.sinPmd} note="requiere localización o elaboración" tone="orange" />
      </section>

      <section className="overview-grid">
        <article className="panel map-panel">
          <div className="panel-heading">
            <div>
              <span className="section-kicker">Vista territorial</span>
              <h2>Cobertura de PMD por región</h2>
            </div>
            <span className="legend">PDF recuperado / total</span>
          </div>
          <div className="region-map">
            {regionStats.map((item) => (
              <button
                key={item.name}
                className={`region-cell ${region === item.name ? "is-selected" : ""}`}
                onClick={() => {
                  setRegion(item.name);
                  setProvince("Todas");
                }}
                aria-pressed={region === item.name}
              >
                <span>{item.name}</span>
                <strong>
                  {item.withPmd}/{item.total}
                </strong>
                <i>
                  <b style={{ width: `${item.percentage}%` }} />
                </i>
              </button>
            ))}
          </div>
        </article>

        <article className="panel guide-panel">
          <div className="panel-heading">
            <div>
              <span className="section-kicker">Ruta recomendada</span>
              <h2>¿Qué acción corresponde?</h2>
            </div>
          </div>
          <ol className="action-guide">
            <li>
              <span className="guide-number">01</span>
              <div>
                <strong>PMD vigente</strong>
                <p>Consultar el plan y sus evidencias en SISMAP.</p>
              </div>
              <em>Ver PMD</em>
            </li>
            <li>
              <span className="guide-number">02</span>
              <div>
                <strong>Borrador o proceso abierto</strong>
                <p>Retomar el expediente y completar la validación.</p>
              </div>
              <em>Continuar</em>
            </li>
            <li>
              <span className="guide-number">03</span>
              <div>
                <strong>PMD histórico</strong>
                <p>Reutilizar la base y actualizar el diagnóstico.</p>
              </div>
              <em>Actualizar</em>
            </li>
            <li>
              <span className="guide-number">04</span>
              <div>
                <strong>Sin PMD confirmado</strong>
                <p>Iniciar un Paquete Mínimo trazable.</p>
              </div>
              <em>Elaborar</em>
            </li>
          </ol>
        </article>
      </section>

      <section className="panel municipality-panel">
        <div className="table-toolbar">
          <div>
            <span className="section-kicker">Directorio nacional</span>
            <h2>Municipios</h2>
          </div>
          <div className="scope-tabs" role="group" aria-label="Filtrar por disponibilidad de PMD">
            {[
              ["todos", "Todos"],
              ["con-pmd", "Con PDF"],
              ["sin-pmd", "Sin PDF"],
            ].map(([value, label]) => (
              <button
                key={value}
                className={scope === value ? "active" : ""}
                onClick={() => setScope(value as typeof scope)}
              >
                {label}
              </button>
            ))}
          </div>
          <span className="result-count">{filtered.length} resultados</span>
        </div>

        <div className="table-scroll">
          <table>
            <thead>
              <tr>
                <th>Municipio</th>
                <th>Territorio</th>
                <th>OMPP</th>
                <th>CDM</th>
                <th>PMD</th>
                <th>Acción</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((item) => (
                <tr key={item.id}>
                  <td>
                    <button className="municipality-name" onClick={() => setSelected(item)}>
                      <span>{String(item.id).padStart(3, "0")}</span>
                      <strong>{item.municipio}</strong>
                    </button>
                  </td>
                  <td>
                    <strong>{item.provincia}</strong>
                    <small>{item.region}</small>
                  </td>
                  <td>
                    <StatusPill level={item.ompp.level}>{item.ompp.label}</StatusPill>
                  </td>
                  <td>
                    <StatusPill level={item.cdm.level}>{item.cdm.label}</StatusPill>
                    <small className="score-note">
                      {item.cdm.score === null ? "Sin dato" : `${item.cdm.score}% SISMAP`}
                    </small>
                  </td>
                  <td>
                    <StatusPill level={item.pmd.level}>{item.pmd.label}</StatusPill>
                    <small className="score-note">
                      {item.pmd.period || `${item.pmd.documentCount} PDF recuperados`}
                    </small>
                  </td>
                  <td>
                    <button className={`action-button action-${item.action}`} onClick={() => handleAction(item)}>
                      {item.actionLabel}
                      <span aria-hidden="true">→</span>
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      <footer>
        <div>
          <strong>Portal de Planificación Municipal</strong>
          <p>Prototipo basado en SISMAP Municipal y documentos PMD recuperados.</p>
        </div>
        <p>
          La disponibilidad documental no sustituye la validación de OMPP, CDM ni
          del Concejo de Regidores.
        </p>
      </footer>

      {selected && (
        <div className="modal-backdrop" role="presentation" onMouseDown={() => setSelected(null)}>
          <section
            className="detail-drawer"
            role="dialog"
            aria-modal="true"
            aria-labelledby="municipality-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <button className="close-button" onClick={() => setSelected(null)} aria-label="Cerrar">
              ×
            </button>
            <span className="detail-region">{selected.region}</span>
            <h2 id="municipality-title">{selected.municipio}</h2>
            <p className="detail-province">Provincia {selected.provincia}</p>

            <div className="detail-status-grid">
              <article>
                <span>OMPP</span>
                <StatusPill level={selected.ompp.level}>{selected.ompp.label}</StatusPill>
                <small>{selected.ompp.evidenceCount} evidencias 1-12</small>
              </article>
              <article>
                <span>CDM</span>
                <StatusPill level={selected.cdm.level}>{selected.cdm.label}</StatusPill>
                <small>{selected.cdm.score ?? 0}% SISMAP</small>
              </article>
              <article>
                <span>PMD</span>
                <StatusPill level={selected.pmd.level}>{selected.pmd.label}</StatusPill>
                <small>{selected.pmd.period || "Período no confirmado"}</small>
              </article>
            </div>

            <div className="document-box">
              <span>Documentación recuperada</span>
              <strong>{selected.pmd.documentCount} PDF</strong>
              <p>
                {selected.pmd.pdfTitle ||
                  "No se ha identificado un documento principal para este municipio."}
              </p>
            </div>

            <div className="detail-links">
              {selected.sismapUrl && (
                <a href={selected.sismapUrl} target="_blank" rel="noreferrer">
                  Ver página municipal en SISMAP
                </a>
              )}
              {selected.pmd.pdfUrl && (
                <a href={selected.pmd.pdfUrl} target="_blank" rel="noreferrer">
                  Abrir PDF del PMD
                </a>
              )}
            </div>

            <button className={`primary-cta action-${selected.action}`} onClick={() => handleAction(selected)}>
              {selected.actionLabel}
              <span aria-hidden="true">→</span>
            </button>
          </section>
        </div>
      )}

      {wizard && (
        <div className="modal-backdrop wizard-backdrop" role="presentation" onMouseDown={() => setWizard(null)}>
          <section
            className="wizard-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="wizard-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <button className="close-button" onClick={() => setWizard(null)} aria-label="Cerrar">
              ×
            </button>
            <span className="detail-region">Paquete Mínimo</span>
            <h2 id="wizard-title">{wizard.actionLabel}</h2>
            <p>
              {wizard.municipio} · {wizard.provincia}
            </p>
            <div className="wizard-progress">
              <span style={{ width: wizard.action === "actualizar" ? "38%" : wizard.action === "continuar" ? "58%" : "12%" }} />
            </div>
            <ol className="wizard-steps">
              {[
                ["01", "Verificar OMPP", wizard.ompp.level === "complete"],
                ["02", "Verificar CDM", (wizard.cdm.score ?? 0) >= 80],
                ["03", "Recopilar fuentes", wizard.pmd.documentCount > 0],
                ["04", "Información General", false],
                ["05", "Diagnóstico Municipal", false],
                ["06", "FODA, Visión y Proyectos", false],
                ["07", "Revisión OMPP / CDM", false],
                ["08", "Exportar Borrador", false],
              ].map(([number, label, done]) => (
                <li key={String(number)} className={done ? "done" : ""}>
                  <span>{done ? "✓" : number}</span>
                  <strong>{label}</strong>
                </li>
              ))}
            </ol>
            <div className="wizard-note">
              <strong>Principio de trabajo</strong>
              <p>
                La IA prepara un borrador con fuentes. La OMPP y el CDM revisan,
                corrigen y validan el contenido.
              </p>
            </div>
            <button className="primary-cta" onClick={() => setWizard(null)}>
              Preparar espacio de trabajo
              <span aria-hidden="true">→</span>
            </button>
          </section>
        </div>
      )}
    </main>
  );
}
