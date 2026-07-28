# Portal de Planificación Municipal

Portal cartográfico para consultar el estado del PMD, borradores, CDM y OMPP
de los 162 municipios de la República Dominicana.

## Desarrollo

Requiere Node.js 22 o posterior.

```bash
pnpm install
pnpm dev
```

## Verificación

```bash
pnpm test
pnpm build:static
```

`pnpm test` valida la versión vinext. `pnpm build:static` genera en
`dist-static/` la versión pública preparada para:

https://prodecare.net/DDPT/planificacion-municipal/

## Despliegue en prodecare.net

Copie `.env.example` como `.env`, complete las credenciales FTP y ejecute:

```bash
pnpm deploy:prodecare
```

El despliegue solamente acepta como destino la carpeta
`/DDPT/planificacion-municipal`.

## Datos

- Estados municipales: `app/data/municipios.json`
- Cartografía municipal: `public/data/adm2.geojson`
- Fuente principal de estado: SISMAP Municipal

La cartografía base se complementa con las geometrías de Villa Central, Tireo,
La Caleta y La Victoria procedentes de la capa DPA Distritos Municipales del
IGN/ONE. Para reconstruir y aligerar el archivo:

```bash
pnpm geo:merge-former-dm
pnpm geo:optimize
```

La condición de PMD oficial aplica cuando SISMAP 2.02 alcanza 100 % o existe
evidencia 8-12 confirmada. Un PMD oficial implica también borrador, CDM y OMPP
para la visualización agregada.

## Borradores técnicos Word

El portal distribuye un borrador técnico editable para los 104 municipios que
no tienen PMD oficial ni borrador 7-12 confirmado en los datos actuales.

- Período propuesto: `2025-2028`
- Extensión de referencia: 10 páginas
- Fuentes permitidas: Dashboard territorial, PMD históricos inventariados y
  Wikipedia solamente para contexto geográfico básico
- Exclusiones: prensa local, cifras sin fuente, reuniones, participación,
  aprobación, presupuesto o ejecución no documentados
- Manifiesto público: `public/data/generated-pmd-drafts.json`
- Archivos: `public/downloads/pmd-borradores/`

Los documentos se identifican como borradores no aprobados y reservan la
validación de FODA, visión y proyectos para la OMPP y el CDM.
