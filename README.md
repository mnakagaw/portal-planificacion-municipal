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

La condición de PMD oficial aplica cuando SISMAP 2.02 alcanza 100 % o existe
evidencia 8-12 confirmada. Un PMD oficial implica también borrador, CDM y OMPP
para la visualización agregada.
