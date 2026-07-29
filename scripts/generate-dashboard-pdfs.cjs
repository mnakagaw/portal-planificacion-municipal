const fs = require("node:fs");
const path = require("node:path");
const { chromium } = require("playwright-core");

const projectRoot = path.resolve(__dirname, "..");
const workspaceRoot = path.resolve(projectRoot, "..", "..");
const dashboardIndexPath =
  process.env.DASHBOARD_INDEX_PATH ||
  path.join(
    workspaceRoot,
    "reference-dashboard-municipal",
    "public",
    "data",
    "municipios_index.json",
  );
const portalDataPath = path.join(
  projectRoot,
  "app",
  "data",
  "municipios.json",
);
const outputDirectory = path.join(
  projectRoot,
  "public",
  "downloads",
  "diagnosticos",
);
const manifestPath = path.join(
  projectRoot,
  "app",
  "data",
  "diagnosticos.json",
);

const args = new Set(process.argv.slice(2));
const argumentValue = (name) => {
  const position = process.argv.indexOf(name);
  return position >= 0 ? process.argv[position + 1] : "";
};
const baseUrl =
  argumentValue("--base-url") ||
  process.env.DASHBOARD_URL ||
  "https://prodecare.net/dashboard/";
const includeNarrative = !args.has("--no-narrative");
const overwrite = args.has("--overwrite");
const onlyCode = argumentValue("--only");
const limit = Number.parseInt(argumentValue("--limit"), 10) || null;
const concurrency = Math.max(
  1,
  Math.min(
    4,
    Number.parseInt(argumentValue("--concurrency"), 10) ||
      Number.parseInt(process.env.PDF_CONCURRENCY, 10) ||
      3,
  ),
);

function cleanName(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

const municipalityAliases = {
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

function canonicalMunicipality(value) {
  const cleaned = cleanName(value);
  return municipalityAliases[cleaned] ?? cleaned;
}

function canonicalProvince(value) {
  const cleaned = cleanName(value);
  return cleaned === "bahoruco" ? "baoruco" : cleaned;
}

function territoryKey(municipality, province) {
  return `${canonicalMunicipality(municipality)}|${canonicalProvince(province)}`;
}

function fileSlug(value) {
  return String(value ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase();
}

function findBrowserExecutable() {
  const candidates = [
    process.env.CHROME_EXECUTABLE_PATH,
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/usr/bin/google-chrome",
    "/usr/bin/chromium",
    "/usr/bin/chromium-browser",
  ].filter(Boolean);
  const executable = candidates.find((candidate) => fs.existsSync(candidate));
  if (!executable) {
    throw new Error(
      "No se encontró Chrome/Edge. Defina CHROME_EXECUTABLE_PATH.",
    );
  }
  return executable;
}

function readJson(filePath) {
  return JSON.parse(fs.readFileSync(filePath, "utf8"));
}

function writeJson(filePath, value) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true });
  fs.writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

async function chooseMunicipality(page, item) {
  await page.goto(baseUrl, {
    waitUntil: "domcontentloaded",
    timeout: 90_000,
  });
  await page.locator("h1").waitFor({ state: "visible", timeout: 60_000 });
  const selects = page.locator("select");
  await selects.first().waitFor({ state: "visible" });
  const selectCount = await selects.count();

  let provinceSelect;
  let municipalitySelect;
  if (selectCount === 3) {
    const regionIdByName = {
      Ozama: "ozama",
      "Ozama o Metropolitana": "ozama",
      "Cibao Norte": "cibao_norte",
      "Cibao Sur": "cibao_sur",
      "Cibao Nordeste": "cibao_nordeste",
      "Cibao Noroeste": "cibao_noroeste",
      Valdesia: "valdesia",
      Enriquillo: "enriquillo",
      "El Valle": "el_valle",
      Yuma: "yuma",
      Higuamo: "higuamo",
    };
    await selects.nth(0).selectOption(regionIdByName[item.region]);
    provinceSelect = selects.nth(1);
    municipalitySelect = selects.nth(2);
  } else {
    const regionButton = page.locator('button[aria-haspopup="listbox"]').first();
    const selectedRegion = (await regionButton.innerText()).trim();
    if (selectedRegion !== item.region) {
      await regionButton.click();
      const selectedOptions = page.locator(
        '[role="option"][aria-selected="true"]',
      );
      const selectedCount = await selectedOptions.count();
      for (let index = selectedCount - 1; index >= 0; index -= 1) {
        await selectedOptions.nth(index).click();
      }
      await page
        .getByRole("option", { name: item.region, exact: true })
        .click();
      await page.keyboard.press("Escape");
    }
    provinceSelect = selects.nth(0);
    municipalitySelect = selects.nth(1);
  }

  await provinceSelect.waitFor({ state: "visible" });
  await provinceSelect.selectOption(item.provincia);
  await municipalitySelect.waitFor({ state: "visible" });
  await municipalitySelect.selectOption(item.adm2_code);

  await page.waitForFunction(
    (municipality) =>
      document.querySelector("h1")?.textContent?.includes(municipality),
    item.municipio,
    { timeout: 60_000 },
  );
  await page.waitForFunction(
    (municipality) => {
      const root = document.querySelector("#dashboard-pdf");
      return (
        root?.textContent?.includes(municipality) &&
        root.querySelectorAll("svg").length >= 8
      );
    },
    item.municipio,
    { timeout: 60_000 },
  );
  await page.evaluate(() => document.fonts.ready);
  await page.waitForTimeout(1_200);
}

async function generateNarrative(page) {
  const button = page.getByRole("button", {
    name: "Crear Diagnóstico Narrativo",
    exact: true,
  });
  if ((await button.count()) === 0) {
    throw new Error("No se encontró el botón de resumen narrativo.");
  }
  await button.click();
  await page.waitForFunction(
    () => {
      const text =
        document.querySelector("#dashboard-pdf")?.textContent ?? "";
      return (
        text.includes("1. Panorama general") &&
        !text.includes("Aún no se ha generado el resumen") &&
        !text.includes("Error al generar resumen")
      );
    },
    null,
    { timeout: 180_000 },
  );
  await page.waitForTimeout(500);
}

async function validatePage(page, municipality) {
  const text = await page.locator("#dashboard-pdf").innerText();
  const required = [
    municipality,
    "Información básica",
    "Pirámide de población 2022",
    "Resumen de Comparación",
  ];
  const missing = required.filter((label) => !text.includes(label));
  if (missing.length) {
    throw new Error(`Contenido incompleto: ${missing.join(", ")}`);
  }
  if (
    includeNarrative &&
    (text.includes("Aún no se ha generado el resumen") ||
      text.includes("Error al generar resumen"))
  ) {
    throw new Error("El resumen narrativo no está disponible.");
  }
}

async function main() {
  const dashboardMunicipalities = readJson(dashboardIndexPath);
  const portalMunicipalities = readJson(portalDataPath);
  const portalLookup = new Map(
    portalMunicipalities.map((item) => [
      territoryKey(item.municipio, item.provincia),
      item,
    ]),
  );

  const mapped = dashboardMunicipalities.map((item) => {
    const portal = portalLookup.get(
      territoryKey(item.municipio, item.provincia),
    );
    if (!portal) {
      throw new Error(
        `Municipio no encontrado en el portal: ${item.municipio}, ${item.provincia}`,
      );
    }
    return { ...item, portalId: portal.id };
  });

  if (mapped.length !== 158) {
    throw new Error(`Se esperaban 158 municipios y se encontraron ${mapped.length}.`);
  }

  let queue = mapped;
  if (onlyCode) {
    queue = queue.filter((item) => item.adm2_code === onlyCode);
  }
  if (limit) queue = queue.slice(0, limit);
  if (!queue.length) {
    throw new Error("No hay municipios para procesar.");
  }

  fs.mkdirSync(outputDirectory, { recursive: true });
  const previousManifest = fs.existsSync(manifestPath)
    ? readJson(manifestPath)
    : [];
  const manifestByCode = new Map(
    previousManifest.map((entry) => [entry.adm2Code, entry]),
  );

  const browser = await chromium.launch({
    executablePath: findBrowserExecutable(),
    headless: true,
  });
  const context = await browser.newContext({
    locale: "es-DO",
    viewport: { width: 1440, height: 1100 },
  });

  const failures = [];
  let nextPosition = 0;
  const runWorker = async (workerNumber) => {
    const page = await context.newPage();
    page.setDefaultTimeout(60_000);
    await page.addInitScript(() => {
      localStorage.clear();
      sessionStorage.clear();
    });

    while (nextPosition < queue.length) {
      const position = nextPosition;
      nextPosition += 1;
      const item = queue[position];
      const filename = `${item.adm2_code}_diagnostico-territorial-${fileSlug(
        item.municipio,
      )}.pdf`;
      const outputPath = path.join(outputDirectory, filename);
      const existing =
        fs.existsSync(outputPath) && fs.statSync(outputPath).size > 50_000;

      if (existing && !overwrite) {
        console.log(
          `[${position + 1}/${queue.length}] Conservado: ${item.municipio}`,
        );
      } else {
        console.log(
          `[${position + 1}/${queue.length}] Generando: ${item.municipio} (worker ${workerNumber})`,
        );
        try {
          await chooseMunicipality(page, item);
          if (includeNarrative) await generateNarrative(page);
          await validatePage(page, item.municipio);
          await page.pdf({
            path: outputPath,
            format: "Letter",
            preferCSSPageSize: true,
            printBackground: true,
            displayHeaderFooter: false,
          });
          if (fs.statSync(outputPath).size <= 50_000) {
            throw new Error("El PDF generado es demasiado pequeño.");
          }
        } catch (error) {
          if (fs.existsSync(outputPath)) fs.unlinkSync(outputPath);
          failures.push({
            adm2Code: item.adm2_code,
            municipio: item.municipio,
            error: error instanceof Error ? error.message : String(error),
          });
          console.error(`  Error: ${failures.at(-1).error}`);
          continue;
        }
      }

      manifestByCode.set(item.adm2_code, {
        id: item.portalId,
        municipio: item.municipio,
        provincia: item.provincia,
        region: item.region,
        adm2Code: item.adm2_code,
        url: `downloads/diagnosticos/${filename}`,
        filename,
        includesNarrative: includeNarrative,
        generatedAt: new Date().toISOString(),
      });
      writeJson(
        manifestPath,
        [...manifestByCode.values()].sort((a, b) => a.id - b.id),
      );
    }

    await page.close();
  };

  try {
    await Promise.all(
      Array.from(
        { length: Math.min(concurrency, queue.length) },
        (_, index) => runWorker(index + 1),
      ),
    );
  } finally {
    await browser.close();
  }

  const currentManifest = [...manifestByCode.values()].filter((entry) =>
    fs.existsSync(path.join(projectRoot, "public", entry.url)),
  );
  writeJson(
    manifestPath,
    currentManifest.sort((a, b) => a.id - b.id),
  );

  console.log(
    `Completados: ${currentManifest.length}/158. Fallos de esta ejecución: ${failures.length}.`,
  );
  if (failures.length) {
    writeJson(
      path.join(projectRoot, "outputs", "diagnosticos-errors.json"),
      failures,
    );
    process.exitCode = 1;
  } else {
    const previousErrors = path.join(
      projectRoot,
      "outputs",
      "diagnosticos-errors.json",
    );
    if (fs.existsSync(previousErrors)) fs.unlinkSync(previousErrors);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
