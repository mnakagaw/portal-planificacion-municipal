import "dotenv/config";
import { access } from "node:fs/promises";
import path from "node:path";
import { Client } from "basic-ftp";

const buildDirectory = path.resolve("dist-static");
const requiredSuffix = "/DDPT/planificacion-municipal";
const remoteRoot =
  process.env.FTP_REMOTE_ROOT ??
  "/public_html/prodecare.net/DDPT/planificacion-municipal";

const normalizedRemoteRoot = remoteRoot.replaceAll("\\", "/").replace(/\/+$/, "");

if (!normalizedRemoteRoot.endsWith(requiredSuffix)) {
  throw new Error(
    `Destino FTP rechazado. La carpeta debe terminar en ${requiredSuffix}.`,
  );
}

for (const key of ["FTP_HOST", "FTP_USER", "FTP_PASS"]) {
  if (!process.env[key]) {
    throw new Error(`Falta la variable ${key}.`);
  }
}

await access(path.join(buildDirectory, "index.html"));

const client = new Client();

try {
  await client.access({
    host: process.env.FTP_HOST,
    user: process.env.FTP_USER,
    password: process.env.FTP_PASS,
    secure: process.env.FTP_SECURE === "true",
  });

  await client.ensureDir(normalizedRemoteRoot);
  const resolvedRemoteRoot = (await client.pwd()).replace(/\/+$/, "");

  if (!resolvedRemoteRoot.endsWith(requiredSuffix)) {
    throw new Error(`La carpeta FTP resuelta no coincide: ${resolvedRemoteRoot}`);
  }

  await client.clearWorkingDir();
  await client.uploadFromDir(buildDirectory);
  console.log(`Publicado en ${resolvedRemoteRoot}`);
} finally {
  client.close();
}
