import "dotenv/config";
import { access } from "node:fs/promises";
import path from "node:path";
import { Client } from "basic-ftp";

const localDirectory = path.resolve("public/downloads/diagnosticos");
const requiredSuffix = "/DDPT/planificacion-municipal";
const remoteRoot =
  process.env.FTP_REMOTE_ROOT ??
  "/public_html/prodecare.net/DDPT/planificacion-municipal";
const normalizedRemoteRoot = remoteRoot.replaceAll("\\", "/").replace(/\/+$/, "");
const remoteDirectory = path.posix.join(
  normalizedRemoteRoot,
  "downloads",
  "diagnosticos",
);

if (!normalizedRemoteRoot.endsWith(requiredSuffix)) {
  throw new Error(
    `Destino FTP rechazado. La carpeta debe terminar en ${requiredSuffix}.`,
  );
}

for (const key of ["FTP_HOST", "FTP_USER", "FTP_PASS"]) {
  if (!process.env[key]) throw new Error(`Falta la variable ${key}.`);
}
await access(localDirectory);

const client = new Client();
client.ftp.verbose = false;

try {
  await client.access({
    host: process.env.FTP_HOST,
    user: process.env.FTP_USER,
    password: process.env.FTP_PASS,
    secure: process.env.FTP_SECURE === "true",
  });
  await client.ensureDir(remoteDirectory);
  const resolvedRemoteDirectory = (await client.pwd()).replace(/\/+$/, "");
  if (!resolvedRemoteDirectory.endsWith("/downloads/diagnosticos")) {
    throw new Error(
      `La carpeta FTP resuelta no coincide: ${resolvedRemoteDirectory}`,
    );
  }
  await client.uploadFromDir(localDirectory);
  console.log(`Diagnósticos publicados en ${resolvedRemoteDirectory}`);
} finally {
  client.close();
}
