import { ConfidentialClientApplication } from "@azure/msal-node";
import { cfg } from "./config";
import fs from "fs";
import path from "path";

const spHost = cfg.siteId.split(",")[0];

const certPath = process.env.BERCIA_CERT_PATH;
const thumbprint = process.env.BERCIA_CERT_THUMBPRINT;

if (!certPath || !thumbprint) {
  throw new Error("Faltan BERCIA_CERT_PATH / BERCIA_CERT_THUMBPRINT en .env");
}

const privateKey = fs.readFileSync(path.resolve(certPath), "utf8");

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: cfg.clientId,
    authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
    clientCertificate: { thumbprint, privateKey },
  },
});

export async function getGraphToken() {
  const r = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  if (!r?.accessToken) throw new Error("No se pudo obtener token Graph");
  return r.accessToken;
}

export async function getSharePointToken() {
  const r = await cca.acquireTokenByClientCredential({
    scopes: [`https://${spHost}/.default`],
  });
  if (!r?.accessToken) throw new Error("No se pudo obtener token SharePoint");
  return r.accessToken;
}
