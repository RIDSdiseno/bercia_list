"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getGraphToken = getGraphToken;
exports.getSharePointToken = getSharePointToken;
const msal_node_1 = require("@azure/msal-node");
const config_1 = require("./config");
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const spHost = config_1.cfg.siteId.split(",")[0];
const certPath = process.env.BERCIA_CERT_PATH;
const thumbprint = process.env.BERCIA_CERT_THUMBPRINT;
if (!certPath || !thumbprint) {
    throw new Error("Faltan BERCIA_CERT_PATH / BERCIA_CERT_THUMBPRINT en .env");
}
const privateKey = fs_1.default.readFileSync(path_1.default.resolve(certPath), "utf8");
const cca = new msal_node_1.ConfidentialClientApplication({
    auth: {
        clientId: config_1.cfg.clientId,
        authority: `https://login.microsoftonline.com/${config_1.cfg.tenantId}`,
        clientCertificate: { thumbprint, privateKey },
    },
});
async function getGraphToken() {
    const r = await cca.acquireTokenByClientCredential({
        scopes: ["https://graph.microsoft.com/.default"],
    });
    if (!r?.accessToken)
        throw new Error("No se pudo obtener token Graph");
    return r.accessToken;
}
async function getSharePointToken() {
    const r = await cca.acquireTokenByClientCredential({
        scopes: [`https://${spHost}/.default`],
    });
    if (!r?.accessToken)
        throw new Error("No se pudo obtener token SharePoint");
    return r.accessToken;
}
