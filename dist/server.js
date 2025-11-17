"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// src/server.ts
require("dotenv/config");
const express_1 = __importDefault(require("express"));
const body_parser_1 = __importDefault(require("body-parser"));
const graph_1 = require("./scripts/graph");
const mail_1 = require("./scripts/mail");
const sharepoint_1 = require("./scripts/sharepoint");
const folders_1 = require("./scripts/folders");
const subscription_1 = require("./scripts/subscription");
/* ===================== Configuración & Utils ===================== */
const { PORT = "4000", TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAILBOX_USER_ID, SITE_ID, LIST_ID, WEBHOOK_URL, TARGET_FOLDER_PATH = "Prueba-Flujo-list", BERCIA_DOMAIN = "@bercia.cl", PA_SHARED_KEY, } = process.env;
function requireEnv(vars) {
    const missing = vars.filter(([, v]) => !v).map(([k]) => k);
    if (missing.length) {
        throw new Error(`Faltan variables de entorno: ${missing.join(", ")}. Revisa tu .env o configuración.`);
    }
}
function requireGraphBase() {
    requireEnv([
        ["TENANT_ID", TENANT_ID],
        ["CLIENT_ID", CLIENT_ID],
        ["CLIENT_SECRET", CLIENT_SECRET],
    ]);
}
function requireSharePointBase() {
    requireEnv([
        ["SITE_ID", SITE_ID],
        ["LIST_ID", LIST_ID],
        ["MAILBOX_USER_ID", MAILBOX_USER_ID],
    ]);
}
function requireWebhookBase() {
    requireEnv([["WEBHOOK_URL", WEBHOOK_URL]]);
}
function requirePAKey() {
    requireEnv([["PA_SHARED_KEY", PA_SHARED_KEY]]);
}
// Normaliza string de correos separados por ;
function normalizeToCc(raw) {
    if (typeof raw !== "string")
        return "";
    return raw.split(";").map((s) => s.trim()).filter(Boolean).join(";");
}
// Extrae emails desde “Nombre <correo>”, comas o ;
function extractEmails(input) {
    const s = String(input ?? "");
    const matches = s.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [];
    return Array.from(new Set(matches.map((e) => e.toLowerCase().trim())));
}
/* ===================== App ===================== */
const app = (0, express_1.default)();
// PA a veces manda content-types raros
app.use(body_parser_1.default.json({ type: "*/*", limit: "2mb" }));
/* ========= Health ========= */
app.get("/", (_req, res) => res.send("OK"));
app.get("/health", (_req, res) => res.status(200).send("ok"));
app.get("/api/graph/health", (_req, res) => res.json({ ok: true, mailbox: MAILBOX_USER_ID, site: SITE_ID, list: LIST_ID }));
/* ========= Intake (Power Automate / Postman → SharePoint) ========= */
app.post("/api/intake/email", async (req, res) => {
    const startedAt = new Date().toISOString();
    try {
        // 1) Seguridad header
        requirePAKey();
        const headerKey = String(req.headers["x-pa-key"] || "").trim();
        const sharedKey = String(PA_SHARED_KEY || "").trim();
        if (!headerKey || headerKey !== sharedKey) {
            return res.status(401).json({ error: "unauthorized" });
        }
        // 2) Envs necesarias
        requireGraphBase();
        requireSharePointBase();
        // 3) Cuerpo
        const { subject, from, toCcBercia, // puede venir con comas, ; o “Nombre <correo>”
        bodyPreview, bodyHtml, receivedDateTime, // opcional
         } = req.body ?? {};
        const token = await (0, graph_1.getAppToken)(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
        const texto = `${subject ?? ""}\n${bodyHtml || bodyPreview || ""}`;
        const prioridad = (0, mail_1.guessPrioridad)(texto);
        const tipoTarea = (0, mail_1.guessTipoTarea)(texto);
        const fechaSolicitada = (0, mail_1.extractFirstDateISO)(bodyHtml || bodyPreview || "");
        const clienteProyecto = (0, mail_1.extractClientProject)(subject ?? "", bodyHtml || bodyPreview || "");
        // ====== Emails LIMPIOS ======
        const solicitanteEmail = extractEmails(from)[0] ?? (typeof from === "string" ? from.trim().toLowerCase() : "");
        // Preferimos extraer por regex (acepta comas/; y “Nombre <mail>”)
        let responsablesArr = extractEmails(toCcBercia);
        if (responsablesArr.length === 0) {
            // Si PA te lo manda como “a,b,c” o con ; mezclados
            const raw = String(toCcBercia ?? "").replace(/,/g, ";");
            responsablesArr = normalizeToCc(raw)
                .split(";")
                .map((s) => s.trim().toLowerCase())
                .filter(Boolean);
        }
        // Fallback obligatorio: centro de correo siempre como responsable principal
        if (responsablesArr.length === 0) {
            responsablesArr = ["administrador@bercia.cl"];
        }
        // (Opcional) advertencia por dominio
        if (BERCIA_DOMAIN && typeof from === "string" && BERCIA_DOMAIN.length > 2) {
            if (!from.toLowerCase().includes(BERCIA_DOMAIN.toLowerCase())) {
                console.warn(`[WARN] Remitente distinto de dominio esperado (${BERCIA_DOMAIN}):`, from);
            }
        }
        // 4) Construir payload
        const fields = {
            Title: subject ?? "(sin asunto)",
            Observaciones: (0, mail_1.truncate)(bodyPreview || "", 1800),
            Notificado: Boolean(from),
            Cliente_x002f_Proyecto: clienteProyecto ?? "",
            // Si necesitas guardar la fecha de recepción original:
            // ReceivedDateTime: receivedDateTime ?? undefined,
        };
        // Fecha solicitada (si viene en el correo)
        if (fechaSolicitada) {
            const iso = new Date(fechaSolicitada).toISOString();
            if (!isNaN(Date.parse(iso)))
                fields.Fechasolicitada = iso;
        }
        // Choices
        const ESTADO_CHOICES = ["Pendiente", "En revisión", "Completado"];
        if (ESTADO_CHOICES.includes("Pendiente"))
            fields.Estadoderevisi_x00f3_n = "Pendiente";
        if (mail_1.PRIORIDAD_CHOICES.includes(prioridad))
            fields.Prioridad = prioridad;
        if (mail_1.TIPO_TAREA_CHOICES.includes(tipoTarea))
            fields.Tipodetarea = tipoTarea;
        // ====== Campos Persona (SharePoint People) vía Graph ======
        // Solicitante: single person (UPN/email)
        if (solicitanteEmail) {
            fields["Solicitante@odata.type"] = "String";
            fields["Solicitante"] = solicitanteEmail;
        }
        // Responsable: por ahora single (toma el primero). Si configuras la columna como múltiple, usa el bloque de abajo.
        if (responsablesArr.length > 0) {
            fields["Responsable@odata.type"] = "String";
            fields["Responsable"] = responsablesArr[0];
        }
        // === Si cambias “Responsable” a MÚLTIPLE en SharePoint, usa esto en su lugar:
        // if (responsablesArr.length > 1) {
        //   fields["Responsable@odata.type"] = "Collection(Edm.String)";
        //   fields["Responsable"] = responsablesArr;
        // }
        console.log("INTAKE responsables:", { toCcBercia, responsablesArr });
        await (0, sharepoint_1.createListItem)(token, { siteId: SITE_ID, listId: LIST_ID, fields });
        // 5) Notificación opcional al solicitante
        if (from) {
            await (0, mail_1.sendConfirmationEmail)(token, MAILBOX_USER_ID, solicitanteEmail || String(from), subject ?? "(sin asunto)");
        }
        return res.json({ ok: true, startedAt, finishedAt: new Date().toISOString() });
    }
    catch (e) {
        const payload = e?.response?.data ?? e?.message ?? String(e);
        console.error("intake error:", payload);
        return res.status(500).json({ error: payload });
    }
});
/* ========= Webhook Graph (opcional, para suscripciones) ========= */
app.get("/api/graph/webhook", (req, res) => {
    const token = req.query.validationToken;
    if (token)
        return res.status(200).type("text/plain").send(token);
    return res.sendStatus(200);
});
app.post("/api/graph/webhook", async (req, res) => {
    const validationToken = req.query?.validationToken;
    if (validationToken) {
        return res.status(200).type("text/plain").send(validationToken);
    }
    res.sendStatus(202);
    // TODO: procesar notificaciones de Graph aquí si vuelves al modo webhook
});
/* ========= Crear suscripciones (Inbox + carpeta objetivo) ========= */
app.post("/api/graph/subscribe", async (_req, res) => {
    try {
        requirePAKey();
        requireGraphBase();
        requireSharePointBase();
        requireWebhookBase();
        const token = await (0, graph_1.getAppToken)(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
        const folderId = await (0, folders_1.ensureFolderPath)(token, MAILBOX_USER_ID, TARGET_FOLDER_PATH);
        const subs = [];
        subs.push(await (0, subscription_1.createSubscription)(token, `/users/${MAILBOX_USER_ID}/mailFolders('inbox')/messages`, WEBHOOK_URL));
        subs.push(await (0, subscription_1.createSubscription)(token, `/users/${MAILBOX_USER_ID}/mailFolders/${folderId}/messages`, WEBHOOK_URL));
        return res.json({ ok: true, folderId, subs });
    }
    catch (e) {
        const payload = e?.response?.data ?? e?.message ?? String(e);
        console.error("subscribe error:", payload);
        return res.status(500).json({ error: payload });
    }
});
/* ===================== Start ===================== */
app.listen(Number(PORT), "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
});
exports.default = app;
//# sourceMappingURL=server.js.map