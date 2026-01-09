// src/mailProcessor.ts
import { cfg } from "./config.js";
import { graphGet } from "./graph.js";
import { createListItem } from "./sharepoint.js";
import { parseMail } from "./parser.js";
import { sendMailNuevaSolicitud } from "./sendMail.js";
import { getSiteUserLookupId } from "./spUsers.js";
import { spSetResponsables } from "./spList.js";
import fs from "node:fs";
import path from "node:path";
// ================== ANTI-DUPLICADOS ==================
const PROCESSED_FILE = path.resolve(process.cwd(), "processed-mails.json");
let processedIds = new Set();
try {
    if (fs.existsSync(PROCESSED_FILE)) {
        const raw = fs.readFileSync(PROCESSED_FILE, "utf8");
        const arr = JSON.parse(raw);
        if (Array.isArray(arr))
            processedIds = new Set(arr.filter(Boolean));
        console.log(`🧠 processedIds cargados: ${processedIds.size}`);
    }
}
catch {
    console.warn("⚠️ No pude leer processed-mails.json, se reinicia vacío.");
    processedIds = new Set();
}
function markProcessed(id) {
    processedIds.add(id);
    try {
        fs.writeFileSync(PROCESSED_FILE, JSON.stringify([...processedIds], null, 2));
    }
    catch {
        console.warn("⚠️ No pude guardar processed-mails.json");
    }
}
// =====================================================
const TIPOS_VALIDOS = new Set([
    "envio",
    "instalacion",
    "Cubicación por planos",
    "Cubicación en terreno",
    "Costeo de proyecto",
    "Productos (requerimientos)",
    "Evaluación postventa",
    "Post venta en terreno",
    "Producto interno",
]);
function normalizeTipodetarea(raw) {
    if (!raw)
        return null;
    const s = raw.trim();
    const lower = s.toLowerCase();
    if (lower.includes("env"))
        return "envio";
    if (lower.includes("instal"))
        return "instalacion";
    if (lower.includes("postventa"))
        return "Evaluación postventa";
    if (lower.includes("producto interno"))
        return "Producto interno";
    if (lower.includes("costeo"))
        return "Costeo de proyecto";
    if (lower.includes("planos"))
        return "Cubicación por planos";
    if (lower.includes("terreno"))
        return "Cubicación en terreno";
    if (lower.includes("requer"))
        return "Productos (requerimientos)";
    if (TIPOS_VALIDOS.has(s))
        return s;
    return null;
}
function normalizePrioridad(raw) {
    if (!raw)
        return null;
    const s = raw.trim().toLowerCase();
    if (s.startsWith("a"))
        return "Alta";
    if (s.startsWith("m"))
        return "Media";
    if (s.startsWith("b"))
        return "Baja";
    return null;
}
// 🔹 Fecha/hora local para SharePoint (sin "Z")
function nowForSharePoint() {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const dd = String(now.getDate()).padStart(2, "0");
    const hh = String(now.getHours()).padStart(2, "0");
    const mi = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}T${hh}:${mi}:${ss}`;
}
function htmlToText(html) {
    if (!html)
        return "";
    let text = html
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<\/p>/gi, "\n")
        .replace(/<\/div>/gi, "\n")
        .replace(/<\/li>/gi, "\n")
        .replace(/<\/tr>/gi, "\n");
    text = text
        .replace(/<style[\s\S]*?<\/style>/gi, " ")
        .replace(/<script[\s\S]*?<\/script>/gi, " ")
        .replace(/<[^>]+>/g, " ");
    text = text.replace(/&nbsp;/g, " ").replace(/&amp;/g, "&");
    const lines = text
        .replace(/\r/g, "")
        .split("\n")
        .map((l) => l.replace(/\s+/g, " ").trim())
        .filter(Boolean);
    return lines.join("\n");
}
// ====== MAIN POLLING ======
export async function processInboxOnce() {
    const folderId = await resolveFolderIdByName();
    const res = await graphGet(`/users/${cfg.mailboxUserId}/mailFolders/${folderId}/messages?$top=25`);
    if (!res.value?.length)
        return;
    const site = await graphGet(`/sites/${cfg.siteId}`);
    const webUrl = site?.webUrl;
    if (!webUrl)
        throw new Error("No pude obtener webUrl del sitio");
    for (const m of res.value) {
        if (processedIds.has(m.id))
            continue;
        const subjectLower = (m.subject || "").toLowerCase();
        if (!subjectLower.includes("list"))
            continue;
        const fromEmail = m.from?.emailAddress?.address?.trim().toLowerCase() || "";
        const solicitanteMail = fromEmail;
        const responsablesMails = Array.from(new Set((m.ccRecipients ?? [])
            .map((r) => r.emailAddress?.address)
            .filter((mail) => typeof mail === "string")
            .map((mail) => mail.trim().toLowerCase())
            .filter((mail) => mail !== cfg.adminEmail)
            .filter((mail) => mail !== solicitanteMail)));
        const bodyHtml = m.body?.content || "";
        const bodyText = htmlToText(bodyHtml);
        const parsed = parseMail(bodyText);
        const fechaSolicitadaValue = parsed.fechaSolicitada
            ? normalizeDate(parsed.fechaSolicitada)
            : nowForSharePoint();
        const fechaConfirmadaValue = parsed.fechaConfirmada
            ? normalizeDate(parsed.fechaConfirmada)
            : undefined;
        const fields = {
            Title: m.subject || "Solicitud",
            Cliente_x002f_Proyecto: parsed.clienteProyecto || "Sin cliente",
            Observaciones: parsed.observaciones || "Sin observaciones",
            Estadoderevisi_x00f3_n: "Pendiente",
            // ✅ texto
            Solicitante: solicitanteMail || "",
            Responsable: responsablesMails.join("; "),
            Fechasolicitada: fechaSolicitadaValue,
        };
        if (fechaConfirmadaValue)
            fields.FechaConfirmada = fechaConfirmadaValue;
        fields.Tipodetarea = normalizeTipodetarea(parsed.tipodetarea) ?? "envio";
        const prioOk = normalizePrioridad(parsed.prioridad);
        if (prioOk)
            fields.Prioridad = prioOk;
        // Solicitante persona
        try {
            if (solicitanteMail) {
                const solicitanteId = await getSiteUserLookupId(solicitanteMail, webUrl);
                if (solicitanteId)
                    fields.Solicitante0LookupId = solicitanteId;
            }
        }
        catch {
            console.warn("⚠️ No se pudo resolver solicitante como persona:", solicitanteMail);
        }
        const created = await createListItem(fields);
        const itemId = Number(created?.id);
        const itemUrl = created?.webUrl;
        // set multi-persona Responsables
        if (Number.isFinite(itemId) && responsablesMails.length) {
            const responsablesIds = [];
            for (const mail of responsablesMails) {
                try {
                    const rid = await getSiteUserLookupId(mail, webUrl);
                    if (rid && !responsablesIds.includes(rid))
                        responsablesIds.push(rid);
                }
                catch {
                    console.warn("⚠️ Responsable no resolvió como persona:", mail);
                }
            }
            if (responsablesIds.length) {
                await spSetResponsables(webUrl, cfg.listId, itemId, responsablesIds);
            }
        }
        // mail al solicitante
        if (solicitanteMail) {
            await sendMailNuevaSolicitud({
                to: solicitanteMail,
                titulo: fields.Title,
                cliente: fields.Cliente_x002f_Proyecto,
                fechaSolicitada: fields.Fechasolicitada,
                tipodetarea: fields.Tipodetarea,
                webUrl: itemUrl,
            });
        }
        markProcessed(m.id);
        console.log("✅ Item creado + correo enviado:", m.subject);
    }
}
async function resolveFolderIdByName() {
    const folders = await graphGet(`/users/${cfg.mailboxUserId}/mailFolders?$top=250`);
    const match = folders.value.find((f) => (f.displayName || "").toLowerCase() ===
        cfg.targetFolderPath.toLowerCase());
    if (!match?.id) {
        throw new Error(`No encontré carpeta "${cfg.targetFolderPath}". Revisa TARGET_FOLDER_PATH.`);
    }
    return match.id;
}
export async function processSimulatedMail(input) {
    const subject = (input.subject || "").trim();
    const fromEmail = (input.from || "").trim().toLowerCase();
    const ccMailsRaw = (input.cc || []).map((x) => (x || "").trim().toLowerCase());
    const bodyRaw = input.body || "";
    if (!subject)
        throw new Error("subject es obligatorio");
    if (!fromEmail || !fromEmail.includes("@")) {
        throw new Error("from (email) es obligatorio y debe ser válido");
    }
    const solicitanteMail = fromEmail;
    const responsablesMails = Array.from(new Set(ccMailsRaw
        .filter((mail) => mail !== cfg.adminEmail)
        .filter((mail) => mail !== solicitanteMail)));
    const bodyText = htmlToText(bodyRaw);
    const parsed = parseMail(bodyText);
    const site = await graphGet(`/sites/${cfg.siteId}`);
    const webUrl = site?.webUrl;
    if (!webUrl)
        throw new Error("No pude obtener webUrl del sitio");
    const fechaSolicitadaValue = parsed.fechaSolicitada
        ? normalizeDate(parsed.fechaSolicitada)
        : nowForSharePoint();
    const fechaConfirmadaValue = parsed.fechaConfirmada
        ? normalizeDate(parsed.fechaConfirmada)
        : undefined;
    const fields = {
        Title: subject || "Solicitud",
        Cliente_x002f_Proyecto: parsed.clienteProyecto || "Sin cliente",
        Observaciones: parsed.observaciones || "Sin observaciones",
        Estadoderevisi_x00f3_n: "Pendiente",
        Solicitante: solicitanteMail || "",
        Responsable: responsablesMails.join("; "),
        Fechasolicitada: fechaSolicitadaValue,
    };
    if (fechaConfirmadaValue)
        fields.FechaConfirmada = fechaConfirmadaValue;
    fields.Tipodetarea = normalizeTipodetarea(parsed.tipodetarea) ?? "envio";
    const prioOk = normalizePrioridad(parsed.prioridad);
    if (prioOk)
        fields.Prioridad = prioOk;
    try {
        const solicitanteId = await getSiteUserLookupId(solicitanteMail, webUrl);
        if (solicitanteId)
            fields.Solicitante0LookupId = solicitanteId;
    }
    catch {
        console.warn("⚠️ No se pudo resolver solicitante como persona (test):", solicitanteMail);
    }
    const created = await createListItem(fields);
    const itemId = Number(created?.id);
    const itemUrl = created?.webUrl;
    if (Number.isFinite(itemId) && responsablesMails.length) {
        const responsablesIds = [];
        for (const mail of responsablesMails) {
            try {
                const rid = await getSiteUserLookupId(mail, webUrl);
                if (rid && !responsablesIds.includes(rid))
                    responsablesIds.push(rid);
            }
            catch {
                console.warn("⚠️ Responsable no resolvió como persona (test):", mail);
            }
        }
        if (responsablesIds.length) {
            await spSetResponsables(webUrl, cfg.listId, itemId, responsablesIds);
        }
    }
    if (solicitanteMail) {
        await sendMailNuevaSolicitud({
            to: solicitanteMail,
            titulo: fields.Title,
            cliente: fields.Cliente_x002f_Proyecto,
            fechaSolicitada: fields.Fechasolicitada,
            tipodetarea: fields.Tipodetarea,
            webUrl: itemUrl,
        });
    }
    console.log("🧪 [test-create] Item creado + correo enviado:", subject);
    return created;
}
function normalizeDate(input) {
    const s = input.trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(s))
        return s;
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m)
        return s;
    const dd = m[1].padStart(2, "0");
    const mm = m[2].padStart(2, "0");
    const yyyy = m[3];
    // recomendado para SharePoint local (sin Z)
    return `${yyyy}-${mm}-${dd}T00:00:00`;
}
