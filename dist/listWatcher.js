"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.processEstadoListOnce = processEstadoListOnce;
// src/listWatcher.ts
const config_1 = require("./config");
const graph_1 = require("./graph");
const sendMail_1 = require("./sendMail");
const fs_1 = __importDefault(require("fs"));
const path_1 = __importDefault(require("path"));
const ESTADO_FILE = path_1.default.resolve(process.cwd(), "estado-notificaciones.json");
let estadoNotificado = {};
try {
    if (fs_1.default.existsSync(ESTADO_FILE)) {
        const raw = fs_1.default.readFileSync(ESTADO_FILE, "utf8");
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === "object") {
            for (const [id, val] of Object.entries(parsed)) {
                if (typeof val === "string") {
                    estadoNotificado[id] = { estado: val };
                }
                else {
                    estadoNotificado[id] = val;
                }
            }
        }
    }
}
catch {
    console.warn("⚠️ No pude leer estado-notificaciones.json, se reinicia vacío.");
    estadoNotificado = {};
}
function saveEstadoNotificado() {
    try {
        fs_1.default.writeFileSync(ESTADO_FILE, JSON.stringify(estadoNotificado, null, 2));
    }
    catch {
        console.warn("⚠️ No pude guardar estado-notificaciones.json");
    }
}
function normalizeEstado(raw) {
    const s = String(raw ?? "").trim();
    const n = s.toLowerCase();
    const isConfirmada = n === "confirmada" || n === "confirmado";
    const isRechazada = n === "rechazada" || n === "rechazado";
    const isFechaModificada = n === "fecha modificada" ||
        n === "fecha_modificada" ||
        n === "fechamodificada";
    const estadoParaCorreo = isConfirmada
        ? "Confirmada"
        : isRechazada
            ? "Rechazada"
            : isFechaModificada
                ? "Fecha modificada"
                : null;
    return { raw: s, norm: n, isConfirmada, isRechazada, isFechaModificada, estadoParaCorreo };
}
async function processEstadoListOnce() {
    const res = await (0, graph_1.graphGet)(`/sites/${config_1.cfg.siteId}/lists/${config_1.cfg.listId}/items?$expand=fields&$top=500`);
    if (!res.value?.length)
        return;
    for (const item of res.value) {
        const id = item.id;
        const f = item.fields || {};
        const titulo = f.Title || `Solicitud #${id}`;
        const cliente = f.Cliente_x002f_Proyecto;
        const fechaSolicitada = f.Fechasolicitada;
        const fechaConfirmada = f.FechaConfirmada;
        const comentario = f.Comentariodelencargado;
        const webUrl = item.webUrl;
        const email = String(f.SolicitanteEmail ?? f.Solicitante ?? "")
            .trim()
            .toLowerCase();
        if (!email || !email.includes("@"))
            continue;
        const prev = estadoNotificado[id] || {};
        const estadoPrevio = String(prev.estado ?? "");
        const comentarioPrevio = String(prev.comentario ?? "").trim();
        const comentarioActual = String(comentario ?? "").trim();
        const nuevo = { ...prev };
        const estadoInfo = normalizeEstado(f.Estadoderevisi_x00f3_n);
        // ✅ 1) Notificación por cambio de estado (Confirmada / Rechazada / Fecha modificada)
        if (estadoInfo.estadoParaCorreo && estadoPrevio !== estadoInfo.raw) {
            await (0, sendMail_1.sendMailCambioEstado)({
                to: email,
                titulo,
                estado: estadoInfo.estadoParaCorreo,
                cliente,
                fechaSolicitada,
                fechaConfirmada,
                comentarioEncargado: comentarioActual || undefined,
                webUrl,
            });
            nuevo.estado = estadoInfo.raw;
            console.log(`✉️ Notificación de estado "${estadoInfo.raw}" enviada para item ${id}`);
        }
        // ✅ 2) Notificación por nuevo/actualizado comentario del encargado
        if (comentarioActual && comentarioActual !== comentarioPrevio) {
            await (0, sendMail_1.sendMailComentarioEncargado)({
                to: email,
                titulo,
                cliente,
                comentarioEncargado: comentarioActual,
                webUrl,
            });
            nuevo.comentario = comentarioActual;
            console.log(`✉️ Notificación de comentario enviada para item ${id}`);
        }
        // guardar cambios si hubo notificación
        if (nuevo.estado !== estadoPrevio ||
            String(nuevo.comentario ?? "").trim() !== comentarioPrevio) {
            estadoNotificado[id] = nuevo;
        }
    }
    saveEstadoNotificado();
}
