"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.createListItem = createListItem;
exports.createBerciaSolicitudItem = createBerciaSolicitudItem;
// src/sharepoint.ts
const graph_1 = require("./graph");
const config_1 = require("./config");
/** Helper genérico para crear items en cualquier lista */
async function createListItem(fields) {
    const safeFields = { ...fields };
    // Solo mandamos ContentTypeId si viene configurado y no vacío
    if (config_1.cfg.contentTypeId && config_1.cfg.contentTypeId.trim()) {
        safeFields.ContentTypeId = config_1.cfg.contentTypeId.trim();
    }
    // Limpia null/undefined
    for (const k of Object.keys(safeFields)) {
        if (safeFields[k] === null || safeFields[k] === undefined) {
            delete safeFields[k];
        }
    }
    try {
        // Primer intento: con todos los campos
        return await (0, graph_1.graphPost)(`/sites/${config_1.cfg.siteId}/lists/${config_1.cfg.listId}/items`, { fields: safeFields });
    }
    catch (e) {
        const msg = e?.response?.data?.error?.message || e?.message || "";
        // Si el problema es específicamente Observaciones, reintenta sin ese campo
        if (msg.includes("Field 'Observaciones' is not recognized") ||
            msg.includes("Field 'Obervaciones' is not recognized")) {
            console.warn("⚠️ Campo Observaciones no reconocido por Graph, reintento sin ese campo");
            delete safeFields.Observaciones;
            delete safeFields.Obervaciones;
            return await (0, graph_1.graphPost)(`/sites/${config_1.cfg.siteId}/lists/${config_1.cfg.listId}/items`, { fields: safeFields });
        }
        // Cualquier otro error se propaga
        throw e;
    }
}
/** 🔹 Wrapper para la lista de Bercia con los internal names reales */
async function createBerciaSolicitudItem(input) {
    const fields = {
        Title: input.title ?? "Solicitud desde correo",
        // Internal names EXACTOS de la lista
        Cliente_x002f_Proyecto: input.clienteProyecto,
        Tipodetarea: input.tipoTarea,
        Fechasolicitada: input.fechaSolicitadaIso,
        Observaciones: input.observaciones,
        Estadoderevisi_x00f3_n: input.estadoRevision ?? "Pendiente",
        FechaConfirmada: input.fechaConfirmadaIso,
        Comentariodelencargado: input.comentarioEncargado,
        Responsable: input.responsable,
        Solicitante: input.solicitante,
        Prioridad: input.prioridad,
        Identificador: input.identificador,
        Documentos: input.documentos,
    };
    return createListItem(fields);
}
