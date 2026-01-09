// src/sharepoint.ts
import { graphPost } from "./graph.js";
import { cfg } from "./config.js";

/** Helper genérico para crear items en cualquier lista */
export async function createListItem(fields: any) {
  const safeFields: any = { ...fields };

  // Solo mandamos ContentTypeId si viene configurado y no vacío
  if (cfg.contentTypeId && cfg.contentTypeId.trim()) {
    safeFields.ContentTypeId = cfg.contentTypeId.trim();
  }

  // Limpia null/undefined
  for (const k of Object.keys(safeFields)) {
    if (safeFields[k] === null || safeFields[k] === undefined) {
      delete safeFields[k];
    }
  }

  try {
    // Primer intento: con todos los campos
    return await graphPost(
      `/sites/${cfg.siteId}/lists/${cfg.listId}/items`,
      { fields: safeFields }
    );
  } catch (e: any) {
    const msg: string =
      e?.response?.data?.error?.message || e?.message || "";

    // Si el problema es específicamente Observaciones, reintenta sin ese campo
    if (
      msg.includes("Field 'Observaciones' is not recognized") ||
      msg.includes("Field 'Obervaciones' is not recognized")
    ) {
      console.warn(
        "⚠️ Campo Observaciones no reconocido por Graph, reintento sin ese campo"
      );

      delete safeFields.Observaciones;
      delete safeFields.Obervaciones;

      return await graphPost(
        `/sites/${cfg.siteId}/lists/${cfg.listId}/items`,
        { fields: safeFields }
      );
    }

    // Cualquier otro error se propaga
    throw e;
  }
}

/** 🔹 Tipo “bonito” para tu backend (sin nombres raros) */
export type BerciaSolicitudInput = {
  title?: string;
  clienteProyecto?: string;
  tipoTarea?: string;
  fechaSolicitadaIso?: string;
  observaciones?: string;
  estadoRevision?: "Pendiente" | "Fecha modificada" | "Rechazada" | "Confirmada";
  fechaConfirmadaIso?: string;
  comentarioEncargado?: string;
  responsable?: string;
  solicitante?: string;
  prioridad?: "Alta" | "Media" | "Baja";
  identificador?: string;
  documentos?: string;
};

/** 🔹 Wrapper para la lista de Bercia con los internal names reales */
export async function createBerciaSolicitudItem(input: BerciaSolicitudInput) {
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
