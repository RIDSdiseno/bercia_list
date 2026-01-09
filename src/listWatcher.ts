// src/listWatcher.ts
import { cfg } from "./config.js";
import { graphGet, graphPatch } from "./graph.js";
import {
  sendMailCambioEstado,
  sendMailComentarioEncargado,
  type EstadoNotificable,
} from "./sendMail.js";
import fs from "node:fs";
import path from "node:path";

type ListItem = {
  id: string;
  webUrl?: string;
  fields: {
    Title?: string;
    Estadoderevisi_x00f3_n?: string;

    Solicitante?: string; // email texto
    SolicitanteEmail?: string; // recomendado

    Cliente_x002f_Proyecto?: string;
    Fechasolicitada?: string;
    FechaConfirmada?: string;
    Comentariodelencargado?: string;
  };
};

type EstadoNotificado = {
  estado?: string;
  comentario?: string | null;
};

const ESTADO_FILE = path.resolve(process.cwd(), "estado-notificaciones.json");
let estadoNotificado: Record<string, EstadoNotificado> = {};

try {
  if (fs.existsSync(ESTADO_FILE)) {
    const raw = fs.readFileSync(ESTADO_FILE, "utf8");
    const parsed = JSON.parse(raw);

    if (parsed && typeof parsed === "object") {
      for (const [id, val] of Object.entries(parsed as any)) {
        if (typeof val === "string") {
          estadoNotificado[id] = { estado: val };
        } else {
          estadoNotificado[id] = val as EstadoNotificado;
        }
      }
    }
  }
} catch {
  console.warn("⚠️ No pude leer estado-notificaciones.json, se reinicia vacío.");
  estadoNotificado = {};
}

function saveEstadoNotificado() {
  try {
    fs.writeFileSync(ESTADO_FILE, JSON.stringify(estadoNotificado, null, 2));
  } catch {
    console.warn("⚠️ No pude guardar estado-notificaciones.json");
  }
}

function normalizeEstado(raw: unknown): {
  raw: string;
  norm: string;
  isConfirmada: boolean;
  isRechazada: boolean;
  isFechaModificada: boolean;
  estadoParaCorreo: EstadoNotificable | null;
} {
  const s = String(raw ?? "").trim();
  const n = s.toLowerCase();

  const isConfirmada = n === "confirmada" || n === "confirmado";
  const isRechazada = n === "rechazada" || n === "rechazado";
  const isFechaModificada =
    n === "fecha modificada" ||
    n === "fecha_modificada" ||
    n === "fechamodificada";

  const estadoParaCorreo: EstadoNotificable | null = isConfirmada
    ? "Confirmada"
    : isRechazada
    ? "Rechazada"
    : isFechaModificada
    ? "Fecha modificada"
    : null;

  return {
    raw: s,
    norm: n,
    isConfirmada,
    isRechazada,
    isFechaModificada,
    estadoParaCorreo,
  };
}

export async function processEstadoListOnce() {
  const res = await graphGet<{ value: ListItem[] }>(
    `/sites/${cfg.siteId}/lists/${cfg.listId}/items?$expand=fields&$top=500`
  );

  if (!res.value?.length) return;

  for (const item of res.value) {
    const id = item.id;
    const f = item.fields || {};

    const titulo = f.Title || `Solicitud #${id}`;
    const cliente = f.Cliente_x002f_Proyecto;
    const fechaSolicitada = f.Fechasolicitada;
    let fechaConfirmada = f.FechaConfirmada; // 👈 let para poder actualizarlo
    const comentario = f.Comentariodelencargado;
    const webUrl = item.webUrl;

    const email = String(f.SolicitanteEmail ?? f.Solicitante ?? "")
      .trim()
      .toLowerCase();

    if (!email || !email.includes("@")) continue;

    const prev = estadoNotificado[id] || {};
    const estadoPrevio = String(prev.estado ?? "");

    const comentarioPrevio = String(prev.comentario ?? "").trim();
    const comentarioActual = String(comentario ?? "").trim();

    const nuevo: EstadoNotificado = { ...prev };

    const estadoInfo = normalizeEstado(f.Estadoderevisi_x00f3_n);

    // ==========================================================
    // ✅ AUTO-FECHA CONFIRMADA (solo si confirman manual y está vacía)
    // - Si ya viene fecha confirmada por correo o ya estaba seteada, NO se pisa.
    // ==========================================================
    const fechaConfirmadaActual = String(fechaConfirmada ?? "").trim();

    if (estadoInfo.isConfirmada && !fechaConfirmadaActual) {
      const nowIso = new Date().toISOString(); // ISO con Z (Graph-friendly)

      try {
        await graphPatch(
          `/sites/${cfg.siteId}/lists/${cfg.listId}/items/${id}/fields`,
          { FechaConfirmada: nowIso }
        );

        fechaConfirmada = nowIso;
        console.log(`🗓️ FechaConfirmada auto-set en item ${id}: ${nowIso}`);
      } catch (e: any) {
        console.warn(
          `⚠️ No pude setear FechaConfirmada automáticamente en item ${id}:`,
          e?.message || e
        );
      }
    }

    // ✅ 1) Notificación por cambio de estado (Confirmada / Rechazada / Fecha modificada)
    if (estadoInfo.estadoParaCorreo && estadoPrevio !== estadoInfo.raw) {
      await sendMailCambioEstado({
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
      console.log(
        `✉️ Notificación de estado "${estadoInfo.raw}" enviada para item ${id}`
      );
    }

    // ✅ 2) Notificación por nuevo/actualizado comentario del encargado
    if (comentarioActual && comentarioActual !== comentarioPrevio) {
      await sendMailComentarioEncargado({
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
    if (
      nuevo.estado !== estadoPrevio ||
      String(nuevo.comentario ?? "").trim() !== comentarioPrevio
    ) {
      estadoNotificado[id] = nuevo;
    }
  }

  saveEstadoNotificado();
}
