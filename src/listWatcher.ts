// src/listWatcher.ts
import { cfg } from "./config";
import { graphGet } from "./graph";
import {
  sendMailCambioEstado,
  sendMailComentarioEncargado,
  type EstadoNotificable,
} from "./sendMail";
import fs from "fs";
import path from "path";

type ListItem = {
  id: string;
  webUrl?: string;
  fields: {
    Title?: string;
    Estadoderevisi_x00f3_n?: string;

    Solicitante?: string;       // email texto
    SolicitanteEmail?: string;  // recomendado

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

  return { raw: s, norm: n, isConfirmada, isRechazada, isFechaModificada, estadoParaCorreo };
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
    const fechaConfirmada = f.FechaConfirmada;
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
      console.log(`✉️ Notificación de estado "${estadoInfo.raw}" enviada para item ${id}`);
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
