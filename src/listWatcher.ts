// src/listWatcher.ts
import { cfg } from "./config";
import { graphGet } from "./graph";
import {
  sendMailCambioEstado,
  sendMailComentarioEncargado,
} from "./sendMail";
import fs from "fs";
import path from "path";

type ListItem = {
  id: string;
  webUrl?: string;
  fields: {
    Title?: string;
    Estadoderevisi_x00f3_n?: string;
    Solicitante?: string; // email texto
    Cliente_x002f_Proyecto?: string;
    Fechasolicitada?: string;
    FechaConfirmada?: string;
    Comentariodelencargado?: string;
  };
};

type EstadoNotificado = {
  estado?: string;            // último estado notificado (Confirmada/Rechazada)
  comentario?: string | null; // último comentario notificado
};

// ======= MEMORIA DE ESTADOS / COMENTARIOS NOTIFICADOS =======
const ESTADO_FILE = path.resolve(process.cwd(), "estado-notificaciones.json");

let estadoNotificado: Record<string, EstadoNotificado> = {};

try {
  if (fs.existsSync(ESTADO_FILE)) {
    const raw = fs.readFileSync(ESTADO_FILE, "utf8");
    const parsed = JSON.parse(raw);

    if (parsed && typeof parsed === "object") {
      for (const [id, val] of Object.entries(parsed as any)) {
        if (typeof val === "string") {
          // compatibilidad con versión antigua: solo guardaba el estado
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

export async function processEstadoListOnce() {
  const res = await graphGet<{ value: ListItem[] }>(
    `/sites/${cfg.siteId}/lists/${cfg.listId}/items?$expand=fields&$top=500`
  );

  if (!res.value?.length) return;

  for (const item of res.value) {
    const id = item.id;
    const f = item.fields || {};
    const estado = f.Estadoderevisi_x00f3_n || "";
    const email = (f.Solicitante || "").trim().toLowerCase();
    const titulo = f.Title || `Solicitud #${id}`;
    const cliente = f.Cliente_x002f_Proyecto;
    const fechaSolicitada = f.Fechasolicitada;
    const fechaConfirmada = f.FechaConfirmada;
    const comentario = f.Comentariodelencargado;
    const webUrl = item.webUrl;

    if (!email) continue;

    const previo: EstadoNotificado = estadoNotificado[id] || {};
    const estadoPrevio = previo.estado || "";
    const comentarioPrevio = (previo.comentario || "").trim();
    const comentarioActual = (comentario || "").trim();

    const nuevo: EstadoNotificado = { ...previo };

    // 1) Notificación por cambio de estado (solo Confirmada / Rechazada)
    if ((estado === "Confirmada" || estado === "Rechazada") && estadoPrevio !== estado) {
      await sendMailCambioEstado({
        to: email,
        titulo,
        estado: estado as "Confirmada" | "Rechazada",
        cliente,
        fechaSolicitada,
        fechaConfirmada,
        comentarioEncargado: comentario,
        webUrl,
      });

      nuevo.estado = estado;
      console.log(`✉️ Notificación de estado "${estado}" enviada para item ${id}`);
    }

    // 2) Notificación por nuevo/actualizado comentario del encargado
    if (comentarioActual && comentarioActual !== comentarioPrevio) {
      await sendMailComentarioEncargado({
        to: email,
        titulo,
        cliente,
        comentarioEncargado: comentarioActual,
        webUrl,
      });

      nuevo.comentario = comentarioActual;
      console.log(
        `✉️ Notificación de comentario del encargado enviada para item ${id}`
      );
    }

    // si hubo cambios, guardar
    if (
      nuevo.estado !== estadoPrevio ||
      (nuevo.comentario || "").trim() !== comentarioPrevio
    ) {
      estadoNotificado[id] = nuevo;
    }
  }

  saveEstadoNotificado();
}
