// src/listWatcher.ts
import { cfg } from "./config";
import { graphGet, graphPatch } from "./graph";
import { sendMailCambioEstado } from "./sendMail";
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

// ======= MEMORIA DE ESTADOS NOTIFICADOS =======
const ESTADO_FILE = path.resolve(process.cwd(), "estado-notificaciones.json");

let estadoNotificado: Record<string, string> = {};
try {
  if (fs.existsSync(ESTADO_FILE)) {
    estadoNotificado = JSON.parse(fs.readFileSync(ESTADO_FILE, "utf8"));
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

// 🔸 Fecha/hora en formato “local” para SharePoint (sin Z)
function nowForSharePoint(): string {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");
  // 👇 importante: NO poner "Z" al final
  return `${yyyy}-${mm}-${dd}T${hh}:${mi}:${ss}`;
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

    if (!email) continue;

    // Solo nos interesan Confirmada / Rechazada
    if (estado !== "Confirmada" && estado !== "Rechazada") {
      continue;
    }

    const yaNotificado = estadoNotificado[id];
    const titulo = f.Title || `Solicitud #${id}`;
    const cliente = f.Cliente_x002f_Proyecto;
    const fechaSolicitada = f.Fechasolicitada;
    let fechaConfirmada = f.FechaConfirmada;
    const comentario = f.Comentariodelencargado;
    const webUrl = item.webUrl;

    // ✅ Si el estado es Confirmada y aún no hay FechaConfirmada, la seteamos ahora
    if (estado === "Confirmada" && !fechaConfirmada) {
      const ahora = nowForSharePoint();
      try {
        await graphPatch(
          `/sites/${cfg.siteId}/lists/${cfg.listId}/items/${id}/fields`,
          {
            FechaConfirmada: ahora,
          }
        );
        fechaConfirmada = ahora;
        console.log(`📅 FechaConfirmada seteada para item ${id}: ${ahora}`);
      } catch (e: any) {
        console.warn(
          `⚠️ No pude actualizar FechaConfirmada en item ${id}:`,
          e?.message || e
        );
      }
    }

    // Si ya enviamos correo para este mismo estado, no repetir
    if (yaNotificado === estado) {
      continue;
    }

    // ✉️ Enviar correo de cambio de estado
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

    estadoNotificado[id] = estado;
    saveEstadoNotificado();

    console.log(
      `✉️ Notificación de estado "${estado}" enviada para item ${id}`
    );
  }
}
