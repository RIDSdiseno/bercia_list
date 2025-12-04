// src/mailProcessor.ts
import { cfg } from "./config";
import { graphGet } from "./graph";
import { createListItem } from "./sharepoint";
import { parseMail } from "./parser";
import { getSiteUserLookupId } from "./spUsers";
import { spSetResponsables } from "./spList";
import { sendMailNuevaSolicitud } from "./sendMail";
import fs from "fs";
import path from "path";

type EmailAddress = { address?: string; name?: string };
type Recipient = { emailAddress?: EmailAddress };

type Msg = {
  id: string;
  subject: string;
  from?: Recipient;
  ccRecipients?: Recipient[];
  body?: { contentType?: string; content?: string };
  isRead?: boolean;
};

// ================== ANTI-DUPLICADOS ==================
const PROCESSED_FILE = path.resolve(process.cwd(), "processed-mails.json");

let processedIds = new Set<string>();
try {
  if (fs.existsSync(PROCESSED_FILE)) {
    const raw = fs.readFileSync(PROCESSED_FILE, "utf8");
    const arr = JSON.parse(raw);
    if (Array.isArray(arr)) processedIds = new Set(arr.filter(Boolean));
    console.log(`🧠 processedIds cargados: ${processedIds.size}`);
  }
} catch {
  console.warn("⚠️ No pude leer processed-mails.json, se reinicia vacío.");
  processedIds = new Set();
}

function markProcessed(id: string) {
  processedIds.add(id);
  try {
    fs.writeFileSync(PROCESSED_FILE, JSON.stringify([...processedIds], null, 2));
  } catch {
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

function normalizeTipodetarea(raw?: string) {
  if (!raw) return null;
  const s = raw.trim();
  const lower = s.toLowerCase();

  if (lower.includes("env")) return "envio";
  if (lower.includes("instal")) return "instalacion";
  if (lower.includes("postventa")) return "Evaluación postventa";
  if (lower.includes("producto interno")) return "Producto interno";
  if (lower.includes("costeo")) return "Costeo de proyecto";
  if (lower.includes("planos")) return "Cubicación por planos";
  if (lower.includes("terreno")) return "Cubicación en terreno";
  if (lower.includes("requer")) return "Productos (requerimientos)";

  if (TIPOS_VALIDOS.has(s)) return s;
  return null;
}

function normalizePrioridad(raw?: string) {
  if (!raw) return null;
  const s = raw.trim().toLowerCase();
  if (s.startsWith("a")) return "Alta";
  if (s.startsWith("m")) return "Media";
  if (s.startsWith("b")) return "Baja";
  return null;
}

// 🔹 Fecha/hora local para SharePoint (sin "Z")
function nowForSharePoint(): string {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}T${hh}:${mi}:${ss}`;
}

function htmlToText(html: string) {
  if (!html) return "";

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
    .map(l => l.replace(/\s+/g, " ").trim())
    .filter(Boolean);

  return lines.join("\n");
}

// ====== MAIN POLLING ======
export async function processInboxOnce() {
  const folderId = await resolveFolderIdByName();

  const res = await graphGet<{ value: Msg[] }>(
    `/users/${cfg.mailboxUserId}/mailFolders/${folderId}/messages?$top=25`
  );

  if (!res.value?.length) return;

  const site = await graphGet<any>(`/sites/${cfg.siteId}`);
  const webUrl: string = site?.webUrl;
  if (!webUrl) throw new Error("No pude obtener webUrl del sitio");

  for (const m of res.value) {
    // anti-duplicados
    if (processedIds.has(m.id)) continue;

    const subjectLower = (m.subject || "").toLowerCase();
    if (!subjectLower.includes("list")) continue;

    const fromEmail =
      m.from?.emailAddress?.address?.trim().toLowerCase() || "";

    const solicitanteMail =
      fromEmail && fromEmail !== cfg.adminEmail ? fromEmail : "";

    const responsablesMails = Array.from(
      new Set(
        (m.ccRecipients ?? [])
          .map(r => r.emailAddress?.address)
          .filter((mail): mail is string => typeof mail === "string")
          .map(mail => mail.trim().toLowerCase())
          .filter(mail => mail.endsWith(cfg.berciaDomain))
          .filter(mail => mail !== cfg.adminEmail)
          .filter(mail => mail !== solicitanteMail)
      )
    );

    const bodyHtml = m.body?.content || "";
    const bodyText = htmlToText(bodyHtml);
    const parsed = parseMail(bodyText);

    const solicitanteId = solicitanteMail
      ? await getSiteUserLookupId(solicitanteMail, webUrl)
      : null;

    const responsablesIds: number[] = [];
    for (const mail of responsablesMails) {
      try {
        const id = await getSiteUserLookupId(mail, webUrl);
        if (id && !responsablesIds.includes(id)) responsablesIds.push(id);
      } catch {
        console.warn("⚠️ Responsable no resolvió:", mail);
      }
    }

    // 🟢 Fecha solicitada: la que viene en el correo o ahora
    const fechaSolicitadaValue = parsed.fechaSolicitada
      ? normalizeDate(parsed.fechaSolicitada)
      : nowForSharePoint();

    // 🟢 Fecha confirmada: la que pone el solicitante en el correo (si viene)
    const fechaConfirmadaValue = parsed.fechaConfirmada
      ? normalizeDate(parsed.fechaConfirmada)
      : undefined;

    const fields: any = {
      Title: m.subject || "Solicitud",
      Cliente_x002f_Proyecto: parsed.clienteProyecto || "Sin cliente",
      Observaciones: parsed.observaciones || "Sin observaciones",
      Estadoderevisi_x00f3_n: "Pendiente",
      Notificado: true,
      Solicitante: solicitanteMail || "",
      Responsable: responsablesMails.join("; "),
      Fechasolicitada: fechaSolicitadaValue,
    };

    // solo setear FechaConfirmada si viene en el correo
    if (fechaConfirmadaValue) {
      fields.FechaConfirmada = fechaConfirmadaValue;
    }

    fields.Tipodetarea =
      normalizeTipodetarea(parsed.tipodetarea) ?? "envio";

    const prioOk = normalizePrioridad(parsed.prioridad);
    if (prioOk) fields.Prioridad = prioOk;

    if (solicitanteId) fields.Solicitante0LookupId = solicitanteId;

    // 1) crear item
    const created = await createListItem(fields);
    const itemId = Number(created?.id);
    const itemUrl: string | undefined = (created as any)?.webUrl;

    // 2) setear responsables persona
    if (Number.isFinite(itemId) && responsablesIds.length) {
      await spSetResponsables(webUrl, cfg.listId, itemId, responsablesIds);
    }

    // 3) mail al solicitante
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

    // 4) marcar este mensaje como procesado
    markProcessed(m.id);

    console.log("✅ Item creado + correo enviado:", m.subject);
  }
}

async function resolveFolderIdByName() {
  const folders = await graphGet<{ value: any[] }>(
    `/users/${cfg.mailboxUserId}/mailFolders?$top=250`
  );

  const match = folders.value.find(
    f =>
      (f.displayName || "").toLowerCase() ===
      cfg.targetFolderPath.toLowerCase()
  );

  if (!match?.id) {
    throw new Error(
      `No encontré carpeta "${cfg.targetFolderPath}". Revisa TARGET_FOLDER_PATH.`
    );
  }

  return match.id;
}

// ===== Simulación desde Postman =====
export type SimulatedMailInput = {
  subject: string;
  from: string;
  cc?: string[];
  body?: string;
};

export async function processSimulatedMail(input: SimulatedMailInput) {
  const subject = (input.subject || "").trim();
  const fromEmail = (input.from || "").trim().toLowerCase();
  const ccMailsRaw = (input.cc || []).map(x =>
    (x || "").trim().toLowerCase()
  );
  const bodyRaw = input.body || "";

  if (!subject) throw new Error("subject es obligatorio");
  if (!fromEmail || !fromEmail.includes("@")) {
    throw new Error("from (email) es obligatorio y debe ser válido");
  }

  const solicitanteMail =
    fromEmail !== cfg.adminEmail ? fromEmail : "";

  const responsablesMails = Array.from(
    new Set(
      ccMailsRaw
        .filter(mail => mail.endsWith(cfg.berciaDomain))
        .filter(mail => mail !== cfg.adminEmail)
        .filter(mail => mail !== solicitanteMail)
    )
  );

  const bodyText = htmlToText(bodyRaw);
  const parsed = parseMail(bodyText);

  const site = await graphGet<any>(`/sites/${cfg.siteId}`);
  const webUrl: string = site?.webUrl;
  if (!webUrl) throw new Error("No pude obtener webUrl del sitio");

  const solicitanteId = solicitanteMail
    ? await getSiteUserLookupId(solicitanteMail, webUrl)
    : null;

  const responsablesIds: number[] = [];
  for (const mail of responsablesMails) {
    try {
      const id = await getSiteUserLookupId(mail, webUrl);
      if (id && !responsablesIds.includes(id)) responsablesIds.push(id);
    } catch {
      console.warn("⚠️ Responsable no resolvió:", mail);
    }
  }

  // 🟢 Fecha solicitada: del body o ahora
  const fechaSolicitadaValue = parsed.fechaSolicitada
    ? normalizeDate(parsed.fechaSolicitada)
    : nowForSharePoint();

  // 🟢 Fecha confirmada: también desde el body si viene
  const fechaConfirmadaValue = parsed.fechaConfirmada
    ? normalizeDate(parsed.fechaConfirmada)
    : undefined;

  const fields: any = {
    Title: subject || "Solicitud",
    Cliente_x002f_Proyecto: parsed.clienteProyecto || "Sin cliente",
    Observaciones: parsed.observaciones || "Sin observaciones",
    Estadoderevisi_x00f3_n: "Pendiente",
    Notificado: true,
    Solicitante: solicitanteMail || "",
    Responsable: responsablesMails.join("; "),
    Fechasolicitada: fechaSolicitadaValue,
  };

  if (fechaConfirmadaValue) {
    fields.FechaConfirmada = fechaConfirmadaValue;
  }

  fields.Tipodetarea =
    normalizeTipodetarea(parsed.tipodetarea) ?? "envio";

  const prioOk = normalizePrioridad(parsed.prioridad);
  if (prioOk) fields.Prioridad = prioOk;

  if (solicitanteId) fields.Solicitante0LookupId = solicitanteId;

  const created = await createListItem(fields);
  const itemId = Number(created?.id);
  const itemUrl: string | undefined = (created as any)?.webUrl;

  if (Number.isFinite(itemId) && responsablesIds.length) {
    await spSetResponsables(webUrl, cfg.listId, itemId, responsablesIds);
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

function normalizeDate(input: string) {
  const s = input.trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s;

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return s;

  const dd = m[1].padStart(2, "0");
  const mm = m[2].padStart(2, "0");
  const yyyy = m[3];
  // puedes dejarla a medianoche, o ajustarla a local si quieres hora también
  return `${yyyy}-${mm}-${dd}T00:00:00Z`;
}
