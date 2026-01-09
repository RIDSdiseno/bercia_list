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

// ================== NORMALIZADOR (no tildes / no mayúsculas) ==================
function norm(input: unknown): string {
  return String(input ?? "")
    .normalize("NFD") // separa letras/tildes
    .replace(/[\u0300-\u036f]/g, "") // elimina tildes
    .toLowerCase()
    .trim()
    .replace(/\s+/g, " "); // colapsa espacios
}
// ============================================================================

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

// ✅ Set en versión NORMALIZADA (sin tildes) para comparación exacta
const TIPOS_VALIDOS_NORM = new Set([
  "envio",
  "instalacion",
  "cubicacion por planos",
  "cubicacion en terreno",
  "costeo de proyecto",
  "productos (requerimientos)",
  "evaluacion postventa",
  "post venta en terreno",
  "producto interno",
]);

// ✅ Mapa: clave normalizada -> valor “bonito” para SharePoint
const TIPOS_MAP: Record<string, string> = {
  envio: "envio",
  instalacion: "instalacion",
  "cubicacion por planos": "Cubicación por planos",
  "cubicacion en terreno": "Cubicación en terreno",
  "costeo de proyecto": "Costeo de proyecto",
  "productos (requerimientos)": "Productos (requerimientos)",
  "evaluacion postventa": "Evaluación postventa",
  "post venta en terreno": "Post venta en terreno",
  "producto interno": "Producto interno",
};

function normalizeTipodetarea(raw?: string) {
  if (!raw) return null;

  const s = norm(raw);

  // 🟦 Match por “contiene” (tolerante)
  if (s.includes("env")) return "envio";
  if (s.includes("instal")) return "instalacion";
  if (s.includes("postventa")) return "Evaluación postventa";
  if (s.includes("producto interno")) return "Producto interno";
  if (s.includes("costeo")) return "Costeo de proyecto";
  if (s.includes("plano")) return "Cubicación por planos";
  if (s.includes("terreno")) return "Cubicación en terreno";
  if (s.includes("requer")) return "Productos (requerimientos)";

  // 🟩 Match exacto (normalizado)
  if (TIPOS_VALIDOS_NORM.has(s)) {
    return TIPOS_MAP[s] ?? null;
  }

  return null;
}

function normalizePrioridad(raw?: string) {
  if (!raw) return null;

  const s = norm(raw);

  // Alta / Media / Baja, también acepta "A", "M", "B", "Alta urgencia", "MÉDIA", etc.
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
    .map((l) => l.replace(/\s+/g, " ").trim())
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
    if (processedIds.has(m.id)) continue;

    const subjectLower = norm(m.subject || "");
    if (!subjectLower.includes("list")) continue;

    const fromEmail = String(m.from?.emailAddress?.address ?? "")
      .trim()
      .toLowerCase();

    const solicitanteMail = fromEmail;

    const responsablesMails = Array.from(
      new Set(
        (m.ccRecipients ?? [])
          .map((r) => r.emailAddress?.address)
          .filter((mail): mail is string => typeof mail === "string")
          .map((mail) => mail.trim().toLowerCase())
          .filter((mail) => mail !== cfg.adminEmail)
          .filter((mail) => mail !== solicitanteMail)
      )
    );

    const bodyHtml = m.body?.content || "";
    const bodyText = htmlToText(bodyHtml);
    const parsed = parseMail(bodyText);

    const fechaSolicitadaValue = parsed.fechaSolicitada
      ? normalizeDate(parsed.fechaSolicitada)
      : nowForSharePoint();

    const fechaConfirmadaValue = parsed.fechaConfirmada
      ? normalizeDate(parsed.fechaConfirmada)
      : undefined;

    const fields: any = {
      Title: m.subject || "Solicitud",
      Cliente_x002f_Proyecto: parsed.clienteProyecto || "Sin cliente",
      Observaciones: parsed.observaciones || "Sin observaciones",
      Estadoderevisi_x00f3_n: "Pendiente",

      // ✅ texto
      Solicitante: solicitanteMail || "",
      Responsable: responsablesMails.join("; "),
      Fechasolicitada: fechaSolicitadaValue,
    };

    if (fechaConfirmadaValue) fields.FechaConfirmada = fechaConfirmadaValue;

    fields.Tipodetarea = normalizeTipodetarea(parsed.tipodetarea) ?? "envio";

    const prioOk = normalizePrioridad(parsed.prioridad);
    if (prioOk) fields.Prioridad = prioOk;

    // Solicitante persona
    try {
      if (solicitanteMail) {
        const solicitanteId = await getSiteUserLookupId(solicitanteMail, webUrl);
        if (solicitanteId) fields.Solicitante0LookupId = solicitanteId;
      }
    } catch {
      console.warn(
        "⚠️ No se pudo resolver solicitante como persona:",
        solicitanteMail
      );
    }

    const created = await createListItem(fields);
    const itemId = Number((created as any)?.id);
    const itemUrl: string | undefined = (created as any)?.webUrl;

    // set multi-persona Responsables
    if (Number.isFinite(itemId) && responsablesMails.length) {
      const responsablesIds: number[] = [];
      for (const mail of responsablesMails) {
        try {
          const rid = await getSiteUserLookupId(mail, webUrl);
          if (rid && !responsablesIds.includes(rid)) responsablesIds.push(rid);
        } catch {
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
  const folders = await graphGet<{ value: any[] }>(
    `/users/${cfg.mailboxUserId}/mailFolders?$top=250`
  );

  const match = folders.value.find(
    (f) =>
      String(f.displayName || "").trim().toLowerCase() ===
      String(cfg.targetFolderPath || "").trim().toLowerCase()
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
  const ccMailsRaw = (input.cc || []).map((x) => (x || "").trim().toLowerCase());
  const bodyRaw = input.body || "";

  if (!subject) throw new Error("subject es obligatorio");
  if (!fromEmail || !fromEmail.includes("@")) {
    throw new Error("from (email) es obligatorio y debe ser válido");
  }

  const solicitanteMail = fromEmail;

  const responsablesMails = Array.from(
    new Set(
      ccMailsRaw
        .filter((mail) => mail !== cfg.adminEmail)
        .filter((mail) => mail !== solicitanteMail)
    )
  );

  const bodyText = htmlToText(bodyRaw);
  const parsed = parseMail(bodyText);

  const site = await graphGet<any>(`/sites/${cfg.siteId}`);
  const webUrl: string = site?.webUrl;
  if (!webUrl) throw new Error("No pude obtener webUrl del sitio");

  const fechaSolicitadaValue = parsed.fechaSolicitada
    ? normalizeDate(parsed.fechaSolicitada)
    : nowForSharePoint();

  const fechaConfirmadaValue = parsed.fechaConfirmada
    ? normalizeDate(parsed.fechaConfirmada)
    : undefined;

  const fields: any = {
    Title: subject || "Solicitud",
    Cliente_x002f_Proyecto: parsed.clienteProyecto || "Sin cliente",
    Observaciones: parsed.observaciones || "Sin observaciones",
    Estadoderevisi_x00f3_n: "Pendiente",

    Solicitante: solicitanteMail || "",
    Responsable: responsablesMails.join("; "),
    Fechasolicitada: fechaSolicitadaValue,
  };

  if (fechaConfirmadaValue) fields.FechaConfirmada = fechaConfirmadaValue;

  fields.Tipodetarea = normalizeTipodetarea(parsed.tipodetarea) ?? "envio";

  const prioOk = normalizePrioridad(parsed.prioridad);
  if (prioOk) fields.Prioridad = prioOk;

  try {
    const solicitanteId = await getSiteUserLookupId(solicitanteMail, webUrl);
    if (solicitanteId) fields.Solicitante0LookupId = solicitanteId;
  } catch {
    console.warn(
      "⚠️ No se pudo resolver solicitante como persona (test):",
      solicitanteMail
    );
  }

  const created = await createListItem(fields);
  const itemId = Number((created as any)?.id);
  const itemUrl: string | undefined = (created as any)?.webUrl;

  if (Number.isFinite(itemId) && responsablesMails.length) {
    const responsablesIds: number[] = [];
    for (const mail of responsablesMails) {
      try {
        const rid = await getSiteUserLookupId(mail, webUrl);
        if (rid && !responsablesIds.includes(rid)) responsablesIds.push(rid);
      } catch {
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

function normalizeDate(input: string) {
  const s = String(input ?? "").trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s;

  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return s;

  const dd = m[1].padStart(2, "0");
  const mm = m[2].padStart(2, "0");
  const yyyy = m[3];

  // recomendado para SharePoint local (sin Z)
  return `${yyyy}-${mm}-${dd}T00:00:00`;
}
