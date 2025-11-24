// src/server.ts
import "dotenv/config";
import express from "express";
import bodyParser from "body-parser";
import axios from "axios";

import { getAppToken } from "./scripts/graph";
import {
  guessPrioridad,
  guessTipoTarea,
  extractFirstDateISO,
  extractClientProject,
  truncate,
  sendConfirmationEmail,
  PRIORIDAD_CHOICES,
  TIPO_TAREA_CHOICES,
} from "./scripts/mail";
import { createListItem } from "./scripts/sharepoint";
import { ensureFolderPath } from "./scripts/folders";
import { createSubscription } from "./scripts/subscription";
import { getSiteUserLookupId } from "./scripts/spUsers";

/* ===================== Configuración & Utils ===================== */

const {
  PORT = "4000",
  TENANT_ID,
  CLIENT_ID,
  CLIENT_SECRET,
  MAILBOX_USER_ID,
  SITE_ID,
  LIST_ID,
  WEBHOOK_URL,
  TARGET_FOLDER_PATH = "Prueba-Flujo-list",
  BERCIA_DOMAIN = "@bercia.cl",
  PA_SHARED_KEY,
} = process.env as Record<string, string | undefined>;

function requireEnv(vars: Array<[string, string | undefined]>) {
  const missing = vars.filter(([, v]) => !v).map(([k]) => k);
  if (missing.length) {
    throw new Error(
      `Faltan variables de entorno: ${missing.join(", ")}. Revisa tu .env o configuración.`
    );
  }
}
function requireGraphBase() {
  requireEnv([
    ["TENANT_ID", TENANT_ID],
    ["CLIENT_ID", CLIENT_ID],
    ["CLIENT_SECRET", CLIENT_SECRET],
  ]);
}
function requireSharePointBase() {
  requireEnv([
    ["SITE_ID", SITE_ID],
    ["LIST_ID", LIST_ID],
    ["MAILBOX_USER_ID", MAILBOX_USER_ID],
  ]);
}
function requireWebhookBase() {
  requireEnv([["WEBHOOK_URL", WEBHOOK_URL]]);
}
function requirePAKey() {
  requireEnv([["PA_SHARED_KEY", PA_SHARED_KEY]]);
}

// Normaliza string de correos separados por ;
function normalizeToCc(raw: unknown): string {
  if (typeof raw !== "string") return "";
  return raw.split(";").map((s) => s.trim()).filter(Boolean).join(";");
}

// Extrae emails desde “Nombre <correo>”, comas o ;
function extractEmails(input: unknown): string[] {
  const s = String(input ?? "");
  const matches = s.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) || [];
  return Array.from(new Set(matches.map((e) => e.toLowerCase().trim())));
}

/**
 * COMBATE HTML:
 * Convierte bodyHtml a texto plano para poder parsear "Responsables:" bien.
 */
function stripHtml(input: string) {
  return input
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<\/(div|p|br|li|tr|td|th|h[1-6])>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * toCcBercia puede venir:
 *  - string "a@x.cl; b@y.cl"
 *  - array de objetos Outlook { emailAddress: { address } }
 *  - array de strings
 */
function parseCc(input: unknown): string[] {
  if (!input) return [];

  if (Array.isArray(input)) {
    const arr = input
      .map((x: any) => x?.emailAddress?.address ?? x)
      .filter(Boolean);
    return extractEmails(arr.join(";"));
  }

  return extractEmails(String(input));
}

/**
 * Lee responsables desde el body (HTML o preview):
 * Ej: "Responsables: a@x.com; b@y.com"
 * Devuelve SOLO correos reales.
 */
function parseResponsablesFromBody(bodyHtmlOrPreview: string): string[] {
  const plain = stripHtml(bodyHtmlOrPreview);
  const m = plain.match(/Responsables\s*:\s*(.+)/i);
  if (!m?.[1]) return [];
  return extractEmails(m[1]);
}

/**
 * Limpia fields para evitar enviar undefined/null a Graph.
 */
function cleanFields(fields: Record<string, any>) {
  const out: Record<string, any> = {};
  for (const [k, v] of Object.entries(fields)) {
    if (v === undefined || v === null) continue;
    if (Array.isArray(v) && v.length === 0) continue;
    out[k] = v;
  }
  return out;
}

/**
 * Trae columnas reales de la lista (internal names).
 */
async function getListColumns(token: string, siteId: string, listId: string) {
  const { data } = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns?$select=name,displayName,hidden,readOnly`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  return (data?.value ?? []) as Array<{
    name: string;
    displayName: string;
    hidden?: boolean;
    readOnly?: boolean;
  }>;
}

/**
 * Deja pasar solo columnas reales.
 * ✅ PERO también deja pasar variantes LookupId
 *    si existe la columna base (ej Responsables → ResponsablesLookupId).
 */
function sanitizeFieldsByColumns(
  fields: Record<string, any>,
  columns: Array<{ name: string; hidden?: boolean; readOnly?: boolean }>
) {
  const allowed = new Set(
    columns
      .filter((c) => !c.hidden && !c.readOnly)
      .map((c) => c.name)
  );

  const out: Record<string, any> = {};
  for (const [k, v] of Object.entries(fields)) {
    const isAllowedExact = allowed.has(k);

    const base = k.endsWith("LookupId") ? k.slice(0, -8) : null;
    const isAllowedLookup = base ? allowed.has(base) : false;

    if (!isAllowedExact && !isAllowedLookup) continue;
    if (v === undefined || v === null) continue;
    if (Array.isArray(v) && v.length === 0) continue;

    out[k] = v;
  }
  return out;
}

/* ===================== App ===================== */

const app = express();
app.use(bodyParser.json({ type: "*/*", limit: "2mb" }));

/* ========= Health ========= */

app.get("/", (_req, res) => res.send("OK"));
app.get("/health", (_req, res) => res.status(200).send("ok"));
app.get("/api/graph/health", (_req, res) =>
  res.json({ ok: true, mailbox: MAILBOX_USER_ID, site: SITE_ID, list: LIST_ID })
);

/* ========= Intake (Power Automate / Postman → SharePoint) ========= */

app.post("/api/intake/email", async (req, res) => {
  const startedAt = new Date().toISOString();

  try {
    // 1) Seguridad header
    requirePAKey();
    const headerKey = String(req.headers["x-pa-key"] || "").trim();
    const sharedKey = String(PA_SHARED_KEY || "").trim();
    if (!headerKey || headerKey !== sharedKey) {
      return res.status(401).json({ error: "unauthorized" });
    }

    // 2) Envs necesarias
    requireGraphBase();
    requireSharePointBase();

    // 3) Cuerpo
    const {
      subject,
      from,
      toCcBercia,
      bodyPreview,
      bodyHtml,
      receivedDateTime,
    } = req.body ?? {};

    const token = await getAppToken(TENANT_ID!, CLIENT_ID!, CLIENT_SECRET!);

    const bodyHtmlText = String(bodyHtml || "");
    const bodyPreviewText = String(bodyPreview || "");
    const bodyPlain = stripHtml(bodyHtmlText || bodyPreviewText);

    const texto = `${subject ?? ""}\n${bodyPlain}`;

    const prioridad = guessPrioridad(texto);
    const tipoTarea = guessTipoTarea(texto);
    const fechaSolicitada = extractFirstDateISO(bodyPlain);
    const clienteProyecto = extractClientProject(subject ?? "", bodyPlain);

    // ====== Solicitante email limpio ======
    const solicitanteEmail =
      extractEmails(from)[0] ??
      (typeof from === "string" ? from.trim().toLowerCase() : "");

    // ========= Responsables =========
    const ADMIN_MAIL = "administrador@bercia.cl";

    // 1) Primero intenta desde body (limpio)
    let responsablesArr = parseResponsablesFromBody(bodyHtmlText || bodyPreviewText);

    // 2) Si no venían en body, usa To+CC del flujo (toCcBercia)
    if (responsablesArr.length === 0) {
      responsablesArr = parseCc(toCcBercia);

      if (responsablesArr.length === 0) {
        const raw = String(toCcBercia ?? "").replace(/,/g, ";");
        responsablesArr = normalizeToCc(raw)
          .split(";")
          .map((s) => s.trim().toLowerCase())
          .filter(Boolean);
      }
    }

    // Limpieza + únicos
    responsablesArr = Array.from(new Set(responsablesArr));

    // Excluir admin si hay más de uno
    if (responsablesArr.length > 1) {
      responsablesArr = responsablesArr.filter((e) => e !== ADMIN_MAIL);
    }

    // Fallback final
    if (responsablesArr.length === 0) {
      responsablesArr = [ADMIN_MAIL];
    }

    // Warn dominio solo si viene from
    if (
      BERCIA_DOMAIN &&
      typeof from === "string" &&
      from.trim().length > 0 &&
      BERCIA_DOMAIN.length > 2
    ) {
      if (!from.toLowerCase().includes(BERCIA_DOMAIN.toLowerCase())) {
        console.warn(
          `[WARN] Remitente distinto de dominio esperado (${BERCIA_DOMAIN}):`,
          from
        );
      }
    }

    // 4) Construir payload base
    const fieldsRaw: Record<string, any> = {
      Title: subject ?? "(sin asunto)",
      Observaciones: truncate(bodyPlain || bodyPreviewText || "", 1800),
      Notificado: Boolean(solicitanteEmail),

      Cliente_x002f_Proyecto: clienteProyecto ?? "",
      Estadoderevisi_x00f3_n: "Pendiente",

      Prioridad: PRIORIDAD_CHOICES.includes(prioridad as any)
        ? prioridad
        : undefined,
      Tipodetarea: TIPO_TAREA_CHOICES.includes(tipoTarea as any)
        ? tipoTarea
        : undefined,

      Solicitante: solicitanteEmail || undefined,
      Responsable: responsablesArr.join(";"),

      // ReceivedDateTime: receivedDateTime ?? undefined,
    };

    if (fechaSolicitada) {
      const iso = new Date(fechaSolicitada).toISOString();
      if (!isNaN(Date.parse(iso))) fieldsRaw.Fechasolicitada = iso;
    }

    /* ================= PEOPLE por LookupId ================= */
    try {
      if (solicitanteEmail) {
        const solicitanteId = await getSiteUserLookupId(
          token,
          SITE_ID!,
          solicitanteEmail
        );
        if (solicitanteId) {
          fieldsRaw["Solicitante0LookupId"] = solicitanteId;
        }
      }

      if (responsablesArr.length > 0) {
        const ids = await Promise.all(
          responsablesArr.map((mail) =>
            getSiteUserLookupId(token, SITE_ID!, mail)
          )
        );

        const responsablesIds = ids.filter(
          (x): x is number => Number.isFinite(x as number)
        );

        if (responsablesIds.length > 0) {
          fieldsRaw["ResponsablesLookupId"] = responsablesIds;
        }
      }
    } catch (err: any) {
      console.warn(
        "Lookup People falló, se creará solo con texto:",
        err?.response?.data || err
      );
    }

    const fieldsClean = cleanFields(fieldsRaw);

    console.log("INTAKE responsables:", {
      toCcBercia,
      responsablesArr,
      desdeBody: parseResponsablesFromBody(bodyHtmlText || bodyPreviewText),
    });

    // 5) Traer columnas reales + sanitizar
    const columns = await getListColumns(token, SITE_ID!, LIST_ID!);

    console.log(
      "COLUMNAS REALES:",
      columns.map((c) => `${c.name} (${c.displayName})`)
    );

    const fieldsSanitized = sanitizeFieldsByColumns(fieldsClean, columns);

    console.log("FIELDS SANITIZADOS:", JSON.stringify(fieldsSanitized, null, 2));

    // 6) Crear item en SharePoint (Graph)
    await createListItem(token, {
      siteId: SITE_ID!,
      listId: LIST_ID!,
      fields: fieldsSanitized,
    });

    // 7) Notificación opcional al solicitante
    if (solicitanteEmail) {
      await sendConfirmationEmail(
        token,
        MAILBOX_USER_ID!,
        solicitanteEmail,
        subject ?? "(sin asunto)"
      );
    }

    return res.json({
      ok: true,
      startedAt,
      finishedAt: new Date().toISOString(),
    });
  } catch (e: any) {
    const status = e?.response?.status || 500;
    const data = e?.response?.data;
    const msg = e?.message;

    console.error("intake error status:", status);
    console.error("intake error data:", data);
    console.error("intake error msg:", msg);

    return res.status(status).json({
      error: data || msg || String(e),
      status,
    });
  }
});

/* ========= Webhook Graph (opcional) ========= */

app.get("/api/graph/webhook", (req, res) => {
  const token = req.query.validationToken as string | undefined;
  if (token) return res.status(200).type("text/plain").send(token);
  return res.sendStatus(200);
});

app.post("/api/graph/webhook", async (req, res) => {
  const validationToken = (req.query as any)?.validationToken as
    | string
    | undefined;
  if (validationToken) {
    return res.status(200).type("text/plain").send(validationToken);
  }
  res.sendStatus(202);
});

/* ========= Crear suscripciones (Inbox + carpeta objetivo) ========= */

app.post("/api/graph/subscribe", async (_req, res) => {
  try {
    requirePAKey();
    requireGraphBase();
    requireSharePointBase();
    requireWebhookBase();

    const token = await getAppToken(TENANT_ID!, CLIENT_ID!, CLIENT_SECRET!);
    const folderId = await ensureFolderPath(
      token,
      MAILBOX_USER_ID!,
      TARGET_FOLDER_PATH!
    );

    const subs: any[] = [];
    subs.push(
      await createSubscription(
        token,
        `/users/${MAILBOX_USER_ID}/mailFolders('inbox')/messages`,
        WEBHOOK_URL!
      )
    );
    subs.push(
      await createSubscription(
        token,
        `/users/${MAILBOX_USER_ID}/mailFolders/${folderId}/messages`,
        WEBHOOK_URL!
      )
    );

    return res.json({ ok: true, folderId, subs });
  } catch (e: any) {
    const payload = e?.response?.data ?? e?.message ?? String(e);
    console.error("subscribe error:", payload);
    return res.status(500).json({ error: payload });
  }
});

/* ===================== Start ===================== */

app.listen(Number(PORT), "0.0.0.0", () => {
  console.log(`Server running on http://localhost:${PORT}`);
});

export default app;
