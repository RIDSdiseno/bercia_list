// src/server.ts
import "dotenv/config";
import express from "express";
import bodyParser from "body-parser";

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
 * toCcBercia puede venir:
 *  - string "a@x.cl; b@y.cl"
 *  - array de objetos Outlook { emailAddress: { address } }
 *  - array de strings
 */
function parseCc(input: unknown): string[] {
  if (!input) return [];

  // Caso array Outlook
  if (Array.isArray(input)) {
    const arr = input
      .map((x: any) => x?.emailAddress?.address ?? x)
      .filter(Boolean);
    return extractEmails(arr.join(";"));
  }

  // Caso string
  return extractEmails(String(input));
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

    const texto = `${subject ?? ""}\n${bodyHtml || bodyPreview || ""}`;
    const prioridad = guessPrioridad(texto);
    const tipoTarea = guessTipoTarea(texto);
    const fechaSolicitada = extractFirstDateISO(bodyHtml || bodyPreview || "");
    const clienteProyecto = extractClientProject(
      subject ?? "",
      bodyHtml || bodyPreview || ""
    );

    // ====== Emails LIMPIOS ======
    const solicitanteEmail =
      extractEmails(from)[0] ??
      (typeof from === "string" ? from.trim().toLowerCase() : "");

    // ========= Responsables desde CC, excluyendo admin =========
    const ADMIN_MAIL = "administrador@bercia.cl";

    let responsablesArr = parseCc(toCcBercia);
    if (responsablesArr.length === 0) {
      const raw = String(toCcBercia ?? "").replace(/,/g, ";");
      responsablesArr = normalizeToCc(raw)
        .split(";")
        .map((s) => s.trim().toLowerCase())
        .filter(Boolean);
    }

    responsablesArr = Array.from(new Set(responsablesArr));
    responsablesArr = responsablesArr.filter((e) => e !== ADMIN_MAIL);

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
    const fields: Record<string, any> = {
      Title: subject ?? "(sin asunto)",
      Observaciones: truncate(bodyPreview || "", 1800),
      Notificado: Boolean(from),
      Cliente_x002f_Proyecto: clienteProyecto ?? "",
      // ReceivedDateTime: receivedDateTime ?? undefined,
    };

    if (fechaSolicitada) {
      const iso = new Date(fechaSolicitada).toISOString();
      if (!isNaN(Date.parse(iso))) fields.Fechasolicitada = iso;
    }

    const ESTADO_CHOICES = ["Pendiente", "En revisión", "Completado"] as const;
    fields.Estadoderevisi_x00f3_n = "Pendiente";
    if (PRIORIDAD_CHOICES.includes(prioridad as any)) fields.Prioridad = prioridad;
    if (TIPO_TAREA_CHOICES.includes(tipoTarea as any)) fields.Tipodetarea = tipoTarea;

    /* ================= BACKUP TEXTO ================= */
    if (solicitanteEmail) {
      fields["Solicitante"] = solicitanteEmail; // SolicitanteEmail (texto)
    }
    fields["Responsable"] = responsablesArr.join(";"); // ResponsablesEmail (texto)

    /* ================= PEOPLE por LookupId =================
       Si lookup falla (permiso/usuario no existe), NO rompe el flujo.
    ========================================================= */
    try {
      if (solicitanteEmail) {
        const solicitanteId = await getSiteUserLookupId(
          token,
          SITE_ID!,
          solicitanteEmail
        );
        if (solicitanteId) {
          fields["Solicitante0LookupId"] = solicitanteId;
        } else {
          console.warn("No LookupId solicitante:", solicitanteEmail);
        }
      }

      if (responsablesArr.length > 0) {
        const ids = await Promise.all(
          responsablesArr.map((mail) =>
            getSiteUserLookupId(token, SITE_ID!, mail)
          )
        );
        const responsablesIds = ids.filter(
          (x): x is number => typeof x === "number"
        );

        if (responsablesIds.length > 0) {
          fields["ResponsablesLookupId"] = responsablesIds;
        } else {
          console.warn("No LookupId responsables:", responsablesArr);
        }
      }
    } catch (err: any) {
      console.warn("Lookup People falló, se creará solo con texto:", err?.response?.data || err);
    }

    console.log("INTAKE responsables:", { toCcBercia, responsablesArr });

    await createListItem(token, {
      siteId: SITE_ID!,
      listId: LIST_ID!,
      fields,
    });

    // 5) Notificación opcional al solicitante
    if (from) {
      await sendConfirmationEmail(
        token,
        MAILBOX_USER_ID!,
        solicitanteEmail || String(from),
        subject ?? "(sin asunto)"
      );
    }

    return res.json({
      ok: true,
      startedAt,
      finishedAt: new Date().toISOString(),
    });
  } catch (e: any) {
    const status = e?.response?.status;
    const data = e?.response?.data;
    const msg = e?.message;

    console.error("intake error status:", status);
    console.error("intake error data:", data);
    console.error("intake error msg:", msg);

    return res.status(500).json({
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
