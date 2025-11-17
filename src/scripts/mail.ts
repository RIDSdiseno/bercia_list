import { gget, gpost } from "./graph";

export type ParsedMail = {
  id: string; subject: string; from: string;
  to: string[]; cc: string[]; bcc: string[];
  receivedDateTime: string; bodyPreview: string; bodyHtml?: string;
};

const normalize = (s: string) =>
  (s || "").toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");

/** === Choices esperadas en SharePoint (ajústalas si tu lista difiere) === */
export const PRIORIDAD_CHOICES = ["Alta", "Media", "Baja"] as const;
export const TIPO_TAREA_CHOICES = [
  "Instalación",
  "Envío",
  "Cubicación por planos",
  "Cubicación en terreno",
  "Costeo de proyecto",
  "Evaluación postventa",
  "Producto interno",
  "Mantención",
  "Producción",
  "Otro",
] as const;

/** Heurística → choice válida (Alta/Media/Baja) */
export function guessPrioridad(text: string): (typeof PRIORIDAD_CHOICES)[number] {
  const s = normalize(text);
  if (/\burgent|urgente|alta|hoy\b/.test(s)) return "Alta";
  if (/\bmedia|normal|proxima semana|pr[oó]xima semana\b/.test(s)) return "Media";
  return "Baja";
}

/** Heurística → choice válida de tipo */
export function guessTipoTarea(text: string): (typeof TIPO_TAREA_CHOICES)[number] {
  const s = normalize(text);
  if (/\binstal/.test(s)) return "Instalación";
  if (/\benvio|env[ií]o\b/.test(s)) return "Envío";
  if (/planos/.test(s)) return "Cubicación por planos";
  if (/terreno/.test(s)) return "Cubicación en terreno";
  if (/costeo|presup/.test(s)) return "Costeo de proyecto";
  if (/postventa/.test(s)) return "Evaluación postventa";
  if (/interno/.test(s)) return "Producto interno";
  if (/mantenc|manteni/.test(s)) return "Mantención";
  if (/producci[oó]n|fabricaci[oó]n/.test(s)) return "Producción";
  return "Otro";
}

/** Extrae la primera fecha dd/mm/yyyy o dd-mm-yyyy y devuelve ISO (UTC 00:00) */
export function extractFirstDateISO(text: string): string | undefined {
  const m = (text || "").match(/\b(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\b/);
  if (!m) return;
  const d = Number(m[1]), mo = Number(m[2]) - 1, y = Number(m[3]);
  const dt = new Date(Date.UTC(y, mo, d, 0, 0, 0));
  return isNaN(dt.getTime()) ? undefined : dt.toISOString();
}

/** Intenta "Cliente: X" y "Proyecto: Y" y retorna "X / Y" o subject limpio */
export function extractClientProject(subject: string, body: string): string {
  const full = `${subject ?? ""}\n${body ?? ""}`;
  const re = /Cliente:\s*([^\n\r<]+).*?Proyecto:\s*([^\n\r<]+)/is;
  const m = full.match(re); // RegExpMatchArray | null
  if (m && m[1] && m[2]) return `${m[1].trim()} / ${m[2].trim()}`;
  return (subject ?? "").trim();
}


export function truncate(s: string, max = 1800) {
  const t = s ?? "";
  return t.length > max ? t.slice(0, max - 3) + "..." : t;
}

/* ====== Helpers opcionales que ya tenías ====== */

export async function fetchMailByIdForUser(token: string, userId: string, messageId: string) {
  const select =
    "$select=subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,parentFolderId";
  const { data: m } = await gget(
    `https://graph.microsoft.com/v1.0/users/${userId}/messages/${messageId}?${select}`, token
  );
  const list = (arr: any[] = []) => arr.map((x) => x?.emailAddress?.address).filter(Boolean);
  const mail: ParsedMail = {
    id: m.id, subject: m.subject ?? "", from: m.from?.emailAddress?.address ?? "",
    to: list(m.toRecipients), cc: list(m.ccRecipients), bcc: list(m.bccRecipients),
    receivedDateTime: m.receivedDateTime, bodyPreview: m.bodyPreview ?? "",
    bodyHtml: m.body?.contentType === "html" ? m.body?.content : undefined,
  };
  return mail;
}

export function hasAnyRecipientInDomain(mail: ParsedMail, domain: string) {
  const d = (domain || "").toLowerCase().trim();
  const all = [...mail.to, ...mail.cc, ...mail.bcc].map((s) => (s || "").toLowerCase().trim());
  return all.some((a) => a.endsWith(d));
}
export function collectRecipientsInDomain(mail: ParsedMail, domain: string, exclude: string[] = []) {
  const d = (domain || "").toLowerCase().trim();
  const ex = new Set(exclude.map(s => s.toLowerCase()));
  const set = new Set(
    [...mail.to, ...mail.cc, ...mail.bcc]
      .map((s) => (s || "").toLowerCase().trim())
      .filter((a) => a.endsWith(d) && !ex.has(a))
  );
  return Array.from(set);
}

export async function sendConfirmationEmail(
  token: string, fromUserId: string, toAddress: string, asuntoOriginal: string
) {
  const body = {
    message: {
      subject: `Recepción de solicitud: ${asuntoOriginal || "(sin asunto)"}`,
      body: {
        contentType: "HTML",
        content: `<p>Hola,</p>
                  <p>Hemos recibido tu solicitud y fue registrada con estado <b>Pendiente</b>.</p>
                  <p><b>Asunto:</b> ${asuntoOriginal || "(sin asunto)"}</p>
                  <p>Te avisaremos cuando cambie el estado.</p>
                  <p>Saludos,<br/>Equipo Bercia</p>`,
      },
      toRecipients: [{ emailAddress: { address: toAddress } }],
    },
    saveToSentItems: true,
  };
  await gpost(`https://graph.microsoft.com/v1.0/users/${fromUserId}/sendMail`, token, body);
}
