// src/scripts/spUsers.ts
import axios from "axios";

/**
 * Obtiene el webUrl real del sitio desde Graph.
 */
async function getSiteWebUrl(token: string, siteId: string): Promise<string> {
  const { data } = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}?$select=webUrl`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  const webUrl: string | undefined = data?.webUrl;
  if (!webUrl) throw new Error("No pude obtener webUrl del sitio con Graph.");
  return webUrl;
}

/**
 * Asegura/resuelve un usuario en el sitio y devuelve su LookupId.
 * - Si el usuario ya existe en el sitio, devuelve su Id igualmente.
 * - Si no existe o no es válido, lanza error (lo capturas arriba).
 */
async function ensureUserId(
  webUrl: string,
  token: string,
  emailOrUpn: string
): Promise<number> {
  const target = (emailOrUpn || "").trim();
  if (!target) throw new Error("emailOrUpn vacío");

  const { data } = await axios.post(
    `${webUrl}/_api/web/ensureuser`,
    { logonName: target },
    {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
      },
    }
  );

  const n = Number(data?.Id);
  if (!Number.isFinite(n)) throw new Error("ensureuser no devolvió Id válido");
  return n;
}

/**
 * Busca LookupId de un usuario dentro del sitio SharePoint.
 * ✅ Versión correcta (sin /sites/{id}/users).
 *
 * Requiere (Application):
 * - Sites.ReadWrite.All  (para ensureuser)
 * - Directory.Read.All o User.Read.All (según tenant)
 */
export async function getSiteUserLookupId(
  token: string,
  siteId: string,
  emailOrUpn: string
): Promise<number | null> {
  const target = (emailOrUpn || "").toLowerCase().trim();
  if (!target) return null;

  try {
    const webUrl = await getSiteWebUrl(token, siteId);
    const id = await ensureUserId(webUrl, token, target);
    return id;
  } catch (e) {
    // no se pudo resolver/asegurar usuario en el sitio
    return null;
  }
}
