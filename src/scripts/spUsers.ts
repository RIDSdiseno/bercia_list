// src/scripts/spUsers.ts
import axios from "axios";

/**
 * Busca/asegura un usuario en el sitio SharePoint y devuelve su LookupId numérico.
 * Usa SharePoint REST porque /sites/{id}/users en Graph no siempre existe.
 */
export async function getSiteUserLookupId(
  token: string,
  siteId: string,
  emailOrUpn: string
): Promise<number | null> {
  const target = (emailOrUpn || "").toLowerCase().trim();
  if (!target) return null;

  // 1) Obtener webUrl real del sitio
  const siteRes = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}?$select=webUrl`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  const webUrl: string | undefined = siteRes.data?.webUrl;
  if (!webUrl) throw new Error("No pude obtener webUrl del sitio.");

  const spHeaders = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json;odata=nometadata",
    "Content-Type": "application/json;odata=nometadata",
  };

  // helper: listar usuarios del sitio
  async function listUsers() {
    const usersRes = await axios.get(
      `${webUrl}/_api/web/siteusers?$select=Id,Email,UserPrincipalName,LoginName`,
      { headers: spHeaders }
    );
    return usersRes.data?.value ?? [];
  }

  // 2) Buscar en usuarios actuales
  let users = await listUsers();

  let match = users.find((u: any) => {
    const email = (u.Email || "").toLowerCase();
    const upn = (u.UserPrincipalName || "").toLowerCase();
    const login = (u.LoginName || "").toLowerCase();

    return (
      email === target ||
      upn === target ||
      login.includes(target) ||
      login.endsWith(target)
    );
  });

  // 3) Si NO existe, asegurar usuario en el sitio
  if (!match) {
    try {
      await axios.post(
        `${webUrl}/_api/web/ensureuser`,
        { logonName: target }, // email o upn
        { headers: spHeaders }
      );
    } catch (e: any) {
      console.warn("ensureuser falló para:", target, e?.response?.data || e);
      // aunque ensureuser falle, seguimos con lo que haya
    }

    // 4) Volver a listar y buscar
    users = await listUsers();
    match = users.find((u: any) => {
      const email = (u.Email || "").toLowerCase();
      const upn = (u.UserPrincipalName || "").toLowerCase();
      const login = (u.LoginName || "").toLowerCase();

      return (
        email === target ||
        upn === target ||
        login.includes(target) ||
        login.endsWith(target)
      );
    });
  }

  if (!match) return null;

  const n = Number(match.Id);
  return Number.isFinite(n) ? n : null;
}
