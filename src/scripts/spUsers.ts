// src/scripts/spUsers.ts
import axios from "axios";

/**
 * Busca/asegura un usuario en el sitio SharePoint y devuelve su LookupId numérico.
 *
 * graphToken: token aud=graph.microsoft.com (solo para obtener webUrl del sitio)
 * spToken:    token aud=<tenant>.sharepoint.com (para siteusers + ensureuser)
 */
export async function getSiteUserLookupId(
  graphToken: string,
  spToken: string,
  siteId: string,
  emailOrUpn: string
): Promise<number | null> {
  const target = (emailOrUpn || "").toLowerCase().trim();
  if (!target) return null;

  // 1) Obtener webUrl real del sitio usando Graph
  const siteRes = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}?$select=webUrl`,
    { headers: { Authorization: `Bearer ${graphToken}` } }
  );

  const webUrl: string | undefined = siteRes.data?.webUrl;
  if (!webUrl) throw new Error("No pude obtener webUrl del sitio.");

  // Headers para SharePoint REST (audience SP)
  const spHeaders = {
    Authorization: `Bearer ${spToken}`,
    Accept: "application/json;odata=nometadata",
    "Content-Type": "application/json;odata=nometadata",
  };

  async function listUsers() {
    const usersRes = await axios.get(
      `${webUrl}/_api/web/siteusers?$select=Id,Email,UserPrincipalName,LoginName`,
      { headers: spHeaders }
    );
    return usersRes.data?.value ?? [];
  }

  function findMatch(users: any[]) {
    return users.find((u: any) => {
      const email = (u.Email || "").toLowerCase();
      const upn = (u.UserPrincipalName || "").toLowerCase();
      const login = (u.LoginName || "").toLowerCase();

      return (
        email === target ||
        upn === target ||
        login.includes(target) || // i:0#.f|membership|user@dominio
        login.endsWith(target)
      );
    });
  }

  // 2) Buscar en usuarios actuales
  let users = await listUsers();
  let match = findMatch(users);

  // 3) Si NO existe, asegurar usuario en el sitio
  if (!match) {
    try {
      await axios.post(
        `${webUrl}/_api/web/ensureuser`,
        { logonName: target },
        { headers: spHeaders }
      );
    } catch (e: any) {
      console.warn("ensureuser falló para:", target, e?.response?.data || e);
    }

    // 4) Reintentar búsqueda
    users = await listUsers();
    match = findMatch(users);
  }

  if (!match) return null;

  const n = Number(match.Id);
  return Number.isFinite(n) ? n : null;
}
