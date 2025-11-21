// src/scripts/spUsers.ts
import axios from "axios";

type SiteUser = {
  id: string;
  email?: string;
  loginName?: string;
  userPrincipalName?: string;
};

/**
 * Busca LookupId de un usuario dentro del sitio SharePoint.
 * Requiere Sites.ReadWrite.All + Directory.Read.All (Application).
 */
export async function getSiteUserLookupId(
  token: string,
  siteId: string,
  emailOrUpn: string
): Promise<number | null> {
  const target = (emailOrUpn || "").toLowerCase().trim();
  if (!target) return null;

  // 1) Listar usuarios del sitio
  const { data } = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/users?$select=id,email,loginName,userPrincipalName`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  const users: SiteUser[] = data?.value || [];

  const match = users.find((u) => {
    const email = (u.email || "").toLowerCase();
    const upn = (u.userPrincipalName || "").toLowerCase();
    const login = (u.loginName || "").toLowerCase();

    return (
      email === target ||
      upn === target ||
      login.includes(target) ||              // a veces viene como i:0#.f|membership|user@dominio
      login.endsWith(target)
    );
  });

  if (!match) return null;

  // Graph entrega id tipo "123" como string -> LookupId numérico
  const n = Number(match.id);
  return Number.isFinite(n) ? n : null;
}
