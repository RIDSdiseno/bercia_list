import axios from "axios";

export async function getSiteUserLookupId(
  token: string,
  siteId: string,
  email: string
): Promise<number | null> {
  const safe = email.toLowerCase().trim();

  const r = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/users?$filter=mail eq '${safe}' or userPrincipalName eq '${safe}'`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  const u = r.data?.value?.[0];
  if (!u?.id) return null;

  const n = Number(u.id);
  return Number.isFinite(n) ? n : null;
}
