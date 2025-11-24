// src/scripts/graph.ts
import axios from "axios";

export async function getAppToken(
  tenantId: string,
  clientId: string,
  clientSecret: string
) {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams();
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("grant_type", "client_credentials");
  body.set("scope", "https://graph.microsoft.com/.default");

  const { data } = await axios.post(url, body.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  return data.access_token as string;
}

// ✅ NUEVO: token para SharePoint REST
export async function getSharePointToken(
  tenantId: string,
  clientId: string,
  clientSecret: string,
  sharepointHost: string // ej: "berciacrm.sharepoint.com"
) {
  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;

  const body = new URLSearchParams();
  body.set("client_id", clientId);
  body.set("client_secret", clientSecret);
  body.set("grant_type", "client_credentials");
  body.set("scope", `https://${sharepointHost}/.default`);

  const { data } = await axios.post(url, body.toString(), {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });

  return data.access_token as string;
}
