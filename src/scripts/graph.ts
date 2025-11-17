import axios from "axios";
import * as qs from "querystring";

const tokenUrl = (tenant: string) =>
  `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;

export async function getAppToken(tenantId: string, clientId: string, clientSecret: string) {
  const data = qs.stringify({
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
    grant_type: "client_credentials",
  });
  const res = await axios.post(tokenUrl(tenantId), data, {
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
  });
  return res.data.access_token as string;
}

export async function gget(url: string, token: string) {
  return axios.get(url, { headers: { Authorization: `Bearer ${token}` } });
}
export async function gpost(url: string, token: string, body: any) {
  return axios.post(url, body, { headers: { Authorization: `Bearer ${token}` } });
}
