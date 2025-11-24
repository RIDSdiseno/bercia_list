// src/scripts/graph.ts
import axios, { AxiosRequestConfig } from "axios";

/**
 * Token app-only para Microsoft Graph
 */
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

/**
 * Token app-only para SharePoint REST
 * (audience: https://<tenant>.sharepoint.com)
 */
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

/* =========================================================
   Helpers que tus otros scripts usan: gget / gpost
   ========================================================= */

/**
 * gget: wrapper GET para Graph.
 * - Si pasas una URL relativa, se pega a https://graph.microsoft.com/v1.0
 * - Si pasas absoluta (https://...), la usa tal cual.
 */
export async function gget<T = any>(
  token: string,
  url: string,
  config: AxiosRequestConfig = {}
) {
  const finalUrl = url.startsWith("http")
    ? url
    : `https://graph.microsoft.com/v1.0${url.startsWith("/") ? "" : "/"}${url}`;

  const { data } = await axios.get<T>(finalUrl, {
    ...config,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(config.headers || {}),
    },
  });

  return data;
}

/**
 * gpost: wrapper POST para Graph.
 */
export async function gpost<T = any>(
  token: string,
  url: string,
  body?: any,
  config: AxiosRequestConfig = {}
) {
  const finalUrl = url.startsWith("http")
    ? url
    : `https://graph.microsoft.com/v1.0${url.startsWith("/") ? "" : "/"}${url}`;

  const { data } = await axios.post<T>(finalUrl, body, {
    ...config,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(config.headers || {}),
    },
  });

  return data;
}
