import "dotenv/config";
import axios from "axios";
import { getAppToken } from "./graph";

function need(name: string, v?: string) {
  if (!v) throw new Error(`Falta env ${name}`);
  return v;
}

async function getWebUrl(siteId: string, token: string) {
  const siteRes = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${siteId}?$select=webUrl`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  const webUrl: string | undefined = siteRes.data?.webUrl;
  if (!webUrl) throw new Error("No pude obtener webUrl del sitio con Graph.");
  return webUrl;
}

async function listSiteUsers(webUrl: string, token: string) {
  const usersRes = await axios.get(
    `${webUrl}/_api/web/siteusers?$select=Id,Email,UserPrincipalName,LoginName,Title`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
      },
    }
  );
  return usersRes.data?.value ?? [];
}

async function ensureUserId(webUrl: string, token: string, email: string) {
  const r = await axios.post(
    `${webUrl}/_api/web/ensureuser`,
    { logonName: email },
    {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
        "Content-Type": "application/json;odata=nometadata",
      },
    }
  );
  return r.data.Id as number;
}

function parseResponsablesFromEnvOrArgs() {
  const envList = process.env.RESPONSABLES || "";
  const envEmails = envList
    .split(/[;,]/)
    .map(x => x.trim().toLowerCase())
    .filter(Boolean);

  const argEmails = process.argv
    .slice(2)
    .map(x => x.trim().toLowerCase())
    .filter(Boolean);

  return argEmails.length ? argEmails : envEmails;
}

(async () => {
  const TENANT_ID = need("TENANT_ID", process.env.TENANT_ID);
  const CLIENT_ID = need("CLIENT_ID", process.env.CLIENT_ID);
  const CLIENT_SECRET = need("CLIENT_SECRET", process.env.CLIENT_SECRET);
  const SITE_ID = need("SITE_ID", process.env.SITE_ID);

  const token = await getAppToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

  // 1) webUrl real del sitio
  const webUrl = await getWebUrl(SITE_ID, token);

  // 2) Listar usuarios actuales del sitio
  const users = await listSiteUsers(webUrl, token);

  console.log("\n=== USUARIOS DEL SITIO ===");
  console.log("Id\tEmail\tUserPrincipalName\tLoginName\tTitle");
  for (const u of users) {
    console.log(
      `${u.Id}\t${u.Email || ""}\t${u.UserPrincipalName || ""}\t${u.LoginName || ""}\t${u.Title || ""}`
    );
  }

  // 3) Si me pasas responsables, los resuelvo a LookupId
  const responsablesEmails = parseResponsablesFromEnvOrArgs();

  if (responsablesEmails.length) {
    console.log("\n=== RESOLUCIÓN RESPONSABLES (ensureuser) ===");
    console.log("Email\tLookupId");

    for (const email of responsablesEmails) {
      try {
        const id = await ensureUserId(webUrl, token, email);
        console.log(`${email}\t${id}`);
      } catch (err: any) {
        console.log(`${email}\tERROR (no existe o sin acceso)`);
      }
    }
  } else {
    console.log(
      "\n(No se pasaron RESPONSABLES por env ni por argumentos, solo se listaron usuarios.)"
    );
  }
})();
