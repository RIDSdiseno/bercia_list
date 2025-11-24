import "dotenv/config";
import axios from "axios";
import { getAppToken } from "./graph";

function need(name: string, v?: string) {
  if (!v) throw new Error(`Falta env ${name}`);
  return v;
}

(async () => {
  const TENANT_ID = need("TENANT_ID", process.env.TENANT_ID);
  const CLIENT_ID = need("CLIENT_ID", process.env.CLIENT_ID);
  const CLIENT_SECRET = need("CLIENT_SECRET", process.env.CLIENT_SECRET);
  const SITE_ID = need("SITE_ID", process.env.SITE_ID);

  const token = await getAppToken(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

  // 1) Obtener webUrl real del sitio (esto sí existe en Graph)
  const siteRes = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}?$select=webUrl`,
    {
      headers: { Authorization: `Bearer ${token}` },
    }
  );

  const webUrl: string | undefined = siteRes.data?.webUrl;
  if (!webUrl) throw new Error("No pude obtener webUrl del sitio con Graph.");

  // 2) Listar usuarios del sitio con SharePoint REST
  const usersRes = await axios.get(
    `${webUrl}/_api/web/siteusers?$select=Id,Email,UserPrincipalName,LoginName,Title`,
    {
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json;odata=nometadata",
      },
    }
  );

  const users = usersRes.data?.value ?? [];

  console.log("Id\tEmail\tUserPrincipalName\tLoginName\tTitle");
  for (const u of users) {
    console.log(
      `${u.Id}\t${u.Email || ""}\t${u.UserPrincipalName || ""}\t${u.LoginName || ""}\t${u.Title || ""}`
    );
  }
})();
