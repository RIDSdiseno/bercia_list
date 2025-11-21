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

  const r = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/users?$select=id,displayName,mail,userPrincipalName`,
    { headers: { Authorization: `Bearer ${token}` } }
  );

  console.log("id\tmail\tuserPrincipalName\tdisplayName");
  for (const u of r.data.value) {
    console.log(
      `${u.id}\t${u.mail || ""}\t${u.userPrincipalName || ""}\t${u.displayName || ""}`
    );
  }
})();
