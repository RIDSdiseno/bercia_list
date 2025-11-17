import "dotenv/config";
import axios from "axios";
import { getAppToken } from "./graph";

(async () => {
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_ID, LIST_ID } = process.env as Record<string, string>;
  const token = await getAppToken(TENANT_ID!, CLIENT_ID!, CLIENT_SECRET!);
  const r = await axios.get(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/columns?$select=id,name,displayName`,
    { headers: { Authorization: `Bearer ${token}` } }
  );
  console.log("name\t=>\tdisplayName");
  for (const c of r.data.value) console.log(`${c.name}\t=>\t${c.displayName}`);
})();
