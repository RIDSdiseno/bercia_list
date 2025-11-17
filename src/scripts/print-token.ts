import "dotenv/config";
import { getAppToken } from "./graph";

(async () => {
  try {
    const token = await getAppToken(process.env.TENANT_ID!, process.env.CLIENT_ID!, process.env.CLIENT_SECRET!);
    process.stdout.write(String(token));
  } catch (err: any) {
    console.error("Error obteniendo token:", err?.response?.data || err);
    process.exit(1);
  }
})();
