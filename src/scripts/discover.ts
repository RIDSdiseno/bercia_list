import path from "path";
import dotenv from "dotenv";
dotenv.config({ path: path.resolve(__dirname, "../../.env") });

import { getAppToken, gget } from "./graph";

const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env as Record<string, string>;
const host = "berciacrm.sharepoint.com";
const sitePath = "/sites/AlfombrasBerciaS.A";
const LIST_NAME_TARGETS = [
  "Solicitudes de Envío e Instalación",
  "Solicitudes de Envio e Instalacion",
  "Solicitudes de Envo e Instalacin"
];

function norm(s: string) {
  return (s || "").toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");
}

async function main() {
  const token = await getAppToken(TENANT_ID!, CLIENT_ID!, CLIENT_SECRET!);
  const { data: site } = await gget(`https://graph.microsoft.com/v1.0/sites/${host}:${sitePath}`, token);
  const SITE_ID = site.id as string;
  console.log("✅ SITE_ID:", SITE_ID, "webUrl:", site.webUrl, "\n");

  const { data: lists } = await gget(
    `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists?$select=id,displayName,webUrl`,
    token
  );

  const targetsNorm = LIST_NAME_TARGETS.map(norm);
  let match = (lists.value as any[]).find((l) => targetsNorm.includes(norm(l.displayName)));
  if (!match) match = (lists.value as any[]).find((l) => targetsNorm.some((t) => norm(l.displayName).includes(t)));

  if (!match) {
    console.log("No se encontró la lista. Disponibles:");
    for (const l of lists.value) console.log(`- ${l.displayName} :: ${l.id} :: ${l.webUrl}`);
    process.exit(2);
  }
  console.log("✅ LIST_ID:", match.id, "displayName:", match.displayName, "webUrl:", match.webUrl);
}

main().catch((e) => {
  console.error(e?.response?.data || e);
  process.exit(1);
});
