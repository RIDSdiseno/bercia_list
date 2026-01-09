// src/debugSiteId.ts
import { graphGet } from "../graph.js";

async function main() {
  try {
    const site = await graphGet<any>(
      "/sites/berciacrm-my.sharepoint.com:/personal/administrador_bercia_cl:/"
    );
    console.log("SITE:", site);
    console.log("ID:", site.id);
  } catch (e: any) {
    console.error("Error obteniendo site:", e?.message || e);
  }
}

main();
