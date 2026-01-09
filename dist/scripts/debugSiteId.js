// src/debugSiteId.ts
import { graphGet } from "../graph";
async function main() {
    try {
        const site = await graphGet("/sites/berciacrm-my.sharepoint.com:/personal/administrador_bercia_cl:/");
        console.log("SITE:", site);
        console.log("ID:", site.id);
    }
    catch (e) {
        console.error("Error obteniendo site:", e?.message || e);
    }
}
main();
