"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
// src/debugSiteId.ts
const graph_1 = require("../graph");
async function main() {
    try {
        const site = await (0, graph_1.graphGet)("/sites/berciacrm-my.sharepoint.com:/personal/administrador_bercia_cl:/");
        console.log("SITE:", site);
        console.log("ID:", site.id);
    }
    catch (e) {
        console.error("Error obteniendo site:", e?.message || e);
    }
}
main();
