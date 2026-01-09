"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const graph_1 = require("../graph");
const config_1 = require("../config");
(async () => {
    const res = await (0, graph_1.graphGet)(`/sites/${config_1.cfg.siteId}/lists/${config_1.cfg.listId}/contentTypes`);
    console.log("=== CONTENT TYPES ===");
    for (const ct of res.value) {
        console.log(`${ct.name} -> ${ct.id}`);
    }
})();
