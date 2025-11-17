"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
require("dotenv/config");
const graph_1 = require("./graph");
(async () => {
    try {
        const token = await (0, graph_1.getAppToken)(process.env.TENANT_ID, process.env.CLIENT_ID, process.env.CLIENT_SECRET);
        process.stdout.write(String(token));
    }
    catch (err) {
        console.error("Error obteniendo token:", err?.response?.data || err);
        process.exit(1);
    }
})();
//# sourceMappingURL=print-token.js.map