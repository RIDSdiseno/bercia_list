"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
require("dotenv/config");
const axios_1 = __importDefault(require("axios"));
const graph_1 = require("./graph");
(async () => {
    const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, SITE_ID, LIST_ID } = process.env;
    const token = await (0, graph_1.getAppToken)(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const r = await axios_1.default.get(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/columns?$select=id,name,displayName`, { headers: { Authorization: `Bearer ${token}` } });
    console.log("name\t=>\tdisplayName");
    for (const c of r.data.value)
        console.log(`${c.name}\t=>\t${c.displayName}`);
})();
//# sourceMappingURL=list-columns.js.map