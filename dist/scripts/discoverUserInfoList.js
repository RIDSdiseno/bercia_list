"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const axios_1 = __importDefault(require("axios"));
const config_1 = require("../config"); // 👈 OJO: ../ porque estás dentro de scripts
const graph_1 = require("../graph");
const auth_1 = require("../auth");
async function main() {
    // 1) Intento v1.0 con includeHiddenLists
    try {
        const v1 = await (0, graph_1.graphGet)(`/sites/${config_1.cfg.siteId}/lists?includeHiddenLists=true&$select=id,displayName`);
        console.log("=== LISTAS (v1.0 includeHiddenLists) ===");
        v1.value.forEach(l => console.log(l.displayName, "->", l.id));
        return;
    }
    catch (e) {
        console.log("v1.0 no devolvió ocultas, probando beta...");
    }
    // 2) Fallback beta
    const token = await (0, auth_1.getGraphToken)();
    const { data } = await axios_1.default.get(`https://graph.microsoft.com/beta/sites/${config_1.cfg.siteId}/lists?includeHiddenLists=true&$select=id,displayName`, { headers: { Authorization: `Bearer ${token}` } });
    console.log("=== LISTAS (beta includeHiddenLists) ===");
    data.value.forEach((l) => console.log(l.displayName, "->", l.id));
}
main().catch(err => {
    console.error(err?.response?.data || err);
    process.exit(1);
});
