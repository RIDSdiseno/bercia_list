"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const path_1 = __importDefault(require("path"));
const dotenv_1 = __importDefault(require("dotenv"));
dotenv_1.default.config({ path: path_1.default.resolve(__dirname, "../../.env") });
const graph_1 = require("./graph");
const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;
const host = "berciacrm.sharepoint.com";
const sitePath = "/sites/AlfombrasBerciaS.A";
const LIST_NAME_TARGETS = [
    "Solicitudes de Envío e Instalación",
    "Solicitudes de Envio e Instalacion",
    "Solicitudes de Envo e Instalacin"
];
function norm(s) {
    return (s || "").toLowerCase().normalize("NFD").replace(/\p{Diacritic}/gu, "");
}
async function main() {
    const token = await (0, graph_1.getAppToken)(TENANT_ID, CLIENT_ID, CLIENT_SECRET);
    const { data: site } = await (0, graph_1.gget)(`https://graph.microsoft.com/v1.0/sites/${host}:${sitePath}`, token);
    const SITE_ID = site.id;
    console.log("✅ SITE_ID:", SITE_ID, "webUrl:", site.webUrl, "\n");
    const { data: lists } = await (0, graph_1.gget)(`https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists?$select=id,displayName,webUrl`, token);
    const targetsNorm = LIST_NAME_TARGETS.map(norm);
    let match = lists.value.find((l) => targetsNorm.includes(norm(l.displayName)));
    if (!match)
        match = lists.value.find((l) => targetsNorm.some((t) => norm(l.displayName).includes(t)));
    if (!match) {
        console.log("No se encontró la lista. Disponibles:");
        for (const l of lists.value)
            console.log(`- ${l.displayName} :: ${l.id} :: ${l.webUrl}`);
        process.exit(2);
    }
    console.log("✅ LIST_ID:", match.id, "displayName:", match.displayName, "webUrl:", match.webUrl);
}
main().catch((e) => {
    console.error(e?.response?.data || e);
    process.exit(1);
});
//# sourceMappingURL=discover.js.map