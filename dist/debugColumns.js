"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.getRequiredColumns = getRequiredColumns;
const graph_1 = require("./graph");
const config_1 = require("./config");
async function getRequiredColumns() {
    const res = await (0, graph_1.graphGet)(`/sites/${config_1.cfg.siteId}/lists/${config_1.cfg.listId}/columns`);
    return res.value.map(c => ({
        name: c.name, // nombre interno Graph
        displayName: c.displayName, // visible
        required: c.required,
        type: c.columnType,
        choices: c.choice?.choices || null,
    }));
}
