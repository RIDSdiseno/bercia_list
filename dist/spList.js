"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.spSetResponsables = spSetResponsables;
// src/spList.ts
const axios_1 = __importDefault(require("axios"));
const auth_1 = require("./auth");
/**
 * Setea columna Persona multi usando SharePoint REST con PATCH (odata=nometadata).
 * En nometadata, los multi-lookup/person se mandan como array directo.
 */
async function spSetResponsables(webUrl, listGuid, itemId, responsablesIds) {
    if (!responsablesIds.length)
        return;
    const token = await (0, auth_1.getSharePointToken)();
    const url = `${webUrl}/_api/web/lists(guid'${listGuid}')/items(${itemId})`;
    // multi-person => array directo
    const body = {
        ResponsablesId: responsablesIds,
    };
    await axios_1.default.patch(url, body, {
        headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json;odata=nometadata",
            "IF-MATCH": "*",
        },
    });
}
