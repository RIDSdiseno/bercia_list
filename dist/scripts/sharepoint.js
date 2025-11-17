"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.createListItem = createListItem;
// src/scripts/sharepoint.ts
const axios_1 = __importDefault(require("axios"));
async function createListItem(token, input) {
    const url = `https://graph.microsoft.com/v1.0/sites/${input.siteId}/lists/${input.listId}/items`;
    const { data } = await axios_1.default.post(url, { fields: input.fields }, { headers: { Authorization: `Bearer ${token}` } });
    return data;
}
//# sourceMappingURL=sharepoint.js.map