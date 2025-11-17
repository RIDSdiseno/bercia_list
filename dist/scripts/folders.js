"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ensureFolderPath = ensureFolderPath;
// src/scripts/folders.ts
const graph_1 = require("./graph");
async function getChildByName(token, userId, parentId, name) {
    const { data } = await (0, graph_1.gget)(`https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/${parentId}/childFolders?$select=id,displayName`, token);
    return data.value.find((f) => f.displayName === name) ?? null;
}
async function createChild(token, userId, parentId, name) {
    const { data } = await (0, graph_1.gpost)(`https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/${parentId}/childFolders`, token, { displayName: name });
    return data;
}
// Asegura ruta bajo Inbox; retorna folderId final
async function ensureFolderPath(token, userId, path) {
    const { data: inbox } = await (0, graph_1.gget)(`https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/Inbox?$select=id,displayName`, token);
    let parentId = inbox.id;
    for (const seg of path.split("/").map((s) => s.trim()).filter(Boolean)) {
        const existing = await getChildByName(token, userId, parentId, seg);
        parentId = existing ? existing.id : (await createChild(token, userId, parentId, seg)).id;
    }
    return parentId;
}
//# sourceMappingURL=folders.js.map