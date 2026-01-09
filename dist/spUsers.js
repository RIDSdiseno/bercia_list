"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getSiteUserLookupId = getSiteUserLookupId;
const axios_1 = __importDefault(require("axios"));
const auth_1 = require("./auth");
function buildClaims(email) {
    return `i:0#.f|membership|${email.trim().toLowerCase()}`;
}
async function spPost(webUrl, path, body) {
    const token = await (0, auth_1.getSharePointToken)();
    const { data } = await axios_1.default.post(`${webUrl}${path}`, body, {
        headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
        },
    });
    return data;
}
async function spGet(webUrl, path) {
    const token = await (0, auth_1.getSharePointToken)();
    const { data } = await axios_1.default.get(`${webUrl}${path}`, {
        headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json;odata=verbose",
        },
    });
    return data;
}
async function getSiteUserLookupId(email, webUrl) {
    const cleanEmail = email.trim().toLowerCase();
    if (!cleanEmail)
        return null;
    const claims = buildClaims(cleanEmail);
    // 1) ensureuser
    const ensured = await spPost(webUrl, "/_api/web/ensureuser", {
        logonName: claims,
    });
    const ensuredId = Number(ensured?.d?.Id);
    if (Number.isFinite(ensuredId))
        return ensuredId;
    // 2) fallback siteusers
    const users = await spGet(webUrl, "/_api/web/siteusers?$top=5000");
    const arr = users?.d?.results || [];
    const match = arr.find((u) => (u.Email || "").toLowerCase() === cleanEmail) ||
        arr.find((u) => (u.UserPrincipalName || "").toLowerCase() === cleanEmail) ||
        arr.find((u) => (u.LoginName || "").toLowerCase() === claims);
    const id = Number(match?.Id);
    return Number.isFinite(id) ? id : null;
}
