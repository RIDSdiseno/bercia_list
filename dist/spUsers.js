import axios from "axios";
import { getSharePointToken } from "./auth.js";
function buildClaims(email) {
    return `i:0#.f|membership|${email.trim().toLowerCase()}`;
}
async function spPost(webUrl, path, body) {
    const token = await getSharePointToken();
    const { data } = await axios.post(`${webUrl}${path}`, body, {
        headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
        },
    });
    return data;
}
async function spGet(webUrl, path) {
    const token = await getSharePointToken();
    const { data } = await axios.get(`${webUrl}${path}`, {
        headers: {
            Authorization: `Bearer ${token}`,
            Accept: "application/json;odata=verbose",
        },
    });
    return data;
}
export async function getSiteUserLookupId(email, webUrl) {
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
