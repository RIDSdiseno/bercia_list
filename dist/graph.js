import axios from "axios";
import { getGraphToken } from "./auth.js";
const base = "https://graph.microsoft.com/v1.0";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
function logError(e, method, url) {
    const status = e?.response?.status;
    const data = e?.response?.data;
    console.error(`❌ ${method} ${base + url} -> ${status}`);
    if (data)
        console.error("Detalles:", JSON.stringify(data, null, 2));
}
export async function graphGet(url) {
    try {
        const token = await getGraphToken();
        const { data } = await axios.get(base + url, {
            headers: { Authorization: `Bearer ${token}` },
        });
        return data;
    }
    catch (e) {
        logError(e, "GET", url);
        throw e;
    }
}
export async function graphPost(url, body) {
    const token = await getGraphToken();
    try {
        const { data } = await axios.post(`${GRAPH_BASE}${url}`, body, {
            headers: {
                Authorization: `Bearer ${token}`,
                "Content-Type": "application/json",
            },
        });
        return data;
    }
    catch (e) {
        console.error("❌ Graph POST error:", JSON.stringify(e?.response?.data, null, 2));
        throw e;
    }
}
export async function graphPatch(url, body) {
    try {
        const token = await getGraphToken();
        await axios.patch(base + url, body, {
            headers: { Authorization: `Bearer ${token}` },
        });
    }
    catch (e) {
        logError(e, "PATCH", url);
        throw e;
    }
}
