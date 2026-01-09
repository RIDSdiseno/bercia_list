"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.graphGet = graphGet;
exports.graphPost = graphPost;
exports.graphPatch = graphPatch;
const axios_1 = __importDefault(require("axios"));
const auth_1 = require("./auth");
const base = "https://graph.microsoft.com/v1.0";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
function logError(e, method, url) {
    const status = e?.response?.status;
    const data = e?.response?.data;
    console.error(`❌ ${method} ${base + url} -> ${status}`);
    if (data)
        console.error("Detalles:", JSON.stringify(data, null, 2));
}
async function graphGet(url) {
    try {
        const token = await (0, auth_1.getGraphToken)();
        const { data } = await axios_1.default.get(base + url, {
            headers: { Authorization: `Bearer ${token}` },
        });
        return data;
    }
    catch (e) {
        logError(e, "GET", url);
        throw e;
    }
}
async function graphPost(url, body) {
    const token = await (0, auth_1.getGraphToken)();
    try {
        const { data } = await axios_1.default.post(`${GRAPH_BASE}${url}`, body, {
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
async function graphPatch(url, body) {
    try {
        const token = await (0, auth_1.getGraphToken)();
        await axios_1.default.patch(base + url, body, {
            headers: { Authorization: `Bearer ${token}` },
        });
    }
    catch (e) {
        logError(e, "PATCH", url);
        throw e;
    }
}
