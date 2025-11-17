"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.getAppToken = getAppToken;
exports.gget = gget;
exports.gpost = gpost;
const axios_1 = __importDefault(require("axios"));
const qs = __importStar(require("querystring"));
const tokenUrl = (tenant) => `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;
async function getAppToken(tenantId, clientId, clientSecret) {
    const data = qs.stringify({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
    });
    const res = await axios_1.default.post(tokenUrl(tenantId), data, {
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
    });
    return res.data.access_token;
}
async function gget(url, token) {
    return axios_1.default.get(url, { headers: { Authorization: `Bearer ${token}` } });
}
async function gpost(url, token, body) {
    return axios_1.default.post(url, body, { headers: { Authorization: `Bearer ${token}` } });
}
//# sourceMappingURL=graph.js.map