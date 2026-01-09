"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
// src/server.ts
const express_1 = __importDefault(require("express"));
const config_1 = require("./config");
const mailProcessor_1 = require("./mailProcessor");
const debugColumns_1 = require("./debugColumns");
const listWatcher_1 = require("./listWatcher");
const app = (0, express_1.default)();
app.use(express_1.default.json());
// Healthcheck simple
app.get("/health", (_, res) => res.json({ ok: true }));
// Fuerza una corrida inmediata del polling
app.post("/run-now", async (_, res) => {
    try {
        await (0, mailProcessor_1.processInboxOnce)();
        await (0, listWatcher_1.processEstadoListOnce)();
        res.json({ ok: true });
    }
    catch (e) {
        res.status(500).json({ ok: false, error: e?.message || String(e) });
    }
});
// Simula un correo desde Postman
app.post("/test-create", async (req, res) => {
    try {
        const created = await (0, mailProcessor_1.processSimulatedMail)(req.body);
        res.json({ ok: true, created });
    }
    catch (e) {
        const details = e?.response?.data || e?.message || e;
        res.status(400).json({ ok: false, error: details });
    }
});
// Debug columns
app.get("/debug-columns", async (_, res) => {
    try {
        const cols = await (0, debugColumns_1.getRequiredColumns)();
        res.json(cols);
    }
    catch (e) {
        res.status(500).json({ ok: false, error: e?.message });
    }
});
app.listen(8080, () => {
    console.log("Server running on http://localhost:8080");
});
// ======================
// Loop de polling (con lock)
// ======================
let runningInbox = false;
let runningEstado = false;
const pollMs = Number(config_1.cfg.pollIntervalMs) || 60_000;
setInterval(async () => {
    if (!runningInbox) {
        runningInbox = true;
        try {
            await (0, mailProcessor_1.processInboxOnce)();
        }
        catch (err) {
            console.error("Polling correo error:", err?.message || err);
        }
        finally {
            runningInbox = false;
        }
    }
    if (!runningEstado) {
        runningEstado = true;
        try {
            await (0, listWatcher_1.processEstadoListOnce)();
        }
        catch (err) {
            console.error("Polling estado lista error:", err?.message || err);
        }
        finally {
            runningEstado = false;
        }
    }
}, pollMs);
