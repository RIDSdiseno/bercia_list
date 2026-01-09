// src/server.ts
import express from "express";
import { cfg } from "./config.js";
import { processInboxOnce, processSimulatedMail } from "./mailProcessor.js";
import { getRequiredColumns } from "./debugColumns.js";
import { processEstadoListOnce } from "./listWatcher.js";
const app = express();
app.use(express.json());
// Healthcheck simple
app.get("/health", (_, res) => res.json({ ok: true }));
// Fuerza una corrida inmediata del polling
app.post("/run-now", async (_, res) => {
    try {
        await processInboxOnce();
        await processEstadoListOnce();
        res.json({ ok: true });
    }
    catch (e) {
        res.status(500).json({ ok: false, error: e?.message || String(e) });
    }
});
// Simula un correo desde Postman
app.post("/test-create", async (req, res) => {
    try {
        const created = await processSimulatedMail(req.body);
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
        const cols = await getRequiredColumns();
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
const pollMs = Number(cfg.pollIntervalMs) || 60_000;
setInterval(async () => {
    if (!runningInbox) {
        runningInbox = true;
        try {
            await processInboxOnce();
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
            await processEstadoListOnce();
        }
        catch (err) {
            console.error("Polling estado lista error:", err?.message || err);
        }
        finally {
            runningEstado = false;
        }
    }
}, pollMs);
