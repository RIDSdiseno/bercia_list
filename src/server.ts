import express from "express";
import { cfg } from "./config";
import { processInboxOnce, processSimulatedMail } from "./mailProcessor";
import { getRequiredColumns } from "./debugColumns";
import { processEstadoListOnce } from "./listWatcher"; // 👈 NUEVO

const app = express();
app.use(express.json());

// Healthcheck simple
app.get("/health", (_, res) => res.json({ ok: true }));

// Fuerza una corrida inmediata del polling
app.post("/run-now", async (_, res) => {
  try {
    // 👇 ahora corre ambas cosas: correos + estado lista
    await processInboxOnce();
    await processEstadoListOnce();

    res.json({ ok: true });
  } catch (e: any) {
    res.status(500).json({ ok: false, error: e?.message || String(e) });
  }
});

// Simula un correo desde Postman
app.post("/test-create", async (req, res) => {
  try {
    const created = await processSimulatedMail(req.body);
    res.json({ ok: true, created });
  } catch (e: any) {
    const details = e?.response?.data || e?.message || e;
    res.status(400).json({ ok: false, error: details });
  }
});

// Debug columns
app.get("/debug-columns", async (_, res) => {
  try {
    const cols = await getRequiredColumns();
    res.json(cols);
  } catch (e: any) {
    res.status(500).json({ ok: false, error: e?.message });
  }
});

app.listen(8080, () => {
  console.log("Server running on http://localhost:8080");
});

// Loop de polling
setInterval(() => {
  // 👇 lee correos y crea LIST
  processInboxOnce().catch(err =>
    console.error("Polling correo error:", err?.message || err)
  );

  // 👇 revisa cambios de estado en la lista y manda correos de Confirmada/Rechazada
  processEstadoListOnce().catch(err =>
    console.error("Polling estado lista error:", err?.message || err)
  );
}, cfg.pollIntervalMs);
