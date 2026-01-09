export function parseMail(bodyText) {
    const lines = bodyText
        .replace(/\r/g, "")
        .split("\n")
        .map(l => l.trim())
        .filter(Boolean);
    const out = {};
    const take = (key) => lines.find(l => l.toLowerCase().startsWith(key.toLowerCase() + ":"));
    const kv = (line) => line ? line.split(":").slice(1).join(":").trim() : undefined;
    out.tipodetarea = kv(take("Tipodetarea")) || kv(take("Tipo de tarea"));
    out.clienteProyecto = kv(take("Cliente")) || kv(take("Cliente/Proyecto"));
    out.prioridad = kv(take("Prioridad"));
    const resp = kv(take("Responsables")) || kv(take("Responsable"));
    if (resp) {
        out.responsables = resp
            .split(/[;, ]+/)
            .map(x => x.trim().toLowerCase())
            .filter(x => x.includes("@"));
    }
    // 🟢 estas ya las tenías
    out.fechaSolicitada =
        kv(take("Fecha solicitada")) || kv(take("Fechasolicitada"));
    // 🆕 nueva: Fecha confirmada escrita por el solicitante
    out.fechaConfirmada =
        kv(take("Fecha confirmada")) || kv(take("Fechaconfirmada"));
    out.observaciones = kv(take("Observaciones")) || bodyText.trim();
    return out;
}
