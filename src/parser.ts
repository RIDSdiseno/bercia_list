export type ParsedMail = {
  tipodetarea?: string;
  clienteProyecto?: string;
  prioridad?: string;
  responsables?: string[];
  fechaSolicitada?: string;
  observaciones?: string;
};

export function parseMail(bodyText: string): ParsedMail {
  const lines = bodyText
    .replace(/\r/g, "")
    .split("\n")
    .map(l => l.trim())
    .filter(Boolean);

  const out: ParsedMail = {};

  const take = (key: string) =>
    lines.find(l => l.toLowerCase().startsWith(key.toLowerCase() + ":"));

  const kv = (line?: string) =>
    line ? line.split(":").slice(1).join(":").trim() : undefined;

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

  out.fechaSolicitada =
    kv(take("Fecha solicitada")) || kv(take("Fechasolicitada"));

  out.observaciones = kv(take("Observaciones")) || bodyText.trim();

  return out;
}
