export type ParsedMail = {
  tipodetarea?: string;
  clienteProyecto?: string;
  prioridad?: string;
  responsables?: string[];
  fechaSolicitada?: string;
  fechaConfirmada?: string;
  observaciones?: string;
};

export function parseMail(bodyText: string): ParsedMail {
  const lines = bodyText
    .replace(/\r/g, "")
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean);

  const out: ParsedMail = {};

  const take = (key: string) =>
    lines.find((l) => l.toLowerCase().startsWith(key.toLowerCase() + ":"));

  const kv = (line?: string) =>
    line ? line.split(":").slice(1).join(":").trim() : undefined;

  out.tipodetarea = kv(take("Tipodetarea")) || kv(take("Tipo de tarea"));
  out.clienteProyecto = kv(take("Cliente")) || kv(take("Cliente/Proyecto"));
  out.prioridad = kv(take("Prioridad"));

  const resp = kv(take("Responsables")) || kv(take("Responsable"));
  if (resp) {
    out.responsables = resp
      .split(/[;, ]+/)
      .map((x) => x.trim().toLowerCase())
      .filter((x) => x.includes("@"));
  }

  out.fechaSolicitada =
    kv(take("Fecha solicitada")) || kv(take("Fechasolicitada"));

  out.fechaConfirmada =
    kv(take("Fecha confirmada")) || kv(take("Fechaconfirmada"));

  // =========================
  // ✅ Observaciones (línea o bloque)
  // =========================
  const obsIdx = lines.findIndex((l) =>
    l.toLowerCase().startsWith("observaciones:")
  );

  // detecta líneas "campo: valor" para cortar el bloque
  const isFieldLine = (l: string) =>
    /^[a-zA-ZáéíóúÁÉÍÓÚñÑ\/\s()_-]+:\s*/.test(l);

  if (obsIdx >= 0) {
    const first = kv(lines[obsIdx]) ?? "";
    const rest: string[] = [];

    for (let i = obsIdx + 1; i < lines.length; i++) {
      const l = lines[i];
      // si aparece otra línea "Campo:" se corta el bloque
      if (isFieldLine(l) && !l.toLowerCase().startsWith("observaciones:")) break;
      rest.push(l);
    }

    const obs = [first, ...rest]
      .map((x) => x.trim())
      .filter(Boolean)
      .join("\n");

    out.observaciones = obs || undefined;
  } else {
    // fallback solo si NO hay Observaciones:
    out.observaciones = bodyText.trim() || undefined;
  }

  return out;
}
