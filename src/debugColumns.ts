import { graphGet } from "./graph.js";
import { cfg } from "./config.js";

export async function getRequiredColumns() {
  const res = await graphGet<{ value: any[] }>(
    `/sites/${cfg.siteId}/lists/${cfg.listId}/columns`
  );

  return res.value.map(c => ({
    name: c.name,                 // nombre interno Graph
    displayName: c.displayName,   // visible
    required: c.required,
    type: c.columnType,
    choices: c.choice?.choices || null,
  }));
}
