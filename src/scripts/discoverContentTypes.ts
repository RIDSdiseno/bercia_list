import { graphGet } from "../graph.js";
import { cfg } from "../config.js";

(async () => {
  const res = await graphGet<{ value: any[] }>(
    `/sites/${cfg.siteId}/lists/${cfg.listId}/contentTypes`
  );

  console.log("=== CONTENT TYPES ===");
  for (const ct of res.value) {
    console.log(`${ct.name} -> ${ct.id}`);
  }
})();
