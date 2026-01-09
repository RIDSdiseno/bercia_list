import { graphGet } from "../graph";
import { cfg } from "../config";
(async () => {
    const res = await graphGet(`/sites/${cfg.siteId}/lists/${cfg.listId}/contentTypes`);
    console.log("=== CONTENT TYPES ===");
    for (const ct of res.value) {
        console.log(`${ct.name} -> ${ct.id}`);
    }
})();
