import { graphPost } from "./graph";
import { cfg } from "./config";

export async function createListItem(fields: any) {
  const safeFields: any = { ...fields };

  // ✅ agrega CT solo si existe (no vacío)
  if (cfg.contentTypeId && cfg.contentTypeId.trim()) {
    safeFields.ContentTypeId = cfg.contentTypeId.trim();
  }

  // ✅ limpia null/undefined
  for (const k of Object.keys(safeFields)) {
    if (safeFields[k] === null || safeFields[k] === undefined) {
      delete safeFields[k];
    }
  }

  return graphPost(
    `/sites/${cfg.siteId}/lists/${cfg.listId}/items`,
    { fields: safeFields }
  );
}
