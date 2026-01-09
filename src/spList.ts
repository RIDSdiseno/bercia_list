// src/spList.ts
import axios from "axios";
import { getSharePointToken } from "./auth.js";

/**
 * Setea columna Persona multi usando SharePoint REST con PATCH (odata=nometadata).
 * En nometadata, los multi-lookup/person se mandan como array directo.
 */
export async function spSetResponsables(
  webUrl: string,
  listGuid: string,
  itemId: number,
  responsablesIds: number[]
) {
  if (!responsablesIds.length) return;

  const token = await getSharePointToken();
  const url = `${webUrl}/_api/web/lists(guid'${listGuid}')/items(${itemId})`;

  // multi-person => array directo
  const body: any = {
    ResponsablesId: responsablesIds,
  };

  await axios.patch(url, body, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=nometadata",
      "Content-Type": "application/json;odata=nometadata",
      "IF-MATCH": "*",
    },
  });
}
