// src/scripts/sharepoint.ts
import axios from "axios";

export type CreateFieldsPayload = {
  siteId: string;
  listId: string;

  /**
   * Campos de la lista (internal names).
   * Para Persona/Lookup usa:
   *  - Single:  ResponsablesLookupId: 12
   *  - Multi:   ResponsablesLookupId: [12, 45, 78]
   */
  fields: Record<string, any>;

  /**
   * Opcional: si quieres que Graph te devuelva los fields ya resueltos
   */
  preferReturnRepresentation?: boolean;
};

export async function createListItem(token: string, input: CreateFieldsPayload) {
  const url = `https://graph.microsoft.com/v1.0/sites/${input.siteId}/lists/${input.listId}/items`;

  const headers: Record<string, string> = {
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  };

  if (input.preferReturnRepresentation) {
    headers["Prefer"] = "return=representation";
  }

  const { data } = await axios.post(
    url,
    { fields: input.fields }, // <- aquí van los fields
    { headers }
  );

  return data;
}
