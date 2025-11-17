// src/scripts/sharepoint.ts
import axios from "axios";

type CreateFieldsPayload = {
  siteId: string;
  listId: string;
  fields: Record<string, any>;
};

export async function createListItem(token: string, input: CreateFieldsPayload) {
  const url = `https://graph.microsoft.com/v1.0/sites/${input.siteId}/lists/${input.listId}/items`;
  const { data } = await axios.post(
    url,
    { fields: input.fields },
    { headers: { Authorization: `Bearer ${token}` } }
  );
  return data;
}
