// src/scripts/folders.ts
import { gget, gpost } from "./graph";

type MailFolder = { id: string; displayName: string };

async function getChildByName(token: string, userId: string, parentId: string, name: string) {
  const { data } = await gget(
    `https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/${parentId}/childFolders?$select=id,displayName`,
    token
  );
  return (data.value as MailFolder[]).find((f) => f.displayName === name) ?? null;
}

async function createChild(token: string, userId: string, parentId: string, name: string) {
  const { data } = await gpost(
    `https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/${parentId}/childFolders`,
    token,
    { displayName: name }
  );
  return data as MailFolder;
}

// Asegura ruta bajo Inbox; retorna folderId final
export async function ensureFolderPath(token: string, userId: string, path: string) {
  const { data: inbox } = await gget(
    `https://graph.microsoft.com/v1.0/users/${userId}/mailFolders/Inbox?$select=id,displayName`,
    token
  );
  let parentId = inbox.id as string;

  for (const seg of path.split("/").map((s) => s.trim()).filter(Boolean)) {
    const existing = await getChildByName(token, userId, parentId, seg);
    parentId = existing ? existing.id : (await createChild(token, userId, parentId, seg)).id;
  }
  return parentId;
}
