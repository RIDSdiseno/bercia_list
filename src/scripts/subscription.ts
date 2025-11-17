// src/scripts/subscription.ts
import { gpost } from "./graph";

/** Crea suscripción a Graph para Outlook Mail */
export async function createSubscription(
  token: string,
  resource: string,
  webhookUrl: string
) {
  const body = {
    changeType: "created,updated",
    notificationUrl: webhookUrl,
    resource, // p.ej.: /users/{userId}/mailFolders('inbox')/messages o /mailFolders/{id}/messages
    expirationDateTime: new Date(Date.now() + 36 * 60 * 60 * 1000).toISOString(),
    clientState: "bercia-secret",
  };
  const { data } = await gpost("https://graph.microsoft.com/v1.0/subscriptions", token, body);
  return data;
}
