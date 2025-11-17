"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.createSubscription = createSubscription;
// src/scripts/subscription.ts
const graph_1 = require("./graph");
/** Crea suscripción a Graph para Outlook Mail */
async function createSubscription(token, resource, webhookUrl) {
    const body = {
        changeType: "created,updated",
        notificationUrl: webhookUrl,
        resource, // p.ej.: /users/{userId}/mailFolders('inbox')/messages o /mailFolders/{id}/messages
        expirationDateTime: new Date(Date.now() + 36 * 60 * 60 * 1000).toISOString(),
        clientState: "bercia-secret",
    };
    const { data } = await (0, graph_1.gpost)("https://graph.microsoft.com/v1.0/subscriptions", token, body);
    return data;
}
//# sourceMappingURL=subscription.js.map