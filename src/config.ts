import "dotenv/config";

export const cfg = {
  contentTypeId: process.env.CONTENT_TYPE_ID || "",

  // Azure app (tenant Bercia)
  tenantId: process.env.BERCIA_TENANT_ID!,
  clientId: process.env.BERCIA_CLIENT_ID!,
  clientSecret: process.env.BERCIA_CLIENT_SECRET!,

  // Mailbox administrador
  mailboxUserId: process.env.MAILBOX_USER_ID!,
  targetFolderPath: process.env.TARGET_FOLDER_PATH || "Prueba-Flujo-list",

  // ✅ NUEVO: carpeta donde se moverán los correos ya procesados
  processedFolderPath:
    process.env.PROCESSED_FOLDER_PATH || "Procesados-Flujo-list",

  berciaDomain: (process.env.BERCIA_DOMAIN || "@bercia.cl").toLowerCase(),
  adminEmail: (process.env.ADMIN_EMAIL || "administrador@bercia.cl")
    .toLowerCase(),

  // SharePoint destino
  siteId: process.env.SITE_ID!,
  listId: process.env.LIST_ID!,

  pollIntervalMs: Number(process.env.POLL_INTERVAL_MS || 60000),
};

function req(v: string | undefined, name: string) {
  if (!v) throw new Error(`Falta ${name} en .env`);
}

req(cfg.tenantId, "BERCIA_TENANT_ID");
req(cfg.clientId, "BERCIA_CLIENT_ID");
req(cfg.clientSecret, "BERCIA_CLIENT_SECRET");
req(cfg.mailboxUserId, "MAILBOX_USER_ID");
req(cfg.siteId, "SITE_ID");
req(cfg.listId, "LIST_ID");
