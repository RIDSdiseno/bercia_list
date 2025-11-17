export type ParsedMail = {
    id: string;
    subject: string;
    from: string;
    to: string[];
    cc: string[];
    bcc: string[];
    receivedDateTime: string;
    bodyPreview: string;
    bodyHtml?: string;
};
/** === Choices esperadas en SharePoint (ajústalas si tu lista difiere) === */
export declare const PRIORIDAD_CHOICES: readonly ["Alta", "Media", "Baja"];
export declare const TIPO_TAREA_CHOICES: readonly ["Instalación", "Envío", "Cubicación por planos", "Cubicación en terreno", "Costeo de proyecto", "Evaluación postventa", "Producto interno", "Mantención", "Producción", "Otro"];
/** Heurística → choice válida (Alta/Media/Baja) */
export declare function guessPrioridad(text: string): (typeof PRIORIDAD_CHOICES)[number];
/** Heurística → choice válida de tipo */
export declare function guessTipoTarea(text: string): (typeof TIPO_TAREA_CHOICES)[number];
/** Extrae la primera fecha dd/mm/yyyy o dd-mm-yyyy y devuelve ISO (UTC 00:00) */
export declare function extractFirstDateISO(text: string): string | undefined;
/** Intenta "Cliente: X" y "Proyecto: Y" y retorna "X / Y" o subject limpio */
export declare function extractClientProject(subject: string, body: string): string;
export declare function truncate(s: string, max?: number): string;
export declare function fetchMailByIdForUser(token: string, userId: string, messageId: string): Promise<ParsedMail>;
export declare function hasAnyRecipientInDomain(mail: ParsedMail, domain: string): boolean;
export declare function collectRecipientsInDomain(mail: ParsedMail, domain: string, exclude?: string[]): string[];
export declare function sendConfirmationEmail(token: string, fromUserId: string, toAddress: string, asuntoOriginal: string): Promise<void>;
//# sourceMappingURL=mail.d.ts.map