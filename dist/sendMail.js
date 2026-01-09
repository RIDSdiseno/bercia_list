"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.sendMailNuevaSolicitud = sendMailNuevaSolicitud;
exports.sendMailCambioEstado = sendMailCambioEstado;
exports.sendMailComentarioEncargado = sendMailComentarioEncargado;
// src/sendMail.ts
const graph_1 = require("./graph");
const config_1 = require("./config");
async function sendMailBase(to, subject, htmlBody) {
    const cleanTo = String(to ?? "").trim();
    if (!cleanTo)
        return;
    const body = {
        message: {
            subject,
            body: {
                contentType: "HTML",
                content: htmlBody,
            },
            toRecipients: [
                {
                    emailAddress: { address: cleanTo },
                },
            ],
        },
        saveToSentItems: true,
    };
    await (0, graph_1.graphPost)(`/users/${config_1.cfg.mailboxUserId}/sendMail`, body);
}
// 🔹 correo cuando se crea la solicitud
async function sendMailNuevaSolicitud(params) {
    const { to, titulo, cliente, fechaSolicitada, tipodetarea, webUrl } = params;
    const subject = `Tu solicitud "${titulo}" ha sido creada`;
    const html = `
    <p>Hola,</p>
    <p>Tu solicitud ha sido creada correctamente en el sistema de Bercia.</p>
    <ul>
      <li><strong>Título:</strong> ${titulo}</li>
      <li><strong>Cliente/Proyecto:</strong> ${cliente || "No especificado"}</li>
      <li><strong>Tipo de tarea:</strong> ${tipodetarea || "No especificado"}</li>
      <li><strong>Fecha solicitada:</strong> ${fechaSolicitada || "No indicada"}</li>
    </ul>
    ${webUrl
        ? `<p>Puedes ver el detalle aquí: <a href="${webUrl}">${webUrl}</a></p>`
        : ""}
    <p>Saludos,<br/>Alfombras Bercia</p>
  `;
    await sendMailBase(to, subject, html);
}
// 🔹 correo cuando cambia el estado (Confirmada / Rechazada / Fecha modificada)
async function sendMailCambioEstado(params) {
    const { to, titulo, estado, cliente, fechaSolicitada, fechaConfirmada, comentarioEncargado, webUrl, } = params;
    const subject = estado === "Confirmada"
        ? `Tu solicitud "${titulo}" ha sido CONFIRMADA`
        : estado === "Rechazada"
            ? `Tu solicitud "${titulo}" ha sido RECHAZADA`
            : `Tu solicitud "${titulo}" tiene FECHA MODIFICADA`;
    const textoEstado = estado === "Fecha modificada"
        ? `El encargado ha modificado la fecha de tu solicitud.`
        : `El estado de tu solicitud ha cambiado a: <strong>${estado}</strong>.`;
    const html = `
    <p>Hola,</p>
    <p>${textoEstado}</p>
    <ul>
      <li><strong>Título:</strong> ${titulo}</li>
      <li><strong>Cliente/Proyecto:</strong> ${cliente || "No especificado"}</li>
      <li><strong>Fecha solicitada:</strong> ${fechaSolicitada || "No indicada"}</li>
      ${fechaConfirmada
        ? `<li><strong>Fecha confirmada:</strong> ${fechaConfirmada}</li>`
        : ""}
    </ul>
    ${comentarioEncargado
        ? `<p><strong>Comentario del encargado:</strong> ${comentarioEncargado}</p>`
        : ""}
    ${webUrl
        ? `<p>Puedes ver el detalle aquí: <a href="${webUrl}">${webUrl}</a></p>`
        : ""}
    <p>Saludos,<br/>Alfombras Bercia</p>
  `;
    await sendMailBase(to, subject, html);
}
// 🔹 correo cuando el encargado escribe / modifica su comentario
async function sendMailComentarioEncargado(params) {
    const { to, titulo, cliente, comentarioEncargado, webUrl } = params;
    const subject = `Nuevo comentario en tu solicitud "${titulo}"`;
    const html = `
    <p>Hola,</p>
    <p>El encargado ha agregado/modificado un comentario en tu solicitud.</p>
    <ul>
      <li><strong>Título:</strong> ${titulo}</li>
      <li><strong>Cliente/Proyecto:</strong> ${cliente || "No especificado"}</li>
    </ul>
    ${comentarioEncargado
        ? `<p><strong>Comentario del encargado:</strong><br/>${comentarioEncargado.replace(/\r?\n/g, "<br/>")}</p>`
        : ""}
    ${webUrl
        ? `<p>Puedes ver el detalle aquí: <a href="${webUrl}">${webUrl}</a></p>`
        : ""}
    <p>Saludos,<br/>Alfombras Bercia</p>
  `;
    await sendMailBase(to, subject, html);
}
