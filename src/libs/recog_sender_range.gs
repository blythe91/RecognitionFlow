/**
 * RecognitionFlow - Librería para generación y envío de Reconocimientos digitales
 * © 2025 Oscar Giovanni Castro Contreras
 * 
 * Licencia dual:
 * - MIT License (LICENSE-MIT)
 * - GNU GPLv3 (LICENSE-GPL)
 * 
 * El usuario puede optar por cualquiera de estas licencias.
 */

/**
 * Envía Reconocimientos por correo para un rango de filas dado.
 *
 * @param {number} startRow     Fila inicial (1-based).
 * @param {number} endRow       Fila final (1-based).
 * @param {string} sheetId      ID de la hoja de cálculo.
 * @param {string} folderId     ID de la carpeta de Reconocimientos PDF.
 * @param {string} mensajeEmail Texto adicional a incluir en el cuerpo del correo.
 */
function enviarReconocimientosEmailPorRango(startRow, endRow, sheetId, folderId, mensajeEmail) {
  try {
    Logger.log("📄 Rango: " + startRow + " a " + endRow);
    if (endRow - startRow + 1 < 5 || endRow - startRow + 1 > 30) {
      Logger.log("❌ El rango debe tener entre 5 y 30 filas.");
      return;
    }

    const sheet = SpreadsheetApp.openById(sheetId).getSheetByName("data");
    const data = sheet.getDataRange().getValues();
    const folder = DriveApp.getFolderById(folderId);

    for (let i = startRow; i <= endRow; i++) {
      if (i <= 0 || i >= data.length) {
        Logger.log("⚠️ Índice de fila inválido: " + (i + 1));
        continue;
      }

      const fila = data[i];
      const primerNombre = fila[1];
      const segundoNombre = fila[2];
      const primerApellido = fila[3];
      const segundoApellido = fila[4];
      const correo = fila[7];
      const textoReconocimiento = fila[8]; // COL-I
      const textoFecha = fila[9];          // COL-J
      const codigoCertificado = fila[11];  // COL-L

      const nombreCompleto = (primerNombre + " " + (segundoNombre || "") + " " + primerApellido + " " + (segundoApellido || "")).trim();
      const nombreCertificado = `Certificado_${codigoCertificado}_${primerNombre}_${primerApellido}.pdf`;

      const archivos = folder.getFilesByName(nombreCertificado);
      if (archivos.hasNext()) {
        const archivo = archivos.next();

        const asunto = `Reconocimiento digital - ${codigoCertificado}`;
        const cuerpo = `Estimado(a) ${nombreCompleto}, reciba un cordial saludo en nombre de la Universidad Nacional Experimental del Táchira UNET.\n\n` +
                       `Nos permitimos informarle que en el archivo adjunto podrá obtener el certificado digital correspondiente al reconocimiento otorgado a usted, ${textoFecha}.\n\n` +
                       `El reconocimiento fue conferido ${textoReconocimiento}.\n\n` +
                       `${mensajeEmail ? mensajeEmail + "\n\n" : ""}` +
                       `Nota: Ante cualquier duda o aclaratoria relacionada con su certificado, escribir al correo electrónico: soporteinv@unet.edu.ve; sugiriendo en tal caso que sea como respuesta a este correo electrónico.\n`;

        MailApp.sendEmail({
          to: correo,
          subject: asunto,
          body: cuerpo,
          attachments: [archivo.getAs(MimeType.PDF)]
        });

        Logger.log("📧 " + nombreCertificado + " enviado a: " + correo);
      } else {
        Logger.log("⚠️ No se encontró el certificado para: " + nombreCompleto);
      }
    }

    Logger.log("✅ Proceso completado para las filas del " + startRow + " al " + endRow);

  } catch (e) {
    Logger.log("❌ Error en enviarReconocimientosEmailPorRango: " + e.toString());
  }
}
