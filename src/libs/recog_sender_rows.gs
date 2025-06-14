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

function enviarReconocimientosEmailPorFilas(sheet_Id, folder_Id, filasSeleccionadas, textEmail) {
  Logger.log("📄 sheet_ID: " + sheet_Id);
  Logger.log("📄 folder_Id: " + folder_Id);
  Logger.log("📄 filasSeleccionadas: " + JSON.stringify(filasSeleccionadas));

  try {
    var sheet = SpreadsheetApp.openById(sheet_Id).getSheetByName("data");
    var data = sheet.getDataRange().getValues();
    var folder = DriveApp.getFolderById(folder_Id);

    // Convierte "4,7,25" → [4, 7, 25]
    var filas = filasSeleccionadas.split(",").map(f => parseInt(f.trim()));
    for (var j = 0; j < filas.length; j++) {
      var i = filas[j]; // índice de fila (basado en índice 0)

      if (i <= 0 || i >= data.length) {
        Logger.log("❌ Índice de fila inválido: " + i);
        continue;
      }

      var fila = data[i];
      var primerNombre = fila[1];
      var segundoNombre = fila[2];
      var primerApellido = fila[3];
      var segundoApellido = fila[4];
      var correo = fila[7];
      var textoReconocimiento = fila[8]; // COL-I
      var textoFecha = fila[9];          // COL-J
      var codigoCertificado = fila[11];  // COL-L

      var nombreCompleto = (primerNombre + " " + (segundoNombre || "") + " " + primerApellido + " " + (segundoApellido || "")).trim();
      var nombreCertificado = "Certificado_" + codigoCertificado + "_" + primerNombre + "_" + primerApellido + ".pdf";

      var archivos = folder.getFilesByName(nombreCertificado);
      if (archivos.hasNext()) {
        var archivo = archivos.next();

        var asunto = "Reconocimiento digital - " + codigoCertificado;

        var cuerpo = "Estimado(a) " + nombreCompleto + ", reciba un cordial saludo en nombre de la Universidad Nacional Experimental del Táchira UNET.\n\n" +
                     "Nos permitimos informarle que en el archivo adjunto podrá obtener el certificado digital correspondiente al reconocimiento otorgado a usted, " + textoFecha + ".\n\n" +
                     "El reconocimiento fue conferido " + textoReconocimiento + ".\n\n" +
                     (textEmail ? textEmail + "\n\n" : "") +
                     "Nota: Ante cualquier duda o aclaratoria relacionada con su certificado, escribir al correo electrónico: soporteinv@unet.edu.ve; sugiriendo en tal caso que sea como respuesta a este correo electrónico.\n";

        MailApp.sendEmail({
          to: correo,
          subject: asunto,
          body: cuerpo,
          attachments: [archivo.getAs(MimeType.PDF)]
        });

        Logger.log("📧 Certificado enviado a: " + correo);
      } else {
        Logger.log("⚠️ No se encontró el certificado para: " + nombreCompleto);
      }
    }

    Logger.log("✅ Proceso completado para las filas seleccionadas.");

  } catch (e) {
    Logger.log("❌ Error en enviarReconocimientosEmailPorFilas: " + e.toString());
  }
}
