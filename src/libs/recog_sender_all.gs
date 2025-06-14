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

function enviarReconocimientosEmail(sheet_Id, folder_Id, batchSize, textEmail) {
  Logger.log("📄 sheet_ID recibido en enviarReconocimientosEmail: " + sheet_Id);
  Logger.log("📄 folder_Id recibido en enviarReconocimientosEmail: " + folder_Id);
  Logger.log("📄 batchSize recibido en enviarReconocimientosEmail: " + batchSize);

  try {
    var sheet = SpreadsheetApp.openById(sheet_Id).getSheetByName("data");
    var data = sheet.getDataRange().getValues();
    var folder = DriveApp.getFolderById(folder_Id);

    PropertiesService.getScriptProperties().setProperty("totalEnviados", data.length - 1); // Excluye encabezado

    var props = PropertiesService.getScriptProperties();
    var startIndex = parseInt(props.getProperty("lastProcessedIndexEnvio")) || 1;
    var endIndex = Math.min(startIndex + batchSize, data.length);

    for (var i = startIndex; i < endIndex; i++) {
      var primerNombre = data[i][1];
      var segundoNombre = data[i][2];
      var primerApellido = data[i][3];
      var segundoApellido = data[i][4];
      var correo = data[i][7];
      var textoReconocimiento = data[i][8]; // COL-I
      var textoFecha = data[i][9];          // COL-J
      var codigoCertificado = data[i][11];

      var nombreCompleto = (primerNombre + " " + (segundoNombre || "") + " " + primerApellido + " " + (segundoApellido || "")).trim();
      var nombreCertificado = "Certificado_" + codigoCertificado + "_" + primerNombre + "_" + primerApellido + ".pdf";

      var archivos = folder.getFilesByName(nombreCertificado);
      if (archivos.hasNext()) {
        var archivo = archivos.next();

        var asunto = "Reconocimiento digital - " + codigoCertificado;

        var cuerpo = "Estimado(a) " + nombreCompleto + ", reciba un cordial saludo en nombre de la Universidad Nacional Experimental del Táchira UNET.\n\n" +
                     "Nos permitimos informarle que en el archivo adjunto podrá obtener el certificado digital correspondiente al reconocimiento otorgado a usted, " + textoFecha + ".\n\n" +
                     "El reconocimiento fue conferido " + textoReconocimiento + ".\n\n" +
                     `${mensajeEmail ? mensajeEmail + "\n\n" : ""}` +
                     "Nota: Ante cualquier duda o aclaratoria relacionada con su certificado, escribir al correo electrónico: soporteinv@unet.edu.ve; sugiriendo en tal caso que sea como respuesta a este correo electrónico.\n";

        MailApp.sendEmail({
          to: correo,
          subject: asunto,
          body: cuerpo,
          attachments: [archivo.getAs(MimeType.PDF)]
        });

        Logger.log("Certificado enviado a: " + correo);
      } else {
        Logger.log("No se encontró el certificado para: " + nombreCompleto);
      }
    }

    if (endIndex < data.length) {
      props.setProperty("lastProcessedIndexEnvio", endIndex);
      Logger.log("Proceso pausado. Continuará desde la fila: " + endIndex);
      triggerEnviarReconocimientos();
    } else {
      props.deleteProperty("lastProcessedIndexEnvio");
      props.deleteProperty("totalEnviados");
      eliminarTriggerEnvio();
      Logger.log("Todos los Reconocimientos fueron enviados.");
      return;
    }

  } catch (e) {
    Logger.log("Error en enviarReconocimientosEmail (dentro de enviarReconocimientosEmail): " + e.toString());
  }
}

function triggerEnviarReconocimientos() {
  eliminarTriggerEnvio();

  var triggers = ScriptApp.getProjectTriggers();
  var existe = triggers.some(t => t.getHandlerFunction() === "continuarEnvioReconocimientos");
  if (existe) return;

  ScriptApp.newTrigger("continuarEnvioReconocimientos")
    .timeBased()
    .after(50000)
    .create();
}

function continuarEnvioReconocimientos() {
  var props = PropertiesService.getScriptProperties();
  var sheet_Id = props.getProperty("sheet_Id_envio");
  var folder_Id = props.getProperty("folder_Id_envio");
  var batchSize = parseInt(props.getProperty("batch_size_envio"));

  if (sheet_Id && folder_Id && batchSize) {
    enviarReconocimientosEmail(sheet_Id, folder_Id, batchSize);
  }
}

function eliminarTriggerEnvio() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "continuarEnvioReconocimientos") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function obtenerProgresoEnvio() {
  var props = PropertiesService.getScriptProperties();
  var last = parseInt(props.getProperty("lastProcessedIndexEnvio")) || 0;
  var total = parseInt(props.getProperty("totalEnviados")) || 0;

  var porcentaje = total ? Math.floor((last / total) * 100) : 0;

  return {
    porcentaje: porcentaje,
    mensaje: "Enviando Reconocimientos...",
    generados: last,
    total: total
  };
}
