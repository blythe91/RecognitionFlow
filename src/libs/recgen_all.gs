/**
 * RecognitionFlow - Librería para generación y envío de Reconocimientos digitales
 * © 2025 Oscar Giovanni Castro Contreras
 * 
 * Licencia dual:
 * - MIT License (LICENSE-MIT)
 * - GNU GPLv3 (LICENSE-GPL)
 */

/**
 * Genera Reconocimientos PDF personalizados a partir de una plantilla de Google Slides.
 * Trabaja en lotes para no exceder límites de tiempo y guarda progreso para continuar después.
 * 
 * @param {string} sheet_Id    ID de la hoja de cálculo donde están los datos.
 * @param {string} template_Id ID de la plantilla de Google Slides.
 * @param {string} folder_Id   ID de la carpeta de Google Drive para guardar los Reconocimientos generados.
 * @param {number} batch_size  Número de Reconocimientos a generar por ejecución.
 */
function generarReconocimientos(sheet_Id, template_Id, folder_Id, batch_size) {
  try {
    const sheet = SpreadsheetApp.openById(sheet_Id).getSheetByName("data");
    const data = sheet.getDataRange().getValues();
    const folder = DriveApp.getFolderById(folder_Id);

    PropertiesService.getScriptProperties().setProperty("totalReconocimientos", data.length - 1);

    let lastProcessedIndex = PropertiesService.getScriptProperties().getProperty("lastProcessedIndex");
    let startIndex = lastProcessedIndex ? parseInt(lastProcessedIndex) + 1 : 1;
    let endIndex = Math.min(startIndex + batch_size, data.length);

    for (let i = startIndex; i < endIndex; i++) {
      const row = data[i];
      const nombreCompleto = `${row[1]} ${row[2] || ""} ${row[3]} ${row[4] || ""}`.trim();
      const documentoIdentidad = (row[5] ? row[5] + " " : "") + row[6];
      const codigoCertificado = row[11];
      const textoReconocimiento = row[8];
      const textoFecha = row[9];
      const nombreCertificado = `Certificado_${codigoCertificado}_${row[1]}_${row[3]}.pdf`;

      if (folder.getFilesByName(nombreCertificado).hasNext()) {
        Logger.log("Certificado ya existe: " + nombreCertificado);
        continue;
      }

      const slideCopy = DriveApp.getFileById(template_Id).makeCopy("Temp_" + nombreCertificado, folder);
      const presentation = SlidesApp.openById(slideCopy.getId());
      const slide = presentation.getSlides()[0];

      slide.replaceAllText("{{nombre-participante}}", nombreCompleto);
      slide.replaceAllText("{{di-participante}}", documentoIdentidad);
      slide.replaceAllText("{{texto-reconocimiento}}", textoReconocimiento);
      slide.replaceAllText("{{texto-fecha}}", textoFecha);
      slide.replaceAllText("{{cod-certificado}}", codigoCertificado);

      presentation.saveAndClose();

      const pdfUrl = "https://docs.google.com/presentation/d/" + slideCopy.getId() + "/export/pdf";
      const response = UrlFetchApp.fetch(pdfUrl, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
      });

      const pdfBlob = response.getBlob().setName(nombreCertificado);
      folder.createFile(pdfBlob);
      sheet.getRange(i + 1, 20).setValue(pdfUrl);

      DriveApp.getFileById(slideCopy.getId()).setTrashed(true);

      PropertiesService.getScriptProperties().setProperty("lastProcessedIndex", i);
    }

    if (endIndex < data.length) {
      triggerGenerarReconocimientos(sheet_Id, template_Id, folder_Id, batch_size);
    } else {
      eliminarTrigger();
      PropertiesService.getScriptProperties().deleteProperty("lastProcessedIndex");
      PropertiesService.getScriptProperties().deleteProperty("totalReconocimientos");
    }
  } catch (e) {
    Logger.log("❌ Error en generarReconocimientos: " + e.toString());
  }
}

function triggerGenerarReconocimientos(sheet_Id, template_Id, folder_Id, batch_size) {
  eliminarTrigger();

  var triggers = ScriptApp.getProjectTriggers();
  var existe = triggers.some(t => t.getHandlerFunction() === "continuarGeneracionReconocimientos");
  if (existe) return;

  PropertiesService.getScriptProperties().setProperty("sheet_Id", sheet_Id);
  PropertiesService.getScriptProperties().setProperty("template_Id", template_Id);
  PropertiesService.getScriptProperties().setProperty("folder_Id", folder_Id);
  PropertiesService.getScriptProperties().setProperty("batch_size", batch_size.toString());

  ScriptApp.newTrigger("continuarGeneracionReconocimientos")
    .timeBased()
    .after(50000)
    .create();
}

function continuarGeneracionReconocimientos() {
  var sheet_Id = PropertiesService.getScriptProperties().getProperty("sheet_Id");
  var template_Id = PropertiesService.getScriptProperties().getProperty("template_Id");
  var folder_Id = PropertiesService.getScriptProperties().getProperty("folder_Id");
  var batch_size = parseInt(PropertiesService.getScriptProperties().getProperty("batch_size"));

  if (sheet_Id && template_Id && folder_Id && batch_size) {
    generarReconocimientos(sheet_Id, template_Id, folder_Id, batch_size);
  }
}

function eliminarTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "continuarGeneracionReconocimientos") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function obtenerProgresoReconocimientos() {
  var props = PropertiesService.getScriptProperties();
  var last = parseInt(props.getProperty("lastProcessedIndex")) || 0;
  var total = parseInt(props.getProperty("totalReconocimientos")) || 0;
  var porcentaje = total ? Math.floor((last / total) * 100) : 0;

  return {
    porcentaje: porcentaje,
    mensaje: "Generando Reconocimientos...",
    generados: last,
    total: total
  };
}
