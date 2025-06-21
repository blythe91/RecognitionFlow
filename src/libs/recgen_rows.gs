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

function generarReconocimientosPorFilas(filasCSV, sheet_Id, template_Id, folder_Id) {
  // limpia las propiedades almacenadas y los triggers antes de comenzar a generar.
  detenerGeneracionReconocimientos();

  try {
    var sheet = SpreadsheetApp.openById(sheet_Id).getSheetByName("data");
    var data = sheet.getDataRange().getValues();
    var folder = DriveApp.getFolderById(folder_Id);

    // Convierte "4,7,25" → [4, 7, 25]
    var filas = filasCSV.split(",").map(f => parseInt(f.trim()));

    filas.forEach(fila => {
      var i = fila;

      if (i <= 0 || i >= data.length) {
        Logger.log("Fila " + fila + " fuera de rango. Se omite.");
        return;
      }

      var filaData = data[i];

      if (!filaData[1] || !filaData[3] || !filaData[6]) {
        Logger.log("Fila " + fila + " con datos incompletos. Se omite.");
        return;
      }

      var primerNombre = filaData[1];
      var segundoNombre = filaData[2];
      var primerApellido = filaData[3];
      var segundoApellido = filaData[4];
      var prefijoDoc = filaData[5];
      var docIdentidad = filaData[6];
      var textoReconocimiento = filaData[8];  // COL-I
      var textoFecha = filaData[9];           // COL-J
      var codigoCertificado = filaData[11];

      var nombreCompleto = primerNombre + " " + (segundoNombre || "") + " " + primerApellido + " " + (segundoApellido || "");
      nombreCompleto = nombreCompleto.trim();
      var documentoIdentidad = (prefijoDoc ? prefijoDoc + " " : "") + docIdentidad;
      var nombreCertificado = "Certificado_" + codigoCertificado + "_" + primerNombre + "_" + primerApellido + ".pdf";

      var archivos = folder.getFilesByName(nombreCertificado);
      if (archivos.hasNext()) {
        Logger.log("Certificado ya existe para la fila " + fila + ". Se omite.");
        return;
      }

      var slideCopy = DriveApp.getFileById(template_Id).makeCopy("Temp_" + nombreCertificado, folder);
      var presentation = SlidesApp.openById(slideCopy.getId());
      var slide = presentation.getSlides()[0];

      slide.replaceAllText("{{nombre-participante}}", nombreCompleto);
      slide.replaceAllText("{{di-participante}}", documentoIdentidad);

      // slide.replaceAllText("{{texto-reconocimiento}}", textoReconocimiento);
      
      slide.replaceAllText("{{texto-fecha}}", textoFecha);
      slide.replaceAllText("{{cod-certificado}}", codigoCertificado);

      const shapes = slide.getShapes();
      for (let j = 0; j < shapes.length; j++) {
        const shape = shapes[j];
        if (typeof shape.getText === "function") {
          const text = shape.getText().asString();
          if (text.includes("{{texto-reconocimiento}}")) {
            insertarTextoFormateado(shape, "{{texto-reconocimiento}}", textoReconocimiento);
            break;
          }
        }
      }

      presentation.saveAndClose();

      var pdfUrl = "https://docs.google.com/presentation/d/" + slideCopy.getId() + "/export/pdf";
      var response = UrlFetchApp.fetch(pdfUrl, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
      });

      var pdfBlob = response.getBlob().setName(nombreCertificado);
      folder.createFile(pdfBlob);

      sheet.getRange(i + 1, 20).setValue(pdfUrl);

      DriveApp.getFileById(slideCopy.getId()).setTrashed(true);

      Logger.log("Certificado generado para la fila " + fila);
    });

  } catch (e) {
    Logger.log("Error en generarReconocimientosPorFilas: " + e.toString());
  }
}

function insertarTextoFormateado(shape, marcador, contenido) {
  const textRange = shape.getText();
  const textoOriginal = textRange.asString();

  if (!textoOriginal.includes(marcador)) return;

  textRange.setText("");

  // Este regex captura:
  // 1. Grupo 1: marcadores (*, _, *_ o #)
  // 2. Grupo 2: texto dentro de los marcadores
  // 3. Grupo 3: cualquier otro texto plano
  const regex = /(\*{1,2}_?|\_{1,2}\*?|#)([^\*_#]+?)\1|([^*_#]+)/g;
  let match;
  let cursor = 0;

  while ((match = regex.exec(contenido)) !== null) {
    const textoPlano = match[2] || match[3];
    const formato = match[1] || "";

    const appended = textRange.appendText(textoPlano);
    const style = appended.getTextStyle();

    // Evaluar formato
    const bold = formato.includes("*") || formato === "#";
    const italic = formato.includes("_") || formato === "#";

    style.setBold(bold);
    style.setItalic(italic);

    cursor += textoPlano.length;
  }
}

