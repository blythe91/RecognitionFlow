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

function procesarYEnviarTotalReconocimientos(sheetUrl, folderUrl, batchSize, mensajeEmail, digitoControl) {
  const sheetId = obtenerIdHojaCalculo(sheetUrl);
  const folderId = obtenerIdCarpeta(folderUrl);

  Logger.log("URL de hoja de cálculo: " + sheetUrl);
  Logger.log("ID de hoja de cálculo: " + sheetId);
  Logger.log("////////////////");
  Logger.log("URL de carpeta destino: " + folderUrl);
  Logger.log("ID de carpeta destino: " + folderId);

  var props = PropertiesService.getScriptProperties();
  props.setProperty("sheet_Id_envio", sheetId);
  props.setProperty("folder_Id_envio", folderId);
  props.setProperty("batch_size_envio", batchSize);
  props.setProperty("mensajeEmail_envio", mensajeEmail);


  if (!sheetId || !folderId) {
    throw new Error("Alguno de los IDs no pudo extraerse correctamente.");
  }
  try {
    enviarReconocimientosEmail(sheetId, folderId, batchSize, mensajeEmail);
    return;
  } catch (e) {
    Logger.log("❌ Error en enviarReconocimientosEmail: " + e.message);
    throw e;
  }
}



function procesarYEnviarReconocimientosPorFilas(filasSeleccionadas, sheetUrl, folderUrl, mensajeEmail) {
  const sheetId = obtenerIdHojaCalculo(sheetUrl);
  const folderId = obtenerIdCarpeta(folderUrl);

  Logger.log("URL de hoja de cálculo: " + sheetUrl);
  Logger.log("ID de hoja de cálculo: " + sheetId);
  Logger.log("////////////////");
  Logger.log("URL de carpeta destino: " + folderUrl);
  Logger.log("ID de carpeta destino: " + folderId);

  if (!sheetId || !folderId) {
    throw new Error("Alguno de los IDs no pudo extraerse correctamente.");
  }
  try {
    enviarReconocimientosEmailPorFilas(sheetId, folderId, filasSeleccionadas, mensajeEmail);
    return;
  } catch (e) {
    Logger.log("❌ Error en enviarReconocimientosEmailPorFilas: " + e.message);
    throw e;
  }
}

function procesarYEnviarReconocimientosPorRango(startRow, endRow, sheetUrl, folderUrl, mensajeEmail){
  const sheetId = obtenerIdHojaCalculo(sheetUrl);
  const folderId = obtenerIdCarpeta(folderUrl);

  Logger.log("URL de hoja de cálculo: " + sheetUrl);
  Logger.log("ID de hoja de cálculo: " + sheetId);
  Logger.log("////////////////");
  Logger.log("URL de carpeta destino: " + folderUrl);
  Logger.log("ID de carpeta destino: " + folderId);

  if (!sheetId || !folderId) {
    throw new Error("Alguno de los IDs no pudo extraerse correctamente.");
  }
  try {
    enviarReconocimientosEmailPorRango(startRow, endRow, sheetId, folderId, mensajeEmail);
    return;
  } catch (e) {
    Logger.log("❌ Error en enviarReconocimientosEmailPorFilas: " + e.message);
    throw e;
  }
}

function procesarYGenerarReconocimientos(sheetUrl, templateUrl, folderUrl, batchSize) {
  const sheetId = obtenerIdHojaCalculo(sheetUrl);
  const templateId = obtenerIdPlantillaSlide(templateUrl);
  const folderId = obtenerIdCarpeta(folderUrl);

  Logger.log("URL de hoja de cálculo: " + sheetUrl);
  Logger.log("ID de hoja de cálculo: " + sheetId);
  Logger.log("////////////////");
  Logger.log("URL de la plantilla del certificado: " + templateUrl);
  Logger.log("ID de la plantilla del certificado: " + templateId);
  Logger.log("////////////////");
  Logger.log("URL de carpeta destino: " + folderUrl);
  Logger.log("ID de carpeta destino: " + folderId);

  if (!sheetId || !templateId || !folderId) {
    throw new Error("Alguno de los IDs no pudo extraerse correctamente.");
  }

  generarReconocimientos(sheetId, templateId, folderId, batchSize);
}


function procesarYGenerarReconocimientosPorFilas(filasCSV, sheetUrl, templateUrl, folderUrl) {
  const sheetId = obtenerIdHojaCalculo(sheetUrl);
  const templateId = obtenerIdPlantillaSlide(templateUrl);
  const folderId = obtenerIdCarpeta(folderUrl);

  Logger.log("URL de hoja de cálculo: " + sheetUrl);
  Logger.log("ID de hoja de cálculo: " + sheetId);
  Logger.log("////////////////");
  Logger.log("URL de la plantilla del certificado: " + templateUrl);
  Logger.log("ID de la plantilla del certificado: " + templateId);
  Logger.log("////////////////");
  Logger.log("URL de carpeta destino: " + folderUrl);
  Logger.log("ID de carpeta destino: " + folderId);

  if (!sheetId || !templateId || !folderId) {
    throw new Error("Alguno de los IDs no pudo extraerse correctamente.");
  }

  generarReconocimientosPorFilas(filasCSV, sheetId, templateId, folderId);
}

function detenerGeneracionReconocimientos() {
  eliminarTrigger();
  PropertiesService.getScriptProperties().deleteAllProperties();
  PropertiesService.getScriptProperties().deleteProperty("lastProcessedIndex");
  PropertiesService.getScriptProperties().deleteProperty("totalReconocimientos");
}