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

const DEFAULT_BATCH_SIZE = 30;

function obtenerBatchSizePorDefecto() {
  return DEFAULT_BATCH_SIZE;
}

function obtenerIdHojaCalculo(url) {
  // Extrae el ID de una URL de Google Sheets
  var regex = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
  var match = url.match(regex);

  Logger.log("////////////////");
  Logger.log("URL de hoja de cálculo: " + url);
  Logger.log("ID de hoja de cálculo: " + (match ? match[1] : "NO DETECTADO"));
  Logger.log("////////////////");

  return match ? match[1] : null;

  
}
function obtenerIdPlantillaSlide(url) {
  // Extrae el ID de una URL de Google Slides
  var regex = /\/presentation\/d\/([a-zA-Z0-9-_]+)/;
  var match = url.match(regex);

  Logger.log("////////////////");
  Logger.log("URL de plantillas slide: " + url);
  Logger.log("ID de plantillas slide: " + (match ? match[1] : "NO DETECTADO"));
  Logger.log("////////////////");

  return match ? match[1] : null;
}
function obtenerIdCarpeta(url) {
  // Extrae el ID de una URL de Google Drive Folder
  var regex = /\/folders\/([a-zA-Z0-9-_]+)/;
  var match = url.match(regex);

  Logger.log("////////////////");
  Logger.log("URL de carpeta: " + url);
  Logger.log("ID de carpeta: " + (match ? match[1] : "NO DETECTADO"));
  Logger.log("////////////////");

  return match ? match[1] : null;
}

/**
 * Detiene el proceso de envío de reconocimientos:
 * - Elimina propiedades almacenadas (si existen)
 * - Elimina el trigger de continuación (si existe)
 */
function detenerEnvioReconocimientos() {
  try {
    const props = PropertiesService.getScriptProperties();
    const claves = [
      "sheet_Id_envio",
      "folder_Id_envio",
      "batch_size_envio",
      "lastProcessedIndexEnvio",
      "totalEnviados"
    ];

    claves.forEach(clave => {
      if (props.getProperty(clave) !== null) {
        props.deleteProperty(clave);
        Logger.log(`✅ Propiedad eliminada: ${clave}`);
      } else {
        Logger.log(`ℹ️ Propiedad no encontrada: ${clave}`);
      }
    });

    const triggers = ScriptApp.getProjectTriggers();
    let eliminado = false;
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === "continuarEnvioReconocimientos") {
        ScriptApp.deleteTrigger(trigger);
        eliminado = true;
        Logger.log("✅ Trigger 'continuarEnvioReconocimientos' eliminado.");
      }
    });

    if (!eliminado) {
      Logger.log("ℹ️ No se encontró trigger para eliminar.");
    }

    Logger.log("🛑 Proceso de envío detenido.");
  } catch (error) {
    Logger.log("❌ Error al detener el proceso de envío: " + error.toString());
  }
}

/**
 * Detiene el proceso de generación de reconocimientos:
 * - Elimina propiedades almacenadas (si existen)
 * - Elimina el trigger de generación (si existe)
 */
function detenerGeneracionReconocimientos() {
  try {
    const props = PropertiesService.getScriptProperties();
    const claves = [
      "sheet_Id",
      "template_Id",
      "folder_Id",
      "batch_size",
      "lastProcessedIndex",
      "totalReconocimientos"
    ];

    claves.forEach(clave => {
      if (props.getProperty(clave) !== null) {
        props.deleteProperty(clave);
        Logger.log(`✅ Propiedad eliminada: ${clave}`);
      } else {
        Logger.log(`ℹ️ Propiedad no encontrada: ${clave}`);
      }
    });

    const triggers = ScriptApp.getProjectTriggers();
    let eliminado = false;
    triggers.forEach(trigger => {
      if (trigger.getHandlerFunction() === "continuarGeneracionReconocimientos") {
        ScriptApp.deleteTrigger(trigger);
        eliminado = true;
        Logger.log("✅ Trigger 'continuarGeneracionReconocimientos' eliminado.");
      }
    });

    if (!eliminado) {
      Logger.log("ℹ️ No se encontró trigger para eliminar.");
    }

    Logger.log("🛑 Proceso de generación detenido.");
  } catch (error) {
    Logger.log("❌ Error al detener la generación: " + error.toString());
  }
}
/**
 * Formatea un número con separador de miles usando punto (.) en lugar de coma (,)
 * @param {string|number} numero - El número a formatear
 * @returns {string} Número formateado
 */
function formatearCedulaConPuntos(numero) {
  if (!numero) return "";
  const limpio = numero.toString().replace(/\D/g, ""); // Elimina cualquier carácter no numérico
  return limpio.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}
