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

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Reconocimientos')
    .addSubMenu(
      ui.createMenu('Generar Reconocimientos')
        .addItem('Todos', 'mostrarModalReconocimientos')
        .addItem('Por Filas', 'mostrarModalReconocimientosPorFilas')
    )
    .addSubMenu(
      ui.createMenu('Enviar Reconocimientos')
        .addItem('Todos', 'mostrarModalEnviarTodos')
        .addItem('Por Filas', 'mostrarModalEnviarPorFilas')
        .addItem('Por Rango de Filas', 'mostrarModalEnviarPorRango')
    )
    .addToUi();
}

function mostrarModalEnviarTodos() {
  const html = HtmlService.createHtmlOutputFromFile('modal_send_all')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enviar Reconocimientos');
}

function mostrarModalReconocimientos() {
  const html = HtmlService.createHtmlOutputFromFile('modal_cert_all')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generar Reconocimientos');
}

function mostrarModalReconocimientosPorFilas() {
  const html = HtmlService.createHtmlOutputFromFile('modal_cert_rows')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generar Reconocimientos');
}

function mostrarModalEnviarPorFilas() {
  const html = HtmlService.createHtmlOutputFromFile('modal_send_rows')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enviar Reconocimientos');
}

function mostrarModalEnviarPorRango() {
  const html = HtmlService.createHtmlOutputFromFile('modal_send_range')
    .setWidth(450)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enviar Reconocimientos');
}