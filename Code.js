function doGet(e) {
  asegurarHojaUsuarios();

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SCM Piura - Panel de Control')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
