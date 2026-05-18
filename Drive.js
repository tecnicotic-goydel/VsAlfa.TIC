function guardarPadronesEnDrive(periodo, archivosArray) {
  try {
    const carpetaMes = obtenerOCrearCarpetaPadrones(periodo);
    const resultados = [];

    for (let i = 0; i < archivosArray.length; i++) {
      const objArchivo = archivosArray[i];
      const blob = Utilities.newBlob(
        Utilities.base64Decode(objArchivo.data),
        objArchivo.mimeType,
        objArchivo.fileName
      );

      const archivoGuardado = carpetaMes.createFile(blob);
      resultados.push({
        nombre: objArchivo.fileName,
        url: archivoGuardado.getUrl(),
        id: archivoGuardado.getId()
      });
    }

    return {
      success: true,
      mensaje: 'Se guardaron exitosamente ' + archivosArray.length + ' archivo(s) en la carpeta ' + periodo + '.',
      carpeta: {
        nombre: carpetaMes.getName(),
        id: carpetaMes.getId(),
        url: carpetaMes.getUrl()
      },
      archivos: resultados
    };
  } catch (e) {
    return { success: false, mensaje: e.message };
  }
}

function obtenerOCrearCarpetaPadrones(periodo) {
  const periodoNormalizado = normalizarPeriodo(periodo, '');
  if (!periodoNormalizado) {
    throw new Error('No se recibió un período válido para guardar el padrón.');
  }

  const carpetaBase = obtenerOCrearCarpetaBasePadrones();
  const subcarpetas = carpetaBase.getFoldersByName(periodoNormalizado);

  if (subcarpetas.hasNext()) {
    return subcarpetas.next();
  }

  return carpetaBase.createFolder(periodoNormalizado);
}

function obtenerOCrearCarpetaBasePadrones() {
  const carpetasBase = DriveApp.getFoldersByName(CONFIG.NOMBRE_CARPETA_PADRONES);
  if (carpetasBase.hasNext()) {
    return carpetasBase.next();
  }

  const carpetaContenedora = obtenerCarpetaContenedoraDelMaestro();
  if (carpetaContenedora) {
    return carpetaContenedora.createFolder(CONFIG.NOMBRE_CARPETA_PADRONES);
  }

  return DriveApp.createFolder(CONFIG.NOMBRE_CARPETA_PADRONES);
}

function obtenerCarpetaContenedoraDelMaestro() {
  try {
    const archivoMaestro = DriveApp.getFileById(CONFIG.ID_HOJA_MAESTRO);
    const padres = archivoMaestro.getParents();
    return padres.hasNext() ? padres.next() : null;
  } catch (e) {
    return null;
  }
}
