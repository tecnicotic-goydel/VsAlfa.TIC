function listarCuadrillas() {
  return resolverCuadrillasConfiguradas().map(function(cuadrilla) {
    return {
      key: cuadrilla.key,
      label: cuadrilla.label,
      folderName: cuadrilla.folderName,
      folderId: cuadrilla.folderId,
      folderDisplayName: cuadrilla.folderDisplayName,
      disponible: cuadrilla.disponible
    };
  });
}

function procesarTodo() {
  return procesarLecturas();
}

function procesarCuadrilla(cuadrillaKey) {
  return procesarLecturas(cuadrillaKey ? [cuadrillaKey] : null);
}

function procesarLecturas(cuadrillaKeys) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.NOMBRE_HOJA_PADRON);

  if (!sheet) {
    throw new Error("No existe la hoja '" + CONFIG.NOMBRE_HOJA_PADRON + "'.");
  }

  asegurarEstructuraPadron(sheet);
  const valoresMaestro = sheet.getDataRange().getDisplayValues();

  if (valoresMaestro.length <= 1) {
    return {
      success: true,
      periodo: '',
      timestamp: obtenerMarcaTiempo(),
      resumen: {
        cuadrillasSolicitadas: 0,
        cuadrillasProcesadas: 0,
        archivosEncontrados: 0,
        procesados: 0,
        leidosSinCarga: 0,
        omitidos: 0,
        errores: 1
      },
      cuadrillas: [{
        key: 'general',
        label: 'General',
        folderName: '',
        folderId: '',
        folderDisplayName: '',
        disponible: false,
        archivosEncontrados: 0,
        procesados: 0,
        leidosSinCarga: 0,
        omitidos: 0,
        errores: 1,
        detalles: ['El padrón maestro no tiene registros para procesar.']
      }]
    };
  }

  const cabeceras = valoresMaestro[0].map(function(h) {
    return normalizarTexto(h).toUpperCase();
  });

  const col = {
    sum: cabeceras.indexOf('SUMINISTRO'),
    per: cabeceras.indexOf('PERIODO'),
    est: cabeceras.indexOf('ESTADO'),
    lecAntEat: cabeceras.indexOf('LEC_ANT_EAT'),
    lecAntEahp: cabeceras.indexOf('LEC_ANT_EAHP'),
    lecAntEafp: cabeceras.indexOf('LEC_ANT_EAFP'),
    lecAntEr: cabeceras.indexOf('LEC_ANT_ER'),
    lecAntPhp: cabeceras.indexOf('LEC_ANT_PHP'),
    lecAntPfp: cabeceras.indexOf('LEC_ANT_PFP'),
    eat: cabeceras.indexOf('EAT'),
    eahp: cabeceras.indexOf('EAHP'),
    eafp: cabeceras.indexOf('EAFP'),
    er: cabeceras.indexOf('ER'),
    php: cabeceras.indexOf('PHP'),
    pfp: cabeceras.indexOf('PFP'),
    fecha: cabeceras.indexOf('FECHA REGISTRO'),
    obs: cabeceras.indexOf('OBS'),
    origen: cabeceras.indexOf('ORIGEN_LECTURA'),
    archivoLectura: cabeceras.indexOf('ARCHIVO_LECTURA'),
    tipoArchivoLectura: cabeceras.indexOf('TIPO_ARCHIVO_LECTURA')
  };

  const periodoObjetivo = obtenerPeriodoActivoDesdeValores(valoresMaestro, col.per);
  const cuadrillasDisponibles = resolverCuadrillasConfiguradas();
  const cuadrillasSeleccionadas = filtrarCuadrillas(cuadrillasDisponibles, cuadrillaKeys);
  const resultados = [];
  const resumen = {
    cuadrillasSolicitadas: cuadrillasSeleccionadas.length,
    cuadrillasProcesadas: 0,
    archivosEncontrados: 0,
    procesados: 0,
    leidosSinCarga: 0,
    omitidos: 0,
    errores: 0
  };

  cuadrillasSeleccionadas.forEach(function(cuadrilla) {
    const resultadoCuadrilla = procesarArchivosDeCuadrilla(
      cuadrilla,
      sheet,
      valoresMaestro,
      col,
      periodoObjetivo
    );

    resultados.push(resultadoCuadrilla);
    resumen.archivosEncontrados += resultadoCuadrilla.archivosEncontrados;
    resumen.procesados += resultadoCuadrilla.procesados;
    resumen.leidosSinCarga += resultadoCuadrilla.leidosSinCarga;
    resumen.omitidos += resultadoCuadrilla.omitidos;
    resumen.errores += resultadoCuadrilla.errores;

    if (resultadoCuadrilla.disponible) {
      resumen.cuadrillasProcesadas++;
    }
  });

  return {
    success: true,
    periodo: periodoObjetivo,
    timestamp: obtenerMarcaTiempo(),
    resumen: resumen,
    cuadrillas: resultados
  };
}

function intentarActualizarFila(sheet, valores, col, datos, periodoObjetivo) {
  const sumBuscado = normalizarSuministro(datos.suministro);
  const archivoLectura = normalizarTexto(datos.archivoLectura);
  const tipoArchivoLectura = normalizarTexto(datos.tipoArchivoLectura);

  for (let i = 1; i < valores.length; i++) {
    const sumEnHoja = normalizarSuministro(valores[i][col.sum]);
    const perEnHoja = normalizarPeriodo(valores[i][col.per], '');
    const estadoEnHoja = normalizarTexto(valores[i][col.est]).toUpperCase();

    if (sumEnHoja === sumBuscado && perEnHoja === periodoObjetivo) {
      if (esEstadoFinalizado(estadoEnHoja)) {
        return { exito: false, motivo: 'Ya estaba marcado como ' + resolverEtiquetaEstadoFinal(estadoEnHoja) + '.' };
      }

      const fila = i + 1;
      escribirValorMedicionSiCorresponde(sheet, fila, col.lecAntEat, datos.lecAntEat);
      escribirValorMedicionSiCorresponde(sheet, fila, col.lecAntEahp, datos.lecAntEahp);
      escribirValorMedicionSiCorresponde(sheet, fila, col.lecAntEafp, datos.lecAntEafp);
      escribirValorMedicionSiCorresponde(sheet, fila, col.lecAntEr, datos.lecAntEr);
      escribirValorMedicionSiCorresponde(sheet, fila, col.lecAntPhp, datos.lecAntPhp);
      escribirValorMedicionSiCorresponde(sheet, fila, col.lecAntPfp, datos.lecAntPfp);
      if (col.eat !== -1) escribirValorMedicion(sheet.getRange(fila, col.eat + 1), datos.eat);
      if (col.eahp !== -1) escribirValorMedicion(sheet.getRange(fila, col.eahp + 1), datos.eahp);
      if (col.eafp !== -1) escribirValorMedicion(sheet.getRange(fila, col.eafp + 1), datos.eafp);
      if (col.er !== -1) escribirValorMedicion(sheet.getRange(fila, col.er + 1), datos.er);
      if (col.php !== -1) escribirValorMedicion(sheet.getRange(fila, col.php + 1), datos.php);
      if (col.pfp !== -1) escribirValorMedicion(sheet.getRange(fila, col.pfp + 1), datos.pfp);
      if (col.fecha !== -1) sheet.getRange(fila, col.fecha + 1).setValue(new Date());
      if (col.obs !== -1) sheet.getRange(fila, col.obs + 1).setValue(datos.obs);
      if (col.origen !== -1) {
        sheet.getRange(fila, col.origen + 1).setValue(CONFIG.ORIGEN_LECTURA_CAMPO);
        valores[i][col.origen] = CONFIG.ORIGEN_LECTURA_CAMPO;
      }
      if (col.archivoLectura !== -1 && archivoLectura) {
        sheet.getRange(fila, col.archivoLectura + 1).setValue(archivoLectura);
        valores[i][col.archivoLectura] = archivoLectura;
      }
      if (col.tipoArchivoLectura !== -1 && tipoArchivoLectura) {
        sheet.getRange(fila, col.tipoArchivoLectura + 1).setValue(tipoArchivoLectura);
        valores[i][col.tipoArchivoLectura] = tipoArchivoLectura;
      }

      sheet.getRange(fila, col.est + 1).setValue(CONFIG.ESTADO_COMPLETO);
      valores[i][col.est] = CONFIG.ESTADO_COMPLETO;
      return { exito: true };
    }
  }

  return {
    exito: false,
    motivo: 'No se halló el suministro ' + sumBuscado + ' para el período ' + periodoObjetivo + '.'
  };
}

function intentarMarcarFilaLeidaSinCarga(sheet, valores, col, datos, periodoObjetivo) {
  const sumBuscado = normalizarSuministro(datos.suministro);
  const observacion = normalizarTexto(datos.obs) || 'Archivo leído sin parámetros.';
  const tipoArchivo = normalizarTexto(datos.tipoArchivoLectura) || 'SIN CARGA';
  const archivoLectura = normalizarTexto(datos.archivoLectura) || '-';

  for (let i = 1; i < valores.length; i++) {
    const sumEnHoja = normalizarSuministro(valores[i][col.sum]);
    const perEnHoja = normalizarPeriodo(valores[i][col.per], '');
    const estadoEnHoja = normalizarTexto(valores[i][col.est]).toUpperCase();

    if (sumEnHoja !== sumBuscado || perEnHoja !== periodoObjetivo) {
      continue;
    }

    if (esEstadoFinalizado(estadoEnHoja)) {
      return { exito: false, motivo: 'Ya estaba marcado como ' + resolverEtiquetaEstadoFinal(estadoEnHoja) + '.' };
    }

    const fila = i + 1;
    if (col.fecha !== -1) {
      sheet.getRange(fila, col.fecha + 1).setValue(new Date());
      valores[i][col.fecha] = obtenerMarcaTiempo();
    }
    if (col.obs !== -1) {
      sheet.getRange(fila, col.obs + 1).setValue(observacion);
      valores[i][col.obs] = observacion;
    }
    if (col.origen !== -1) {
      sheet.getRange(fila, col.origen + 1).setValue(CONFIG.ORIGEN_LECTURA_CAMPO);
      valores[i][col.origen] = CONFIG.ORIGEN_LECTURA_CAMPO;
    }
    if (col.archivoLectura !== -1) {
      sheet.getRange(fila, col.archivoLectura + 1).setValue(archivoLectura);
      valores[i][col.archivoLectura] = archivoLectura;
    }
    if (col.tipoArchivoLectura !== -1) {
      sheet.getRange(fila, col.tipoArchivoLectura + 1).setValue(tipoArchivo);
      valores[i][col.tipoArchivoLectura] = tipoArchivo;
    }

    sheet.getRange(fila, col.est + 1).setValue(CONFIG.ESTADO_LEIDO_SIN_CARGA);
    valores[i][col.est] = CONFIG.ESTADO_LEIDO_SIN_CARGA;

    return {
      exito: true,
      mensaje: 'Leído sin parámetros.'
    };
  }

  return {
    exito: false,
    motivo: 'No se halló el suministro ' + sumBuscado + ' para el período ' + periodoObjetivo + '.'
  };
}

function escribirValorMedicion(range, valor) {
  const numero = parsearNumeroMedicion(valor, 0);
  const formato = resolverFormatoNumeroMedicion(valor, numero);

  range
    .setValue(numero)
    .setNumberFormat(formato || '0.############');
}

function escribirValorMedicionSiCorresponde(sheet, fila, indiceColumna, valor) {
  if (indiceColumna === -1 || valor === null || valor === undefined || valor === '') {
    return;
  }

  escribirValorMedicion(sheet.getRange(fila, indiceColumna + 1), valor);
}

function obtenerCarpetaProcesadosPeriodo(periodoObjetivo, crearSiNoExiste) {
  const periodo = normalizarPeriodo(periodoObjetivo, '');
  if (!periodo) {
    throw new Error('No se pudo resolver el período para la carpeta de procesados.');
  }

  const nombreCarpeta = formatearNombreCarpetaProcesadosPeriodo(periodo);
  const carpetaPadre = DriveApp.getFolderById(CONFIG.ID_CARPETA_PROCESADOS);
  const subcarpetas = carpetaPadre.getFoldersByName(nombreCarpeta);

  if (subcarpetas.hasNext()) {
    return subcarpetas.next();
  }

  if (crearSiNoExiste === false) {
    return null;
  }

  return carpetaPadre.createFolder(nombreCarpeta);
}

function formatearNombreCarpetaProcesadosPeriodo(periodoObjetivo) {
  const periodo = normalizarPeriodo(periodoObjetivo, '');
  const meses = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const anio = periodo.slice(0, 4);
  const mesIndice = parseInt(periodo.slice(4, 6), 10) - 1;
  const mesNombre = meses[mesIndice] || periodo.slice(4, 6);
  return mesNombre + '_' + anio;
}

function construirNombreCarpetaProcesadosCuadrilla(cuadrilla) {
  const prefijo = normalizarTexto(cuadrilla && cuadrilla.folderName) || '00';
  const etiqueta = normalizarTexto(cuadrilla && cuadrilla.label) || 'Cuadrilla';
  return prefijo + '_' + etiqueta;
}

function resolverCarpetaProcesadosCuadrilla(periodoObjetivo, cuadrilla, crearSiNoExiste) {
  const carpetaPeriodo = obtenerCarpetaProcesadosPeriodo(periodoObjetivo, crearSiNoExiste);
  if (!carpetaPeriodo) {
    return null;
  }

  const nombreCarpeta = construirNombreCarpetaProcesadosCuadrilla(cuadrilla);
  const subcarpetas = carpetaPeriodo.getFoldersByName(nombreCarpeta);

  if (subcarpetas.hasNext()) {
    return subcarpetas.next();
  }

  if (crearSiNoExiste === false) {
    return null;
  }

  return carpetaPeriodo.createFolder(nombreCarpeta);
}

function construirNombreCarpetaNoProcesadosCategoria(categoria) {
  const clave = normalizarTexto(categoria).toUpperCase();

  if (clave === 'ERRORES') {
    return 'Errores';
  }

  if (clave === 'OMITIDOS') {
    return 'Omitidos';
  }

  return 'No_Leidos';
}

function resolverCarpetaNoProcesadosCuadrilla(periodoObjetivo, cuadrilla, categoria, crearSiNoExiste) {
  const carpetaCuadrilla = resolverCarpetaProcesadosCuadrilla(periodoObjetivo, cuadrilla, crearSiNoExiste);
  if (!carpetaCuadrilla) {
    return null;
  }

  const nombreCarpeta = construirNombreCarpetaNoProcesadosCategoria(categoria);
  const subcarpetas = carpetaCuadrilla.getFoldersByName(nombreCarpeta);

  if (subcarpetas.hasNext()) {
    return subcarpetas.next();
  }

  if (crearSiNoExiste === false) {
    return null;
  }

  return carpetaCuadrilla.createFolder(nombreCarpeta);
}

function moverArchivoANoProcesados(archivo, periodoObjetivo, cuadrilla, categoria) {
  const carpetaDestino = resolverCarpetaNoProcesadosCuadrilla(periodoObjetivo, cuadrilla, categoria, true);
  archivo.moveTo(carpetaDestino);
  return carpetaDestino.getName();
}

function procesarArchivosDeCuadrilla(cuadrilla, sheet, valoresMaestro, col, periodoObjetivo) {
  const resultado = {
    key: cuadrilla.key,
    label: cuadrilla.label,
    folderName: cuadrilla.folderName,
    folderId: cuadrilla.folderId,
    folderDisplayName: cuadrilla.folderDisplayName,
    disponible: cuadrilla.disponible,
    archivosEncontrados: 0,
    procesados: 0,
    leidosSinCarga: 0,
    omitidos: 0,
    errores: 0,
    detalles: []
  };

  if (!cuadrilla.disponible || !cuadrilla.folder) {
    resultado.errores = 1;
    resultado.detalles.push('No se encontró la carpeta configurada para esta cuadrilla.');
    return resultado;
  }

  const carpetaDestino = resolverCarpetaProcesadosCuadrilla(periodoObjetivo, cuadrilla, true);
  const archivos = cuadrilla.folder.getFiles();

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    resultado.archivosEncontrados++;

    try {
      const datos = extraerDatosDeArchivo(archivo);

      if (!datos || !datos.suministro) {
        resultado.errores++;
        const carpetaNoLeidos = moverArchivoANoProcesados(archivo, periodoObjetivo, cuadrilla, 'NO_LEIDOS');
        resultado.detalles.push(archivo.getName() + ': No se halló suministro. Movido a ' + carpetaNoLeidos + '.');
        continue;
      }

      if (datos.sinCargaEspecial) {
        const actualizacionSinCarga = intentarMarcarFilaLeidaSinCarga(
          sheet,
          valoresMaestro,
          col,
          datos,
          periodoObjetivo
        );

        if (actualizacionSinCarga.exito) {
          archivo.moveTo(carpetaDestino);
          resultado.procesados++;
          resultado.leidosSinCarga++;
          resultado.detalles.push(archivo.getName() + ': ' + (actualizacionSinCarga.mensaje || 'Leido sin parametros.'));
        } else {
          resultado.omitidos++;
          const carpetaOmitidos = moverArchivoANoProcesados(archivo, periodoObjetivo, cuadrilla, 'OMITIDOS');
          resultado.detalles.push(archivo.getName() + ': ' + actualizacionSinCarga.motivo + ' Movido a ' + carpetaOmitidos + '.');
        }
        continue;
      }

      if (!tieneLecturasExtraidasValidas(datos)) {
        resultado.omitidos++;
        const carpetaNoLeidos = moverArchivoANoProcesados(archivo, periodoObjetivo, cuadrilla, 'NO_LEIDOS');
        resultado.detalles.push(archivo.getName() + ': No se encontraron lecturas utiles en el archivo. Movido a ' + carpetaNoLeidos + '.');
        continue;
      }

      const actualizacion = intentarActualizarFila(
        sheet,
        valoresMaestro,
        col,
        datos,
        periodoObjetivo
      );

      if (actualizacion.exito) {
        archivo.moveTo(carpetaDestino);
        resultado.procesados++;
        resultado.detalles.push(archivo.getName() + ': Procesado correctamente.');
      } else {
        resultado.omitidos++;
        const carpetaOmitidos = moverArchivoANoProcesados(archivo, periodoObjetivo, cuadrilla, 'OMITIDOS');
        resultado.detalles.push(archivo.getName() + ': ' + actualizacion.motivo + ' Movido a ' + carpetaOmitidos + '.');
      }
    } catch (error) {
      resultado.errores++;
      try {
        const carpetaErrores = moverArchivoANoProcesados(archivo, periodoObjetivo, cuadrilla, 'ERRORES');
        resultado.detalles.push(archivo.getName() + ': Error durante el procesamiento. ' + error.message + ' Movido a ' + carpetaErrores + '.');
      } catch (moveError) {
        resultado.detalles.push(archivo.getName() + ': Error durante el procesamiento. ' + error.message + ' Adicionalmente no se pudo mover a Errores: ' + moveError.message);
      }
    }
  }

  if (resultado.archivosEncontrados === 0) {
    resultado.detalles.push('No hay archivos pendientes en la carpeta.');
  }

  return resultado;
}

function descargarZipSinCargaCuadrilla(cuadrillaKey, periodoFiltro) {
  const periodoObjetivo = normalizarPeriodo(periodoFiltro, obtenerPeriodoActual());
  if (!periodoObjetivo) {
    throw new Error('Seleccione el período que desea descargar.');
  }

  const cuadrillas = filtrarCuadrillas(resolverCuadrillasConfiguradas(), [cuadrillaKey]);
  const cuadrilla = cuadrillas.length ? cuadrillas[0] : null;
  if (!cuadrilla) {
    throw new Error('No se pudo resolver la cuadrilla solicitada.');
  }

  const carpetaCuadrilla = resolverCarpetaProcesadosCuadrilla(periodoObjetivo, cuadrilla, false);
  if (!carpetaCuadrilla) {
    throw new Error('No existe carpeta de procesados para la cuadrilla en el período seleccionado.');
  }

  const archivosEspeciales = listarArchivosSinCargaEspeciales(carpetaCuadrilla);
  if (!archivosEspeciales.length) {
    throw new Error('No existen archivos .msr o .RP3 procesados para esa cuadrilla en el período seleccionado.');
  }

  const resumen = construirResumenSinCargaCuadrilla(periodoObjetivo, cuadrilla, archivosEspeciales);
  const archivoZip = crearZipSinCargaCuadrilla(periodoObjetivo, cuadrilla, carpetaCuadrilla, archivosEspeciales, resumen);

  return {
    success: true,
    periodo: periodoObjetivo,
    cuadrilla: cuadrilla.label,
    archivos: archivosEspeciales.length,
    suministros: resumen.itemsManual.length,
    nombre: archivoZip.getName(),
    url: archivoZip.getUrl(),
    urlDescarga: 'https://drive.google.com/uc?export=download&id=' + archivoZip.getId(),
    mensaje: 'ZIP consolidado generado correctamente.'
  };
}

function listarArchivosSinCargaEspeciales(carpeta) {
  const archivos = [];
  const iterador = carpeta.getFiles();

  while (iterador.hasNext()) {
    const archivo = iterador.next();
    if (esArchivoSinCargaEspecial(normalizarTexto(archivo.getName()).toLowerCase())) {
      archivos.push(archivo);
    }
  }

  archivos.sort(function(a, b) {
    return a.getName().localeCompare(b.getName());
  });

  return archivos;
}

function construirResumenSinCargaCuadrilla(periodoObjetivo, cuadrilla, archivosEspeciales) {
  const resumenMensual = getResumenMensual(periodoObjetivo);
  const detallePorSuministro = {};
  const filasIndice = [];
  const itemsManual = [];
  const vistos = {};

  (resumenMensual.suministros || []).forEach(function(item) {
    detallePorSuministro[normalizarSuministro(item.suministro)] = item;
  });

  archivosEspeciales.forEach(function(archivo) {
    const datos = extraerDatosDeArchivo(archivo);
    const suministro = normalizarTexto(datos && datos.suministro);
    const suministroClave = normalizarSuministro(suministro);
    const tipo = normalizarTexto(datos && datos.tipoArchivoLectura) || resolverTipoArchivoSinCargaEspecial(archivo.getName());
    const observacion = normalizarTexto(datos && datos.obs) || ('Archivo ' + tipo + ' leído sin parámetros - ' + archivo.getName());
    const detalleSuministro = detallePorSuministro[suministroClave] || null;

    filasIndice.push([
      suministro || '-',
      cuadrilla.label,
      archivo.getName(),
      tipo,
      formatearPeriodo(periodoObjetivo),
      CONFIG.ESTADO_LEIDO_SIN_CARGA,
      observacion
    ]);

    if (!suministroClave || vistos[suministroClave]) {
      return;
    }

    if (detalleSuministro && esEstadoFinalizado(detalleSuministro.estado)) {
      vistos[suministroClave] = true;
      return;
    }

    vistos[suministroClave] = true;
    itemsManual.push(detalleSuministro || {
      suministro: suministro,
      campos: []
    });
  });

  itemsManual.sort(function(a, b) {
    return normalizarSuministro(a.suministro).localeCompare(normalizarSuministro(b.suministro));
  });

  return {
    itemsManual: itemsManual,
    filasIndice: filasIndice
  };
}

function crearZipSinCargaCuadrilla(periodoObjetivo, cuadrilla, carpetaCuadrilla, archivosEspeciales, resumen) {
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'GMT-5',
    'yyyyMMdd_HHmmss'
  );
  const nombreBase = 'Sin_Parámetros_' + periodoObjetivo + '_' + (cuadrilla.folderName || cuadrilla.key) + '_' + timestamp;
  const blobResumen = crearExcelResumenSinCargaCuadrilla(nombreBase, resumen.itemsManual, resumen.filasIndice);
  const blobs = [blobResumen];

  archivosEspeciales.forEach(function(archivo) {
    blobs.push(archivo.getBlob().copyBlob().setName(archivo.getName()));
  });

  const zipBlob = Utilities.zip(blobs, nombreBase + '.zip');
  return carpetaCuadrilla.createFile(zipBlob.setName(nombreBase + '.zip'));
}

function crearExcelResumenSinCargaCuadrilla(nombreBase, itemsManual, filasIndice) {
  const spreadsheet = SpreadsheetApp.create(nombreBase);

  try {
    const hojaDatos = spreadsheet.getSheets()[0];
    hojaDatos.setName(NOMBRE_HOJA_PLANTILLA_CARGA_MANUAL);
    cargarContenidoPlantillaManual(hojaDatos, itemsManual);

    const hojaIndice = spreadsheet.insertSheet('Indice');
    cargarContenidoIndiceSinCarga(hojaIndice, filasIndice);

    return exportarSpreadsheetAXlsx(spreadsheet.getId(), nombreBase + '.xlsx').setName(nombreBase + '.xlsx');
  } finally {
    try {
      DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
    } catch (error) {
      // No interrumpir la descarga si la limpieza falla.
    }
  }
}

function cargarContenidoIndiceSinCarga(sheet, filasIndice) {
  const headers = ['Suministro', 'Cuadrilla', 'Archivo', 'Tipo', 'Período', 'Estado', 'Observación'];
  const filas = Array.isArray(filasIndice) ? filasIndice : [];

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  if (filas.length) {
    sheet.getRange(2, 1, filas.length, headers.length).setValues(filas);
  }

  sheet.autoResizeColumns(1, headers.length);
}

function tieneLecturasExtraidasValidas(datos) {
  const mediciones = [
    datos.eat, datos.eahp, datos.eafp, datos.er, datos.php, datos.pfp,
    datos.lecAntEat, datos.lecAntEahp, datos.lecAntEafp,
    datos.lecAntEr, datos.lecAntPhp, datos.lecAntPfp
  ];

  for (let i = 0; i < mediciones.length; i++) {
    const valor = parsearNumeroMedicion(mediciones[i], null);
    if (typeof valor === 'number' && !isNaN(valor) && Math.abs(valor) > 0) {
      return true;
    }
  }

  return false;
}

function filtrarCuadrillas(cuadrillasDisponibles, cuadrillaKeys) {
  if (!Array.isArray(cuadrillaKeys) || cuadrillaKeys.length === 0) {
    return cuadrillasDisponibles;
  }

  const permitidas = {};
  cuadrillaKeys.forEach(function(key) {
    permitidas[key] = true;
  });

  return cuadrillasDisponibles.filter(function(cuadrilla) {
    return permitidas[cuadrilla.key];
  });
}

function resolverCuadrillasConfiguradas() {
  const carpetaBase = DriveApp.getFolderById(CONFIG.ID_CARPETA_CUADRILLAS);
  const subcarpetas = obtenerSubcarpetasCuadrillas(carpetaBase);

  return CONFIG.CUADRILLAS.map(function(config, index) {
    const folder = resolverCarpetaCuadrilla(config, subcarpetas, index);

    return {
      key: config.key,
      label: config.label,
      folderName: config.folderName,
      folderId: folder ? folder.getId() : config.folderId,
      folderDisplayName: folder ? folder.getName() : (config.folderName || config.label),
      folder: folder,
      disponible: Boolean(folder)
    };
  });
}

function obtenerSubcarpetasCuadrillas(carpetaBase) {
  const folders = [];
  const iterator = carpetaBase.getFolders();

  while (iterator.hasNext()) {
    folders.push(iterator.next());
  }

  folders.sort(function(a, b) {
    return a.getName().localeCompare(b.getName());
  });

  return folders;
}

function resolverCarpetaCuadrilla(config, subcarpetas, index) {
  if (config.folderId) {
    try {
      return DriveApp.getFolderById(config.folderId);
    } catch (error) {
      return null;
    }
  }

  if (config.folderName) {
    const porNombre = subcarpetas.find(function(folder) {
      return folder.getName() === config.folderName;
    });
    if (porNombre) {
      return porNombre;
    }

    const porPrefijo = subcarpetas.find(function(folder) {
      return folder.getName().indexOf(config.folderName) === 0;
    });
    if (porPrefijo) {
      return porPrefijo;
    }
  }

  return subcarpetas[index] || null;
}
