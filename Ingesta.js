function ingestarNuevoPadron(fileData, periodoCarga) {
  const ID_MI_EXCEL = CONFIG.ID_HOJA_MAESTRO;
  const NOMBRE_HOJA = CONFIG.NOMBRE_HOJA_PADRON;

  const MAPEO_ALIAS = {
    SUMINISTRO: ['SUMINISTRO', 'CODIGO', 'ID', 'SUMI'],
    CARTERA: ['CARTERA'],
    NOMBRE_CLIENTE: ['NOMBRE', 'CLIENTE', 'RAZON SOCIAL'],
    DIRECCION: ['DIRECCION', 'DIR', 'UBICACION'],
    SERIE_FAB: ['SERIE_FAB', 'SERIE', 'SERIE_MEDIDOR', 'NRO_SERIE'],
    FACTOR: ['FACTOR', 'FACTOR_K'],
    NOMB_RUTA: ['NOMB_RUTA', 'NOMBRE_RUTA', 'DESCRIPCION_RUTA'],
    PERIODO: ['PERIODO', 'MES', 'ANIO_MES', 'ANO_MES', 'CICLO', 'PFACTURA', 'P_FACTURA'],
    LEC_ANT_EAT: ['LEC_ANT_EAT', 'LECTURA_ANTERIOR_EAT', 'EAT_ANT'],
    LEC_ANT_EAHP: ['LEC_ANT_EAHP', 'LECTURA_ANTERIOR_EAHP', 'EAHP_ANT'],
    LEC_ANT_EAFP: ['LEC_ANT_EAFP', 'LECTURA_ANTERIOR_EAFP', 'EAFP_ANT'],
    LEC_ANT_ER: ['LEC_ANT_ER', 'LECTURA_ANTERIOR_ER', 'ER_ANT'],
    LEC_ANT_PHP: ['LEC_ANT_PHP', 'LECTURA_ANTERIOR_PHP', 'PHP_ANT'],
    LEC_ANT_PFP: ['LEC_ANT_PFP', 'LECTURA_ANTERIOR_PFP', 'PFP_ANT'],
    CUADRILLA: ['CUADRILLA', 'EQUIPO', 'GRUPO_TRABAJO'],
    ORIGEN_LECTURA: ['ORIGEN_LECTURA', 'ORIGEN', 'TIPO_LECTURA_ORIGEN'],
    ESTADO_LECTURA: ['ESTADO_LECTURA', 'ESTADO LECTURA', 'ESTADO_MEDIDOR', 'ESTADO MEDIDOR'],
    TARIFA: ['TARIFA', 'TIPO_TARIFA'],
    ESTADO: ['ESTADO', 'STATUS']
  };

  let tempFile;

  try {
    const periodoSeleccionado = normalizarPeriodo(periodoCarga, '');
    if (!periodoSeleccionado) {
      throw new Error('Debe seleccionar el mes del padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n antes de cargar el archivo.');
    }

    const ss = SpreadsheetApp.openById(ID_MI_EXCEL);
    let sheet = ss.getSheetByName(NOMBRE_HOJA);
    if (!sheet) {
      sheet = ss.insertSheet(NOMBRE_HOJA);
    }
    const estructuraPadron = asegurarEstructuraPadron(sheet);
    const columnasRequeridas = estructuraPadron.headersUpper.slice();

    const blob = Utilities.newBlob(
      Utilities.base64Decode(fileData.data),
      fileData.mimeType,
      fileData.fileName
    );
    const resource = { name: 'temp_ingesta', mimeType: MimeType.GOOGLE_SHEETS };
    tempFile = Drive.Files.create
      ? Drive.Files.create(resource, blob)
      : Drive.Files.insert({ title: resource.name, mimeType: resource.mimeType }, blob);

    const tempSS = SpreadsheetApp.openById(tempFile.id);
    const tempSheet = tempSS.getSheets()[0];
    const matrizOrigen = tempSheet.getDataRange().getValues();

    if (matrizOrigen.length === 0) {
      throw new Error('El archivo no contiene datos.');
    }

    const cabeceraOrigen = matrizOrigen[0].map(function(celda) {
      return normalizarTexto(celda).toUpperCase();
    });
    const diagnosticoPeriodo = detectarPeriodoPadron(cabeceraOrigen, matrizOrigen, MAPEO_ALIAS, fileData.fileName);
    validarPeriodoPadronSeleccionado(periodoSeleccionado, diagnosticoPeriodo);
    const periodoDetectado = diagnosticoPeriodo.periodo;
    const periodoAplicado = periodoSeleccionado;
    const transformacion = transformarFilasPadronOrigen(
      matrizOrigen,
      cabeceraOrigen,
      columnasRequeridas,
      MAPEO_ALIAS,
      periodoAplicado,
      fileData.fileName
    );
    const filasUnicas = transformacion.filasUnicas;

    if (!filasUnicas.length) {
      throw new Error('El archivo no contiene suministros vÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¡lidos para cargar.');
    }

    if (transformacion.duplicadosDentroArchivo.length) {
      throw new Error(
        construirMensajeCoincidenciasArchivo(periodoAplicado, transformacion.duplicadosDentroArchivo)
      );
    }

    const estadoPadron = obtenerEstadoActualPadron(sheet);
    const validacionCoincidencias = validarCoincidenciasPadronExistente(
      estadoPadron,
      filasUnicas,
      periodoAplicado
    );

    if (validacionCoincidencias.coincidencias.length) {
      throw new Error(
        construirMensajeCoincidenciasPadron(periodoAplicado, validacionCoincidencias.coincidencias)
      );
    }

    const guardadoDrive = guardarPadronesEnDrive(periodoSeleccionado, [fileData]);
    if (!guardadoDrive.success) {
      throw new Error(guardadoDrive.mensaje);
    }

    if (estadoPadron.hasHeaders && !estadoPadron.rows.length && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }

    const filaInicio = estadoPadron.hasHeaders && !estadoPadron.rows.length
      ? 2
      : Math.max(sheet.getLastRow(), 1) + 1;
    sheet.getRange(filaInicio, 1, filasUnicas.length, columnasRequeridas.length)
      .setValues(filasUnicas);

    return {
      success: true,
      mensaje: transformacion.mensaje || 'Padrón cargado correctamente.',
      archivo: fileData.fileName,
      periodo: periodoAplicado,
      periodoSeleccionado: periodoSeleccionado,
      periodoDetectado: periodoDetectado,
      fuentePeriodoDetectado: diagnosticoPeriodo.fuente,
      registros: filasUnicas.length,
      registrosAgregados: filasUnicas.length,
      filasOrigen: transformacion.filasOrigen || 0,
      modoTransformacion: transformacion.modo || 'directo',
      totalPeriodo: validacionCoincidencias.totalPeriodo + filasUnicas.length,
      hoja: NOMBRE_HOJA,
      carpeta: guardadoDrive.carpeta,
      timestamp: obtenerMarcaTiempo()
    };
  } catch (e) {
    throw new Error('Error en Ingesta: ' + e.message);
  } finally {
    if (tempFile && tempFile.id) {
      Drive.Files.remove(tempFile.id);
    }
  }
}

function transformarFilasPadronOrigen(matrizOrigen, cabeceraOrigen, columnasRequeridas, mapeoAlias, periodoAplicado, nombreArchivo) {
  if (esArchivoPadronPorConceptos(cabeceraOrigen)) {
    return consolidarFilasPadronPorConceptos(
      matrizOrigen,
      cabeceraOrigen,
      columnasRequeridas,
      mapeoAlias,
      periodoAplicado,
      nombreArchivo
    );
  }

  return construirFilasPadronDirectas(
    matrizOrigen,
    cabeceraOrigen,
    columnasRequeridas,
    mapeoAlias,
    periodoAplicado,
    nombreArchivo
  );
}

function construirFilasPadronDirectas(matrizOrigen, cabeceraOrigen, columnasRequeridas, mapeoAlias, periodoAplicado, nombreArchivo) {
  const filas = Array.isArray(matrizOrigen) ? matrizOrigen : [];
  const filasUnicas = [];
  const indicePorClave = {};
  const duplicadosDentroArchivo = {};
  const indicesOrigen = construirIndicesOrigenPorDestino(cabeceraOrigen, columnasRequeridas, mapeoAlias);
  const indiceSuministro = columnasRequeridas.indexOf('SUMINISTRO');
  const indicePeriodo = columnasRequeridas.indexOf('PERIODO');

  for (let i = 1; i < filas.length; i++) {
    const filaNueva = construirFilaPadronDesdeOrigen(
      filas[i],
      columnasRequeridas,
      indicesOrigen,
      periodoAplicado,
      nombreArchivo
    );

    if (indiceSuministro === -1 || !normalizarTexto(filaNueva[indiceSuministro])) {
      continue;
    }

    const claveFila = construirClaveSuministroPeriodo(
      filaNueva[indiceSuministro],
      (indicePeriodo === -1 ? '' : filaNueva[indicePeriodo]) || periodoAplicado,
      'fila_' + i
    );

    if (Object.prototype.hasOwnProperty.call(indicePorClave, claveFila)) {
      duplicadosDentroArchivo[normalizarSuministro(filaNueva[indiceSuministro])] = true;
      continue;
    }

    indicePorClave[claveFila] = filasUnicas.length;
    filasUnicas.push(filaNueva);
  }

  return {
    filasUnicas: filasUnicas,
    duplicadosDentroArchivo: Object.keys(duplicadosDentroArchivo).sort(),
    filasOrigen: Math.max(filas.length - 1, 0),
    modo: 'directo',
    mensaje: 'Padrón cargado correctamente.'
  };
}

function consolidarFilasPadronPorConceptos(matrizOrigen, cabeceraOrigen, columnasRequeridas, mapeoAlias, periodoAplicado, nombreArchivo) {
  const filas = Array.isArray(matrizOrigen) ? matrizOrigen : [];
  const filasUnicas = [];
  const indicePorClave = {};
  const conflictos = {};
  const indicesOrigen = construirIndicesOrigenPorDestino(cabeceraOrigen, columnasRequeridas, mapeoAlias);
  const indiceSuministro = columnasRequeridas.indexOf('SUMINISTRO');
  const indicePeriodoPadron = columnasRequeridas.indexOf('PERIODO');
  const indiceConcepto = buscarIndiceCabeceraPadron(cabeceraOrigen, ['CONCEPTO']);
  const indiceConceptoDetalle = buscarIndiceCabeceraPadron(cabeceraOrigen, ['CONCEPTO_D', 'CONCEPTO DETALLE', 'CONCEPTODETALLE']);
  const indiceLecturaAnterior = buscarIndiceCabeceraPadron(cabeceraOrigen, ['LECANT', 'LEC_ANT', 'LECTURA ANTERIOR', 'LECTURA_ANTERIOR']);
  const indicesLecturas = {
    LEC_ANT_EAT: columnasRequeridas.indexOf('LEC_ANT_EAT'),
    LEC_ANT_EAHP: columnasRequeridas.indexOf('LEC_ANT_EAHP'),
    LEC_ANT_EAFP: columnasRequeridas.indexOf('LEC_ANT_EAFP'),
    LEC_ANT_ER: columnasRequeridas.indexOf('LEC_ANT_ER'),
    LEC_ANT_PHP: columnasRequeridas.indexOf('LEC_ANT_PHP'),
    LEC_ANT_PFP: columnasRequeridas.indexOf('LEC_ANT_PFP')
  };

  for (let i = 1; i < filas.length; i++) {
    const filaOrigen = filas[i];
    const filaBase = construirFilaPadronDesdeOrigen(
      filaOrigen,
      columnasRequeridas,
      indicesOrigen,
      periodoAplicado,
      nombreArchivo
    );

    if (indiceSuministro === -1 || !normalizarTexto(filaBase[indiceSuministro])) {
      continue;
    }

    const claveFila = construirClaveSuministroPeriodo(
      filaBase[indiceSuministro],
      (indicePeriodoPadron === -1 ? '' : filaBase[indicePeriodoPadron]) || periodoAplicado,
      'fila_' + i
    );

    if (!Object.prototype.hasOwnProperty.call(indicePorClave, claveFila)) {
      indicePorClave[claveFila] = filasUnicas.length;
      filasUnicas.push(filaBase);
    } else {
      combinarFilaPadronBase(filasUnicas[indicePorClave[claveFila]], filaBase);
    }

    const filaConsolidada = filasUnicas[indicePorClave[claveFila]];
    const campoLectura = resolverCampoLecturaPorConcepto(
      filaOrigen,
      indiceConcepto,
      indiceConceptoDetalle
    );
    const valorLectura = indiceLecturaAnterior === -1 ? '' : filaOrigen[indiceLecturaAnterior];

    if (!campoLectura || indicesLecturas[campoLectura] === -1) {
      conflictos[normalizarSuministro(filaConsolidada[0])] = true;
      continue;
    }

    asignarValorConsolidadoPadron(
      filaConsolidada,
      indicesLecturas[campoLectura],
      valorLectura,
      conflictos
    );
  }

  return {
    filasUnicas: filasUnicas,
    duplicadosDentroArchivo: Object.keys(conflictos).sort(),
    filasOrigen: Math.max(filas.length - 1, 0),
    modo: 'conceptos',
    mensaje: 'Padrón cargado correctamente. El archivo fue consolidado a una fila por suministro y período.'
  };
}

function construirIndicesOrigenPorDestino(cabeceraOrigen, columnasRequeridas, mapeoAlias) {
  const indices = {};
  const columnas = Array.isArray(columnasRequeridas) ? columnasRequeridas : [];

  columnas.forEach(function(colDestino) {
    const aliasPosibles = (mapeoAlias[colDestino] || [colDestino]).slice();
    if (aliasPosibles.indexOf(colDestino) === -1) {
      aliasPosibles.unshift(colDestino);
    }
    if (colDestino === 'SERIE_FAB') {
      aliasPosibles.push('SERIEFAB');
    }
    if (colDestino === 'NOMB_RUTA') {
      aliasPosibles.push('NOMBRUTA');
    }
    indices[colDestino] = buscarIndiceCabeceraPadron(cabeceraOrigen, aliasPosibles);
  });

  return indices;
}

function construirFilaPadronDesdeOrigen(filaOrigen, columnasRequeridas, indicesOrigen, periodoAplicado, nombreArchivo) {
  const filaNueva = [];
  const fila = Array.isArray(filaOrigen) ? filaOrigen : [];
  const columnas = Array.isArray(columnasRequeridas) ? columnasRequeridas : [];
  const indices = indicesOrigen || {};

  columnas.forEach(function(colDestino) {
    const indiceEncontrado = Object.prototype.hasOwnProperty.call(indices, colDestino)
      ? indices[colDestino]
      : -1;

    if (indiceEncontrado !== -1) {
      let valor = fila[indiceEncontrado];
      if (colDestino === 'PERIODO') {
        valor = periodoAplicado;
      }
      filaNueva.push(valor);
      return;
    }

    if (colDestino === 'PERIODO') {
      filaNueva.push(periodoAplicado);
    } else if (colDestino === 'ESTADO') {
      filaNueva.push(CONFIG.ESTADO_PENDIENTE);
    } else if (colDestino === 'FUENTE_ARCHIVO') {
      filaNueva.push('Excel ENOSA');
    } else if (colDestino === 'NOMBRE_ARCHIVO') {
      filaNueva.push(nombreArchivo);
    } else {
      filaNueva.push('');
    }
  });
  const indicePeriodo = columnas.indexOf('PERIODO');
  const indiceEstado = columnas.indexOf('ESTADO');
  const indiceFuente = columnas.indexOf('FUENTE_ARCHIVO');
  const indiceNombreArchivo = columnas.indexOf('NOMBRE_ARCHIVO');

  if (indicePeriodo !== -1) {
    filaNueva[indicePeriodo] = filaNueva[indicePeriodo] || periodoAplicado;
  }
  if (indiceEstado !== -1) {
    filaNueva[indiceEstado] = filaNueva[indiceEstado] || CONFIG.ESTADO_PENDIENTE;
  }
  if (indiceFuente !== -1) {
    filaNueva[indiceFuente] = filaNueva[indiceFuente] || 'Excel ENOSA';
  }
  if (indiceNombreArchivo !== -1) {
    filaNueva[indiceNombreArchivo] = filaNueva[indiceNombreArchivo] || nombreArchivo;
  }

  return filaNueva;
}

function obtenerColumnasPadronRequeridas() {
  return (CONFIG.COLUMNAS_PADRON_REQUERIDAS || []).slice();
}

function asegurarEstructuraPadron(sheet, columnasEsperadas) {
  const hoja = sheet;
  const columnasBase = (Array.isArray(columnasEsperadas) && columnasEsperadas.length
    ? columnasEsperadas
    : obtenerColumnasPadronRequeridas()).slice();

  if (!hoja) {
    throw new Error("No existe la hoja '" + CONFIG.NOMBRE_HOJA_PADRON + "'.");
  }

  const lastRow = hoja.getLastRow();
  const lastColumn = hoja.getLastColumn();

  if (!lastRow || !lastColumn) {
    hoja.clearContents();
    hoja.getRange(1, 1, 1, columnasBase.length).setValues([columnasBase]);
    return {
      headers: columnasBase.slice(),
      headersUpper: columnasBase.map(function(columna) {
        return normalizarTexto(columna).toUpperCase();
      }),
      added: columnasBase.slice()
    };
  }

  const headers = hoja.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(valor) {
    return normalizarTexto(valor);
  });
  const headersUpper = headers.map(function(valor) {
    return valor.toUpperCase();
  });
  const faltantes = [];

  columnasBase.forEach(function(columna) {
    const header = normalizarTexto(columna).toUpperCase();
    if (headersUpper.indexOf(header) === -1) {
      faltantes.push(columna);
      headers.push(columna);
      headersUpper.push(header);
    }
  });

  if (faltantes.length) {
    hoja.getRange(1, lastColumn + 1, 1, faltantes.length).setValues([faltantes]);
  }

  return {
    headers: headers,
    headersUpper: headersUpper,
    added: faltantes
  };
}

function combinarFilaPadronBase(destino, candidata) {
  const filaDestino = Array.isArray(destino) ? destino : [];
  const filaCandidata = Array.isArray(candidata) ? candidata : [];
  const longitud = Math.max(filaDestino.length, filaCandidata.length);

  for (let i = 0; i < longitud; i++) {
    if (!esValorInformado(filaDestino[i]) && esValorInformado(filaCandidata[i])) {
      filaDestino[i] = filaCandidata[i];
    }
  }
}

function resolverCampoLecturaPorConcepto(filaOrigen, indiceConcepto, indiceConceptoDetalle) {
  const fila = Array.isArray(filaOrigen) ? filaOrigen : [];
  const concepto = indiceConcepto === -1 ? '' : normalizarCabeceraPadron(fila[indiceConcepto]);
  const conceptoDetalle = indiceConceptoDetalle === -1 ? '' : normalizarCabeceraPadron(fila[indiceConceptoDetalle]);
  const mapaConceptos = {
    '1': 'LEC_ANT_EAT',
    '2': 'LEC_ANT_EAHP',
    '3': 'LEC_ANT_EAFP',
    '4': 'LEC_ANT_ER',
    '5': 'LEC_ANT_PHP',
    '6': 'LEC_ANT_PFP',
    'ENERGIAACTIVATOTAL': 'LEC_ANT_EAT',
    'ENERGIAACTIVAHORAPUNTA': 'LEC_ANT_EAHP',
    'ENERGIAACTIVAFUERAPUNTA': 'LEC_ANT_EAFP',
    'ENERGIAACTIVAFUERADEPUNTA': 'LEC_ANT_EAFP',
    'ENERGIAREACTIVA': 'LEC_ANT_ER',
    'POTENCIAHORAPUNTA': 'LEC_ANT_PHP',
    'POTENCIAFUERAPUNTA': 'LEC_ANT_PFP',
    'POTENCIAFUERADEPUNTA': 'LEC_ANT_PFP'
  };

  return mapaConceptos[concepto] || mapaConceptos[conceptoDetalle] || '';
}

function asignarValorConsolidadoPadron(filaConsolidada, indiceDestino, valor, conflictos) {
  const fila = Array.isArray(filaConsolidada) ? filaConsolidada : [];
  if (!fila.length || indiceDestino < 0) {
    return;
  }

  if (!esValorInformado(valor)) {
    return;
  }

  const actual = fila[indiceDestino];
  if (!esValorInformado(actual)) {
    fila[indiceDestino] = valor;
    return;
  }

  const actualTexto = normalizarTexto(actual);
  const nuevoTexto = normalizarTexto(valor);
  const actualNumero = parsearNumeroMedicion(actual, null);
  const nuevoNumero = parsearNumeroMedicion(valor, null);
  const coincidenNumeros = typeof actualNumero === 'number' && !isNaN(actualNumero) &&
    typeof nuevoNumero === 'number' && !isNaN(nuevoNumero) &&
    actualNumero === nuevoNumero;

  if (actualTexto !== nuevoTexto && !coincidenNumeros) {
    conflictos[normalizarSuministro(fila[0])] = true;
  }
}

function esArchivoPadronPorConceptos(cabeceraOrigen) {
  return buscarIndiceCabeceraPadron(cabeceraOrigen, ['SUMINISTRO']) !== -1 &&
    buscarIndiceCabeceraPadron(cabeceraOrigen, ['CONCEPTO']) !== -1 &&
    buscarIndiceCabeceraPadron(cabeceraOrigen, ['LECANT', 'LEC_ANT', 'LECTURA ANTERIOR', 'LECTURA_ANTERIOR']) !== -1;
}

function buscarIndiceCabeceraPadron(cabeceraOrigen, aliases) {
  const headers = Array.isArray(cabeceraOrigen) ? cabeceraOrigen : [];
  const aliasList = Array.isArray(aliases) ? aliases : [];
  const normalizados = headers.map(normalizarCabeceraPadron);

  for (let i = 0; i < aliasList.length; i++) {
    const aliasNormalizado = normalizarCabeceraPadron(aliasList[i]);
    const indice = normalizados.indexOf(aliasNormalizado);
    if (indice !== -1) {
      return indice;
    }
  }

  return -1;
}

function normalizarCabeceraPadron(valor) {
  return normalizarTextoBusqueda(valor).replace(/[^A-Z0-9]/g, '');
}

function inferirPeriodoArchivo(nombreArchivo) {
  return normalizarPeriodo(nombreArchivo, '');
}

function detectarPeriodoPadron(cabeceraOrigen, matrizOrigen, mapeoAlias, nombreArchivo) {
  const candidatos = [];
  agregarPeriodosDetectadosPadron(
    candidatos,
    detectarPeriodosPadronEnCabeceras(cabeceraOrigen),
    'cabecera del archivo'
  );
  agregarPeriodosDetectadosPadron(
    candidatos,
    detectarPeriodosPadronEnColumna(cabeceraOrigen, matrizOrigen, mapeoAlias),
    'columna de periodo'
  );

  const periodoEnNombre = inferirPeriodoArchivo(nombreArchivo);
  if (!candidatos.length && periodoEnNombre) {
    agregarPeriodosDetectadosPadron(candidatos, [periodoEnNombre], 'nombre del archivo');
  }

  const periodosEncontrados = [];
  const fuentesPorPeriodo = {};

  candidatos.forEach(function(candidato) {
    if (!candidato.periodo) {
      return;
    }

    if (!fuentesPorPeriodo[candidato.periodo]) {
      fuentesPorPeriodo[candidato.periodo] = candidato.fuente;
      periodosEncontrados.push(candidato.periodo);
    }
  });

  return {
    periodo: periodosEncontrados[0] || '',
    fuente: periodosEncontrados.length ? fuentesPorPeriodo[periodosEncontrados[0]] : '',
    periodosEncontrados: periodosEncontrados,
    fuentesPorPeriodo: fuentesPorPeriodo
  };
}

function detectarPeriodosPadronEnCabeceras(cabeceraOrigen) {
  const headers = Array.isArray(cabeceraOrigen) ? cabeceraOrigen : [];
  const periodos = {};

  headers.forEach(function(header) {
    const periodo = normalizarPeriodo(header, '');
    if (periodo) {
      periodos[periodo] = true;
    }
  });

  return Object.keys(periodos).sort();
}

function detectarPeriodosPadronEnColumna(cabeceraOrigen, matrizOrigen, mapeoAlias) {
  const aliasPeriodo = mapeoAlias.PERIODO || ['PERIODO'];
  const headers = Array.isArray(cabeceraOrigen) ? cabeceraOrigen : [];
  const filas = Array.isArray(matrizOrigen) ? matrizOrigen : [];
  const indicePeriodo = buscarIndiceCabeceraPadron(headers, aliasPeriodo);

  if (indicePeriodo === -1) {
    return [];
  }

  const periodos = {};

  if (indicePeriodo !== -1) {
    for (let fila = 1; fila < filas.length; fila++) {
      const periodo = normalizarPeriodo(filas[fila][indicePeriodo], '');
      if (periodo) {
        periodos[periodo] = true;
      }
    }
  }

  return Object.keys(periodos).sort();
}

function agregarPeriodosDetectadosPadron(destino, periodos, fuente) {
  const listaDestino = Array.isArray(destino) ? destino : [];
  const listaPeriodos = Array.isArray(periodos) ? periodos : [];

  listaPeriodos.forEach(function(periodo) {
    const periodoNormalizado = normalizarPeriodo(periodo, '');
    if (!periodoNormalizado) {
      return;
    }

    listaDestino.push({
      periodo: periodoNormalizado,
      fuente: fuente
    });
  });
}

function validarPeriodoPadronSeleccionado(periodoSeleccionado, diagnosticoPeriodo) {
  const diagnostico = diagnosticoPeriodo || {};
  const periodosEncontrados = Array.isArray(diagnostico.periodosEncontrados)
    ? diagnostico.periodosEncontrados.filter(Boolean)
    : [];

  if (!periodosEncontrados.length) {
    throw new Error(
      'No se pudo identificar el perÃ­odo del archivo. El Excel debe indicar un ÃƒÆ’Ã‚Âºnico mes, por ejemplo 202604, y ese valor debe coincidir con el mes seleccionado.'
    );
  }

  if (periodosEncontrados.length > 1) {
    throw new Error(
      'El archivo contiene mÃ¡s de un perÃ­odo: ' +
      periodosEncontrados.join(', ') +
      '. Debe existir un ÃƒÆ’Ã‚Âºnico mes en el archivo y debe coincidir con el mes seleccionado.'
    );
  }

  if (periodosEncontrados[0] !== periodoSeleccionado) {
    throw new Error(
      'El mes seleccionado ' +
      periodoSeleccionado +
      ' no coincide con el perÃ­odo detectado en el archivo ' +
      periodosEncontrados[0] +
      '.'
    );
  }
}

function obtenerEstadoActualPadron(sheet) {
  const hoja = sheet;
  const lastRow = hoja ? hoja.getLastRow() : 0;
  const lastColumn = hoja ? hoja.getLastColumn() : 0;

  if (!hoja || lastRow === 0 || lastColumn === 0) {
    return {
      hasHeaders: false,
      headersUpper: [],
      rows: []
    };
  }

  const headersUpper = hoja.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(valor) {
    return normalizarTexto(valor).toUpperCase();
  });
  const hasHeaders = headersUpper.indexOf('SUMINISTRO') !== -1 && headersUpper.indexOf('PERIODO') !== -1;
  const rowsCrudas = lastRow > 1
    ? hoja.getRange(2, 1, lastRow - 1, lastColumn).getValues()
    : [];
  const rows = rowsCrudas.filter(filaPadronTieneContenido);

  if (!hasHeaders && rows.length) {
    throw new Error('La hoja Padron_Maestro no tiene cabeceras vÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¡lidas para SUMINISTRO y PERIODO.');
  }

  if (!hasHeaders && !rows.length) {
    return {
      hasHeaders: false,
      headersUpper: [],
      rows: []
    };
  }

  return {
    hasHeaders: hasHeaders,
    headersUpper: headersUpper,
    rows: rows
  };
}

function validarCoincidenciasPadronExistente(estadoPadron, filasNuevas, periodoAplicado) {
  const headersUpper = estadoPadron && Array.isArray(estadoPadron.headersUpper)
    ? estadoPadron.headersUpper
    : [];
  const rows = estadoPadron && Array.isArray(estadoPadron.rows)
    ? estadoPadron.rows
    : [];
  const indiceSuministro = headersUpper.indexOf('SUMINISTRO');
  const indicePeriodo = headersUpper.indexOf('PERIODO');
  const clavesExistentes = {};
  const clavesPeriodo = {};
  const coincidenciasMap = {};

  if (indiceSuministro === -1 || indicePeriodo === -1) {
    return {
      coincidencias: [],
      totalPeriodo: 0
    };
  }

  rows.forEach(function(row) {
    const suministro = indiceSuministro < row.length ? row[indiceSuministro] : '';
    const periodo = indicePeriodo < row.length ? row[indicePeriodo] : '';
    const periodoNormalizado = normalizarPeriodo(periodo, '');
    const clave = construirClaveSuministroPeriodo(suministro, periodoNormalizado, '');

    if (!clave) {
      return;
    }

    clavesExistentes[clave] = true;

    if (periodoNormalizado === periodoAplicado) {
      clavesPeriodo[clave] = true;
    }
  });

  filasNuevas.forEach(function(fila) {
    const suministro = fila[0];
    const periodo = fila[9] || periodoAplicado;
    const clave = construirClaveSuministroPeriodo(suministro, periodo, '');

    if (clave && clavesExistentes[clave]) {
      coincidenciasMap[normalizarSuministro(suministro)] = true;
    }
  });

  return {
    coincidencias: Object.keys(coincidenciasMap).sort(),
    totalPeriodo: Object.keys(clavesPeriodo).length
  };
}

function filaPadronTieneContenido(row) {
  const values = Array.isArray(row) ? row : [];

  for (let i = 0; i < values.length; i++) {
    if (normalizarTexto(values[i]) !== '') {
      return true;
    }
  }

  return false;
}

function construirMensajeCoincidenciasPadron(periodo, coincidencias) {
  const lista = coincidencias.slice(0, 10).join(', ');
  const restantes = coincidencias.length > 10
    ? ' y ' + (coincidencias.length - 10) + ' mÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¡s'
    : '';

  return 'No se aceptÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³ el archivo porque ya existen ' + coincidencias.length +
    ' suministro(s) cargado(s) en el perÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â­odo ' + periodo + '. Coincidencias: ' +
    lista + restantes + '.';
}

function construirMensajeCoincidenciasArchivo(periodo, coincidencias) {
  const lista = coincidencias.slice(0, 10).join(', ');
  const restantes = coincidencias.length > 10
    ? ' y ' + (coincidencias.length - 10) + ' mÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¡s'
    : '';

  return 'No se aceptÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³ el archivo porque trae ' + coincidencias.length +
    ' suministro(s) repetido(s) dentro del mismo perÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â­odo ' + periodo + '. Coincidencias: ' +
    lista + restantes + '.';
}

function borrarCargaPadron(periodo, usuario, password) {
  const periodoObjetivo = normalizarPeriodo(periodo, '');
  if (!periodoObjetivo) {
    throw new Error('Seleccione el mes del padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n que desea borrar.');
  }

  validarClaveUsuarioActual(usuario, password);

  const ss = SpreadsheetApp.openById(CONFIG.ID_HOJA_MAESTRO);
  const sheet = ss.getSheetByName(CONFIG.NOMBRE_HOJA_PADRON);

  if (!sheet || sheet.getLastRow() <= 1 || sheet.getLastColumn() === 0) {
    throw new Error('No existen registros de padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n para borrar.');
  }

  const data = sheet.getDataRange().getValues();
  const headersUpper = data[0].map(function(valor) {
    return normalizarTexto(valor).toUpperCase();
  });
  const indicePeriodo = headersUpper.indexOf('PERIODO');

  if (indicePeriodo === -1) {
    throw new Error('La hoja Padron_Maestro no tiene la columna PERIODO.');
  }

  const filasConservadas = [];
  let registrosEliminados = 0;

  for (let i = 1; i < data.length; i++) {
    const fila = data[i];
    if (!filaPadronTieneContenido(fila)) {
      continue;
    }

    const periodoFila = indicePeriodo < fila.length
      ? normalizarPeriodo(fila[indicePeriodo], '')
      : '';

    if (periodoFila === periodoObjetivo) {
      registrosEliminados++;
      continue;
    }

    filasConservadas.push(fila);
  }

  if (!registrosEliminados) {
    throw new Error('No existen registros cargados en el perÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â­odo ' + periodoObjetivo + '.');
  }

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }

  if (filasConservadas.length) {
    sheet.getRange(2, 1, filasConservadas.length, data[0].length).setValues(filasConservadas);
  }

  return {
    success: true,
    archivo: 'Borrado del padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n',
    periodo: periodoObjetivo,
    periodoSeleccionado: periodoObjetivo,
    periodoDetectado: periodoObjetivo,
    carpeta: null,
    registros: registrosEliminados,
    timestamp: obtenerMarcaTiempo(),
    mensaje: 'Se eliminaron ' + registrosEliminados + ' registro(s) del padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n del perÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â­odo ' + periodoObjetivo + '.',
    accion: 'borrado',
    usuario: normalizarTexto(usuario)
  };
}

function exportarPadronEnosa(periodoFiltro) {
  const periodoObjetivo = normalizarPeriodo(periodoFiltro, '');
  if (!periodoObjetivo) {
    throw new Error('Seleccione el mes del padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n que desea exportar.');
  }

  const padron = obtenerContextoPadron();
  if (!padron.success) {
    throw new Error('No existe informaciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n de padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n disponible para exportar.');
  }

  const filasPeriodo = (padron.rows || []).filter(function(row) {
    return row.periodo === periodoObjetivo;
  });

  if (!filasPeriodo.length) {
    throw new Error('No existen registros de padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n en el perÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â­odo ' + periodoObjetivo + '.');
  }

  const nombreArchivoOrigen = resolverNombreArchivoOriginalPadron(padron.headersUpper, filasPeriodo);
  const carpetaPeriodo = obtenerOCrearCarpetaPadrones(periodoObjetivo);
  const archivoOrigen = buscarArchivoOrigenPadron(carpetaPeriodo, nombreArchivoOrigen);
  if (!archivoOrigen) {
    throw new Error('No se encontrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³ el archivo original del padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n en la carpeta del perÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â­odo ' + periodoObjetivo + '.');
  }

  const tempFile = crearSpreadsheetTemporalDesdeBlobPadron(archivoOrigen.getBlob(), 'temp_export_padron_' + periodoObjetivo);

  try {
    const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    const tempSheet = tempSpreadsheet.getSheets()[0];
    const matriz = tempSheet.getDataRange().getDisplayValues();

    if (!matriz.length) {
      throw new Error('El archivo original del padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n no contiene datos.');
    }

    const cabecera = matriz[0].map(function(celda) {
      return normalizarTexto(celda);
    });
    const cabeceraUpper = cabecera.map(function(celda) {
      return celda.toUpperCase();
    });
    const indiceSuministro = buscarIndiceCabeceraPadron(cabeceraUpper, ['SUMINISTRO', 'CODIGO', 'ID', 'SUMI']);
    const indiceConcepto = buscarIndiceCabeceraPadron(cabeceraUpper, ['CONCEPTO']);
    const indiceConceptoDetalle = buscarIndiceCabeceraPadron(cabeceraUpper, ['CONCEPTO_D', 'CONCEPTO DETALLE', 'CONCEPTODETALLE']);
    const indiceLecturaO = buscarIndiceCabeceraPadron(cabeceraUpper, ['LECTURAO', 'LECTURA_O', 'LECTURA O']);

    if (indiceSuministro === -1) {
      throw new Error('El archivo original no tiene una columna SUMINISTRO reconocible.');
    }
    if (indiceLecturaO === -1) {
      throw new Error('El archivo original no tiene la columna LecturaO.');
    }

    const mapaCargas = construirMapaCargasParaExportPadron(padron.headersUpper, filasPeriodo);
    const filasOrdenadas = [];
    let ultimoSuministro = '';

    for (let i = 1; i < matriz.length; i++) {
      const fila = (matriz[i] || []).slice();
      let suministroFila = indiceSuministro < fila.length ? normalizarTexto(fila[indiceSuministro]) : '';

      if (suministroFila) {
        ultimoSuministro = suministroFila;
      } else if (ultimoSuministro) {
        suministroFila = ultimoSuministro;
        fila[indiceSuministro] = ultimoSuministro;
      }

      const ordenConcepto = resolverOrdenConceptoExportPadron(
        fila,
        indiceConcepto,
        indiceConceptoDetalle
      );
      if (indiceConcepto !== -1 && ordenConcepto >= 1 && ordenConcepto <= 6) {
        fila[indiceConcepto] = String(ordenConcepto);
      }

      const cargaSuministro = mapaCargas[normalizarSuministro(suministroFila)] || null;
      const lecturaActual = cargaSuministro
        ? resolverLecturaActualExportPadron(cargaSuministro, ordenConcepto)
        : '';
      if (indiceLecturaO !== -1) {
        fila[indiceLecturaO] = lecturaActual;
      }

      filasOrdenadas.push({
        fila: completarFilaExportPadron(fila, cabecera.length),
        suministroClave: normalizarSuministro(suministroFila),
        ordenConcepto: ordenConcepto,
        posicionOriginal: i
      });
    }

    filasOrdenadas.sort(function(a, b) {
      if (a.suministroClave < b.suministroClave) {
        return -1;
      }
      if (a.suministroClave > b.suministroClave) {
        return 1;
      }
      if (a.ordenConcepto !== b.ordenConcepto) {
        return a.ordenConcepto - b.ordenConcepto;
      }
      return a.posicionOriginal - b.posicionOriginal;
    });

    const salida = [cabecera].concat(filasOrdenadas.map(function(item) {
      return item.fila;
    }));

    tempSheet.clearContents();
    tempSheet.getRange(1, 1, salida.length, cabecera.length).setValues(salida);

    const nombreBase = construirNombreExportPadron(archivoOrigen.getName(), periodoObjetivo);
    const blobXlsx = exportarSpreadsheetAXlsx(tempSpreadsheet.getId(), nombreBase + '.xlsx');
    const archivoExportado = carpetaPeriodo.createFile(blobXlsx);

    return {
      success: true,
      periodo: periodoObjetivo,
      registros: filasOrdenadas.length,
      archivoOrigen: archivoOrigen.getName(),
      archivoExportado: archivoExportado.getName(),
      url: archivoExportado.getUrl(),
      urlDescarga: 'https://drive.google.com/uc?export=download&id=' + archivoExportado.getId(),
      mensaje: 'ExportaciÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n del padrÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â³n generada correctamente.'
    };
  } finally {
    if (tempFile && tempFile.id) {
      Drive.Files.remove(tempFile.id);
    }
  }
}

function resolverNombreArchivoOriginalPadron(headersUpper, filasPeriodo) {
  const headers = Array.isArray(headersUpper) ? headersUpper : [];
  const filas = Array.isArray(filasPeriodo) ? filasPeriodo : [];
  const indiceNombreArchivo = headers.indexOf('NOMBRE_ARCHIVO');
  const contador = {};
  let mejorNombre = '';
  let mejorTotal = 0;

  if (indiceNombreArchivo === -1) {
    return '';
  }

  filas.forEach(function(row) {
    const values = row && Array.isArray(row.values) ? row.values : [];
    const nombre = indiceNombreArchivo < values.length ? normalizarTexto(values[indiceNombreArchivo]) : '';
    if (!nombre) {
      return;
    }
    contador[nombre] = (contador[nombre] || 0) + 1;
    if (contador[nombre] > mejorTotal) {
      mejorTotal = contador[nombre];
      mejorNombre = nombre;
    }
  });

  return mejorNombre;
}

function buscarArchivoOrigenPadron(carpetaPeriodo, nombrePreferido) {
  const folder = carpetaPeriodo;
  const preferido = normalizarTexto(nombrePreferido);

  if (preferido) {
    const coincidencias = folder.getFilesByName(preferido);
    while (coincidencias.hasNext()) {
      const archivo = coincidencias.next();
      if (esArchivoExcelPadronOrigen(archivo)) {
        return archivo;
      }
    }
  }

  const archivos = folder.getFiles();
  let mejorArchivo = null;
  let mejorFecha = 0;

  while (archivos.hasNext()) {
    const archivo = archivos.next();
    if (!esArchivoExcelPadronOrigen(archivo)) {
      continue;
    }
    const nombre = normalizarTexto(archivo.getName()).toUpperCase();
    if (nombre.indexOf('EXPORT_PADRON_') === 0) {
      continue;
    }

    const modificado = archivo.getLastUpdated().getTime();
    if (!mejorArchivo || modificado > mejorFecha) {
      mejorArchivo = archivo;
      mejorFecha = modificado;
    }
  }

  return mejorArchivo;
}

function esArchivoExcelPadronOrigen(archivo) {
  const nombre = normalizarTexto(archivo && archivo.getName ? archivo.getName() : '').toLowerCase();
  return nombre.endsWith('.xlsx') || nombre.endsWith('.xls');
}

function crearSpreadsheetTemporalDesdeBlobPadron(blob, nombreBase) {
  const resource = {
    name: nombreBase,
    mimeType: MimeType.GOOGLE_SHEETS
  };

  return Drive.Files.create
    ? Drive.Files.create(resource, blob)
    : Drive.Files.insert({ title: resource.name, mimeType: resource.mimeType }, blob);
}

function construirMapaCargasParaExportPadron(headersUpper, filasPeriodo) {
  const mapa = {};
  const filas = Array.isArray(filasPeriodo) ? filasPeriodo : [];

  filas.forEach(function(row) {
    const detalle = construirDetalleSuministro(headersUpper, row);
    mapa[normalizarSuministro(detalle.suministro)] = {
      1: detalle.carga.EAT,
      2: detalle.carga.EAHP,
      3: detalle.carga.EAFP,
      4: detalle.carga.ER,
      5: detalle.carga.PHP,
      6: detalle.carga.PFP
    };
  });

  return mapa;
}

function resolverOrdenConceptoExportPadron(fila, indiceConcepto, indiceConceptoDetalle) {
  const valorConcepto = indiceConcepto === -1 ? '' : normalizarTexto(fila[indiceConcepto]);
  const matchConcepto = valorConcepto.match(/\d+/);
  if (matchConcepto) {
    const numero = parseInt(matchConcepto[0], 10);
    if (!isNaN(numero) && numero >= 1 && numero <= 6) {
      return numero;
    }
  }

  const campoLectura = resolverCampoLecturaPorConcepto(fila, indiceConcepto, indiceConceptoDetalle);
  const mapaOrden = {
    LEC_ANT_EAT: 1,
    LEC_ANT_EAHP: 2,
    LEC_ANT_EAFP: 3,
    LEC_ANT_ER: 4,
    LEC_ANT_PHP: 5,
    LEC_ANT_PFP: 6
  };

  return mapaOrden[campoLectura] || 999;
}

function resolverLecturaActualExportPadron(cargaSuministro, ordenConcepto) {
  if (!cargaSuministro || !Object.prototype.hasOwnProperty.call(cargaSuministro, ordenConcepto)) {
    return '';
  }

  const valor = cargaSuministro[ordenConcepto];
  return valor === '-' ? '' : valor;
}

function completarFilaExportPadron(fila, longitud) {
  const salida = Array.isArray(fila) ? fila.slice() : [];
  while (salida.length < longitud) {
    salida.push('');
  }
  return salida;
}

function construirNombreExportPadron(nombreArchivoOrigen, periodoObjetivo) {
  const nombre = normalizarTexto(nombreArchivoOrigen).replace(/\.(xlsx|xls)$/i, '');
  return 'EXPORT_PADRON_' + (nombre || 'ENOSA') + '_' + periodoObjetivo;
}
