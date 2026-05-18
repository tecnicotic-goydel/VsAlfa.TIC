const CAMPOS_CARGA_MANUAL = ['EAT', 'EAHP', 'EAFP', 'ER', 'PHP', 'PFP'];
const CAMPOS_CARGA_PREVIA = ['LEC_ANT_EAT', 'LEC_ANT_EAHP', 'LEC_ANT_EAFP', 'LEC_ANT_ER', 'LEC_ANT_PHP', 'LEC_ANT_PFP'];
const MIME_TYPE_XLSX_CARGA_MANUAL = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const NOMBRE_HOJA_PLANTILLA_CARGA_MANUAL = 'Datos';
const NOMBRE_SUBCARPETA_CARGA_MANUAL = 'CARGA_MANUAL';

function listarPendientesCargaManual(periodoFiltro) {
  const resumen = getResumenMensual(periodoFiltro);
  const suministros = Array.isArray(resumen.suministros) ? resumen.suministros : [];
  const pendientes = suministros.filter(function(item) {
    const estado = normalizarTexto(item.estado).toUpperCase();
    return estado === CONFIG.ESTADO_PENDIENTE || estado === CONFIG.ESTADO_LEIDO_SIN_CARGA;
  });

  return {
    success: true,
    periodo: resumen.periodo,
    periodoLabel: resumen.periodoLabel,
    periodosDisponibles: resumen.periodosDisponibles || [],
    totalPendientes: pendientes.length,
    suministros: pendientes
  };
}

function registrarCargaManual(payload) {
  const data = payload || {};
  const suministro = normalizarTexto(data.suministro);
  const periodo = normalizarPeriodo(data.periodo, '');
  const observacion = normalizarTexto(data.observacion || data.obs);
  const mediciones = construirMedicionesCargaManualDesdePayload(data);

  if (!suministro) {
    throw new Error('Debe seleccionar un suministro pendiente.');
  }

  if (!periodo) {
    throw new Error('No se pudo resolver el período del suministro seleccionado.');
  }

  if (!tieneLecturasManualesInformadas(mediciones)) {
    throw new Error('Ingrese al menos una lectura para registrar la carga manual.');
  }

  const sheet = obtenerHojaPadronManual();
  const contexto = obtenerContextoActualizacionCargaManual(sheet);
  const actualizacion = actualizarFilaPadronConCargaManual(
    sheet,
    contexto.valores,
    contexto.columnas,
    {
      suministro: suministro,
      periodo: periodo,
      mediciones: mediciones,
      observacion: observacion
    },
    {
      observacionPorDefecto: 'Carga manual registrada desde el módulo.',
      origenLectura: CONFIG.ORIGEN_LECTURA_CAMPO
    }
  );

  if (!actualizacion.exito) {
    throw new Error(actualizacion.motivo);
  }

  return {
    success: true,
    mensaje: 'La carga manual se registró correctamente.',
    suministro: suministro,
    periodo: periodo,
    periodoLabel: formatearPeriodo(periodo),
    fila: actualizacion.fila,
    timestamp: obtenerMarcaTiempo()
  };
}

function generarPlantillaCargaManual(periodoFiltro, suministroFiltro) {
  const respuesta = listarPendientesCargaManual(periodoFiltro);
  const periodo = normalizarPeriodo(respuesta.periodo, '');
  const pendientes = filtrarPendientesCargaManualPorSuministro(respuesta.suministros, suministroFiltro);

  if (!periodo) {
    throw new Error('No existe un período activo para generar la plantilla.');
  }

  if (!pendientes.length) {
    throw new Error('No hay suministros pendientes para generar la plantilla con el filtro actual.');
  }

  const nombreBase = 'Plantilla_Carga_Manual_' + periodo + '_' + Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'GMT-5',
    'yyyyMMdd_HHmmss'
  );
  const tempSpreadsheet = SpreadsheetApp.create(nombreBase);

  try {
    const hojaDatos = tempSpreadsheet.getSheets()[0];
    hojaDatos.setName(NOMBRE_HOJA_PLANTILLA_CARGA_MANUAL);
    cargarContenidoPlantillaManual(hojaDatos, pendientes);

    const carpetaDestino = obtenerOCrearCarpetaCargaManual(periodo);
    const blobXlsx = exportarSpreadsheetAXlsx(tempSpreadsheet.getId(), nombreBase + '.xlsx');
    const archivo = carpetaDestino.createFile(blobXlsx);

    return {
      success: true,
      mensaje: 'Plantilla generada correctamente.',
      nombre: archivo.getName(),
      url: archivo.getUrl(),
      urlDescarga: 'https://drive.google.com/uc?export=download&id=' + archivo.getId(),
      id: archivo.getId(),
      periodo: periodo,
      periodoLabel: formatearPeriodo(periodo),
      registros: pendientes.length,
      timestamp: obtenerMarcaTiempo()
    };
  } finally {
    try {
      DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
    } catch (error) {
      // No interrumpir la respuesta si la limpieza falla.
    }
  }
}

function procesarPlantillaCargaManual(fileData, periodoFallback) {
  if (!fileData || !fileData.data) {
    throw new Error('Seleccione un archivo Excel con la plantilla de carga manual.');
  }

  const periodoRespaldo = normalizarPeriodo(periodoFallback, '');
  let tempFile;

  try {
    tempFile = crearSpreadsheetTemporalDesdeArchivo(fileData, 'temp_carga_manual');
    const tempSpreadsheet = SpreadsheetApp.openById(tempFile.id);
    const hoja = tempSpreadsheet.getSheets()[0];
    const rawData = hoja.getDataRange().getValues();
    const displayData = hoja.getDataRange().getDisplayValues();

    if (!rawData.length) {
      throw new Error('La plantilla no contiene datos.');
    }

    const estructuraPlantilla = resolverEstructuraPlantillaCargaManual(displayData);
    const columnasPlantilla = estructuraPlantilla.columnas;

    if (columnasPlantilla.sum === -1) {
      throw new Error('La plantilla debe incluir la columna NroServicio.');
    }

    if (columnasPlantilla.per === -1 && !periodoRespaldo) {
      throw new Error('Seleccione el mes del padrón antes de subir la plantilla de carga manual.');
    }

    const registros = [];
    const indicePorClave = {};
    const duplicadosConsolidados = [];
    let filasConLectura = 0;

    for (let i = estructuraPlantilla.dataStartRow; i < rawData.length; i++) {
      const filaRaw = rawData[i] || [];
      const filaDisplay = displayData[i] || [];
      const suministro = normalizarTexto(obtenerValorDeFilaPorIndice(filaDisplay, columnasPlantilla.sum));

      if (!suministro) {
        continue;
      }

      const periodo = normalizarPeriodo(
        obtenerValorDeFilaPorIndice(filaRaw, columnasPlantilla.per),
        normalizarPeriodo(obtenerValorDeFilaPorIndice(filaDisplay, columnasPlantilla.per), periodoRespaldo)
      );
      const mediciones = obtenerMedicionesDesdeFilaPlantilla(filaRaw, filaDisplay, columnasPlantilla);
      const observacion = normalizarTexto(obtenerValorDeFilaPorIndice(filaDisplay, columnasPlantilla.obs));

      if (!periodo) {
        registros.push({
          error: true,
          detalle: 'Fila ' + (i + 1) + ': no se pudo resolver el período del suministro ' + suministro + '.'
        });
        continue;
      }

      if (!tieneLecturasManualesInformadas(mediciones)) {
        continue;
      }

      filasConLectura++;

      const clave = construirClaveSuministroPeriodo(suministro, periodo, '');
      const registro = {
        suministro: suministro,
        periodo: periodo,
        mediciones: mediciones,
        observacion: observacion
      };

      if (clave && Object.prototype.hasOwnProperty.call(indicePorClave, clave)) {
        const indiceExistente = indicePorClave[clave];
        registros[indiceExistente] = combinarRegistroCargaManual(
          registros[indiceExistente],
          registro
        );
        duplicadosConsolidados.push(suministro + ' (' + periodo + ')');
        continue;
      }

      if (clave) {
        indicePorClave[clave] = registros.length;
      }

      registros.push(registro);
    }

    if (!registros.length) {
      throw new Error('La plantilla no contiene lecturas manuales para procesar.');
    }

    const sheet = obtenerHojaPadronManual();
    const contexto = obtenerContextoActualizacionCargaManual(sheet);
    const detalles = [];
    let procesados = 0;
    let omitidos = 0;
    let errores = 0;
    const periodosAfectados = {};
    const periodosProcesados = {};
    const clavesTelemetria = {};
    const noEncontrados = [];
    const yaRealizados = [];
    const omitidosPorOtraCausa = [];

    registros.forEach(function(registro) {
      if (registro.error) {
        errores++;
        detalles.push(registro.detalle);
        return;
      }

      periodosAfectados[registro.periodo] = true;

      const actualizacion = actualizarFilaPadronConCargaManual(
        sheet,
        contexto.valores,
        contexto.columnas,
        registro,
        {
          observacionPorDefecto: 'Carga masiva registrada desde plantilla.',
          origenLectura: CONFIG.ORIGEN_LECTURA_TELEMETRIA
        }
      );

      if (actualizacion.exito) {
        procesados++;
        periodosProcesados[registro.periodo] = true;
        if (actualizacion.clave) {
          clavesTelemetria[actualizacion.clave] = true;
        }
        detalles.push(registro.suministro + ' (' + registro.periodo + '): carga aplicada.');
      } else {
        omitidos++;
        detalles.push(registro.suministro + ' (' + registro.periodo + '): ' + actualizacion.motivo);
        clasificarOmitidoCargaManual(registro, actualizacion.motivo, noEncontrados, yaRealizados, omitidosPorOtraCausa);
      }
    });

    const clasificacionOrigen = clasificarOrigenLecturaTrasPlantilla(
      sheet,
      contexto.valores,
      contexto.columnas,
      Object.keys(periodosProcesados),
      clavesTelemetria
    );

    return {
      success: true,
      mensaje: 'Plantilla procesada correctamente.',
      archivo: fileData.fileName,
      periodos: Object.keys(periodosAfectados).sort(),
      timestamp: obtenerMarcaTiempo(),
      resumen: {
        filasConLectura: filasConLectura,
        registrosLeidos: registros.length,
        procesados: procesados,
        omitidos: omitidos,
        errores: errores,
        telemetria: clasificacionOrigen.telemetria || 0,
        campo: clasificacionOrigen.campo || 0
      },
      detalles: detalles.slice(0, 50),
      duplicadosConsolidados: deduplicarListaTexto(duplicadosConsolidados),
      noEncontrados: noEncontrados,
      yaRealizados: yaRealizados,
      omitidosPorOtraCausa: omitidosPorOtraCausa
    };
  } finally {
    if (tempFile && tempFile.id) {
      Drive.Files.remove(tempFile.id);
    }
  }
}

function obtenerHojaPadronManual() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.NOMBRE_HOJA_PADRON);
  if (!sheet) {
    throw new Error("No existe la hoja '" + CONFIG.NOMBRE_HOJA_PADRON + "'.");
  }

  return sheet;
}

function obtenerContextoActualizacionCargaManual(sheet) {
  asegurarEstructuraPadron(sheet);
  const valores = sheet.getDataRange().getDisplayValues();
  if (!valores.length) {
    throw new Error("La hoja '" + CONFIG.NOMBRE_HOJA_PADRON + "' no tiene cabeceras.");
  }

  const headersUpper = (valores[0] || []).map(function(valor) {
    return normalizarTexto(valor).toUpperCase();
  });
  const columnas = {
    sum: headersUpper.indexOf('SUMINISTRO'),
    per: headersUpper.indexOf('PERIODO'),
    est: headersUpper.indexOf('ESTADO'),
    estadoLectura: headersUpper.indexOf('ESTADO_LECTURA'),
    origen: headersUpper.indexOf('ORIGEN_LECTURA'),
    eat: headersUpper.indexOf('EAT'),
    eahp: headersUpper.indexOf('EAHP'),
    eafp: headersUpper.indexOf('EAFP'),
    er: headersUpper.indexOf('ER'),
    php: headersUpper.indexOf('PHP'),
    pfp: headersUpper.indexOf('PFP'),
    lecAntEat: headersUpper.indexOf('LEC_ANT_EAT'),
    lecAntEahp: headersUpper.indexOf('LEC_ANT_EAHP'),
    lecAntEafp: headersUpper.indexOf('LEC_ANT_EAFP'),
    lecAntEr: headersUpper.indexOf('LEC_ANT_ER'),
    lecAntPhp: headersUpper.indexOf('LEC_ANT_PHP'),
    lecAntPfp: headersUpper.indexOf('LEC_ANT_PFP'),
    fecha: headersUpper.indexOf('FECHA REGISTRO'),
    obs: headersUpper.indexOf('OBS'),
    archivoLectura: headersUpper.indexOf('ARCHIVO_LECTURA'),
    tipoArchivoLectura: headersUpper.indexOf('TIPO_ARCHIVO_LECTURA')
  };

  if (columnas.sum === -1 || columnas.per === -1 || columnas.est === -1) {
    throw new Error("La hoja '" + CONFIG.NOMBRE_HOJA_PADRON + "' debe incluir las columnas SUMINISTRO, PERIODO y ESTADO.");
  }

  return {
    valores: valores,
    columnas: columnas
  };
}

function actualizarFilaPadronConCargaManual(sheet, valores, columnas, payload, opciones) {
  const suministroBuscado = normalizarSuministro(payload && payload.suministro);
  const periodoObjetivo = normalizarPeriodo(payload && payload.periodo, '');
  const mediciones = payload && payload.mediciones ? payload.mediciones : {};
  const observacion = normalizarTexto(payload && payload.observacion);
  const config = opciones || {};
  const periodosAlternativos = [];

  if (!suministroBuscado || !periodoObjetivo) {
    return {
      exito: false,
      motivo: 'No se pudo identificar el suministro y el período a actualizar.'
    };
  }

  if (!tieneLecturasManualesInformadas(mediciones)) {
    return {
      exito: false,
      motivo: 'No se encontraron lecturas manuales para registrar.'
    };
  }

  for (let i = 1; i < valores.length; i++) {
    const suministroFila = normalizarSuministro(obtenerValorDeFilaPorIndice(valores[i], columnas.sum));
    const periodoFila = normalizarPeriodo(obtenerValorDeFilaPorIndice(valores[i], columnas.per), '');
    const estadoFila = normalizarTexto(obtenerValorDeFilaPorIndice(valores[i], columnas.est)).toUpperCase();

    if (suministroFila === suministroBuscado && periodoFila && periodoFila !== periodoObjetivo) {
      periodosAlternativos.push(periodoFila);
    }

    if (suministroFila !== suministroBuscado || periodoFila !== periodoObjetivo) {
      continue;
    }

    if (esEstadoFinalizado(estadoFila)) {
      return {
        exito: false,
        motivo: 'Ya estaba marcado como ' + resolverEtiquetaEstadoFinal(estadoFila) + '.'
      };
    }

    const filaHoja = i + 1;
    escribirMedicionesManualesEnFila(sheet, filaHoja, columnas, mediciones);

    if (columnas.fecha !== -1) {
      sheet.getRange(filaHoja, columnas.fecha + 1).setValue(new Date());
      valores[i][columnas.fecha] = obtenerMarcaTiempo();
    }

    if (columnas.obs !== -1) {
      const observacionActual = normalizarTexto(obtenerValorDeFilaPorIndice(valores[i], columnas.obs));
      const observacionFinal = observacion || observacionActual || normalizarTexto(config.observacionPorDefecto);

      if (observacionFinal !== observacionActual) {
        sheet.getRange(filaHoja, columnas.obs + 1).setValue(observacionFinal);
        valores[i][columnas.obs] = observacionFinal;
      }
    }

    if (columnas.origen !== -1 && normalizarTexto(config.origenLectura)) {
      sheet.getRange(filaHoja, columnas.origen + 1).setValue(config.origenLectura);
      valores[i][columnas.origen] = config.origenLectura;
    }

    sheet.getRange(filaHoja, columnas.est + 1).setValue(CONFIG.ESTADO_COMPLETO);
    valores[i][columnas.est] = CONFIG.ESTADO_COMPLETO;

    return {
      exito: true,
      fila: filaHoja,
      clave: construirClaveSuministroPeriodo(payload.suministro, payload.periodo, '')
    };
  }

  const periodosUnicos = deduplicarListaTexto(periodosAlternativos).map(function(periodo) {
    return formatearPeriodo(periodo);
  });

  if (periodosUnicos.length) {
    return {
      exito: false,
      motivo: 'El suministro existe en el padrón, pero no en el período seleccionado ' +
        formatearPeriodo(periodoObjetivo) + '. Períodos encontrados: ' + periodosUnicos.join(', ') + '.'
    };
  }

  return {
    exito: false,
    motivo: 'No se halló el suministro para el período seleccionado.'
  };
}

function actualizarSuministroDesdePlataforma(payload) {
  const data = payload || {};
  const suministro = normalizarTexto(data.suministro);
  const periodo = normalizarPeriodo(data.periodo, '');

  if (!suministro) {
    throw new Error('Debe seleccionar un suministro para editar.');
  }

  if (!periodo) {
    throw new Error('No se pudo resolver el período del suministro seleccionado.');
  }

  const sheet = obtenerHojaPadronManual();
  const contexto = obtenerContextoActualizacionCargaManual(sheet);
  const actualizacion = actualizarFilaPadronDesdePlataforma(
    sheet,
    contexto.valores,
    contexto.columnas,
    {
      suministro: suministro,
      periodo: periodo,
      mediciones: construirMedicionesCargaManualDesdePayload(data),
      previas: construirMedicionesPreviasDesdePayload(data),
      estadoLectura: data.estadoLectura,
      origenLectura: normalizarOrigenLectura(data.origenLectura != null ? data.origenLectura : data.origen),
      observacion: data.observacion != null ? data.observacion : data.obs
    }
  );

  if (!actualizacion.exito) {
    throw new Error(actualizacion.motivo);
  }

  return {
    success: true,
    mensaje: 'El suministro se actualizó correctamente desde la plataforma.',
    suministro: suministro,
    periodo: periodo,
    periodoLabel: formatearPeriodo(periodo),
    fila: actualizacion.fila,
    estado: actualizacion.estado,
    timestamp: obtenerMarcaTiempo()
  };
}

function marcarEstadoLecturaSuministro(payload) {
  const data = payload || {};
  const suministro = normalizarTexto(data.suministro);
  const periodo = normalizarPeriodo(data.periodo, '');
  const estadoLectura = normalizarEstadoLecturaOperativa(data.estadoLectura);

  if (!suministro) {
    throw new Error('Debe seleccionar un suministro para actualizar su estado.');
  }

  if (!periodo) {
    throw new Error('No se pudo resolver el período del suministro seleccionado.');
  }

  if (!estadoLectura) {
    throw new Error('Seleccione un estado válido: CORTADO o RETIRADO.');
  }

  const sheet = obtenerHojaPadronManual();
  const contexto = obtenerContextoActualizacionCargaManual(sheet);
  const actualizacion = actualizarEstadoLecturaDesdePlataforma(
    sheet,
    contexto.valores,
    contexto.columnas,
    {
      suministro: suministro,
      periodo: periodo,
      estadoLectura: estadoLectura
    }
  );

  if (!actualizacion.exito) {
    throw new Error(actualizacion.motivo);
  }

  return {
    success: true,
    mensaje: 'Estado de lectura actualizado correctamente desde la plataforma.',
    suministro: suministro,
    periodo: periodo,
    periodoLabel: formatearPeriodo(periodo),
    fila: actualizacion.fila,
    estado: actualizacion.estado,
    estadoLectura: estadoLectura,
    timestamp: obtenerMarcaTiempo()
  };
}

function actualizarFilaPadronDesdePlataforma(sheet, valores, columnas, payload) {
  const suministroBuscado = normalizarSuministro(payload && payload.suministro);
  const periodoObjetivo = normalizarPeriodo(payload && payload.periodo, '');
  const mediciones = payload && payload.mediciones ? payload.mediciones : {};
  const previas = payload && payload.previas ? payload.previas : {};
  const actualizarEstadoLectura = Boolean(payload) && Object.prototype.hasOwnProperty.call(payload, 'estadoLectura');
  const estadoLectura = normalizarEstadoLecturaOperativa(payload && payload.estadoLectura);
  const observacion = normalizarTexto(payload && payload.observacion);
  const origenLectura = normalizarOrigenLectura(payload && payload.origenLectura);
  const periodosAlternativos = [];

  if (!suministroBuscado || !periodoObjetivo) {
    return {
      exito: false,
      motivo: 'No se pudo identificar el suministro y el período a actualizar.'
    };
  }

  for (let i = 1; i < valores.length; i++) {
    const filaValores = valores[i] || [];
    const suministroFila = normalizarSuministro(obtenerValorDeFilaPorIndice(filaValores, columnas.sum));
    const periodoFila = normalizarPeriodo(obtenerValorDeFilaPorIndice(filaValores, columnas.per), '');

    if (suministroFila === suministroBuscado && periodoFila && periodoFila !== periodoObjetivo) {
      periodosAlternativos.push(periodoFila);
    }

    if (suministroFila !== suministroBuscado || periodoFila !== periodoObjetivo) {
      continue;
    }

    const filaHoja = i + 1;
    escribirMedicionesEditablesEnFila(sheet, filaHoja, filaValores, columnas, mediciones);
    escribirPreviasEditablesEnFila(sheet, filaHoja, filaValores, columnas, previas);

    if (actualizarEstadoLectura && columnas.estadoLectura !== -1) {
      filaValores[columnas.estadoLectura] = escribirTextoEditableEnFila(sheet, filaHoja, columnas.estadoLectura, estadoLectura);
    }

    if (columnas.origen !== -1) {
      filaValores[columnas.origen] = escribirTextoEditableEnFila(sheet, filaHoja, columnas.origen, origenLectura);
    }

    if (columnas.obs !== -1) {
      filaValores[columnas.obs] = escribirTextoEditableEnFila(sheet, filaHoja, columnas.obs, observacion);
    }

    if (columnas.fecha !== -1) {
      sheet.getRange(filaHoja, columnas.fecha + 1).setValue(new Date());
      filaValores[columnas.fecha] = obtenerMarcaTiempo();
    }

    const estadoFinal = resolverEstadoTrasCrud(filaValores, columnas);
    sheet.getRange(filaHoja, columnas.est + 1).setValue(estadoFinal);
    filaValores[columnas.est] = estadoFinal;

    return {
      exito: true,
      fila: filaHoja,
      estado: estadoFinal
    };
  }

  const periodosUnicos = deduplicarListaTexto(periodosAlternativos).map(function(periodo) {
    return formatearPeriodo(periodo);
  });

  if (periodosUnicos.length) {
    return {
      exito: false,
      motivo: 'El suministro existe en el padrón, pero no en el período seleccionado ' +
        formatearPeriodo(periodoObjetivo) + '. Períodos encontrados: ' + periodosUnicos.join(', ') + '.'
    };
  }

  return {
    exito: false,
    motivo: 'No se halló el suministro para el período seleccionado.'
  };
}

function actualizarEstadoLecturaDesdePlataforma(sheet, valores, columnas, payload) {
  const suministroBuscado = normalizarSuministro(payload && payload.suministro);
  const periodoObjetivo = normalizarPeriodo(payload && payload.periodo, '');
  const estadoLectura = normalizarEstadoLecturaOperativa(payload && payload.estadoLectura);
  const periodosAlternativos = [];

  if (!suministroBuscado || !periodoObjetivo) {
    return {
      exito: false,
      motivo: 'No se pudo identificar el suministro y el período a actualizar.'
    };
  }

  if (!estadoLectura) {
    return {
      exito: false,
      motivo: 'El estado de lectura debe ser CORTADO o RETIRADO.'
    };
  }

  for (let i = 1; i < valores.length; i++) {
    const filaValores = valores[i] || [];
    const suministroFila = normalizarSuministro(obtenerValorDeFilaPorIndice(filaValores, columnas.sum));
    const periodoFila = normalizarPeriodo(obtenerValorDeFilaPorIndice(filaValores, columnas.per), '');

    if (suministroFila === suministroBuscado && periodoFila && periodoFila !== periodoObjetivo) {
      periodosAlternativos.push(periodoFila);
    }

    if (suministroFila !== suministroBuscado || periodoFila !== periodoObjetivo) {
      continue;
    }

    const filaHoja = i + 1;

    if (columnas.estadoLectura !== -1) {
      filaValores[columnas.estadoLectura] = escribirTextoEditableEnFila(sheet, filaHoja, columnas.estadoLectura, estadoLectura);
    }

    if (columnas.origen !== -1) {
      filaValores[columnas.origen] = escribirTextoEditableEnFila(sheet, filaHoja, columnas.origen, CONFIG.ORIGEN_LECTURA_CAMPO);
    }

    if (columnas.fecha !== -1) {
      sheet.getRange(filaHoja, columnas.fecha + 1).setValue(new Date());
      filaValores[columnas.fecha] = obtenerMarcaTiempo();
    }

    sheet.getRange(filaHoja, columnas.est + 1).setValue(estadoLectura);
    filaValores[columnas.est] = estadoLectura;

    return {
      exito: true,
      fila: filaHoja,
      estado: estadoLectura
    };
  }

  const periodosUnicos = deduplicarListaTexto(periodosAlternativos).map(function(periodo) {
    return formatearPeriodo(periodo);
  });

  if (periodosUnicos.length) {
    return {
      exito: false,
      motivo: 'El suministro existe en el padrón, pero no en el período seleccionado ' +
        formatearPeriodo(periodoObjetivo) + '. Períodos encontrados: ' + periodosUnicos.join(', ') + '.'
    };
  }

  return {
    exito: false,
    motivo: 'No se halló el suministro para el período seleccionado.'
  };
}

function escribirMedicionesManualesEnFila(sheet, filaHoja, columnas, mediciones) {
  const mapa = {
    EAT: columnas.eat,
    EAHP: columnas.eahp,
    EAFP: columnas.eafp,
    ER: columnas.er,
    PHP: columnas.php,
    PFP: columnas.pfp
  };

  Object.keys(mapa).forEach(function(clave) {
    const indice = mapa[clave];
    const valor = mediciones[clave];

    if (indice === -1 || !esValorCargaManualInformado(valor)) {
      return;
    }

    escribirValorMedicion(sheet.getRange(filaHoja, indice + 1), valor);
  });
}

function escribirMedicionesEditablesEnFila(sheet, filaHoja, filaValores, columnas, mediciones) {
  const mapa = {
    EAT: columnas.eat,
    EAHP: columnas.eahp,
    EAFP: columnas.eafp,
    ER: columnas.er,
    PHP: columnas.php,
    PFP: columnas.pfp
  };

  Object.keys(mapa).forEach(function(clave) {
    const indice = mapa[clave];
    if (indice === -1) {
      return;
    }

    filaValores[indice] = escribirMedicionEditableEnFila(sheet, filaHoja, indice, mediciones[clave]);
  });
}

function escribirPreviasEditablesEnFila(sheet, filaHoja, filaValores, columnas, previas) {
  const mapa = {
    EAT: columnas.lecAntEat,
    EAHP: columnas.lecAntEahp,
    EAFP: columnas.lecAntEafp,
    ER: columnas.lecAntEr,
    PHP: columnas.lecAntPhp,
    PFP: columnas.lecAntPfp
  };

  Object.keys(mapa).forEach(function(clave) {
    const indice = mapa[clave];
    if (indice === -1) {
      return;
    }

    filaValores[indice] = escribirMedicionEditableEnFila(sheet, filaHoja, indice, previas[clave]);
  });
}

function escribirMedicionEditableEnFila(sheet, filaHoja, indiceColumna, valor) {
  const rango = sheet.getRange(filaHoja, indiceColumna + 1);

  if (!esValorCargaManualInformado(valor)) {
    rango.setValue('');
    return '';
  }

  const numero = parsearNumeroMedicion(valor, 0);
  const formato = resolverFormatoNumeroMedicion(valor, numero);
  const textoExacto = obtenerTextoExactoMedicion(valor, numero);
  rango
    .setValue(numero)
    .setNumberFormat(formato || '0.############');

  return textoExacto || formatearNumeroVisible(numero) || String(numero);
}

function escribirTextoEditableEnFila(sheet, filaHoja, indiceColumna, valor) {
  if (indiceColumna === -1) {
    return '';
  }

  const texto = normalizarTexto(valor);
  sheet.getRange(filaHoja, indiceColumna + 1).setValue(texto);
  return texto;
}

function resolverEstadoTrasCrud(filaValores, columnas) {
  const fila = Array.isArray(filaValores) ? filaValores : [];
  const estadoLectura = columnas.estadoLectura === -1
    ? ''
    : normalizarEstadoLecturaOperativa(obtenerValorDeFilaPorIndice(fila, columnas.estadoLectura));
  const indicesLectura = [
    columnas.eat,
    columnas.eahp,
    columnas.eafp,
    columnas.er,
    columnas.php,
    columnas.pfp,
    columnas.lecAntEat,
    columnas.lecAntEahp,
    columnas.lecAntEafp,
    columnas.lecAntEr,
    columnas.lecAntPhp,
    columnas.lecAntPfp
  ];

  if (estadoLectura) {
    return estadoLectura;
  }

  for (let i = 0; i < indicesLectura.length; i++) {
    const indice = indicesLectura[i];
    if (indice === -1) {
      continue;
    }

    if (normalizarTexto(obtenerValorDeFilaPorIndice(fila, indice)) !== '') {
      return CONFIG.ESTADO_COMPLETO;
    }
  }

  const archivoLectura = columnas.archivoLectura === -1 ? '' : normalizarTexto(obtenerValorDeFilaPorIndice(fila, columnas.archivoLectura));
  const tipoArchivoLectura = columnas.tipoArchivoLectura === -1 ? '' : normalizarTexto(obtenerValorDeFilaPorIndice(fila, columnas.tipoArchivoLectura));

  if (archivoLectura || tipoArchivoLectura) {
    return CONFIG.ESTADO_LEIDO_SIN_CARGA;
  }

  return CONFIG.ESTADO_PENDIENTE;
}

function clasificarOrigenLecturaTrasPlantilla(sheet, valores, columnas, periodosObjetivo, clavesTelemetria) {
  const periodos = {};
  const claves = clavesTelemetria || {};
  const data = Array.isArray(valores) ? valores : [];
  let actualizados = 0;
  let telemetria = 0;
  let campo = 0;

  if (columnas.origen === -1 || data.length <= 1) {
    return {
      actualizados: 0,
      telemetria: 0,
      campo: 0
    };
  }

  (Array.isArray(periodosObjetivo) ? periodosObjetivo : []).forEach(function(periodo) {
    const valor = normalizarPeriodo(periodo, '');
    if (valor) {
      periodos[valor] = true;
    }
  });

  if (!Object.keys(periodos).length) {
    return {
      actualizados: 0,
      telemetria: 0,
      campo: 0
    };
  }

  const columnaActualizada = [];

  for (let i = 1; i < data.length; i++) {
    const fila = data[i] || [];
    const periodoFila = normalizarPeriodo(obtenerValorDeFilaPorIndice(fila, columnas.per), '');
    const suministroFila = normalizarSuministro(obtenerValorDeFilaPorIndice(fila, columnas.sum));
    const actual = normalizarTexto(obtenerValorDeFilaPorIndice(fila, columnas.origen)).toUpperCase();
    let finalValue = actual;

    if (periodos[periodoFila] && suministroFila) {
      const clave = construirClaveSuministroPeriodo(suministroFila, periodoFila, '');
      finalValue = claves[clave]
        ? CONFIG.ORIGEN_LECTURA_TELEMETRIA
        : CONFIG.ORIGEN_LECTURA_CAMPO;

      if (finalValue === CONFIG.ORIGEN_LECTURA_TELEMETRIA) {
        telemetria++;
      } else {
        campo++;
      }

      if (finalValue !== actual) {
        actualizados++;
      }

      fila[columnas.origen] = finalValue;
    }

    columnaActualizada.push([finalValue]);
  }

  sheet.getRange(2, columnas.origen + 1, columnaActualizada.length, 1).setValues(columnaActualizada);

  return {
    actualizados: actualizados,
    telemetria: telemetria,
    campo: campo
  };
}

function construirMedicionesCargaManualDesdePayload(data) {
  const origen = data || {};

  return {
    EAT: origen.EAT != null ? origen.EAT : origen.eat,
    EAHP: origen.EAHP != null ? origen.EAHP : origen.eahp,
    EAFP: origen.EAFP != null ? origen.EAFP : origen.eafp,
    ER: origen.ER != null ? origen.ER : origen.er,
    PHP: origen.PHP != null ? origen.PHP : origen.php,
    PFP: origen.PFP != null ? origen.PFP : origen.pfp
  };
}

function construirMedicionesPreviasDesdePayload(data) {
  const origen = data || {};

  return {
    EAT: origen.LEC_ANT_EAT != null ? origen.LEC_ANT_EAT : origen.lecAntEat,
    EAHP: origen.LEC_ANT_EAHP != null ? origen.LEC_ANT_EAHP : origen.lecAntEahp,
    EAFP: origen.LEC_ANT_EAFP != null ? origen.LEC_ANT_EAFP : origen.lecAntEafp,
    ER: origen.LEC_ANT_ER != null ? origen.LEC_ANT_ER : origen.lecAntEr,
    PHP: origen.LEC_ANT_PHP != null ? origen.LEC_ANT_PHP : origen.lecAntPhp,
    PFP: origen.LEC_ANT_PFP != null ? origen.LEC_ANT_PFP : origen.lecAntPfp
  };
}

function normalizarOrigenLectura(valor) {
  const origen = normalizarTexto(valor).toUpperCase();

  if (origen === CONFIG.ORIGEN_LECTURA_TELEMETRIA || origen === CONFIG.ORIGEN_LECTURA_CAMPO) {
    return origen;
  }

  return '';
}

function normalizarEstadoLecturaOperativa(valor) {
  const estado = normalizarTexto(valor).toUpperCase();

  if (estado === 'CORTADO' || estado === 'RETIRADO') {
    return estado;
  }

  return '';
}

function tieneLecturasManualesInformadas(mediciones) {
  const data = mediciones || {};

  return CAMPOS_CARGA_MANUAL.some(function(clave) {
    return esValorCargaManualInformado(data[clave]);
  });
}

function esValorCargaManualInformado(valor) {
  if (typeof valor === 'number') {
    return !isNaN(valor);
  }

  return normalizarTexto(valor) !== '';
}

function obtenerColumnasPlantillaCargaManual(headersUpper) {
  const headers = (Array.isArray(headersUpper) ? headersUpper : []).map(normalizarEncabezadoPlantillaCargaManual);
  const buscar = function(alias) {
    for (let i = 0; i < alias.length; i++) {
      const indice = headers.indexOf(normalizarEncabezadoPlantillaCargaManual(alias[i]));
      if (indice !== -1) {
        return indice;
      }
    }
    return -1;
  };

  return {
    sum: buscar(['NroServicio', 'Nro Servicio', 'NROSERVICIO', 'SUMINISTRO', 'SUMI']),
    serie: buscar(['SerieMedidor', 'Serie Medidor', 'SERIEMEDIDOR', 'SERIE_FAB', 'SERIEFAB', 'SERIE']),
    per: buscar(['PERIODO', 'MES']),
    obs: buscar(['OBS', 'OBSERVACION', 'OBSERVACIONES']),
    eat: buscar(['EAT', 'EAT (KWH)']),
    eahp: buscar(['EAHP', 'EAHP (KWH)']),
    eafp: buscar(['EAFP', 'EAFP (KWH)']),
    er: buscar(['ER', 'ER (KVARH)']),
    php: buscar(['PHP', 'PHP (KW)']),
    pfp: buscar(['PFP', 'PFP (KW)'])
  };
}

function resolverEstructuraPlantillaCargaManual(displayData) {
  const data = Array.isArray(displayData) ? displayData : [];
  const limite = Math.min(data.length, 10);

  for (let i = 0; i < limite; i++) {
    const headersUpper = (data[i] || []).map(normalizarEncabezadoPlantillaCargaManual);
    const columnas = obtenerColumnasPlantillaCargaManual(headersUpper);
    const tieneSuministro = columnas.sum !== -1;
    const tieneLecturas = [columnas.eat, columnas.eahp, columnas.eafp, columnas.er, columnas.php, columnas.pfp]
      .some(function(indice) {
        return indice !== -1;
      });

    if (!tieneSuministro || !tieneLecturas) {
      continue;
    }

    return {
      headerRowIndex: i,
      dataStartRow: i + 1,
      columnas: columnas
    };
  }

  throw new Error(
    'El archivo no coincide con el formato estándar esperado. Debe incluir la hoja con encabezados NroServicio, SerieMedidor, EAT, EAHP, EAFP, ER, PHP y PFP.'
  );
}

function normalizarEncabezadoPlantillaCargaManual(valor) {
  return normalizarTextoBusqueda(valor).replace(/[^A-Z0-9]/g, '');
}

function obtenerMedicionesDesdeFilaPlantilla(filaRaw, filaDisplay, columnas) {
  const raw = Array.isArray(filaRaw) ? filaRaw : [];
  const display = Array.isArray(filaDisplay) ? filaDisplay : [];

  return {
    EAT: obtenerValorPlantillaPriorizado(raw, display, columnas.eat),
    EAHP: obtenerValorPlantillaPriorizado(raw, display, columnas.eahp),
    EAFP: obtenerValorPlantillaPriorizado(raw, display, columnas.eafp),
    ER: obtenerValorPlantillaPriorizado(raw, display, columnas.er),
    PHP: obtenerValorPlantillaPriorizado(raw, display, columnas.php),
    PFP: obtenerValorPlantillaPriorizado(raw, display, columnas.pfp)
  };
}

function obtenerValorPlantillaPriorizado(rawRow, displayRow, indice) {
  if (typeof indice !== 'number' || indice < 0) {
    return '';
  }

  const valorDisplay = obtenerValorDeFilaPorIndice(displayRow, indice);
  if (normalizarTexto(valorDisplay) !== '') {
    return valorDisplay;
  }

  const valorRaw = obtenerValorDeFilaPorIndice(rawRow, indice);
  if (typeof valorRaw === 'number' && !isNaN(valorRaw)) {
    return valorRaw;
  }

  return obtenerValorDeFilaPorIndice(displayRow, indice);
}

function obtenerValorDeFilaPorIndice(row, indice) {
  if (!Array.isArray(row) || typeof indice !== 'number' || indice < 0 || indice >= row.length) {
    return '';
  }

  return row[indice];
}

function filtrarPendientesCargaManualPorSuministro(items, filtro) {
  const listado = Array.isArray(items) ? items : [];
  const termino = normalizarTexto(filtro);

  if (!termino) {
    return listado.slice();
  }

  const terminoNormalizado = normalizarSuministro(termino);
  const terminoTexto = termino.toUpperCase();

  return listado.filter(function(item) {
    const suministro = normalizarTexto(item && item.suministro);
    const suministroNormalizado = normalizarSuministro(suministro);
    const suministroTexto = suministro.toUpperCase();

    return suministroNormalizado === terminoNormalizado ||
      suministroTexto === terminoTexto ||
      suministroNormalizado.indexOf(terminoNormalizado) !== -1 ||
      suministroTexto.indexOf(terminoTexto) !== -1;
  });
}

function cargarContenidoPlantillaManual(sheet, pendientes) {
  const headers = [
    'NroServicio',
    'SerieMedidor',
    'EAT',
    'EAHP',
    'EAFP',
    'ER',
    'PHP',
    'PFP'
  ];
  const elaborado = 'Elaborado: ' + Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'GMT-5',
    'd/M/yyyy H:mm:ss'
  );
  const responsable = 'Responsable: SCM PIURA';
  const rows = pendientes.map(function(item) {
    return [
      item.suministro || '',
      obtenerSerieMedidorPendiente(item),
      '',
      '',
      '',
      '',
      '',
      ''
    ];
  });

  sheet.clear();
  sheet.getRange(1, 1).setValue(elaborado);
  sheet.getRange(2, 1).setValue(responsable);
  sheet.getRange(3, 1, 1, headers.length).setValues([['Resultado', '', '', '', '', '', '', '']]);
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);

  if (rows.length) {
    sheet.getRange(5, 1, rows.length, headers.length).setValues(rows);
  }

  aplicarFormatoPlantillaManual(sheet, rows.length, headers.length);
}

function aplicarFormatoPlantillaManual(sheet, totalRows, totalColumns) {
  const cantidadDatos = Math.max(totalRows, 1);
  const filaInicioDatos = 5;

  sheet.setFrozenRows(4);
  sheet.setColumnWidth(1, 84);
  sheet.setColumnWidth(2, 98);
  sheet.setColumnWidth(3, 92);
  sheet.setColumnWidth(4, 92);
  sheet.setColumnWidth(5, 92);
  sheet.setColumnWidth(6, 92);
  sheet.setColumnWidth(7, 56);
  sheet.setColumnWidth(8, 64);

  sheet.getRange(3, 1, 1, totalColumns)
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('left');
  sheet.getRange(4, 1, 1, totalColumns)
    .setFontWeight('bold')
    .setBackground('#c8c864')
    .setHorizontalAlignment('center')
    .setBorder(true, true, true, true, true, true);
  sheet.getRange(filaInicioDatos, 1, cantidadDatos, totalColumns)
    .setBorder(true, true, true, true, true, true);
  sheet.getRange(filaInicioDatos, 3, cantidadDatos, 6)
    .setNumberFormat('0.############');
}

function obtenerSerieMedidorPendiente(item) {
  const campos = item && Array.isArray(item.campos) ? item.campos : [];
  const etiquetas = ['SERIEFAB', 'SERIEFAB', 'SERIE', 'SERIEMEDIDOR', 'SERIEMED'];

  for (let i = 0; i < campos.length; i++) {
    const definicion = campos[i] || {};
    const nombre = normalizarEncabezadoPlantillaCargaManual(definicion.campo);
    if (etiquetas.indexOf(nombre) === -1) {
      continue;
    }

    const valor = normalizarTexto(definicion.valor);
    if (valor && valor !== '-') {
      return valor;
    }
  }

  return '';
}

function clasificarOmitidoCargaManual(registro, motivo, noEncontrados, yaRealizados, omitidosPorOtraCausa) {
  const suministro = registro && registro.suministro ? registro.suministro : '-';
  const periodo = registro && registro.periodo ? registro.periodo : '-';
  const detalle = suministro + ' (' + periodo + ')';
  const textoMotivo = normalizarTextoBusqueda(motivo);

  if (textoMotivo.indexOf('NO SE HALLO EL SUMINISTRO') !== -1) {
    noEncontrados.push(detalle);
    return;
  }

  if (textoMotivo.indexOf('YA ESTABA MARCADO COMO REALIZADO') !== -1) {
    yaRealizados.push(detalle);
    return;
  }

  if (textoMotivo.indexOf('YA ESTABA MARCADO COMO CORTADO') !== -1 || textoMotivo.indexOf('YA ESTABA MARCADO COMO RETIRADO') !== -1) {
    yaRealizados.push(detalle);
    return;
  }

  omitidosPorOtraCausa.push(detalle + ': ' + (motivo || 'Sin detalle.'));
}

function combinarRegistroCargaManual(actual, candidato) {
  const base = clonarRegistroCargaManual(actual);
  const nuevo = clonarRegistroCargaManual(candidato);
  const lecturasBase = contarLecturasCargaManualInformadas(base.mediciones);
  const lecturasNuevo = contarLecturasCargaManualInformadas(nuevo.mediciones);
  const principal = lecturasNuevo > lecturasBase ? nuevo : base;
  const complemento = lecturasNuevo > lecturasBase ? base : nuevo;

  CAMPOS_CARGA_MANUAL.forEach(function(clave) {
    if (!esValorCargaManualInformado(principal.mediciones[clave]) && esValorCargaManualInformado(complemento.mediciones[clave])) {
      principal.mediciones[clave] = complemento.mediciones[clave];
    }
  });

  if (!normalizarTexto(principal.observacion) && normalizarTexto(complemento.observacion)) {
    principal.observacion = complemento.observacion;
  }

  return principal;
}

function clonarRegistroCargaManual(registro) {
  const origen = registro || {};
  const mediciones = origen.mediciones || {};

  return {
    suministro: origen.suministro || '',
    periodo: origen.periodo || '',
    observacion: origen.observacion || '',
    mediciones: {
      EAT: mediciones.EAT,
      EAHP: mediciones.EAHP,
      EAFP: mediciones.EAFP,
      ER: mediciones.ER,
      PHP: mediciones.PHP,
      PFP: mediciones.PFP
    }
  };
}

function contarLecturasCargaManualInformadas(mediciones) {
  const data = mediciones || {};
  let total = 0;

  CAMPOS_CARGA_MANUAL.forEach(function(clave) {
    if (esValorCargaManualInformado(data[clave])) {
      total++;
    }
  });

  return total;
}

function deduplicarListaTexto(items) {
  const lista = Array.isArray(items) ? items : [];
  const mapa = {};
  const salida = [];

  lista.forEach(function(item) {
    const texto = normalizarTexto(item);
    if (!texto || mapa[texto]) {
      return;
    }

    mapa[texto] = true;
    salida.push(texto);
  });

  return salida;
}

function crearSpreadsheetTemporalDesdeArchivo(fileData, nombreBase) {
  const blob = Utilities.newBlob(
    Utilities.base64Decode(fileData.data),
    fileData.mimeType,
    fileData.fileName
  );
  const resource = {
    name: nombreBase,
    mimeType: MimeType.GOOGLE_SHEETS
  };

  return Drive.Files.create
    ? Drive.Files.create(resource, blob)
    : Drive.Files.insert({ title: resource.name, mimeType: resource.mimeType }, blob);
}

function exportarSpreadsheetAXlsx(spreadsheetId, fileName) {
  const url = 'https://www.googleapis.com/drive/v3/files/' + spreadsheetId +
    '/export?mimeType=' + encodeURIComponent(MIME_TYPE_XLSX_CARGA_MANUAL);
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() >= 300) {
    throw new Error('No se pudo exportar la plantilla a Excel.');
  }

  return response.getBlob().setName(fileName);
}

function obtenerOCrearCarpetaCargaManual(periodo) {
  const carpetaPeriodo = obtenerOCrearCarpetaPadrones(periodo);
  const existentes = carpetaPeriodo.getFoldersByName(NOMBRE_SUBCARPETA_CARGA_MANUAL);

  if (existentes.hasNext()) {
    return existentes.next();
  }

  return carpetaPeriodo.createFolder(NOMBRE_SUBCARPETA_CARGA_MANUAL);
}
