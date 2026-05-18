const CAMPOS_DETALLE_SUMINISTRO = [
  { label: 'Suministro', aliases: ['SUMINISTRO'] },
  { label: 'Cartera', aliases: ['CARTERA'] },
  { label: 'Nombre', aliases: ['NOMBRE_CLIENTE', 'NOMBRE', 'CLIENTE', 'RAZON SOCIAL'] },
  { label: 'Dirección', aliases: ['DIRECCION', 'DIR'] },
  { label: 'Marca', aliases: ['MARCA'] },
  { label: 'Modelo', aliases: ['MODELO'] },
  { label: 'SerieFab', aliases: ['SERIE_FAB', 'SERIEFAB', 'SERIE FAB', 'SERIE_MEDIDOR', 'NRO_SERIE'] },
  { label: 'Factor', aliases: ['FACTOR', 'FACTOR_K'] },
  { label: 'Tarifa', aliases: ['TARIFA', 'TIPO_TARIFA'] },
  { label: 'IP', aliases: ['IP'] },
  { label: 'TLM', aliases: ['TLM', 'TELEFONO', 'CELULAR'] },
  { label: 'Operador', aliases: ['OPERADOR'] },
  { label: 'Marca', aliases: ['MARCA_TLM', 'MARCA_MODEM', 'MARCA_EQUIPO', 'MARCA_HHU', 'MARCA_DISPOSITIVO'] },
  { label: 'Modelo', aliases: ['MODELO_TLM', 'MODELO_MODEM', 'MODELO_EQUIPO', 'MODELO_HHU', 'MODELO_DISPOSITIVO'] },
  { label: 'Serie', aliases: ['SERIE_TLM', 'SERIE_MODEM', 'SERIE_EQUIPO', 'SERIE_HHU', 'SERIE_DISPOSITIVO', 'SERIE'] },
  { label: 'Soft.', aliases: ['SOFT', 'SOFTWARE', 'VERSION_SOFT', 'VERSION'] },
  { label: 'Ubic.', aliases: ['UBIC.', 'UBIC', 'UBICACION_TLM', 'UBICACION_MEDIDOR'] },
  { label: 'Ruta', aliases: ['RUTA', 'NOMB_RUTA', 'NOMBRE_RUTA'] },
  { label: 'Estado', aliases: ['ESTADO'] },
  { label: 'Fecha', aliases: ['FECHA REGISTRO', 'FECHA', 'FECHA_LECTURA'] },
  { label: 'Tipo -Lec.', aliases: ['TIPO -LEC.', 'TIPO_LEC', 'TIPO LEC', 'TIPO_LECTURA', 'TIPOLECTURA'] },
  { label: 'Ruta-Cuadrilla', aliases: ['RUTA-CUADRILLA', 'RUTA_CUADRILLA', 'CUADRILLA'] },
  { label: 'Origen lectura', aliases: ['ORIGEN_LECTURA'] },
  { label: 'Estado', aliases: ['ESTADO_LECTURA', 'ESTADO LECTURA', 'ESTADO_MEDIDOR', 'ESTADO MEDIDOR'] },
  { label: 'Archivo de lectura', aliases: ['ARCHIVO_LECTURA'] },
  { label: 'Tipo de archivo de lectura', aliases: ['TIPO_ARCHIVO_LECTURA'] },
  { label: 'Observación', aliases: ['OBSERVACION', 'OBSERVACIONES', 'OBS'] }
];

function getMetricas(periodoFiltro) {
  const resumen = getResumenMensual(periodoFiltro);

  return {
    total: resumen.total,
    pendientes: resumen.pendientes,
    leidosSinCarga: resumen.leidosSinCarga || 0,
    realizados: resumen.realizados,
    progreso: resumen.progreso,
    periodo: resumen.periodo,
    periodoLabel: resumen.periodoLabel,
    periodosDisponibles: resumen.periodosDisponibles
  };
}

function getResumenMensual(periodoFiltro) {
  const padron = obtenerContextoPadron();

  if (!padron.success) {
    return {
      total: 0,
      pendientes: 0,
      leidosSinCarga: 0,
      realizados: 0,
      progreso: 0,
      periodo: '',
      periodoLabel: '',
      periodosDisponibles: [],
      suministros: [],
      totalSuministros: 0
    };
  }

  const periodoObjetivo = normalizarPeriodo(periodoFiltro, padron.periodoActivo);
  let total = 0;
  let pendientes = 0;
  let leidosSinCarga = 0;
  let realizados = 0;
  const suministros = [];

  padron.rows.forEach(function(row) {
    if (periodoObjetivo && row.periodo !== periodoObjetivo) {
      return;
    }

    const estado = row.estado;
    if (estado === CONFIG.ESTADO_PENDIENTE) {
      pendientes++;
      total++;
    } else if (estado === CONFIG.ESTADO_LEIDO_SIN_CARGA) {
      leidosSinCarga++;
      total++;
    } else if (esEstadoFinalizado(estado)) {
      realizados++;
      total++;
    }

    suministros.push(construirDetalleSuministro(padron.headersUpper, row));
  });

  const progreso = total === 0 ? 0 : Math.round((realizados / total) * 100);
  const suministrosOrdenados = suministros.sort(function(a, b) {
    const suministroA = normalizarSuministro(a.suministro);
    const suministroB = normalizarSuministro(b.suministro);

    if (suministroA < suministroB) {
      return -1;
    }

    if (suministroA > suministroB) {
      return 1;
    }

    return a.rowNumber - b.rowNumber;
  });

  return {
    total: total,
    pendientes: pendientes,
    leidosSinCarga: leidosSinCarga,
    realizados: realizados,
    progreso: progreso,
    periodo: periodoObjetivo,
    periodoLabel: formatearPeriodo(periodoObjetivo),
    periodosDisponibles: padron.periodosDisponibles,
    suministros: suministrosOrdenados,
    totalSuministros: suministrosOrdenados.length
  };
}

function buscarSuministroDetalle(suministro, periodoFiltro) {
  const padron = obtenerContextoPadron();

  if (!padron.success) {
    return {
      success: false,
      mensaje: 'No existe información de padrón disponible.',
      coincidencias: [],
      totalCoincidencias: 0,
      periodo: ''
    };
  }

  const textoBuscado = normalizarTexto(suministro);
  if (!textoBuscado) {
    return {
      success: false,
      mensaje: 'Ingrese un suministro para realizar la búsqueda.',
      coincidencias: [],
      totalCoincidencias: 0,
      periodo: normalizarPeriodo(periodoFiltro, padron.periodoActivo),
      periodoLabel: formatearPeriodo(normalizarPeriodo(periodoFiltro, padron.periodoActivo))
    };
  }

  const periodoObjetivo = normalizarPeriodo(periodoFiltro, padron.periodoActivo);
  const suministroNormalizado = normalizarSuministro(textoBuscado);
  const suministroTexto = textoBuscado.toUpperCase();
  const filasFiltradas = padron.rows.filter(function(row) {
    return !periodoObjetivo || row.periodo === periodoObjetivo;
  });

  let coincidencias = filasFiltradas.filter(function(row) {
    return coincideSuministroExacto(row.suministro, suministroNormalizado, suministroTexto);
  });

  if (!coincidencias.length) {
    coincidencias = filasFiltradas.filter(function(row) {
      return coincideSuministroParcial(row.suministro, suministroNormalizado, suministroTexto);
    });
  }

  if (!coincidencias.length) {
    return {
      success: false,
      mensaje: 'No se encontraron registros para el suministro solicitado en el período seleccionado.',
      coincidencias: [],
      totalCoincidencias: 0,
      periodo: periodoObjetivo,
      periodoLabel: formatearPeriodo(periodoObjetivo)
    };
  }

  return {
    success: true,
    mensaje: coincidencias.length === 1
      ? 'Se encontró 1 suministro.'
      : 'Se encontraron ' + coincidencias.length + ' suministros.',
    coincidencias: coincidencias.slice(0, 5).map(function(row) {
      return construirDetalleSuministro(padron.headersUpper, row);
    }),
    totalCoincidencias: coincidencias.length,
    periodo: periodoObjetivo,
    periodoLabel: formatearPeriodo(periodoObjetivo)
  };
}

function obtenerContextoPadron() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.NOMBRE_HOJA_PADRON);

  if (!sheet) {
    return {
      success: false,
      headers: [],
      rows: [],
      periodoActivo: '',
      periodosDisponibles: []
    };
  }

  const range = sheet.getDataRange();
  const data = range.getDisplayValues();
  const rawData = range.getValues();
  if (data.length <= 1) {
    return {
      success: true,
      headers: data.length ? data[0] : [],
      rows: [],
      periodoActivo: '',
      periodosDisponibles: []
    };
  }

  const headers = data[0].map(function(header) {
    return normalizarTexto(header);
  });
  const headersUpper = headers.map(function(header) {
    return header.toUpperCase();
  });
  const colPeriodo = headersUpper.indexOf('PERIODO');
  const colEstado = headersUpper.indexOf('ESTADO');
  const colSuministro = headersUpper.indexOf('SUMINISTRO');
  const periodoActivo = obtenerPeriodoActivoDesdeValores(data, colPeriodo);
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const values = data[i];
    const rawValues = rawData[i] || [];
    const periodo = colPeriodo === -1
      ? ''
      : normalizarPeriodo(rawValues[colPeriodo], normalizarPeriodo(values[colPeriodo], ''));
    const estado = colEstado === -1 ? '' : normalizarTexto(values[colEstado]).toUpperCase();
    const suministro = colSuministro === -1 ? '' : normalizarTexto(values[colSuministro]);

    rows.push({
      rowNumber: i + 1,
      values: values,
      rawValues: rawValues,
      periodo: periodo,
      estado: estado,
      suministro: suministro
    });
  }

  const rowsConsolidadas = consolidarRegistrosPadron(rows);
  const periodosMap = {};

  rowsConsolidadas.forEach(function(row) {
    if (row.periodo) {
      periodosMap[row.periodo] = (periodosMap[row.periodo] || 0) + 1;
    }
  });

  const periodosDisponibles = Object.keys(periodosMap)
    .sort()
    .reverse()
    .map(function(periodo) {
      return {
        value: periodo,
        label: formatearPeriodo(periodo),
        total: periodosMap[periodo]
      };
    });

  return {
    success: true,
    headers: headers,
    headersUpper: headersUpper,
    rows: rowsConsolidadas,
    periodoActivo: periodoActivo,
    periodosDisponibles: periodosDisponibles
  };
}

function coincideSuministroExacto(suministro, suministroNormalizado, suministroTexto) {
  const raw = normalizarTexto(suministro).toUpperCase();
  const normalized = normalizarSuministro(suministro);
  return normalized === suministroNormalizado || raw === suministroTexto;
}

function coincideSuministroParcial(suministro, suministroNormalizado, suministroTexto) {
  const raw = normalizarTexto(suministro).toUpperCase();
  const normalized = normalizarSuministro(suministro);
  return normalized.indexOf(suministroNormalizado) !== -1 || raw.indexOf(suministroTexto) !== -1;
}

function construirDetalleSuministro(headersUpper, row) {
  const estadoPadron = row.estado || obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['ESTADO']);
  const estadoLectura = obtenerValorDetallePorAlias(
    headersUpper,
    row.values,
    row.rawValues,
    ['ESTADO_LECTURA', 'ESTADO LECTURA', 'ESTADO_MEDIDOR', 'ESTADO MEDIDOR']
  );
  const fechaCarga = obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['FECHA REGISTRO', 'FECHA', 'FECHA_LECTURA']);
  const observacion = obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['OBSERVACION', 'OBSERVACIONES', 'OBS']);
  const archivoLectura = obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['ARCHIVO_LECTURA']);
  const tipoArchivoLectura = obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['TIPO_ARCHIVO_LECTURA']);
  const campos = CAMPOS_DETALLE_SUMINISTRO.map(function(definicion) {
    return {
      campo: definicion.label,
      valor: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, definicion.aliases)
    };
  });

  const carga = {
    EAT: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['EAT', 'EAT (KWH)']),
    EAHP: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['EAHP', 'EAHP (KWH)']),
    EAFP: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['EAFP', 'EAFP (KWH)']),
    ER: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['ER', 'ER (KVARH)']),
    PHP: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['PHP', 'PHP (KW)']),
    PFP: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['PFP', 'PFP (KW)']),
    previas: {
      EAT: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['LEC_ANT_EAT']),
      EAHP: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['LEC_ANT_EAHP']),
      EAFP: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['LEC_ANT_EAFP']),
      ER: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['LEC_ANT_ER']),
      PHP: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['LEC_ANT_PHP']),
      PFP: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['LEC_ANT_PFP'])
    },
    fecha: fechaCarga,
    observacion: observacion
  };
  const estadoNormalizado = normalizarTexto(estadoPadron).toUpperCase();
  const esArchivoRg = esCargaRg(archivoLectura, tipoArchivoLectura, observacion);
  const resumenListado = {
    suministro: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['SUMINISTRO']),
    origen: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['ORIGEN_LECTURA']),
    cartera: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['CARTERA']),
    nombre: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['NOMBRE_CLIENTE', 'NOMBRE', 'CLIENTE', 'RAZON SOCIAL']),
    direccion: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['DIRECCION', 'DIR']),
    marca: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['MARCA']),
    modelo: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['MODELO']),
    serieFab: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['SERIE_FAB', 'SERIEFAB', 'SERIE FAB', 'SERIE_MEDIDOR', 'NRO_SERIE']),
    factor: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['FACTOR', 'FACTOR_K']),
    tarifa: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['TARIFA', 'TIPO_TARIFA'])
  };

  return {
    key: construirClaveSuministroPeriodo(row.suministro, row.periodo, 'fila_' + row.rowNumber),
    rowNumber: row.rowNumber,
    suministro: row.suministro,
    periodo: row.periodo,
    periodoLabel: formatearPeriodo(row.periodo),
    estado: estadoPadron || '-',
    estadoLectura: estadoLectura,
    resumenListado: resumenListado,
    vistaRapida: {
      cartera: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['CARTERA']),
      nombre: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['NOMBRE_CLIENTE', 'NOMBRE', 'CLIENTE', 'RAZON SOCIAL']),
      direccion: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['DIRECCION', 'DIR']),
      ruta: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['RUTA', 'NOMB_RUTA', 'NOMBRE_RUTA']),
      tarifa: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['TARIFA', 'TIPO_TARIFA']),
      fecha: fechaCarga,
      origenLectura: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['ORIGEN_LECTURA']),
      tipoLectura: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['TIPO -LEC.', 'TIPO_LEC', 'TIPO LEC', 'TIPO_LECTURA', 'TIPOLECTURA']),
      rutaCuadrilla: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['RUTA-CUADRILLA', 'RUTA_CUADRILLA', 'CUADRILLA'])
    },
    carga: {
      EAT: carga.EAT,
      EAHP: carga.EAHP,
      EAFP: carga.EAFP,
      ER: carga.ER,
      PHP: carga.PHP,
      PFP: carga.PFP,
      previas: carga.previas,
      esAlphaset: esArchivoRg,
      esArchivoRg: esArchivoRg,
      fecha: carga.fecha,
      observacion: carga.observacion,
      origenLectura: obtenerValorDetallePorAlias(headersUpper, row.values, row.rawValues, ['ORIGEN_LECTURA']),
      archivoLectura: archivoLectura,
      tipoArchivoLectura: tipoArchivoLectura,
      completa: esEstadoFinalizado(estadoNormalizado),
      sinCargaLeida: estadoNormalizado === CONFIG.ESTADO_LEIDO_SIN_CARGA,
      tieneCarga: resolverTieneCargaRegistrada(carga, estadoPadron)
    },
    campos: campos
  };
}

function resolverTieneCargaRegistrada(carga, estadoPadron) {
  const valores = [
    carga.EAT,
    carga.EAHP,
    carga.EAFP,
    carga.ER,
    carga.PHP,
    carga.PFP,
    carga.previas && carga.previas.EAT,
    carga.previas && carga.previas.EAHP,
    carga.previas && carga.previas.EAFP,
    carga.previas && carga.previas.ER,
    carga.previas && carga.previas.PHP,
    carga.previas && carga.previas.PFP
  ];

  for (let i = 0; i < valores.length; i++) {
    const valor = parsearNumeroMedicion(valores[i], null);
    if (typeof valor === 'number' && !isNaN(valor) && Math.abs(valor) > 0) {
      return true;
    }
  }

  return false;
}

function esCargaRg(archivoLectura, tipoArchivoLectura, observacion) {
  const tipo = normalizarTexto(tipoArchivoLectura).toUpperCase();
  const archivo = normalizarTexto(archivoLectura).toUpperCase();

  if (tipo === '.RG' || tipo === 'RG' || /\.RG$/i.test(archivo)) {
    return true;
  }

  return normalizarTextoBusqueda(observacion).indexOf('RG ALPHASET') !== -1;
}

function consolidarRegistrosPadron(rows) {
  const registros = Array.isArray(rows) ? rows : [];
  const consolidados = [];
  const indicePorClave = {};

  registros.forEach(function(row) {
    const clave = construirClaveSuministroPeriodo(
      row.suministro,
      row.periodo,
      'fila_' + row.rowNumber
    );

    if (Object.prototype.hasOwnProperty.call(indicePorClave, clave)) {
      const indice = indicePorClave[clave];
      consolidados[indice] = combinarRegistroPadron(consolidados[indice], row);
      return;
    }

    indicePorClave[clave] = consolidados.length;
    consolidados.push(clonarRegistroPadron(row));
  });

  return consolidados;
}

function clonarRegistroPadron(row) {
  return {
    rowNumber: row.rowNumber,
    values: Array.isArray(row.values) ? row.values.slice() : [],
    rawValues: Array.isArray(row.rawValues) ? row.rawValues.slice() : [],
    periodo: row.periodo,
    estado: row.estado,
    suministro: row.suministro
  };
}

function combinarRegistroPadron(actual, candidata) {
  const actualSeguro = clonarRegistroPadron(actual);
  const candidataSegura = clonarRegistroPadron(candidata);
  const candidataEsMejor = contarValoresInformados(candidataSegura.values) > contarValoresInformados(actualSeguro.values);
  const base = candidataEsMejor ? candidataSegura : actualSeguro;
  const complemento = candidataEsMejor ? actualSeguro : candidataSegura;

  base.values = combinarValoresPriorizandoInformados(base.values, complemento.values);
  base.rawValues = combinarValoresPriorizandoInformados(base.rawValues, complemento.rawValues);
  base.estado = resolverEstadoConsolidado(actualSeguro.estado, candidataSegura.estado, base.estado);
  base.periodo = base.periodo || complemento.periodo;
  base.suministro = base.suministro || complemento.suministro;
  base.rowNumber = Math.min(actualSeguro.rowNumber, candidataSegura.rowNumber);

  return base;
}

function resolverEstadoConsolidado(estadoA, estadoB, fallback) {
  const estados = [estadoA, estadoB, fallback].map(function(valor) {
    return normalizarTexto(valor).toUpperCase();
  });

  if (estados.indexOf(CONFIG.ESTADO_RETIRADO) !== -1) {
    return CONFIG.ESTADO_RETIRADO;
  }

  if (estados.indexOf(CONFIG.ESTADO_CORTADO) !== -1) {
    return CONFIG.ESTADO_CORTADO;
  }

  if (estados.indexOf(CONFIG.ESTADO_COMPLETO) !== -1) {
    return CONFIG.ESTADO_COMPLETO;
  }

  if (estados.indexOf(CONFIG.ESTADO_LEIDO_SIN_CARGA) !== -1) {
    return CONFIG.ESTADO_LEIDO_SIN_CARGA;
  }

  if (estados.indexOf(CONFIG.ESTADO_PENDIENTE) !== -1) {
    return CONFIG.ESTADO_PENDIENTE;
  }

  for (let i = 0; i < estados.length; i++) {
    if (estados[i]) {
      return estados[i];
    }
  }

  return '';
}

function obtenerValorDetallePorAlias(headersUpper, values, rawValues, aliases) {
  const listaHeaders = Array.isArray(headersUpper) ? headersUpper : [];
  const listaValores = Array.isArray(values) ? values : [];
  const listaValoresRaw = Array.isArray(rawValues) ? rawValues : [];
  const listaAliases = Array.isArray(aliases) ? aliases : [];

  for (let i = 0; i < listaAliases.length; i++) {
    const indice = listaHeaders.indexOf(String(listaAliases[i]).toUpperCase());
    if (indice === -1) {
      continue;
    }

    const header = listaHeaders[indice];
    const valorRaw = indice < listaValoresRaw.length ? listaValoresRaw[indice] : '';
    const valor = normalizarTexto(listaValores[indice]);
    if (esCampoDecimalDetallado(header, valorRaw)) {
      return formatearValorDecimalDetallado(valorRaw, valor);
    }

    if (valor) {
      return valor;
    }
  }

  return '-';
}

function esCampoDecimalDetallado(header, valorRaw) {
  const nombre = normalizarTexto(header).toUpperCase();
  if ([
    'EAT',
    'EAT (KWH)',
    'EAHP',
    'EAHP (KWH)',
    'EAFP',
    'EAFP (KWH)',
    'ER',
    'ER (KVARH)',
    'PHP',
    'PHP (KW)',
    'PFP',
    'PFP (KW)',
    'LEC_ANT_EAT',
    'LEC_ANT_EAHP',
    'LEC_ANT_EAFP',
    'LEC_ANT_ER',
    'LEC_ANT_PHP',
    'LEC_ANT_PFP'
  ].indexOf(nombre) !== -1) {
    return true;
  }

  return typeof valorRaw === 'number' && !isNaN(valorRaw);
}

function formatearValorDecimalDetallado(valorRaw, valorDisplay) {
  const textoDisplay = normalizarTexto(valorDisplay);
  if (textoDisplay) {
    return textoDisplay;
  }

  if (typeof valorRaw === 'number' && !isNaN(valorRaw)) {
    const textoNumero = formatearNumeroVisible(valorRaw);
    if (textoNumero) {
      return textoNumero;
    }
  }

  const textoRaw = normalizarTexto(valorRaw);
  if (textoRaw) {
    return textoRaw;
  }

  return '-';
}
