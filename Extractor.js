function extraerDatosDeArchivo(file) {
  const fileName = file.getName();
  const nameLower = fileName.toLowerCase();

  // 1. Extraer Suministro del nombre (primer bloque de numeros)
  const matchSum = fileName.match(/(\d+)/);
  let sumBuscado = matchSum ? matchSum[1] : "";

  // 2. Extraer Periodo del nombre (busca 6 numeros seguidos, ej: 202605)
  const matchPer = fileName.match(/20\d{4}/);
  let periodoBuscado = matchPer ? matchPer[0] : Utilities.formatDate(new Date(), "GMT-5", "yyyyMM");

  let datos = {
    suministro: sumBuscado,
    periodo: periodoBuscado,
    eat: 0, eahp: 0, eafp: 0, er: 0, php: 0, pfp: 0,
    lecAntEat: null, lecAntEahp: null, lecAntEafp: null,
    lecAntEr: null, lecAntPhp: null, lecAntPfp: null,
    fecha: new Date(),
    obs: "",
    archivoLectura: fileName,
    tipoArchivoLectura: resolverTipoArchivoLectura(fileName)
  };

  try {
    if (esArchivoSinCargaEspecial(nameLower)) {
      return parsearArchivoSinCargaEspecial(file, datos);
    }

    if (esArchivoExcelLecturas(nameLower)) {
      return parsearExcel(file, datos);
    }

    // Aqui se mantiene la logica de TXT y RG
    return parsearTextoAntiguo(file, datos);
  } catch (e) {
    datos.obs = "Error de lectura: " + e.message;
    return datos;
  }
}

function esArchivoSinCargaEspecial(fileNameLower) {
  const nombre = String(fileNameLower || '').toLowerCase();
  return nombre.endsWith('.msr') || nombre.endsWith('.rp3');
}

function esArchivoExcelLecturas(fileNameLower) {
  const nombre = String(fileNameLower || '').toLowerCase();
  return nombre.endsWith('.xlsx') || nombre.endsWith('.xls');
}

function parsearArchivoSinCargaEspecial(file, datos) {
  const resultado = datos || {};
  const blob = validarArchivoSinCargaEspecial(file);
  const contenido = leerContenidoArchivoSinCargaEspecial(blob);
  const archivo = file.getName();
  const suministro = resolverSuministroArchivoSinCargaEspecial(archivo, contenido);
  const tipo = resolverTipoArchivoSinCargaEspecial(archivo);

  if (suministro) {
    resultado.suministro = normalizarSuministro(suministro);
  }

  resultado.sinCargaEspecial = true;
  resultado.archivoLectura = archivo;
  resultado.tipoArchivoLectura = tipo;
  resultado.obs = 'Archivo ' + tipo + ' leído sin parámetros - ' + archivo;

  return resultado;
}

function validarArchivoSinCargaEspecial(file) {
  const blob = file && file.getBlob ? file.getBlob() : null;
  const bytes = blob && blob.getBytes ? blob.getBytes() : [];

  if (!blob || !bytes.length) {
    throw new Error('El archivo especial no contiene datos legibles.');
  }

  return blob;
}

function leerContenidoArchivoSinCargaEspecial(blob) {
  try {
    return blob.getDataAsString();
  } catch (error) {
    return '';
  }
}

function resolverTipoArchivoSinCargaEspecial(fileName) {
  const tipo = resolverTipoArchivoLectura(fileName);
  return tipo || 'SIN CARGA';
}

function resolverTipoArchivoLectura(fileName) {
  const nombre = String(fileName || '').toLowerCase();
  if (!nombre) {
    return '';
  }

  if (nombre.endsWith('.rp3')) {
    return '.RP3';
  }
  if (nombre.endsWith('.msr')) {
    return '.MSR';
  }
  if (nombre.endsWith('.rg')) {
    return '.RG';
  }
  if (nombre.endsWith('.txt')) {
    return '.TXT';
  }
  if (nombre.endsWith('.xlsx')) {
    return '.XLSX';
  }
  if (nombre.endsWith('.xls')) {
    return '.XLS';
  }

  return '';
}

function resolverSuministroArchivoSinCargaEspecial(fileName, contenido) {
  const nombre = normalizarTexto(fileName);
  const matchNombre = nombre.match(/\b(\d{7,12})\b/);
  if (matchNombre) {
    return matchNombre[1];
  }

  const texto = String(contenido == null ? '' : contenido);
  const matchContenido = texto.match(/\b(\d{7,12})\b/);
  return matchContenido ? matchContenido[1] : '';
}

function parsearExcel(file, datos) {
  const blob = file.getBlob();
  const resource = {
    name: "temp_" + datos.suministro,
    mimeType: MimeType.GOOGLE_SHEETS
  };

  let tempFile;

  try {
    if (Drive.Files.create) {
      tempFile = Drive.Files.create(resource, blob);
    } else {
      tempFile = Drive.Files.insert({ title: resource.name, mimeType: resource.mimeType }, blob);
    }
  } catch (err) {
    throw new Error("Error al convertir Excel: " + err.message);
  }

  const ss = SpreadsheetApp.openById(tempFile.id);

  try {
    const hojaEnergia = buscarHojaPorPatron(ss, ['ENERG']);
    const hojaDemanda = buscarHojaPorPatron(ss, ['DEMANDA', 'MAX']);
    let usoPatronSrNo = false;

    if (hojaEnergia) {
      const energiaPorSrNo = leerValoresPorSrNo(hojaEnergia, {
        2: 'eat',
        7: 'eahp',
        12: 'eafp',
        17: 'er'
      });

      if (energiaPorSrNo.encontrados > 0) {
        usoPatronSrNo = true;
        aplicarLecturasExtraidas(datos, energiaPorSrNo.valores);
      } else {
        datos.eat = leerValorOBISDesdeHoja(hojaEnergia, "1.8.0.255", { obisCol: 5, valorCol: 6 });
        datos.eahp = leerValorOBISDesdeHoja(hojaEnergia, "1.8.1.255", { obisCol: 5, valorCol: 6 });
        datos.eafp = leerValorOBISDesdeHoja(hojaEnergia, "1.8.2.255", { obisCol: 5, valorCol: 6 });
        datos.er = leerValorOBISDesdeHoja(hojaEnergia, "3.8.0.255", { obisCol: 5, valorCol: 6 });
      }
    }

    if (hojaDemanda) {
      const demandaPorSrNo = leerValoresPorSrNo(hojaDemanda, {
        2: 'php',
        7: 'pfp'
      });

      if (demandaPorSrNo.encontrados > 0) {
        usoPatronSrNo = true;
        aplicarLecturasExtraidas(datos, demandaPorSrNo.valores);
      } else {
        datos.php = leerValorOBISDesdeHoja(hojaDemanda, "1.6.1.255", { obisCol: 4, valorCol: 6 });
        datos.pfp = leerValorOBISDesdeHoja(hojaDemanda, "1.6.2.255", { obisCol: 4, valorCol: 6 });
      }
    }

    datos.obs = usoPatronSrNo
      ? "Procesado Excel por Sr.No y columna Valor/Value"
      : "Procesado Excel con lectura previa";
    return datos;
  } finally {
    Drive.Files.remove(tempFile.id);
  }
}

function buscarHojaPorPatron(ss, patrones) {
  const lista = Array.isArray(patrones) ? patrones : [];
  const sheets = ss.getSheets();

  for (let i = 0; i < sheets.length; i++) {
    const nombre = normalizarTextoBusqueda(sheets[i].getName());
    let coincide = true;

    for (let j = 0; j < lista.length; j++) {
      if (nombre.indexOf(normalizarTextoBusqueda(lista[j])) === -1) {
        coincide = false;
        break;
      }
    }

    if (coincide) {
      return sheets[i];
    }
  }

  return sheets.find(function(sheet) {
    const nombre = normalizarTextoBusqueda(sheet.getName());
    return lista.some(function(patron) {
      return nombre.indexOf(normalizarTextoBusqueda(patron)) !== -1;
    });
  }) || null;
}

function leerValoresPorSrNo(sheet, mapaSeries) {
  const rawData = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();
  const mapa = mapaSeries || {};
  const seriesObjetivo = Object.keys(mapa).map(function(valor) {
    return parseInt(valor, 10);
  }).filter(function(valor) {
    return !isNaN(valor);
  });
  const headerInfo = encontrarColumnasSrNoYValor(displayData, seriesObjetivo);
  const resultado = {
    encontrados: 0,
    valores: {}
  };

  if (!headerInfo) {
    return resultado;
  }

  for (let i = headerInfo.headerRow + 1; i < displayData.length; i++) {
    const srTexto = obtenerTextoCelda(displayData[i], headerInfo.srNoCol);
    const srNo = parseInt(srTexto.replace(/[^\d]/g, ''), 10);

    if (!Object.prototype.hasOwnProperty.call(mapa, srNo)) {
      continue;
    }

    const claveDestino = mapa[srNo];
    const valorRaw = obtenerValorCelda(rawData[i], headerInfo.valorCol);
    const valorDisplay = obtenerTextoCelda(displayData[i], headerInfo.valorCol);
    resultado.valores[claveDestino] = parsearNumeroMedicion(
      typeof valorRaw === 'number' ? valorRaw : valorDisplay,
      0
    );
    resultado.encontrados++;

    if (resultado.encontrados >= seriesObjetivo.length) {
      break;
    }
  }

  return resultado;
}

function encontrarColumnasSrNoYValor(displayData, seriesObjetivo) {
  const limite = Math.min(Array.isArray(displayData) ? displayData.length : 25);
  const objetivos = Array.isArray(seriesObjetivo) ? seriesObjetivo : [];
  let mejor = null;

  for (let i = 0; i < limite; i++) {
    const row = displayData[i] || [];
    let srNoCol = -1;
    const valorCols = [];

    for (let j = 0; j < row.length; j++) {
      const texto = normalizarEncabezadoMedicion(row[j]);
      if (srNoCol === -1 && texto === 'SRNO') {
        srNoCol = j;
      }
      if (esEncabezadoValorMedicion(texto)) {
        valorCols.push(j);
      }
    }

    if (srNoCol === -1 || !valorCols.length) {
      continue;
    }

    for (let k = 0; k < valorCols.length; k++) {
      const candidato = evaluarCandidatoSrNoValor(displayData, i, srNoCol, valorCols[k], objetivos);
      if (!mejor || candidato.score > mejor.score) {
        mejor = candidato;
      }
    }
  }

  if (mejor && mejor.score > 0) {
    return {
      headerRow: mejor.headerRow,
      srNoCol: mejor.srNoCol,
      valorCol: mejor.valorCol
    };
  }

  return mejor ? {
    headerRow: mejor.headerRow,
    srNoCol: mejor.srNoCol,
    valorCol: mejor.valorCol
  } : null;
}

function normalizarEncabezadoMedicion(valor) {
  return normalizarTextoBusqueda(valor).replace(/[^A-Z0-9]/g, '');
}

function esEncabezadoValorMedicion(texto) {
  return texto === 'VALOR' || texto === 'VALUE';
}

function evaluarCandidatoSrNoValor(displayData, headerRow, srNoCol, valorCol, seriesObjetivo) {
  const data = Array.isArray(displayData) ? displayData : [];
  const objetivos = Array.isArray(seriesObjetivo) ? seriesObjetivo : [];
  const vistos = {};
  let score = 0;

  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i] || [];
    const srTexto = obtenerTextoCelda(row, srNoCol);
    const srNo = parseInt(srTexto.replace(/[^\d]/g, ''), 10);

    if (isNaN(srNo) || objetivos.indexOf(srNo) === -1 || vistos[srNo]) {
      continue;
    }

    const valorTexto = obtenerTextoCelda(row, valorCol);
    if (esCeldaNumerica('', valorTexto)) {
      vistos[srNo] = true;
      score++;
      if (score >= objetivos.length) {
        break;
      }
    }
  }

  return {
    headerRow: headerRow,
    srNoCol: srNoCol,
    valorCol: valorCol,
    score: score
  };
}

function aplicarLecturasExtraidas(datos, valores) {
  const origen = valores || {};
  Object.keys(origen).forEach(function(clave) {
    if (Object.prototype.hasOwnProperty.call(datos, clave)) {
      datos[clave] = origen[clave];
    }
  });
}

function obtenerTextoCelda(row, index) {
  if (!Array.isArray(row) || index < 0 || index >= row.length) {
    return '';
  }

  return normalizarTexto(row[index]);
}

function obtenerValorCelda(row, index) {
  if (!Array.isArray(row) || index < 0 || index >= row.length) {
    return '';
  }

  return row[index];
}

function leerValorOBISDesdeHoja(sheet, codigo, fallback) {
  const displayData = sheet.getDataRange().getDisplayValues();
  const rawData = sheet.getDataRange().getValues();
  const estructura = analizarEstructuraHojaMedicion(displayData, fallback || {});
  const filaDato = buscarFilaOBIS(displayData, codigo, estructura.obisColumna);

  if (filaDato === -1) {
    return 0;
  }

  const columnaPrevia = resolverColumnaPrevia(estructura.headers, estructura.obisColumna);
  const columnaGenerica = resolverColumnaValorGenerica(estructura.headers, estructura.obisColumna);
  const columnaFallback = typeof fallback.valorCol === 'number' ? fallback.valorCol : -1;
  const candidatos = [columnaPrevia, columnaGenerica, columnaFallback];

  for (let i = 0; i < candidatos.length; i++) {
    const col = candidatos[i];
    if (col == null || col === -1) {
      continue;
    }

    const valor = rawData[filaDato] && col < rawData[filaDato].length ? rawData[filaDato][col] : '';
    const texto = displayData[filaDato] && col < displayData[filaDato].length ? displayData[filaDato][col] : '';

    if (esCeldaNumerica(valor, texto)) {
      return parsearNumeroMedicion(valor || texto, 0);
    }
  }

  return buscarNumeroCercanoEnFila(
    rawData[filaDato] || [],
    displayData[filaDato] || [],
    estructura.obisColumna
  );
}

function analizarEstructuraHojaMedicion(displayData, fallback) {
  const data = Array.isArray(displayData) ? displayData : [];
  const filaCabecera = encontrarFilaCabeceraMedicion(data);
  const headers = filaCabecera === -1 ? [] : data[filaCabecera].map(function(celda) {
    return normalizarTextoBusqueda(celda);
  });

  return {
    filaCabecera: filaCabecera,
    headers: headers,
    obisColumna: resolverColumnaOBIS(headers, fallback.obisCol)
  };
}

function encontrarFilaCabeceraMedicion(data) {
  const limite = Math.min(Array.isArray(data) ? data.length : 0, 12);
  let mejorFila = -1;
  let mejorPuntaje = -1;

  for (let i = 0; i < limite; i++) {
    const row = (data[i] || []).map(normalizarTextoBusqueda);
    const textoFila = row.join(' | ');
    let puntaje = 0;

    if (textoFila.indexOf('OBIS') !== -1 || textoFila.indexOf('REGISTRO') !== -1 || textoFila.indexOf('CODIGO') !== -1) {
      puntaje += 3;
    }
    if (textoFila.indexOf('PREV') !== -1 || textoFila.indexOf('ANTER') !== -1) {
      puntaje += 3;
    }
    if (textoFila.indexOf('LECTURA') !== -1 || textoFila.indexOf('VALOR') !== -1 || textoFila.indexOf('DEMANDA') !== -1 || textoFila.indexOf('ENERG') !== -1) {
      puntaje += 1;
    }

    if (puntaje > mejorPuntaje) {
      mejorPuntaje = puntaje;
      mejorFila = i;
    }
  }

  return mejorPuntaje > 0 ? mejorFila : -1;
}

function resolverColumnaOBIS(headers, fallback) {
  const lista = Array.isArray(headers) ? headers : [];

  for (let i = 0; i < lista.length; i++) {
    if (lista[i].indexOf('OBIS') !== -1 || lista[i].indexOf('REGISTRO') !== -1 || lista[i].indexOf('CODIGO') !== -1) {
      return i;
    }
  }

  return typeof fallback === 'number' ? fallback : 0;
}

function resolverColumnaPrevia(headers, obisColumna) {
  const lista = Array.isArray(headers) ? headers : [];
  let candidata = -1;

  for (let i = 0; i < lista.length; i++) {
    if (i === obisColumna) {
      continue;
    }

    const texto = lista[i];
    if ((texto.indexOf('PREV') !== -1 || texto.indexOf('ANTER') !== -1) &&
        texto.indexOf('INSTAN') === -1 &&
        texto.indexOf('ACTUAL') === -1) {
      candidata = i;
      if (i > obisColumna) {
        return i;
      }
    }
  }

  return candidata;
}

function resolverColumnaValorGenerica(headers, obisColumna) {
  const lista = Array.isArray(headers) ? headers : [];
  let candidata = -1;

  for (let i = 0; i < lista.length; i++) {
    if (i === obisColumna) {
      continue;
    }

    const texto = lista[i];
    if (!texto || texto.indexOf('INSTAN') !== -1 || texto.indexOf('ACTUAL') !== -1) {
      continue;
    }

    if (
      texto.indexOf('LECTURA') !== -1 ||
      texto.indexOf('VALOR') !== -1 ||
      texto.indexOf('DEMANDA') !== -1 ||
      texto.indexOf('ENERG') !== -1 ||
      texto.indexOf('POTENCIA') !== -1
    ) {
      candidata = i;
      if (i > obisColumna) {
        return i;
      }
    }
  }

  return candidata;
}

function buscarFilaOBIS(displayData, codigo, obisColumna) {
  const data = Array.isArray(displayData) ? displayData : [];
  const objetivo = normalizarTextoBusqueda(codigo);

  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];

    if (typeof obisColumna === 'number' && obisColumna >= 0 && obisColumna < row.length) {
      const textoObis = normalizarTextoBusqueda(row[obisColumna]);
      if (textoObis.indexOf(objetivo) !== -1) {
        return i;
      }
    }

    for (let j = 0; j < row.length; j++) {
      const texto = normalizarTextoBusqueda(row[j]);
      if (texto.indexOf(objetivo) !== -1) {
        return i;
      }
    }
  }

  return -1;
}

function esCeldaNumerica(rawValue, displayValue) {
  if (typeof rawValue === 'number') {
    return !isNaN(rawValue);
  }

  const texto = normalizarTexto(displayValue || rawValue).replace(/\s+/g, '');
  if (!/[\d]/.test(texto)) {
    return false;
  }

  if (/^\d+(\.\d+){2,}$/.test(texto) || texto.indexOf('.255') !== -1) {
    return false;
  }

  return true;
}

function buscarNumeroCercanoEnFila(rawRow, displayRow, obisColumna) {
  const valoresCrudos = Array.isArray(rawRow) ? rawRow : [];
  const valoresTexto = Array.isArray(displayRow) ? displayRow : [];
  const limite = Math.max(valoresCrudos.length, valoresTexto.length);

  for (let offset = 1; offset < limite; offset++) {
    const colDerecha = obisColumna + offset;
    if (colDerecha < limite) {
      const valorDerecha = colDerecha < valoresCrudos.length ? valoresCrudos[colDerecha] : '';
      const textoDerecha = colDerecha < valoresTexto.length ? valoresTexto[colDerecha] : '';
      if (esCeldaNumerica(valorDerecha, textoDerecha)) {
        return parsearNumeroMedicion(valorDerecha || textoDerecha, 0);
      }
    }
  }

  for (let offset = 1; offset <= obisColumna; offset++) {
    const colIzquierda = obisColumna - offset;
    const valorIzquierda = colIzquierda < valoresCrudos.length ? valoresCrudos[colIzquierda] : '';
    const textoIzquierda = colIzquierda < valoresTexto.length ? valoresTexto[colIzquierda] : '';
    if (esCeldaNumerica(valorIzquierda, textoIzquierda)) {
      return parsearNumeroMedicion(valorIzquierda || textoIzquierda, 0);
    }
  }

  return 0;
}

function parsearTextoAntiguo(file, datos) {
  const txt = file.getBlob().getDataAsString();
  const base = datos || {};

  if (esTextoMetcom(txt)) {
    return parsearTextoMetcom(txt, base);
  }

  if (esTextoAlphasetRg(txt)) {
    return parsearTextoAlphasetRg(txt, base);
  }

  return parsearTextoObisGenerico(txt, base);
}

function esTextoMetcom(txt) {
  const contenido = normalizarTexto(txt);
  return contenido.indexOf('/MCS') === 0 || contenido.indexOf('96.1.0(') !== -1;
}

function esTextoAlphasetRg(txt) {
  const contenido = normalizarTextoBusqueda(txt);
  return contenido.indexOf('ALPHASET DATAFILE') === 0 || contenido.indexOf('/ELS') !== -1;
}

function parsearTextoMetcom(txt, datos) {
  const resultado = datos || {};
  const serie = extraerValorTextoObis(txt, '96.1.0');
  const suministroObis = extraerValorTextoObis(txt, '96.1.1');
  const suministroResuelto = resolverSuministroMetcom(resultado.suministro, suministroObis);

  if (suministroResuelto) {
    resultado.suministro = suministroResuelto;
  }

  const mediciones = {
    eat: extraerNumeroTextoObisPreferente(txt, ['1.8.0*05', '1.8.0']),
    eahp: extraerNumeroTextoObisPreferente(txt, ['1.8.2*05', '1.8.2']),
    eafp: extraerNumeroTextoObisPreferente(txt, ['1.8.1*05', '1.8.1']),
    er: extraerNumeroTextoObisPreferente(txt, ['3.8.0*05', '3.8.0']),
    php: extraerNumeroTextoObisPreferente(txt, ['1.6.2*05', '1.6.2']),
    pfp: extraerNumeroTextoObisPreferente(txt, ['1.6.1*05', '1.6.1'])
  };

  aplicarLecturasExtraidas(resultado, mediciones);
  resultado.obs = construirObservacionTextoMetcom(serie, suministroObis, resultado.suministro) +
    ' - Lecturas previas *05';
  return resultado;
}

function parsearTextoObisGenerico(txt, datos) {
  const resultado = datos || {};
  const suministro = extraerValorTextoObis(txt, '96.1.0');

  if (suministro) {
    resultado.suministro = normalizarSuministro(suministro);
  }

  aplicarLecturasExtraidas(resultado, {
    eat: extraerNumeroTextoObis(txt, '1.8.0'),
    eahp: extraerNumeroTextoObis(txt, '1.8.1'),
    eafp: extraerNumeroTextoObis(txt, '1.8.2'),
    er: extraerNumeroTextoObis(txt, '3.8.0'),
    php: extraerNumeroTextoObis(txt, '1.6.1'),
    pfp: extraerNumeroTextoObis(txt, '1.6.2')
  });

  resultado.obs = 'Procesado TXT/RG por OBIS';
  return resultado;
}

function parsearTextoAlphasetRg(txt, datos) {
  const resultado = datos || {};

  aplicarLecturasExtraidas(resultado, construirLecturasAlphasetRg(txt));
  resultado.obs = 'Procesado RG Alphaset - 12 parámetros';
  return resultado;
}

function construirLecturasAlphasetRg(txt) {
  const eat = extraerLecturaTextoObisConAnterior(txt, '1.8.0');
  const eahp = extraerLecturaTextoObisConAnterior(txt, '1.8.1');
  const eafp = extraerLecturaTextoObisConAnterior(txt, '1.8.2');
  const er = extraerLecturaTextoObisConAnterior(txt, '3.8.0');
  const php = extraerLecturaTextoObisConAnterior(txt, '1.6.1');
  const pfp = extraerLecturaTextoObisConAnterior(txt, '1.6.2');

  return {
    eat: eat.actual,
    eahp: eahp.actual,
    eafp: eafp.actual,
    er: er.actual,
    php: php.actual,
    pfp: pfp.actual,
    lecAntEat: eat.anterior,
    lecAntEahp: eahp.anterior,
    lecAntEafp: eafp.anterior,
    lecAntEr: er.anterior,
    lecAntPhp: php.anterior,
    lecAntPfp: pfp.anterior
  };
}

function resolverSuministroMetcom(suministroArchivo, suministroObis) {
  const suministroBase = normalizarSuministro(suministroArchivo);
  const suministroInterno = normalizarSuministro(suministroObis);

  if (!suministroInterno) {
    return suministroBase;
  }

  if (!suministroBase || suministroBase === suministroInterno) {
    return suministroInterno;
  }

  if (!/\d/.test(suministroBase) && /\d/.test(suministroInterno)) {
    return suministroInterno;
  }

  return suministroBase;
}

function construirObservacionTextoMetcom(serie, suministroObis, suministroResuelto) {
  let observacion = 'Procesado TXT Metcom';
  const serieNormalizada = normalizarTexto(serie);
  const suministroInterno = normalizarSuministro(suministroObis);
  const suministroFinal = normalizarSuministro(suministroResuelto);

  if (serieNormalizada) {
    observacion += ' - Serie ' + serieNormalizada;
  }

  if (suministroInterno && suministroFinal && suministroInterno !== suministroFinal) {
    observacion += ' - El suministro OBIS ' + suministroInterno + ' difiere del nombre del archivo';
  }

  return observacion;
}

function extraerValorTextoObis(txt, codigo) {
  const contenido = String(txt == null ? '' : txt).replace(/^\uFEFF/, '');
  const codigoSeguro = escaparCodigoObisRegex(codigo);
  const patrones = [
    new RegExp('(?:^|[\\r\\n])\\s*(?:\\d+-\\d+:)?' + codigoSeguro + '(?:\\.255|\\*255)?\\(([^\\)]*)\\)', 'i'),
    new RegExp('(?:^|[\\r\\n])\\s*' + codigoSeguro + '\\(([^\\)]*)\\)', 'i'),
    new RegExp(codigoSeguro + '(?:\\.255|\\*255)?\\(([^\\)]*)\\)', 'i')
  ];
  let match = null;

  for (let i = 0; i < patrones.length; i++) {
    match = contenido.match(patrones[i]);
    if (match) {
      break;
    }
  }

  return match ? normalizarTexto(match[1]) : '';
}

function extraerNumeroTextoObis(txt, codigo) {
  const valor = extraerValorTextoObis(txt, codigo);
  if (!valor) {
    return 0;
  }

  return extraerNumeroDesdeValorTextoObis(valor, 0);
}

function extraerNumeroTextoObisPreferente(txt, codigos) {
  const lista = Array.isArray(codigos) ? codigos : [codigos];

  for (let i = 0; i < lista.length; i++) {
    const valor = extraerValorTextoObis(txt, lista[i]);
    if (valor) {
      return extraerNumeroDesdeValorTextoObis(valor, 0);
    }
  }

  return 0;
}

function extraerLecturaTextoObisConAnterior(txt, codigoBase) {
  const entradas = listarEntradasTextoObis(txt, codigoBase);
  let actual = null;
  let anterior = null;

  for (let i = 0; i < entradas.length; i++) {
    const entrada = entradas[i];
    const numero = extraerNumeroDesdeValorTextoObis(entrada.valor, null);
    if (numero === null || numero === undefined || isNaN(numero)) {
      continue;
    }

    if (!entrada.sufijo && actual === null) {
      actual = numero;
      continue;
    }

    if (entrada.sufijo && anterior === null) {
      anterior = numero;
    }
  }

  return {
    actual: actual === null ? 0 : actual,
    anterior: anterior
  };
}

function listarEntradasTextoObis(txt, codigoBase) {
  const contenido = String(txt == null ? '' : txt).replace(/^\uFEFF/, '');
  const codigoSeguro = escaparCodigoObisRegex(codigoBase);
  const regex = new RegExp('^\\s*(?:\\d+-\\d+:)?' + codigoSeguro + '(?:\\.255)?(?:\\*(\\d+))?\\s*\\(([^\\)]*)\\)', 'i');
  const lineas = contenido.split(/\r?\n/);
  const entradas = [];

  for (let i = 0; i < lineas.length; i++) {
    const match = lineas[i].match(regex);
    if (match) {
      entradas.push({
        sufijo: match[1] || '',
        valor: normalizarTexto(match[2])
      });
    }
  }

  return entradas;
}

function extraerNumeroDesdeValorTextoObis(valor, fallback) {
  const texto = normalizarTexto(valor);
  if (!texto) {
    return fallback;
  }

  const numero = texto.match(/-?\d+(?:[.,]\d+)?/);
  return parsearNumeroMedicion(numero ? numero[0] : texto, fallback);
}

function escaparCodigoObisRegex(codigo) {
  return String(codigo == null ? '' : codigo).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
