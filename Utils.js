function normalizarTexto(valor) {
  return String(valor == null ? '' : valor).trim();
}

function normalizarTextoBusqueda(valor) {
  return normalizarTexto(valor)
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function esValorInformado(valor) {
  return normalizarTexto(valor) !== '';
}

function normalizarSuministro(valor) {
  const texto = normalizarTexto(valor);
  const soloDigitos = texto.replace(/\D/g, '').replace(/^0+/, '');
  return soloDigitos || texto.toUpperCase();
}

function obtenerPeriodoActual() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'GMT-5',
    'yyyyMM'
  );
}

function normalizarPeriodo(valor, fallback) {
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(
      valor,
      Session.getScriptTimeZone() || 'GMT-5',
      'yyyyMM'
    );
  }

  const texto = normalizarTexto(valor);
  if (!texto) {
    return fallback || '';
  }

  const matchDirecto = texto.match(/20\d{4}/);
  if (matchDirecto) {
    return matchDirecto[0];
  }

  const soloDigitos = texto.replace(/\D/g, '');
  if (/^20\d{4}$/.test(soloDigitos)) {
    return soloDigitos;
  }

  const matchMes = texto.match(/(20\d{2})[-\/]?(\d{2})/);
  if (matchMes) {
    return matchMes[1] + matchMes[2];
  }

  return fallback || '';
}

function obtenerPeriodoActivoDesdeValores(valores, indicePeriodo) {
  if (!Array.isArray(valores) || valores.length <= 1 || indicePeriodo === -1) {
    return obtenerPeriodoActual();
  }

  const periodos = {};

  for (let i = 1; i < valores.length; i++) {
    const periodo = normalizarPeriodo(valores[i][indicePeriodo], '');
    if (periodo) {
      periodos[periodo] = (periodos[periodo] || 0) + 1;
    }
  }

  const lista = Object.keys(periodos).sort();
  return lista.length ? lista[lista.length - 1] : obtenerPeriodoActual();
}

function normalizarEstadoOperativo(valor) {
  return normalizarTexto(valor).toUpperCase();
}

function esEstadoFinalizado(estado) {
  const valor = normalizarEstadoOperativo(estado);
  return valor === CONFIG.ESTADO_COMPLETO ||
    valor === CONFIG.ESTADO_CORTADO ||
    valor === CONFIG.ESTADO_RETIRADO;
}

function esEstadoFinalizadoNoLectura(estado) {
  return normalizarEstadoOperativo(estado) === CONFIG.ESTADO_COMPLETO;
}

function resolverEtiquetaEstadoFinal(estado) {
  const valor = normalizarEstadoOperativo(estado);
  if (valor === CONFIG.ESTADO_CORTADO) {
    return CONFIG.ESTADO_CORTADO;
  }

  if (valor === CONFIG.ESTADO_RETIRADO) {
    return CONFIG.ESTADO_RETIRADO;
  }

  if (valor === CONFIG.ESTADO_COMPLETO) {
    return CONFIG.ESTADO_COMPLETO;
  }

  return valor || 'ESTADO FINAL';
}

function obtenerMarcaTiempo() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'GMT-5',
    'yyyy-MM-dd HH:mm:ss'
  );
}

function contarValoresInformados(valores) {
  if (!Array.isArray(valores)) {
    return 0;
  }

  let total = 0;

  for (let i = 0; i < valores.length; i++) {
    if (esValorInformado(valores[i])) {
      total++;
    }
  }

  return total;
}

function construirClaveSuministroPeriodo(suministro, periodo, fallback) {
  const suministroClave = normalizarSuministro(suministro);
  const periodoClave = normalizarPeriodo(periodo, '');

  if (!suministroClave || !periodoClave) {
    return fallback || '';
  }

  return periodoClave + '|' + suministroClave;
}

function combinarValoresPriorizandoInformados(actual, candidata) {
  const actualSeguro = Array.isArray(actual) ? actual : [];
  const candidataSegura = Array.isArray(candidata) ? candidata : [];
  const candidataEsMejor = contarValoresInformados(candidataSegura) > contarValoresInformados(actualSeguro);
  const base = (candidataEsMejor ? candidataSegura : actualSeguro).slice();
  const complemento = candidataEsMejor ? actualSeguro : candidataSegura;
  const longitud = Math.max(base.length, complemento.length);

  for (let i = 0; i < longitud; i++) {
    if (!esValorInformado(base[i]) && esValorInformado(complemento[i])) {
      base[i] = complemento[i];
    }
  }

  return base;
}

function parsearNumeroMedicion(valor, fallback) {
  const tieneFallbackExplicito = arguments.length >= 2;
  const fallbackSeguro = tieneFallbackExplicito ? fallback : 0;

  if (typeof valor === 'number') {
    return isNaN(valor) ? fallbackSeguro : valor;
  }

  const textoOriginal = normalizarTexto(valor);
  if (!textoOriginal) {
    return fallbackSeguro;
  }

  let texto = textoOriginal.replace(/\s+/g, '').replace(/[^\d,.-]/g, '');
  if (!texto) {
    return fallbackSeguro;
  }

  const ultimaComa = texto.lastIndexOf(',');
  const ultimoPunto = texto.lastIndexOf('.');
  const tieneComa = ultimaComa !== -1;
  const tienePunto = ultimoPunto !== -1;

  if (tieneComa && tienePunto) {
    if (ultimaComa > ultimoPunto) {
      texto = texto.replace(/\./g, '').replace(',', '.');
    } else {
      texto = texto.replace(/,/g, '');
    }
  } else if (tieneComa) {
    if ((texto.match(/,/g) || []).length > 1) {
      const partes = texto.split(',');
      const decimal = partes.pop();
      texto = partes.join('') + '.' + decimal;
    } else {
      texto = texto.replace(',', '.');
    }
  } else if (tienePunto && (texto.match(/\./g) || []).length > 1) {
    const partesPunto = texto.split('.');
    const decimalPunto = partesPunto.pop();
    texto = partesPunto.join('') + '.' + decimalPunto;
  }

  const numero = parseFloat(texto);
  return isNaN(numero) ? fallbackSeguro : numero;
}

function formatearNumeroVisible(valor) {
  if (typeof valor !== 'number' || isNaN(valor)) {
    return '';
  }

  const factor = 1000000000;
  const redondeado = Math.round(valor * factor) / factor;
  const texto = redondeado.toFixed(9).replace(/0+$/, '').replace(/\.$/, '');
  return texto === '-0' ? '0' : texto;
}

function obtenerTextoExactoMedicion(valor, numeroRespaldo) {
  const texto = normalizarTexto(valor);
  if (texto) {
    return texto;
  }

  if (typeof numeroRespaldo === 'number' && !isNaN(numeroRespaldo)) {
    return formatearNumeroVisible(numeroRespaldo);
  }

  return '';
}

function resolverFormatoNumeroMedicion(valor, numeroRespaldo) {
  const texto = obtenerTextoExactoMedicion(valor, numeroRespaldo);
  if (!texto) {
    return '0.############';
  }

  const matchDecimal = texto.match(/[.,](\d+)$/);
  if (!matchDecimal) {
    return '0';
  }

  return '0.' + '0'.repeat(matchDecimal[1].length);
}

function formatearPeriodo(periodo) {
  const valor = normalizarPeriodo(periodo, '');
  if (!valor || valor.length !== 6) {
    return valor;
  }

  return valor.slice(0, 4) + '-' + valor.slice(4, 6);
}
