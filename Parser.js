/**
 * SERVICE PACK V2.1 - TIC LEAD
 * Soluciona: celdas vacias, error de mapeo y limpieza de datos.
 */

function actualizarMaestro(data, mesPeriodo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Padron_Maestro");
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values[0].map(function(h) {
    return String(h).toUpperCase().trim();
  });

  const col = {
    sum: headers.findIndex(function(h) { return h.includes("SUMINISTRO"); }),
    per: headers.findIndex(function(h) { return h.includes("PERIODO"); }),
    est: headers.findIndex(function(h) { return h.includes("ESTADO"); }),
    eat: headers.findIndex(function(h) { return h === "EAT (KWH)" || h === "EAT"; }),
    eahp: headers.findIndex(function(h) { return h === "EAHP (KWH)" || h === "EAHP"; }),
    eafp: headers.findIndex(function(h) { return h === "EAFP (KWH)" || h === "EAFP"; }),
    er: headers.findIndex(function(h) { return h === "ER (KVARH)" || h === "ER"; }),
    php: headers.findIndex(function(h) { return h === "PHP (KW)" || h === "PHP"; }),
    pfp: headers.findIndex(function(h) { return h === "PFP (KW)" || h === "PFP"; })
  };

  if (col.sum === -1 || col.est === -1) {
    return { exito: false, motivo: "No encontre columna Suministro o Estado" };
  }

  const sumArchivo = String(data.suministro).replace(/\D/g, '').replace(/^0+/, '');
  const mesBusqueda = String(mesPeriodo).replace(/\D/g, '');

  for (let i = 1; i < values.length; i++) {
    const sumSheet = String(values[i][col.sum]).replace(/\D/g, '').replace(/^0+/, '');
    const perSheet = String(values[i][col.per]).replace(/\D/g, '');

    if (sumSheet === sumArchivo && perSheet === mesBusqueda) {
      const fila = i + 1;
      const clean = function(val) {
        return parsearNumeroMedicion(val, 0);
      };

      if (col.eat !== -1) sheet.getRange(fila, col.eat + 1).setValue(clean(data.eat));
      if (col.eahp !== -1) sheet.getRange(fila, col.eahp + 1).setValue(clean(data.eahp));
      if (col.eafp !== -1) sheet.getRange(fila, col.eafp + 1).setValue(clean(data.eafp));
      if (col.er !== -1) sheet.getRange(fila, col.er + 1).setValue(clean(data.er));
      if (col.php !== -1) sheet.getRange(fila, col.php + 1).setValue(clean(data.php));
      if (col.pfp !== -1) sheet.getRange(fila, col.pfp + 1).setValue(clean(data.pfp));

      sheet.getRange(fila, col.est + 1).setValue("REALIZADO");
      return { exito: true };
    }
  }

  return {
    exito: false,
    motivo: "Suministro " + sumArchivo + " no hallado en el periodo " + mesPeriodo
  };
}
